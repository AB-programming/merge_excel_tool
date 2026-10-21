#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 数据合并工具
==================

把多个 Excel 文件合并成一张总表，支持三种合并模式：

    concat  纵向拼接 —— 多个文件表头相同，把数据行上下堆叠（最常见）
    merge   横向合并 —— 类似 VLOOKUP，按指定关键列把列拼到一起
    sheets  多表汇总 —— 把工作簿里的多个 sheet 汇总成一张总表

设计要点：
    * 默认把单元格按字符串读取，避免身份证号、手机号、前导零编号被
      转成科学计数法或丢掉前导零；需要数值计算时用 --keep-types 关闭。
    * 表头不一致时取所有列的并集，缺失的填空白，并明确报告哪些文件有差异，
      不做静默丢弃。
    * 输出文件默认写到独立目录并用时间戳命名，绝不覆盖输入文件。

用法示例见 README.md，或直接运行：python3 merge_excel.py --help
不想敲命令就运行：python3 merge_excel.py --gui（桌面端，见 gui.py）
"""

from __future__ import annotations

import argparse
import logging
import sys
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

try:
    import pandas as pd
except ImportError:  # pragma: no cover - 依赖缺失时的友好提示
    sys.exit(
        "[错误] 缺少依赖 pandas，请先安装：\n"
        "       pip install pandas openpyxl xlrd"
    )

# xlrd 为可选依赖，缺失时用一个占位异常，保证下面的 except 语句始终合法
try:
    from xlrd import XLRDError
except ImportError:  # pragma: no cover
    class XLRDError(Exception):
        pass


# ---------------------------------------------------------------------------
# 常量与全局配置
# ---------------------------------------------------------------------------

#: 支持的 Excel 扩展名（.xlsb 需要额外安装 pyxlsb）
EXCEL_EXTS = {".xlsx", ".xlsm", ".xls", ".xlsb"}

#: 来源标记列的默认列名
COL_SOURCE_FILE = "来源文件"
COL_SOURCE_SHEET = "来源工作表"

#: 超过这个行数时提示内存占用（写成普通数字，避免老版本解释器解析失败）
BIG_TABLE_ROWS = 1000000  # 一百万行

#: 结构差异日志最多逐条列出多少个来源 / 多少列，其余只汇总数量
MAX_DIFF_DETAIL = 5
MAX_MISSING_COLS = 8

logger = logging.getLogger("merge_excel")


class ExcelMergeError(Exception):
    """可预期的业务错误，直接以中文提示给用户，不打印堆栈。"""


# ---------------------------------------------------------------------------
# 日志：中文级别前缀
# ---------------------------------------------------------------------------

_LEVEL_CN = {
    "DEBUG": "调试",
    "INFO": "信息",
    "WARNING": "警告",
    "ERROR": "错误",
    "CRITICAL": "严重",
}


class ChineseFormatter(logging.Formatter):
    """把日志级别换成中文，输出形如「[警告] 结构差异 ...」"""

    def format(self, record: logging.LogRecord) -> str:
        prefix = _LEVEL_CN.get(record.levelname, record.levelname)
        return f"[{prefix}] {record.getMessage()}"


def setup_logging(verbose: bool = False, console: bool = True) -> None:
    """重建日志输出。console=False 时由调用方自己挂 handler（GUI 用）。"""
    logger.handlers.clear()
    if console:
        handler = logging.StreamHandler(sys.stdout)
        handler.setFormatter(ChineseFormatter())
        logger.addHandler(handler)
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)
    logger.propagate = False


def add_log_handler(handler: logging.Handler) -> None:
    """追加一个日志 handler，供桌面端把日志转到界面里。"""
    handler.setFormatter(ChineseFormatter())
    logger.addHandler(handler)


# ---------------------------------------------------------------------------
# 文件扫描
# ---------------------------------------------------------------------------

def scan_excel_files(
    directory: Path, pattern: str, recursive: bool
) -> tuple[list[Path], list[Path]]:
    """扫描目录下的 Excel 文件。

    返回 (待合并文件列表, 被跳过的文件列表)。
    跳过 Excel 临时文件（~$ 开头）和隐藏文件（. 开头）。
    """
    if not directory.exists():
        raise ExcelMergeError(f"目录不存在：{directory}")
    if not directory.is_dir():
        raise ExcelMergeError(f"不是目录，无法扫描：{directory}")

    glob_pattern = pattern or "*.xls*"
    iterator = directory.rglob(glob_pattern) if recursive else directory.glob(glob_pattern)

    files: list[Path] = []
    skipped: list[Path] = []
    for path in sorted(iterator):
        if not path.is_file():
            continue
        name = path.name
        if name.startswith("~$"):
            skipped.append(path)  # Excel 打开时生成的临时锁文件
            continue
        if name.startswith("."):
            skipped.append(path)  # 隐藏文件
            continue
        if path.suffix.lower() in EXCEL_EXTS:
            files.append(path)
        else:
            skipped.append(path)

    return files, skipped


def summarize_extensions(directory: Path, recursive: bool) -> str:
    """统计目录里的文件扩展名，用于「没找到 Excel」时给用户提示。"""
    counter: dict[str, int] = {}
    iterator = directory.rglob("*") if recursive else directory.glob("*")
    for path in iterator:
        if path.is_file():
            ext = path.suffix.lower() or "(无扩展名)"
            counter[ext] = counter.get(ext, 0) + 1
    if not counter:
        return "目录是空的"
    top = sorted(counter.items(), key=lambda kv: -kv[1])[:8]
    return "、".join(f"{ext} × {n}" for ext, n in top)


# ---------------------------------------------------------------------------
# 读取
# ---------------------------------------------------------------------------

def _default_sheet_label(sheet: str | int | None) -> str:
    """未指定工作表名时，给出一个可读的标签（避免出现「（0）」这种显示）。"""
    if isinstance(sheet, int):
        return f"第 {sheet + 1} 个工作表"
    if sheet is None:
        return "第 1 个工作表"
    return str(sheet)


def read_workbook(
    path: Path,
    sheet: str | int | None = 0,
    as_str: bool = True,
    all_sheets: bool = False,
) -> dict[str, pd.DataFrame]:
    """读取一个工作簿，返回 {工作表名: DataFrame}。

    as_str=True 时全部按字符串读取并保留原始字符，空单元格填 ""。
    """
    suffix = path.suffix.lower()
    if suffix == ".xlsb":
        raise ExcelMergeError(
            f"{path.name}：.xlsb 格式需要额外依赖，请执行 pip install pyxlsb 后重试，"
            "或先另存为 .xlsx"
        )

    kwargs: dict = {
        "sheet_name": None if all_sheets else sheet,
        "dtype": str if as_str else None,
        "na_filter": False,
        "keep_default_na": False,
    }

    try:
        obj = pd.read_excel(path, **kwargs)
    except zipfile.BadZipFile:
        raise ExcelMergeError(
            f"{path.name}：文件已损坏或不是有效的 Excel（也可能是被加密的工作簿）"
        ) from None
    except XLRDError as exc:
        raise ExcelMergeError(
            f"{path.name}：无法解析该工作簿（{exc}）。若为旧版 .xls，"
            "可尝试用 Excel/WPS 另存为 .xlsx 后重试"
        ) from None
    except PermissionError:
        raise ExcelMergeError(f"{path.name}：没有读取权限，或文件正被 Excel 占用") from None
    except ValueError as exc:
        # pandas 对伪装成 xlsx 的非 Excel 文件会抛出这类提示，换成用户看得懂的说法
        if "cannot be determined" in str(exc) or "Unknown engine" in str(exc):
            raise ExcelMergeError(
                f"{path.name}：文件已损坏，或只是改了扩展名的非 Excel 文件"
            ) from None
        raise ExcelMergeError(f"{path.name}：读取失败（{exc}）") from None
    except Exception as exc:  # 兜底，不让单个坏文件打断整个流程
        raise ExcelMergeError(
            f"{path.name}：读取失败（{type(exc).__name__}: {exc}）"
        ) from None

    if isinstance(obj, dict):
        frames = {str(name): df for name, df in obj.items()}
    else:
        frames = {_default_sheet_label(sheet): obj}

    # 统一做一次空值清理，保证 NaN/None 都变成空字符串
    return {name: _blank_fill(df) for name, df in frames.items()}


def _blank_fill(df: pd.DataFrame) -> pd.DataFrame:
    """把 NaN / None 统一替换为空字符串，避免输出文件里出现 nan 字样。"""
    if df.empty:
        return df
    return df.astype(object).where(df.notna(), "")


def load_frames(
    files: list[Path],
    sheet: str | int | None,
    as_str: bool,
    all_sheets: bool,
) -> tuple[list[tuple[str, str, pd.DataFrame]], list[tuple[Path, str]]]:
    """批量读取文件。

    返回 ([(文件名, 工作表名, DataFrame)], [(失败文件, 失败原因)])
    """
    items: list[tuple[str, str, pd.DataFrame]] = []
    failures: list[tuple[Path, str]] = []

    total = len(files)
    for index, path in enumerate(files, start=1):
        try:
            frames = read_workbook(path, sheet=sheet, as_str=as_str, all_sheets=all_sheets)
        except ExcelMergeError as exc:
            logger.error("[%d/%d] %s", index, total, exc)
            failures.append((path, str(exc)))
            continue

        empty_sheets = 0
        for sheet_name, df in frames.items():
            if df.shape[1] == 0:
                empty_sheets += 1
                continue
            items.append((path.name, sheet_name, df))

        rows = sum(df.shape[0] for df in frames.values())
        cols = max((df.shape[1] for df in frames.values()), default=0)
        logger.info("[%d/%d] 读取 %s …… %d 行 × %d 列", index, total, path.name, rows, cols)
        if empty_sheets:
            logger.warning("      %s 中有 %d 个完全空白的工作表，已跳过", path.name, empty_sheets)
        if rows == 0:
            logger.warning("      %s 没有任何数据行", path.name)

    return items, failures


# ---------------------------------------------------------------------------
# 通用工具
# ---------------------------------------------------------------------------

def _unique_column(name: str, existing: set[str]) -> str:
    """若列名已存在，追加序号直到不冲突。"""
    if name not in existing:
        return name
    i = 2
    while f"{name}_{i}" in existing:
        i += 1
    return f"{name}_{i}"


def _align_and_concat(
    items: list[tuple[str, str, pd.DataFrame]],
    add_source_file: bool,
    add_source_sheet: bool,
) -> pd.DataFrame:
    """列取并集对齐后纵向拼接，并报告各来源的结构差异。"""
    if not items:
        raise ExcelMergeError("没有任何可合并的数据")

    # 1) 计算列的并集，保持首次出现顺序
    union: list[str] = []
    seen: set[str] = set()
    for _, _, df in items:
        for col in df.columns:
            key = str(col)
            if key not in seen:
                seen.add(key)
                union.append(key)

    # 2) 报告结构差异（不静默丢弃；来源较多时只列举前几个，避免刷屏）
    diff_count = 0
    for name, sheet_name, df in items:
        current = {str(c) for c in df.columns}
        missing = [c for c in union if c not in current]
        if not missing:
            continue
        diff_count += 1
        if diff_count <= MAX_DIFF_DETAIL:
            label = name + (f"/{sheet_name}" if add_source_sheet else "")
            shown = missing[:MAX_MISSING_COLS]
            tail = "" if len(missing) <= MAX_MISSING_COLS else f" 等 {len(missing)} 列"
            logger.warning(
                "结构差异 %s：缺少 %d 列 -> %s%s（将留空）",
                label, len(missing), shown, tail,
            )
    if diff_count > MAX_DIFF_DETAIL:
        logger.warning(
            "…… 另有 %d 个来源存在结构差异（已按同样规则留空处理）",
            diff_count - MAX_DIFF_DETAIL,
        )
    if diff_count:
        logger.info("说明：共 %d 个来源的表头不一致，已取全部列的并集", diff_count)

    # 3) 逐表对齐到并集，并在最前面插入来源标记列
    frames: list[pd.DataFrame] = []
    for name, sheet_name, df in items:
        part = df.copy()
        part.columns = [str(c) for c in part.columns]
        part = part.reindex(columns=union)
        part = _blank_fill(part)
        if add_source_sheet:
            part.insert(0, COL_SOURCE_SHEET, sheet_name)
        if add_source_file:
            part.insert(0, COL_SOURCE_FILE, name)
        frames.append(part)

    result = pd.concat(frames, ignore_index=True, sort=False)
    logger.info("合并完成：共 %d 行 × %d 列", result.shape[0], result.shape[1])
    return result


# ---------------------------------------------------------------------------
# 三种合并模式
# ---------------------------------------------------------------------------

def mode_concat(items: list[tuple[str, str, pd.DataFrame]], add_source: bool) -> pd.DataFrame:
    """纵向拼接：多文件数据行上下堆叠。"""
    return _align_and_concat(
        items, add_source_file=add_source, add_source_sheet=False
    )


def mode_sheets(items: list[tuple[str, str, pd.DataFrame]], add_source: bool) -> pd.DataFrame:
    """多工作表汇总：把各文件的各 sheet 汇总成一张表。"""
    # 只在涉及多个文件时才标记来源文件，单文件场景只标记工作表名
    multi_file = len({name for name, _, _ in items}) > 1
    return _align_and_concat(
        items,
        add_source_file=add_source and multi_file,
        add_source_sheet=add_source,
    )


def mode_merge(
    items: list[tuple[str, str, pd.DataFrame]],
    keys: list[str],
    how: str,
) -> pd.DataFrame:
    """横向合并：按关键列做连接，类似 VLOOKUP。"""
    if not keys:
        raise ExcelMergeError("merge 模式必须用 -k/--keys 指定至少一个关键列")

    # 1) 关键列存在性校验，失败时把现有列名一并报出来
    for name, sheet_name, df in items:
        columns = {str(c) for c in df.columns}
        missing = [k for k in keys if k not in columns]
        if missing:
            raise ExcelMergeError(
                f"{name}（{sheet_name}）缺少关键列 {missing}，"
                f"该文件现有列：{list(df.columns)}"
            )

    # 2) 以第一个表为基准逐个连接
    first_name, first_sheet, result = items[0]
    result = result.copy()
    result.columns = [str(c) for c in result.columns]
    logger.info("以 %s 为基准进行 %s 连接，关键列：%s", first_name, how, keys)

    for name, sheet_name, df in items[1:]:
        right = df.copy()
        right.columns = [str(c) for c in right.columns]

        # 3) 重名列预先加上文件名后缀，比默认的 _x/_y 更可读
        overlap = [
            c for c in right.columns
            if c not in keys and c in set(result.columns)
        ]
        if overlap:
            stem = Path(name).stem
            rename = {c: _unique_column(f"{c}__{stem}", set(result.columns)) for c in overlap}
            right = right.rename(columns=rename)
            logger.warning(
                "%s 中有 %d 列与已合并结果重名，已重命名为：%s",
                name, len(overlap), list(rename.values()),
            )

        before = result.shape[0]
        result = pd.merge(result, right, on=keys, how=how, suffixes=("", "_dup"))
        logger.info(
            "    合并 %s：%d 行 -> %d 行，当前 %d 列",
            name, before, result.shape[0], result.shape[1],
        )

    result = _blank_fill(result)
    logger.info("合并完成：共 %d 行 × %d 列", result.shape[0], result.shape[1])
    return result


# ---------------------------------------------------------------------------
# 输出
# ---------------------------------------------------------------------------

def build_output_path(output: str | None, output_dir: str) -> Path:
    if output:
        return Path(output).expanduser()
    directory = Path(output_dir).expanduser()
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return directory / f"合并结果_{stamp}.xlsx"


def write_result(df: pd.DataFrame, out_path: Path, inputs: list[Path]) -> None:
    """写出结果，拒绝覆盖任何输入文件。"""
    resolved_out = out_path.resolve()
    for src in inputs:
        if src.resolve() == resolved_out:
            raise ExcelMergeError(
                f"输出路径与输入文件相同，已拒绝执行以保护原数据：{out_path}"
            )

    out_path.parent.mkdir(parents=True, exist_ok=True)

    if df.shape[0] > BIG_TABLE_ROWS:
        logger.warning(
            "结果超过 %d 行，写出可能占用较多内存与时间", BIG_TABLE_ROWS
        )
    if df.shape[0] == 0:
        logger.warning("结果没有任何数据行，将只写出表头")

    try:
        df.to_excel(out_path, index=False, engine="openpyxl")
    except PermissionError:
        raise ExcelMergeError(
            f"无法写入 {out_path}：没有写权限，或同名文件正被 Excel 打开"
        ) from None
    except Exception as exc:
        raise ExcelMergeError(f"写入 {out_path} 失败（{type(exc).__name__}: {exc}）") from None


# ---------------------------------------------------------------------------
# 命令行参数
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="merge_excel.py",
        description="合并多个 Excel 文件的数据（支持纵向拼接 / 按关键列横向合并 / 多工作表汇总）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""示例：
  python3 merge_excel.py -i ./数据 -m concat -o ./output
  python3 merge_excel.py -i ./数据 -m merge -k 学号 姓名 --how outer
  python3 merge_excel.py -i ./数据 -m sheets --recursive
  python3 merge_excel.py -f 单文件.xlsx -m sheets
  python3 merge_excel.py --gui          # 打开桌面端窗口
""",
    )

    src = parser.add_argument_group("输入")
    src.add_argument("-i", "--input-dir", default=".", help="要扫描的目录，默认当前目录")
    src.add_argument(
        "-f", "--files", nargs="*", default=None,
        help="直接指定要合并的文件（指定后不再扫描目录）",
    )
    src.add_argument(
        "--recursive", action="store_true",
        help="递归扫描子目录，默认只扫描当前层",
    )
    src.add_argument("--pattern", default="*.xls*", help="文件过滤通配符，默认 *.xls*")
    src.add_argument(
        "--sheet", default=0,
        help="concat/merge 模式读取的工作表：可写表名或序号，默认 0（第一个）",
    )

    mode = parser.add_argument_group("合并方式")
    mode.add_argument(
        "-m", "--mode", default="concat", choices=["concat", "merge", "sheets"],
        help="concat=纵向拼接（默认）；merge=按关键列横向合并；sheets=多工作表汇总",
    )
    mode.add_argument(
        "-k", "--keys", nargs="+", default=None,
        help="merge 模式的关键列，可指定多列，如 -k 学号 姓名",
    )
    mode.add_argument(
        "--how", default="left", choices=["left", "right", "outer", "inner"],
        help="merge 模式的连接方式，默认 left（以第一个文件为基准）",
    )
    mode.add_argument(
        "--keep-types", action="store_true",
        help="保留原始数据类型（数值可求和）；默认全部按文本读取，以保护身份证号等长数字串",
    )
    mode.add_argument(
        "--no-source", action="store_true",
        help="不加「来源文件 / 来源工作表」标记列（对 merge 模式无效）",
    )

    out = parser.add_argument_group("输出")
    out.add_argument(
        "-o", "--output-dir", default="./output",
        help="输出目录，默认 ./output（自动创建，不会覆盖输入文件）",
    )
    out.add_argument("--output", default=None, help="直接指定输出文件路径，优先于 -o")
    out.add_argument("-v", "--verbose", action="store_true", help="输出调试信息")

    ui = parser.add_argument_group("运行方式")
    ui.add_argument(
        "--gui", action="store_true",
        help="打开桌面端窗口（忽略其他参数），适合不想敲命令的情况",
    )

    return parser


# ---------------------------------------------------------------------------
# 统一入口：命令行和桌面端共用同一套流程
# ---------------------------------------------------------------------------

@dataclass
class MergeOptions:
    """一次合并任务的全部输入，字段与命令行参数一一对应。"""

    input_dir: str = "."
    files: list[str] | None = None
    recursive: bool = False
    pattern: str = "*.xls*"
    sheet: str | int = 0
    mode: str = "concat"
    keys: list[str] | None = None
    how: str = "left"
    keep_types: bool = False
    no_source: bool = False
    output_dir: str = "./output"
    output: str | None = None
    verbose: bool = False


@dataclass
class MergeOutcome:
    """一次合并任务的结果，code 与命令行退出码同义。"""

    code: int
    output_path: Path | None = None
    total_files: int = 0
    failures: list[tuple[Path, str]] = field(default_factory=list)


def resolve_inputs(opts: MergeOptions) -> list[Path]:
    """根据参数确定待合并文件列表。"""
    if opts.files:
        files = [Path(f).expanduser() for f in opts.files]
        missing = [str(f) for f in files if not f.is_file()]
        if missing:
            raise ExcelMergeError(f"以下文件不存在：{missing}")
        logger.info("已指定 %d 个文件", len(files))
        return files

    directory = Path(opts.input_dir).expanduser()
    files, skipped = scan_excel_files(directory, opts.pattern, opts.recursive)
    logger.info(
        "扫描目录：%s（%s）", directory, "含子目录" if opts.recursive else "仅当前层"
    )
    if skipped:
        logger.info("跳过 %d 个非目标文件（Excel 临时文件或隐藏文件）", len(skipped))
    if not files:
        raise ExcelMergeError(
            f"在 {directory} 中没有找到任何 Excel 文件。"
            f"该目录下的文件类型统计：{summarize_extensions(directory, opts.recursive)}"
        )
    logger.info("发现 %d 个 Excel 文件", len(files))
    return files


def run_merge(opts: MergeOptions) -> MergeOutcome:
    """执行一次合并。所有可预期错误都转成中文日志 + 非零退出码。"""
    # --sheet 传数字时转成 int，传表名时保持字符串
    sheet: str | int = opts.sheet
    if isinstance(sheet, str) and sheet.lstrip("-").isdigit():
        sheet = int(sheet)

    add_source = not opts.no_source
    as_str = not opts.keep_types

    try:
        files = resolve_inputs(opts)

        all_sheets = opts.mode == "sheets"
        items, failures = load_frames(files, sheet=sheet, as_str=as_str, all_sheets=all_sheets)

        if not items:
            raise ExcelMergeError("所有文件都读取失败或为空，没有可合并的数据")

        if opts.mode == "concat":
            result = mode_concat(items, add_source)
        elif opts.mode == "sheets":
            result = mode_sheets(items, add_source)
        else:
            if add_source:
                logger.info("提示：merge 为横向合并，每行本就来自多个文件，不添加来源列")
            result = mode_merge(items, opts.keys or [], opts.how)

        out_path = build_output_path(opts.output, opts.output_dir)
        write_result(result, out_path, files)

    except ExcelMergeError as exc:
        logger.error(str(exc))
        return MergeOutcome(code=2, total_files=0, failures=[])
    except KeyboardInterrupt:
        logger.error("已取消")
        return MergeOutcome(code=130, total_files=0, failures=[])

    logger.info("输出：%s", out_path.resolve())
    if failures:
        logger.warning(
            "完成，但有 %d 个文件读取失败（其余 %d 个已合并）", len(failures), len(files) - len(failures)
        )
        return MergeOutcome(code=1, output_path=out_path, total_files=len(files), failures=failures)
    logger.info("完成：已合并 %d 个文件", len(files))
    return MergeOutcome(code=0, output_path=out_path, total_files=len(files), failures=[])


# ---------------------------------------------------------------------------
# 命令行入口
# ---------------------------------------------------------------------------

def options_from_args(args: argparse.Namespace) -> MergeOptions:
    return MergeOptions(
        input_dir=args.input_dir,
        files=args.files,
        recursive=args.recursive,
        pattern=args.pattern,
        sheet=args.sheet,
        mode=args.mode,
        keys=args.keys,
        how=args.how,
        keep_types=args.keep_types,
        no_source=args.no_source,
        output_dir=args.output_dir,
        output=args.output,
        verbose=args.verbose,
    )


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.gui:
        from gui import launch
        return launch()

    setup_logging(args.verbose)
    return run_merge(options_from_args(args)).code


if __name__ == "__main__":
    sys.exit(main())

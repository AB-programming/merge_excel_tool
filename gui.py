#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 合并工具 —— 桌面端
=========================

界面只负责收集参数和显示日志，合并流程调用 merge_excel.run_merge，
和命令行入口共用同一套逻辑，不会出现两边行为不一致。

启动方式：
    python gui.py
    python merge_excel.py --gui
    双击 启动桌面版.bat / 打包后的 exe
"""

from __future__ import annotations

import logging
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk
from tkinter.scrolledtext import ScrolledText

import merge_excel as core

MODES = (
    ("concat", "纵向拼接 —— 几个文件表头一样，把数据行上下堆叠"),
    ("merge", "横向合并 —— 类似 VLOOKUP，按关键列把列拼起来"),
    ("sheets", "多表汇总 —— 把每个工作表（sheet）汇总成一张表"),
)

HOWS = ("left", "right", "outer", "inner")


class QueueHandler(logging.Handler):
    """把日志文字塞进队列，由主线程取走显示（Tkinter 不能跨线程操作控件）。"""

    def __init__(self, sink: queue.Queue) -> None:
        super().__init__()
        self.sink = sink

    def emit(self, record: logging.LogRecord) -> None:
        self.sink.put(("log", self.format(record)))


def reveal_in_folder(path: Path) -> None:
    if sys.platform == "win32":
        os.startfile(str(path.parent))
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path.parent)])
    else:
        subprocess.Popen(["xdg-open", str(path.parent)])


class MergeApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.sink: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self.selected_files: tuple[str, ...] = ()
        self.last_output: Path | None = None

        root.title("Excel 合并工具")
        root.minsize(760, 620)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(4, weight=1)

        self._build_ui()
        self._on_mode_change()
        self._append("提示：把要合并的 Excel 放进同一个文件夹，选好文件夹和合并方式，点「开始合并」。")
        root.after(80, self._drain)

    # ------------------------------------------------------------------ 界面

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 6}

        # --- 输入 ---
        in_box = ttk.LabelFrame(self.root, text=" 输入 ")
        in_box.grid(row=0, column=0, sticky="nsew", **pad)
        in_box.columnconfigure(0, weight=1)

        self.input_var = tk.StringVar(value=os.getcwd())
        self.recursive_var = tk.BooleanVar(value=False)
        self.pattern_var = tk.StringVar(value="*.xls*")
        self.sheet_var = tk.StringVar(value="0")
        self.files_label = ttk.Label(in_box, text="未选择具体文件")

        ttk.Label(in_box, text="文件夹：").grid(row=0, column=0, sticky="w", padx=8, pady=(6, 2))
        ttk.Entry(in_box, textvariable=self.input_var).grid(
            row=1, column=0, sticky="ew", padx=8
        )
        ttk.Button(in_box, text="浏览…", command=self._choose_dir).grid(
            row=0, column=1, rowspan=2, padx=8, pady=6, sticky="ns"
        )
        ttk.Button(in_box, text="选几个文件…", command=self._choose_files).grid(
            row=0, column=2, rowspan=2, padx=(0, 8), pady=6, sticky="ns"
        )
        self.files_label.grid(row=2, column=0, columnspan=3, sticky="w", padx=8)

        row3 = ttk.Frame(in_box)
        row3.grid(row=3, column=0, columnspan=3, sticky="w", padx=8, pady=(4, 8))
        ttk.Checkbutton(row3, text="包含子目录", variable=self.recursive_var).pack(side=tk.LEFT)
        ttk.Label(row3, text="   文件匹配：").pack(side=tk.LEFT)
        ttk.Entry(row3, textvariable=self.pattern_var, width=12).pack(side=tk.LEFT)

        # --- 合并方式 ---
        mode_box = ttk.LabelFrame(self.root, text=" 合并方式 ")
        mode_box.grid(row=1, column=0, sticky="nsew", **pad)
        mode_box.columnconfigure(0, weight=1)

        self.mode_var = tk.StringVar(value="concat")
        for i, (value, text) in enumerate(MODES):
            ttk.Radiobutton(
                mode_box, text=text, value=value, variable=self.mode_var,
                command=self._on_mode_change,
            ).grid(row=i, column=0, columnspan=4, sticky="w", padx=10, pady=2)

        self.keys_var = tk.StringVar()
        self.how_var = tk.StringVar(value="left")
        ttk.Label(mode_box, text="关键列（多个用空格分开，如：学号 姓名）：").grid(
            row=3, column=0, columnspan=2, sticky="w", padx=10, pady=(6, 2)
        )
        self.keys_entry = ttk.Entry(mode_box, textvariable=self.keys_var)
        self.keys_entry.grid(row=4, column=0, columnspan=2, sticky="ew", padx=10)

        ttk.Label(mode_box, text="连接方式：").grid(row=3, column=2, sticky="e", padx=(10, 2))
        self.how_combo = ttk.Combobox(
            mode_box, textvariable=self.how_var, values=HOWS, width=8, state="readonly"
        )
        self.how_combo.grid(row=4, column=2, padx=2, sticky="w")

        ttk.Label(mode_box, text="读取工作表：").grid(row=3, column=3, sticky="e", padx=(10, 2))
        self.sheet_entry = ttk.Entry(mode_box, textvariable=self.sheet_var, width=10)
        self.sheet_entry.grid(row=4, column=3, padx=(2, 10), sticky="w")

        # --- 选项与输出 ---
        opt_box = ttk.LabelFrame(self.root, text=" 选项与输出 ")
        opt_box.grid(row=2, column=0, sticky="nsew", **pad)
        opt_box.columnconfigure(0, weight=1)

        self.keep_types_var = tk.BooleanVar(value=False)
        self.no_source_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            opt_box,
            text="保留数字类型（默认按文本读取，可保护身份证号、手机号、前导零编号）",
            variable=self.keep_types_var,
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(6, 2))
        self.no_source_check = ttk.Checkbutton(
            opt_box, text="不添加「来源文件 / 来源工作表」标记列", variable=self.no_source_var
        )
        self.no_source_check.grid(row=1, column=0, columnspan=3, sticky="w", padx=10)

        self.output_var = tk.StringVar(value=str(Path(os.getcwd()) / "output"))
        ttk.Label(opt_box, text="输出到：").grid(row=2, column=0, sticky="w", padx=10, pady=(8, 2))
        ttk.Entry(opt_box, textvariable=self.output_var).grid(
            row=3, column=0, sticky="ew", padx=10
        )
        ttk.Button(opt_box, text="浏览…", command=self._choose_output).grid(
            row=2, column=1, rowspan=2, padx=6, pady=8, sticky="ns"
        )
        ttk.Button(opt_box, text="清空日志", command=self._clear_log).grid(
            row=2, column=2, rowspan=2, padx=(0, 10), pady=8, sticky="ns"
        )

        # --- 执行 ---
        action = ttk.Frame(self.root)
        action.grid(row=3, column=0, sticky="ew", padx=10)
        action.columnconfigure(1, weight=1)
        self.run_button = ttk.Button(action, text="开始合并", command=self._start)
        self.run_button.grid(row=0, column=0, pady=4, sticky="w")
        self.open_button = ttk.Button(
            action, text="打开结果所在文件夹", command=self._open_output, state="disabled"
        )
        self.open_button.grid(row=0, column=1, padx=8, pady=4, sticky="w")
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(action, textvariable=self.status_var).grid(
            row=0, column=2, pady=4, sticky="e"
        )

        # --- 日志 ---
        log_box = ttk.LabelFrame(self.root, text=" 运行日志 ")
        log_box.grid(row=4, column=0, sticky="nsew", **pad)
        log_box.columnconfigure(0, weight=1)
        log_box.rowconfigure(0, weight=1)
        self.log_view = ScrolledText(log_box, height=14, wrap=tk.WORD, state="disabled")
        self.log_view.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)

    # ------------------------------------------------------------ 界面事件

    def _choose_dir(self) -> None:
        path = filedialog.askdirectory(title="选择要合并的文件夹")
        if path:
            self.input_var.set(str(Path(path)))

    def _choose_output(self) -> None:
        path = filedialog.askdirectory(title="选择输出文件夹")
        if path:
            self.output_var.set(str(Path(path)))

    def _choose_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="选择要合并的 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xls *.xlsb"), ("所有文件", "*.*")],
        )
        if not paths:
            return
        self.selected_files = tuple(paths)
        self.files_label.config(text=f"已选择 {len(paths)} 个文件（将忽略上面的文件夹）")

    def _on_mode_change(self) -> None:
        mode = self.mode_var.get()
        self.keys_entry.config(state="normal" if mode == "merge" else "disabled")
        self.how_combo.config(state="readonly" if mode == "merge" else "disabled")
        self.sheet_entry.config(state="normal" if mode != "sheets" else "disabled")
        # merge 是横向拼列，每行本来就来自多个文件，来源列没有意义
        self.no_source_check.config(state="disabled" if mode == "merge" else "normal")

    def _open_output(self) -> None:
        if self.last_output is not None:
            reveal_in_folder(self.last_output)

    def _clear_log(self) -> None:
        self.log_view.config(state="normal")
        self.log_view.delete("1.0", tk.END)
        self.log_view.config(state="disabled")

    def _append(self, text: str) -> None:
        self.log_view.config(state="normal")
        self.log_view.insert(tk.END, text + "\n")
        self.log_view.see(tk.END)
        self.log_view.config(state="disabled")

    # ------------------------------------------------------------ 执行流程

    def collect_options(self) -> core.MergeOptions:
        """把界面上的选择换成一份命令行等价的参数。"""
        mode = self.mode_var.get()
        keys = [k for k in self.keys_var.get().split() if k]
        if mode == "merge" and not keys:
            raise ValueError("横向合并必须填写关键列，例如：学号")

        sheet = self.sheet_var.get().strip() or "0"
        return core.MergeOptions(
            input_dir=self.input_var.get().strip() or ".",
            files=list(self.selected_files) or None,
            recursive=self.recursive_var.get(),
            pattern=self.pattern_var.get().strip() or "*.xls*",
            sheet=sheet,
            mode=mode,
            keys=keys,
            how=self.how_var.get(),
            keep_types=self.keep_types_var.get(),
            no_source=self.no_source_var.get(),
            output_dir=self.output_var.get().strip() or "./output",
            verbose=False,
        )

    def _start(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            return
        try:
            options = self.collect_options()
        except ValueError as exc:
            self.status_var.set("参数不完整")
            self._append(f"[错误] {exc}")
            return

        self.last_output = None
        self.open_button.config(state="disabled")
        self.run_button.config(state="disabled")
        self.status_var.set("进行中…")
        self.worker = threading.Thread(target=self._run, args=(options,), daemon=True)
        self.worker.start()

    def _run(self, options: core.MergeOptions) -> None:
        """后台线程：跑合并，同时把日志塞进队列。"""
        core.setup_logging(console=False)
        handler = QueueHandler(self.sink)
        core.add_log_handler(handler)
        try:
            outcome = core.run_merge(options)
        except Exception as exc:  # 兜底：界面不能因为一次异常就卡死在「进行中」
            core.logger.error("未预期的错误：%s（%s）", type(exc).__name__, exc)
            outcome = core.MergeOutcome(code=3)
        finally:
            core.logger.removeHandler(handler)
        self.sink.put(("done", outcome))

    def _drain(self) -> None:
        try:
            while True:
                kind, payload = self.sink.get_nowait()
                if kind == "log":
                    self._append(payload)
                else:
                    self._finish(payload)
        except queue.Empty:
            pass
        try:
            self.root.after(80, self._drain)
        except tk.TclError:
            pass  # 窗口已关闭

    def _finish(self, outcome: core.MergeOutcome) -> None:
        self.run_button.config(state="normal")
        self.last_output = outcome.output_path
        if outcome.output_path is not None:
            self.open_button.config(state="normal")
        else:
            self.open_button.config(state="disabled")
        if outcome.code == 0:
            self.status_var.set("完成")
        elif outcome.code == 1:
            self.status_var.set(f"完成，但有 {len(outcome.failures)} 个文件读取失败")
        else:
            self.status_var.set("失败，请看日志")


def launch() -> int:
    root = tk.Tk()
    MergeApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(launch())

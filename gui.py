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

import ctypes
import logging
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, font as tkfont
from tkinter.scrolledtext import ScrolledText

import ttkbootstrap as ttk

import merge_excel as core

MODES = (
    ("concat", "纵向拼接 —— 几个文件表头一样，把数据行上下堆叠"),
    ("merge", "横向合并 —— 类似 VLOOKUP，按关键列把列拼起来"),
    ("sheets", "多表汇总 —— 把每个工作表（sheet）汇总成一张表"),
)

HOWS = ("left", "right", "outer", "inner")

LIGHT_THEME = "bootstrap-light"
DARK_THEME = "bootstrap-dark"


def _system_prefers_dark() -> bool:
    """读 Windows 个性化设置判断系统是否处于深色模式，失败时按浅色处理。"""
    if sys.platform != "win32":
        return False
    try:
        import winreg
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            return winreg.QueryValueEx(key, "AppsUseLightTheme")[0] == 0
    except OSError:
        return False


class QueueHandler(logging.Handler):
    """把日志文字塞进队列，由主线程取走显示（Tkinter 不能跨线程操作控件）。"""

    def __init__(self, sink: queue.Queue) -> None:
        super().__init__()
        self.sink = sink

    def emit(self, record: logging.LogRecord) -> None:
        self.sink.put(("log", record.levelname, self.format(record)))


def reveal_in_folder(path: Path) -> None:
    if sys.platform == "win32":
        os.startfile(str(path.parent))
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path.parent)])
    else:
        subprocess.Popen(["xdg-open", str(path.parent)])


class MergeApp:
    def __init__(self, root: ttk.Window) -> None:
        self.root = root
        self.style = root.style
        self.sink: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self.selected_files: tuple[str, ...] = ()
        self.last_output: Path | None = None

        root.columnconfigure(0, weight=1)
        root.rowconfigure(4, weight=1)

        self.theme = self.style.theme_mode or "light"
        self.dark_var = tk.BooleanVar(value=self.theme == "dark")
        self._set_fonts()
        self._sync_root_bg()

        self._build_ui()
        self._on_mode_change()
        self._append("提示：把要合并的 Excel 放进同一个文件夹，选好文件夹和合并方式，点「开始合并」。", "HINT")
        root.bind("<<ThemeChanged>>", self._on_theme_changed)
        root.after(80, self._drain)
        self._fit_to_content()

    def _fit_to_content(self) -> None:
        """按内容需要的高度定初始窗口大小，避免高 DPI 屏幕上内容被裁掉。"""
        self.root.update_idletasks()
        w = self.root.winfo_reqwidth() + 24
        h = self.root.winfo_reqheight() + 20
        w = min(w, int(self.root.winfo_screenwidth() * 0.9))
        h = min(h, int(self.root.winfo_screenheight() * 0.9))
        self.root.geometry(f"{w}x{h}")
        self._apply_titlebar()

    # ------------------------------------------------------------------ 主题

    def _set_fonts(self) -> None:
        """按屏幕 DPI 把默认字体放大到约 10.5 磅，主题切换后需要重新套用。"""
        scale = float(self.root.tk.call("tk", "scaling")) or 1.333
        body = -round(10.5 * scale)  # 负值表示像素，高分屏下等比放大
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
            try:
                font = tkfont.nametofont(name)
                font.configure(size=body)
                if sys.platform == "win32":
                    font.configure(family="Microsoft YaHei UI")
            except tk.TclError:
                pass

    def _log_colors(self) -> dict:
        """日志区配色直接取主题调色板，换主题时自动跟着变。"""
        c = self.style.colors
        return {"bg": c.inputbg, "fg": c.fg, "select": c.selectbg,
                "INFO": c.fg, "WARNING": c.warning, "ERROR": c.danger, "HINT": c.success}

    def _sync_root_bg(self) -> None:
        """让 Tk 根窗口底色跟随主题，避免深色模式下露出一圈灰底。"""
        self.root.configure(bg=self.style.colors.bg)

    def _apply_titlebar(self) -> None:
        """把 Windows 标题栏和 1 像素边框染成主题色，否则深色内容配白色标题栏很割裂。"""
        if sys.platform != "win32":
            return
        try:
            hwnd = int(self.root.wm_frame(), 16)
        except (tk.TclError, ValueError):
            return
        dark = ctypes.c_int(1 if self.theme == "dark" else 0)
        try:
            dwm = ctypes.windll.dwmapi
            for attr in (20, 19):  # Win11 用 20，旧版 Win10 用 19，不支持的会被系统忽略
                dwm.DwmSetWindowAttribute(hwnd, attr, ctypes.byref(dark), ctypes.sizeof(dark))
            # Win11 22000+：窗口边框也染成主题底色（BGR 顺序的 COLORREF）
            bg = self.style.colors.bg
            colorref = ctypes.c_int(int(bg[5:7] + bg[3:5] + bg[1:3], 16))
            dwm.DwmSetWindowAttribute(hwnd, 34, ctypes.byref(colorref), ctypes.sizeof(colorref))
        except (AttributeError, OSError):
            pass

    def _toggle_theme(self) -> None:
        self.style.theme_use(DARK_THEME if self.dark_var.get() else LIGHT_THEME)

    def _on_theme_changed(self, event: tk.Event | None = None) -> None:
        self.theme = self.style.theme_mode or "light"
        self._set_fonts()
        self._sync_root_bg()
        self._style_log()
        self._apply_titlebar()

    def _style_log(self) -> None:
        c = self._log_colors()
        self.log_view.config(
            bg=c["bg"], fg=c["fg"], insertbackground=c["fg"], selectbackground=c["select"]
        )
        for level in ("INFO", "WARNING", "ERROR", "HINT"):
            self.log_view.tag_configure(level, foreground=c[level])

    # ------------------------------------------------------------------ 界面

    def _build_ui(self) -> None:
        pad = {"padx": 12, "pady": 8}

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
        ttk.Button(in_box, text="浏览…", bootstyle="secondary-outline",
                   command=self._choose_dir).grid(
            row=0, column=1, rowspan=2, padx=8, pady=6, sticky="ns"
        )
        ttk.Button(in_box, text="选几个文件…", bootstyle="secondary-outline",
                   command=self._choose_files).grid(
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
        ttk.Checkbutton(
            opt_box, text="深色模式", variable=self.dark_var,
            bootstyle="success-round-toggle", command=self._toggle_theme,
        ).grid(row=1, column=1, columnspan=2, sticky="e", padx=(0, 10))

        self.output_var = tk.StringVar(value=str(Path(os.getcwd()) / "output"))
        ttk.Label(opt_box, text="输出到：").grid(row=2, column=0, sticky="w", padx=10, pady=(8, 2))
        ttk.Entry(opt_box, textvariable=self.output_var).grid(
            row=3, column=0, sticky="ew", padx=10
        )
        ttk.Button(opt_box, text="浏览…", bootstyle="secondary-outline",
                   command=self._choose_output).grid(
            row=2, column=1, rowspan=2, padx=6, pady=8, sticky="ns"
        )
        ttk.Button(opt_box, text="清空日志", bootstyle="secondary-outline",
                   command=self._clear_log).grid(
            row=2, column=2, rowspan=2, padx=(0, 10), pady=8, sticky="ns"
        )

        # --- 执行 ---
        action = ttk.Frame(self.root)
        action.grid(row=3, column=0, sticky="ew", padx=10)
        action.columnconfigure(2, weight=1)
        self.run_button = ttk.Button(
            action, text="开始合并", bootstyle="success", command=self._start
        )
        self.run_button.grid(row=0, column=0, pady=4, sticky="w")
        self.open_button = ttk.Button(
            action, text="打开结果所在文件夹", bootstyle="secondary-outline",
            command=self._open_output, state="disabled",
        )
        self.open_button.grid(row=0, column=1, padx=8, pady=4, sticky="w")
        self.progress = ttk.Progressbar(action, mode="indeterminate", bootstyle="success")
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(action, textvariable=self.status_var).grid(
            row=0, column=3, pady=4, sticky="e"
        )

        # --- 日志 ---
        log_box = ttk.LabelFrame(self.root, text=" 运行日志 ")
        log_box.grid(row=4, column=0, sticky="nsew", **pad)
        log_box.columnconfigure(0, weight=1)
        log_box.rowconfigure(0, weight=1)
        self.log_view = ScrolledText(
            log_box, height=12, wrap=tk.WORD, state="disabled",
            relief="flat", borderwidth=0, font="TkTextFont",
        )
        self.log_view.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        self._style_log()

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

    def _append(self, text: str, level: str = "INFO") -> None:
        tag = level if level in ("INFO", "WARNING", "ERROR", "HINT") else "INFO"
        self.log_view.config(state="normal")
        self.log_view.insert(tk.END, text + "\n", tag)
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
            self._append(f"[错误] {exc}", "ERROR")
            return

        self.last_output = None
        self.open_button.config(state="disabled")
        self.run_button.config(state="disabled")
        self.status_var.set("进行中…")
        self.progress.grid(row=0, column=2, sticky="ew", padx=8)
        self.progress.start(60)
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
                item = self.sink.get_nowait()
                if item[0] == "log":
                    self._append(item[2], item[1])
                else:
                    self._finish(item[1])
        except queue.Empty:
            pass
        try:
            self.root.after(80, self._drain)
        except tk.TclError:
            pass  # 窗口已关闭

    def _finish(self, outcome: core.MergeOutcome) -> None:
        self.progress.stop()
        self.progress.grid_remove()
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
    theme = DARK_THEME if _system_prefers_dark() else LIGHT_THEME
    # high_dpi 让 ttkbootstrap 自己处理高分辨率屏幕的缩放与清晰度
    root = ttk.Window(title="Excel 合并工具", themename=theme, minsize=(860, 680), high_dpi=True)
    MergeApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(launch())

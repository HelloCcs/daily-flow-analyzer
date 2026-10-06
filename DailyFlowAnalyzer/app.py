from __future__ import annotations

import os
import queue
import subprocess
import threading
import traceback
from datetime import date
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from . import APP_NAME, APP_VERSION
from .history_store import HistoryStore
from .paths import app_data_dir
from .providers import fetch_limit_pool, fetch_market_data, fetch_volume_history, validate_tushare_token
from .report_builder import DEFAULT_SORT_OPTION, SORT_OPTIONS, build_report
from .source_parser import SourceValidationError, parse_workbook
from .token_store import load_token, normalize_token, save_token


def _default_output_dir() -> Path:
    desktop = Path(os.environ.get("USERPROFILE", str(Path.home()))) / "Desktop"
    return desktop / "每日资金分析报表"


class DailyFlowApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"{APP_NAME} {APP_VERSION}")
        self.geometry("1120x760")
        self.minsize(940, 650)
        self.store = HistoryStore()
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.busy = False
        self.output_dir = tk.StringVar(value=str(_default_output_dir()))
        self.report_date = tk.StringVar()
        self.sort_option = tk.StringVar(value=DEFAULT_SORT_OPTION)
        self.sort_direction = tk.StringVar(value="降序")
        self.status_text = tk.StringVar(value="准备就绪")
        self.summary_text = tk.StringVar()
        self.token_status_text = tk.StringVar()
        self._configure_style()
        self._build_ui()
        self._refresh_history()
        self.after(100, self._poll_events)

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 18, "bold"), foreground="#17365D")
        style.configure("Sub.TLabel", font=("Microsoft YaHei UI", 10), foreground="#52606D")
        style.configure("Primary.TButton", font=("Microsoft YaHei UI", 10, "bold"), padding=(14, 8))
        style.configure("TButton", font=("Microsoft YaHei UI", 9), padding=(10, 6))
        style.configure("Treeview", font=("Microsoft YaHei UI", 9), rowheight=26)
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 9, "bold"))

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=16)
        root.pack(fill="both", expand=True)

        header = ttk.Frame(root)
        header.pack(fill="x")
        ttk.Label(header, text=APP_NAME, style="Title.TLabel").pack(side="left")
        ttk.Label(header, text="每日导入一张资金表，本机保存并自动滚动计算", style="Sub.TLabel").pack(side="left", padx=(18, 0), pady=(8, 0))
        ttk.Button(header, text="数据目录", command=lambda: self._open_path(app_data_dir())).pack(side="right")
        ttk.Button(header, text="Tushare 设置", command=self._token_dialog).pack(side="right", padx=8)
        ttk.Label(header, textvariable=self.token_status_text, style="Sub.TLabel").pack(side="right", pady=(8, 0))
        self._refresh_token_status()

        action = ttk.LabelFrame(root, text="每日操作", padding=12)
        action.pack(fill="x", pady=(14, 10))
        ttk.Button(action, text="导入每日资金表", style="Primary.TButton", command=self._choose_import).grid(row=0, column=0, rowspan=3, padx=(0, 14), sticky="nsew")
        ttk.Label(action, text="报告交易日").grid(row=0, column=1, sticky="w")
        self.date_combo = ttk.Combobox(action, textvariable=self.report_date, width=15, state="readonly")
        self.date_combo.grid(row=1, column=1, sticky="w", padx=(0, 18))
        ttk.Label(action, text="导出目录").grid(row=0, column=2, sticky="w")
        ttk.Entry(action, textvariable=self.output_dir).grid(row=1, column=2, sticky="ew", padx=(0, 8))
        ttk.Button(action, text="选择", command=self._choose_output).grid(row=1, column=3, padx=(0, 14))
        ttk.Label(action, text="全部工作表统一排序（生成时联动）").grid(row=2, column=1, sticky="w", pady=(10, 0))
        self.sort_combo = ttk.Combobox(action, textvariable=self.sort_option, values=list(SORT_OPTIONS), width=28, state="readonly")
        self.sort_combo.grid(row=2, column=2, sticky="w", pady=(10, 0))
        self.direction_combo = ttk.Combobox(action, textvariable=self.sort_direction, values=("降序", "升序"), width=8, state="readonly")
        self.direction_combo.grid(row=2, column=3, sticky="w", pady=(10, 0))
        ttk.Button(action, text="生成完整报表", style="Primary.TButton", command=self._generate).grid(row=0, column=4, rowspan=3, sticky="nsew")
        action.columnconfigure(2, weight=1)

        info = ttk.Frame(root)
        info.pack(fill="x", pady=(0, 10))
        ttk.Label(info, textvariable=self.summary_text, font=("Microsoft YaHei UI", 10, "bold"), foreground="#17365D").pack(side="left")
        ttk.Label(info, text="首次可一次多选历史文件；以后每天只选当天一张。漏日时补导即可自动重算。", style="Sub.TLabel").pack(side="right")

        middle = ttk.Panedwindow(root, orient="vertical")
        middle.pack(fill="both", expand=True)
        history_frame = ttk.LabelFrame(middle, text="已保存交易日", padding=8)
        log_frame = ttk.LabelFrame(middle, text="运行记录", padding=8)
        middle.add(history_frame, weight=3)
        middle.add(log_frame, weight=2)

        columns = ("date", "rows", "name", "imported")
        self.history = ttk.Treeview(history_frame, columns=columns, show="headings", selectmode="browse")
        self.history.heading("date", text="交易日期")
        self.history.heading("rows", text="股票数")
        self.history.heading("name", text="原文件")
        self.history.heading("imported", text="导入时间")
        self.history.column("date", width=120, anchor="center")
        self.history.column("rows", width=100, anchor="e")
        self.history.column("name", width=560)
        self.history.column("imported", width=180, anchor="center")
        history_scroll = ttk.Scrollbar(history_frame, orient="vertical", command=self.history.yview)
        self.history.configure(yscrollcommand=history_scroll.set)
        self.history.pack(side="left", fill="both", expand=True)
        history_scroll.pack(side="left", fill="y")
        history_buttons = ttk.Frame(history_frame)
        history_buttons.pack(side="right", fill="y", padx=(8, 0))
        ttk.Button(history_buttons, text="刷新", command=self._refresh_history).pack(fill="x", pady=(0, 6))
        ttk.Button(history_buttons, text="删除选中", command=self._delete_selected).pack(fill="x")

        self.log = tk.Text(log_frame, height=9, wrap="word", font=("Consolas", 9), bg="#F7F9FB", relief="flat")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=log_scroll.set)
        self.log.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")

        footer = ttk.Frame(root)
        footer.pack(fill="x", pady=(10, 0))
        self.progress = ttk.Progressbar(footer, mode="indeterminate", length=220)
        self.progress.pack(side="right")
        ttk.Label(footer, textvariable=self.status_text).pack(side="left")

    def _log(self, message: str) -> None:
        def write() -> None:
            self.log.insert("end", message.rstrip() + "\n")
            self.log.see("end")
        if threading.current_thread() is threading.main_thread():
            write()
        else:
            self.events.put(("log", message))

    def _set_busy(self, busy: bool, text: str = "") -> None:
        self.busy = busy
        self.status_text.set(text or ("正在处理..." if busy else "准备就绪"))
        if busy:
            self.progress.start(12)
        else:
            self.progress.stop()

    def _choose_import(self) -> None:
        if self.busy:
            return
        paths = filedialog.askopenfilenames(title="选择每日资金流向表（首次可多选）", filetypes=[("Excel 工作簿", "*.xlsx")])
        if not paths:
            return
        self._set_busy(True, "正在校验并导入...")
        threading.Thread(target=self._import_worker, args=(paths,), daemon=True).start()

    def _import_worker(self, paths: tuple[str, ...]) -> None:
        conflicts = []
        try:
            for index, path in enumerate(paths, start=1):
                self._log(f"[{index}/{len(paths)}] 校验 {Path(path).name}")
                parsed = parse_workbook(path)
                result = self.store.import_workbook(parsed)
                self._log(result.message)
                for warning in parsed.warnings:
                    self._log("提示：" + warning)
                if result.status == "conflict":
                    conflicts.append(parsed)
            self.events.put(("import_done", conflicts))
        except Exception:
            self.events.put(("error", traceback.format_exc()))

    def _resolve_conflicts(self, conflicts) -> None:
        for parsed in conflicts:
            if messagebox.askyesno("同日文件不同", f"{parsed.trade_date} 已有不同文件。\n\n是否用“{parsed.path.name}”替换并重算？"):
                result = self.store.import_workbook(parsed, replace=True)
                self._log(result.message)
        self._refresh_history()
        self._set_busy(False, "导入完成")

    def _refresh_history(self) -> None:
        records = self.store.import_records()
        self.history.delete(*self.history.get_children())
        for record in records:
            self.history.insert("", "end", iid=record["trade_date"], values=(record["trade_date"], f'{record["row_count"]:,}', record["original_name"], record["imported_at"]))
        dates = [record["trade_date"] for record in records]
        self.date_combo["values"] = dates
        if dates and self.report_date.get() not in dates:
            self.report_date.set(dates[0])
        self.summary_text.set(f"已保存 {len(dates)} 个交易日" + (f"，最新 {dates[0]}" if dates else "，请先导入样表"))

    def _delete_selected(self) -> None:
        selection = self.history.selection()
        if not selection or self.busy:
            return
        day = date.fromisoformat(selection[0])
        if messagebox.askyesno("删除交易日", f"确定从历史库删除 {day}？\n原始导入文件副本仍保留，可重新导入。"):
            self.store.delete_date(day)
            self._log(f"已从历史库删除 {day}。")
            self._refresh_history()

    def _choose_output(self) -> None:
        selected = filedialog.askdirectory(title="选择报表导出目录", initialdir=self.output_dir.get())
        if selected:
            self.output_dir.set(selected)

    def _generate(self) -> None:
        if self.busy:
            return
        if not self.report_date.get():
            messagebox.showwarning("缺少交易日", "请先导入至少一个每日资金表。")
            return
        if not load_token():
            messagebox.showwarning(
                "Tushare 尚未验证",
                "完整报表需要先设置并验证 Tushare Token。\n\n"
                "请点击右上角“Tushare 设置”，粘贴 Token 后等待验证成功。",
            )
            return
        report_date = date.fromisoformat(self.report_date.get())
        available_days = len([day for day in self.store.imported_dates() if day <= report_date])
        if available_days < 11 and not messagebox.askyesno(
            "历史交易日不足",
            f"当前只有 {available_days} 个交易日。\n"
            "10日双窗口需要11个交易日，相关两项会留空。\n\n"
            "是否仍然生成？",
        ):
            return
        output = Path(self.output_dir.get()) / f"{report_date:%Y.%m.%d}_每日资金综合分析.xlsx"
        sort_option = self.sort_option.get() or DEFAULT_SORT_OPTION
        sort_descending = self.sort_direction.get() != "升序"
        self._set_busy(True, "正在获取联网数据并生成完整报表...")
        self._log(f"开始生成 {report_date} 完整报表；全部工作表按“{sort_option}”{'降序' if sort_descending else '升序'}同步排列。")
        threading.Thread(target=self._generate_worker, args=(report_date, output, sort_option, sort_descending), daemon=True).start()

    def _generate_worker(self, report_date: date, output: Path, sort_option: str, sort_descending: bool) -> None:
        try:
            token = load_token()
            if token:
                market = fetch_market_data(token, report_date, self._log)
                volume_dates = sorted(
                    (day for day in self.store.imported_dates() if day <= report_date),
                    reverse=True,
                )[:11]
                volume_history = fetch_volume_history(token, volume_dates, self._log)
            else:
                market = {}
                volume_history = {}
                self._log("未设置 Tushare Token：资金分析继续生成，收盘价、总市值和真实成交量留空。")
            limits = fetch_limit_pool(report_date, self._log)
            result = build_report(
                self.store, report_date, market, limits, output, self._log,
                sort_option=sort_option, sort_descending=sort_descending,
                volume_history=volume_history,
            )
            self.events.put(("generated", result))
        except Exception:
            self.events.put(("error", traceback.format_exc()))

    def _token_dialog(self) -> None:
        current = load_token()
        prompt = (
            "粘贴纯 Token、MCP URL 或完整 MCP JSON。\n"
            "软件会先联网验证 daily 和 daily_basic，成功后才在本机加密保存。"
        )
        token = simpledialog.askstring("Tushare 设置", prompt, initialvalue=current, show="*", parent=self)
        if token is None:
            return
        try:
            normalized = normalize_token(token)
            if not normalized:
                save_token("")
                self._refresh_token_status()
                self._log("Tushare Token 已清除。")
                return
            if not self.report_date.get():
                messagebox.showwarning("请先导入数据", "请先导入至少一个交易日文件，再验证 Token。")
                return
            trade_date = date.fromisoformat(self.report_date.get())
            self.token_status_text.set("Tushare：正在验证...")
            self._set_busy(True, "正在验证 Tushare Token...")
            threading.Thread(
                target=self._validate_token_worker,
                args=(normalized, trade_date),
                daemon=True,
            ).start()
        except Exception as exc:
            messagebox.showerror("保存失败", str(exc))

    def _validate_token_worker(self, token: str, trade_date: date) -> None:
        try:
            daily_count, basic_count = validate_tushare_token(token, trade_date, self._log)
            save_token(token)
            self.events.put(("token_validated", (daily_count, basic_count)))
        except Exception:
            self.events.put(("token_error", traceback.format_exc()))

    def _refresh_token_status(self) -> None:
        self.token_status_text.set("Tushare：已保存" if load_token() else "Tushare：未设置")

    def _poll_events(self) -> None:
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == "log":
                    self.log.insert("end", str(value).rstrip() + "\n")
                    self.log.see("end")
                elif kind == "import_done":
                    self._resolve_conflicts(value)
                elif kind == "generated":
                    self._set_busy(False, "报表生成并验收通过")
                    messagebox.showinfo("生成完成", f"完整报表已生成并通过结构验收：\n{value}")
                    self._open_path(Path(value).parent)
                elif kind == "token_validated":
                    daily_count, basic_count = value
                    self._set_busy(False, "Tushare Token 验证成功")
                    self.token_status_text.set("Tushare：已验证")
                    self._log(f"Tushare Token 验证成功：日线 {daily_count:,} 只，市值 {basic_count:,} 只。")
                    messagebox.showinfo(
                        "验证成功",
                        f"Token 已验证并在本机加密保存。\n"
                        f"日线 {daily_count:,} 只，市值 {basic_count:,} 只。",
                    )
                elif kind == "token_error":
                    self._set_busy(False, "Tushare Token 验证失败")
                    self.token_status_text.set("Tushare：验证失败")
                    details = str(value)
                    self._log(details)
                    last_line = next((line for line in reversed(details.splitlines()) if line.strip()), "未知错误")
                    messagebox.showerror("Tushare 验证失败", last_line)
                elif kind == "error":
                    self._set_busy(False, "处理失败")
                    details = str(value)
                    self._log(details)
                    last_line = next((line for line in reversed(details.splitlines()) if line.strip()), "未知错误")
                    messagebox.showerror("处理失败", last_line)
        except queue.Empty:
            pass
        self.after(100, self._poll_events)

    @staticmethod
    def _open_path(path: Path) -> None:
        path.mkdir(parents=True, exist_ok=True)
        os.startfile(path)


def main() -> None:
    app = DailyFlowApp()
    app.mainloop()


if __name__ == "__main__":
    main()

import json
import subprocess
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import main

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    DND_FILES = None
    TkinterDnD = None


BaseWindow = TkinterDnD.Tk if TkinterDnD else tk.Tk
GUI_STATE_FILE = main.LOG_DIR / "gui_state.json"
APP_TITLE = "売上データ自動集計・月末売上管理ツール"
RUN_TAB_DESCRIPTION = (
    "CSV・Excelの売上データを読み込み、月末確認用のExcelレポートを作成します。\n"
    "入力データを選び、対象月を指定して、プレビュー確認後にレポートを作成してください。"
)
RUN_TAB_STEPS = (
    "操作手順:\n"
    "1. 入力フォルダまたは売上ファイルを選択\n"
    "2. 出力フォルダを確認\n"
    "3. 対象月を指定\n"
    "4. プレビューで内容を確認\n"
    "5. Excelレポートを作成"
)


class MonthlyReportApp(BaseWindow):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1120x860")
        self.resizable(True, True)
        self.report_history: list[Path] = []
        self.audit_records: dict[str, dict] = {}
        self._init_variables()
        self._build_widgets()
        self._load_state()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _init_variables(self) -> None:
        self.input_dir = tk.StringVar(value=str(main.INPUT_DIR))
        self.output_dir = tk.StringVar(value=str(main.OUTPUT_DIR))
        self.config_file = tk.StringVar()
        self.preset = tk.StringVar()
        self.month = tk.StringVar()
        self.start_date = tk.StringVar()
        self.end_date = tk.StringVar()
        self.product = tk.StringVar()
        self.category = tk.StringVar()
        self.pattern = tk.StringVar(value=main.DEFAULT_PATTERN)
        self.output_name = tk.StringVar()
        self.summary_csv_dir = tk.StringVar()
        self.group_by = tk.StringVar(value="product")
        self.all_summaries = tk.BooleanVar(value=True)
        self.monthly_trend = tk.BooleanVar(value=True)
        self.charts = tk.BooleanVar(value=True)
        self.dry_run = tk.BooleanVar(value=False)
        self.notify = tk.BooleanVar(value=False)
        self.notify_webhook_url = tk.StringVar()
        self.column_aliases_json = tk.StringVar(value="{}")
        self.warning_amount_threshold = tk.StringVar(value="1000000")
        self.style_header_fill = tk.StringVar(value="D9EAF7")
        self.style_total_fill = tk.StringVar(value="EAF4E2")
        self.style_title_size = tk.StringVar(value="14")
        self.style_chart_height = tk.StringVar(value="7")
        self.style_chart_width = tk.StringVar(value="14")
        self.audit_keep_count = tk.StringVar()
        self.audit_keep_days = tk.StringVar()
        self.schedule_day = tk.StringVar(value="1")
        self.schedule_time = tk.StringVar(value="09:00")
        self.audit_filter_text = tk.StringVar()
        self.audit_filter_status = tk.StringVar(value="all")
        self.audit_filter_date_from = tk.StringVar()
        self.audit_filter_date_to = tk.StringVar()
        self.status = tk.StringVar(value="待機中")

    def _build_widgets(self) -> None:
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True)

        run_tab = ttk.Frame(notebook, padding=12)
        history_tab = ttk.Frame(notebook, padding=12)
        audit_tab = ttk.Frame(notebook, padding=12)
        log_tab = ttk.Frame(notebook, padding=12)
        notebook.add(run_tab, text="実行")
        notebook.add(history_tab, text="レポート履歴")
        notebook.add(audit_tab, text="監査履歴")
        notebook.add(log_tab, text="ログ")

        self._build_run_tab(run_tab)
        self._build_history_tab(history_tab)
        self._build_audit_tab(audit_tab)
        self._build_log_tab(log_tab)

    def _build_run_tab(self, root: ttk.Frame) -> None:
        intro = ttk.Label(root, text=RUN_TAB_DESCRIPTION, justify=tk.LEFT, wraplength=940)
        intro.grid(row=0, column=0, columnspan=4, sticky=tk.EW, pady=(0, 8))
        steps = ttk.Label(root, text=RUN_TAB_STEPS, justify=tk.LEFT, relief=tk.GROOVE, padding=10)
        steps.grid(row=1, column=0, columnspan=4, sticky=tk.EW, pady=(0, 12))

        fields = [
            ("設定ファイル", self.config_file, self._choose_config_file, "保存済み設定を使う場合に指定します。"),
            ("入力フォルダ", self.input_dir, self._choose_input_dir, "CSVまたはExcelファイルを置くフォルダを指定します。ファイル選択もできます。"),
            ("出力フォルダ", self.output_dir, self._choose_output_dir, "作成したExcelレポートの保存先です。"),
            ("対象月", self.month, None, "例: 2026-04。月単位で集計する場合に指定します。"),
            ("開始日", self.start_date, None, "例: 2026-04-01。任意期間で集計する場合に指定します。"),
            ("終了日", self.end_date, None, "例: 2026-04-30。開始日とセットで使います。"),
            ("商品名で絞り込み", self.product, None, "空欄の場合はすべての商品を対象にします。"),
            ("カテゴリで絞り込み", self.category, None, "空欄の場合はすべてのカテゴリを対象にします。"),
            ("読み込みファイル名パターン", self.pattern, None, "例: sales_*.csv。CSV・Excelの名前規則を指定します。"),
            ("出力ファイル名", self.output_name, None, "空欄の場合は日時付きのファイル名で自動作成します。"),
            ("集計CSVフォルダ", self.summary_csv_dir, self._choose_summary_csv_dir, "集計結果CSVも出力したい場合だけ指定します。"),
            ("通知Webhook URL", self.notify_webhook_url, None, "通知連携を使う場合だけ指定します。"),
        ]
        field_start_row = 2
        for index, (label, variable, command, help_text) in enumerate(fields):
            row = field_start_row + index
            ttk.Label(root, text=label).grid(row=row, column=0, sticky=tk.W, pady=4)
            ttk.Entry(root, textvariable=variable).grid(row=row, column=1, sticky=tk.EW, pady=4)
            if command:
                ttk.Button(root, text="選択", command=command).grid(row=row, column=2, sticky=tk.EW, padx=(8, 0), pady=4)
            ttk.Label(root, text=help_text, foreground="#555555", wraplength=300).grid(row=row, column=3, sticky=tk.W, padx=(8, 0), pady=4)

        option_row = field_start_row + len(fields)
        ttk.Label(root, text="集計単位").grid(row=option_row, column=0, sticky=tk.W, pady=4)
        ttk.Combobox(root, textvariable=self.group_by, values=["product", "category"], state="readonly").grid(
            row=option_row,
            column=1,
            sticky=tk.W,
            pady=4,
        )
        ttk.Label(root, text="プリセット").grid(row=option_row, column=2, sticky=tk.W, padx=(8, 0), pady=4)
        self.preset_combo = ttk.Combobox(root, textvariable=self.preset, values=[], state="readonly", width=24)
        self.preset_combo.grid(row=option_row, column=3, sticky=tk.EW, pady=4)

        checks = ttk.Frame(root)
        checks.grid(row=option_row + 1, column=0, columnspan=4, sticky=tk.W, pady=8)
        ttk.Checkbutton(checks, text="商品別とカテゴリ別を両方出力", variable=self.all_summaries).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(checks, text="月別推移", variable=self.monthly_trend).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(checks, text="グラフ", variable=self.charts).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(checks, text="検証のみ", variable=self.dry_run).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(checks, text="通知音", variable=self.notify).pack(side=tk.LEFT)

        action_row = option_row + 2
        ttk.Button(root, text="Excelレポートを作成", command=self._run_async).grid(row=action_row, column=0, sticky=tk.EW, pady=8)
        ttk.Button(root, text="データをプレビュー", command=self._preview).grid(row=action_row, column=1, sticky=tk.EW, padx=(8, 0), pady=8)
        ttk.Button(root, text="集計プレビュー", command=self._summary_preview).grid(row=action_row, column=2, sticky=tk.EW, padx=(8, 0), pady=8)
        ttk.Button(root, text="プリセット適用", command=self._apply_selected_preset).grid(row=action_row, column=3, sticky=tk.EW, padx=(8, 0), pady=8)

        action_row += 1
        ttk.Button(root, text="設定読込", command=self._load_config).grid(row=action_row, column=0, sticky=tk.EW, pady=(0, 8))
        ttk.Button(root, text="設定保存", command=self._save_config).grid(row=action_row, column=1, sticky=tk.EW, padx=(8, 0), pady=(0, 8))
        ttk.Button(root, text="CSVテンプレート作成", command=self._write_template).grid(row=action_row, column=2, sticky=tk.EW, padx=(8, 0), pady=(0, 8))
        ttk.Button(root, text="タスク登録", command=self._register_schedule).grid(row=action_row, column=3, sticky=tk.EW, padx=(8, 0), pady=(0, 8))

        ttk.Button(root, text="設定ウィザード", command=self._open_config_wizard).grid(row=action_row + 1, column=0, sticky=tk.EW, pady=(0, 8))

        ttk.Button(root, text="設定内容チェック", command=self._review_current_settings).grid(row=action_row + 1, column=1, sticky=tk.EW, padx=(8, 0), pady=(0, 8))
        ttk.Button(root, text="文字化け修復", command=self._repair_legacy_text_files).grid(row=action_row + 1, column=2, sticky=tk.EW, padx=(8, 0), pady=(0, 8))

        schedule_row = action_row + 2
        ttk.Label(root, text="定期実行 日/時刻").grid(row=schedule_row, column=0, sticky=tk.W, pady=(0, 8))
        schedule_frame = ttk.Frame(root)
        schedule_frame.grid(row=schedule_row, column=1, sticky=tk.W, pady=(0, 8))
        ttk.Entry(schedule_frame, textvariable=self.schedule_day, width=5).pack(side=tk.LEFT)
        ttk.Entry(schedule_frame, textvariable=self.schedule_time, width=8).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(root, text="列名マッピング", command=self._edit_column_aliases).grid(
            row=schedule_row,
            column=2,
            sticky=tk.EW,
            padx=(8, 0),
            pady=(0, 8),
        )
        ttk.Button(root, text="Excel見た目設定", command=self._edit_style_config).grid(
            row=schedule_row,
            column=3,
            sticky=tk.EW,
            padx=(8, 0),
            pady=(0, 8),
        )

        audit_row = schedule_row + 1
        ttk.Label(root, text="監査保持 件数/日数").grid(row=audit_row, column=0, sticky=tk.W, pady=(0, 8))
        audit_frame = ttk.Frame(root)
        audit_frame.grid(row=audit_row, column=1, sticky=tk.W, pady=(0, 8))
        ttk.Entry(audit_frame, textvariable=self.audit_keep_count, width=8).pack(side=tk.LEFT)
        ttk.Entry(audit_frame, textvariable=self.audit_keep_days, width=8).pack(side=tk.LEFT, padx=(8, 0))

        drop_row = audit_row + 1
        self.drop_label = ttk.Label(root, text="CSV/Excelをここへドロップすると入力フォルダを設定します", relief=tk.GROOVE, anchor=tk.CENTER, padding=16)
        self.drop_label.grid(row=drop_row, column=0, columnspan=4, sticky=tk.EW, pady=(4, 8))
        self._setup_drop_target()

        status_row = drop_row + 1
        ttk.Label(root, textvariable=self.status).grid(row=status_row, column=0, columnspan=4, sticky=tk.W, pady=(0, 4))
        self.output = tk.Text(root, height=10)
        self.output.grid(row=status_row + 1, column=0, columnspan=4, sticky=tk.NSEW, pady=(8, 0))

        root.columnconfigure(1, weight=1)
        root.columnconfigure(3, weight=1)
        root.rowconfigure(status_row + 1, weight=1)

    def _build_history_tab(self, root: ttk.Frame) -> None:
        ttk.Label(
            root,
            text="作成済みのExcelレポートを確認できます。レポート作成後に一覧へ表示されます。",
            justify=tk.LEFT,
            wraplength=900,
        ).pack(fill=tk.X, pady=(0, 8))
        buttons = ttk.Frame(root)
        buttons.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(buttons, text="履歴を更新", command=self._refresh_history).pack(side=tk.LEFT)
        ttk.Button(buttons, text="詳細", command=self._show_selected_report_detail).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(buttons, text="レポートを開く", command=self._open_selected_report).pack(side=tk.LEFT, padx=(8, 0))

        self.history_list = tk.Listbox(root, height=12)
        self.history_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.history_list.bind("<Double-Button-1>", lambda _event: self._show_selected_report_detail())
        history_scroll = ttk.Scrollbar(root, orient=tk.VERTICAL, command=self.history_list.yview)
        history_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.history_list.configure(yscrollcommand=history_scroll.set)
        self._refresh_history()

    def _build_audit_tab(self, root: ttk.Frame) -> None:
        ttk.Label(
            root,
            text="実行履歴やエラー内容を確認できます。レポート作成時に問題が起きた場合は、この画面を確認してください。",
            justify=tk.LEFT,
            wraplength=900,
        ).pack(fill=tk.X, pady=(0, 8))
        buttons = ttk.Frame(root)
        buttons.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(buttons, text="履歴を更新", command=self._refresh_audit_table).pack(side=tk.LEFT)
        ttk.Button(buttons, text="CSVエクスポート", command=self._export_audit_csv).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(buttons, text="要約CSV", command=self._export_audit_summary_csv).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(buttons, text="月別要約CSV", command=self._export_audit_monthly_summary_csv).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(buttons, text="バックアップ", command=self._backup_audit_log).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(buttons, text="異常チェック", command=self._show_audit_anomalies).pack(side=tk.LEFT, padx=(8, 0))

        ttk.Button(buttons, text="履歴から条件復元", command=self._restore_selected_audit_record).pack(side=tk.LEFT, padx=(8, 0))

        ttk.Label(buttons, text="検索").pack(side=tk.LEFT, padx=(16, 4))
        ttk.Entry(buttons, textvariable=self.audit_filter_text, width=24).pack(side=tk.LEFT)
        ttk.Combobox(buttons, textvariable=self.audit_filter_status, values=["all", "success", "validation_error", "error"], state="readonly", width=16).pack(side=tk.LEFT, padx=(8, 0))
        filters = ttk.Frame(root)
        filters.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(filters, text="期間").pack(side=tk.LEFT)
        ttk.Entry(filters, textvariable=self.audit_filter_date_from, width=12).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Label(filters, text="〜").pack(side=tk.LEFT, padx=4)
        ttk.Entry(filters, textvariable=self.audit_filter_date_to, width=12).pack(side=tk.LEFT)
        ttk.Button(filters, text="適用", command=self._refresh_audit_table).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(filters, text="修復プレビュー", command=self._preview_legacy_text_repairs).pack(side=tk.RIGHT)

        columns = ("timestamp", "status", "period", "detail_count", "summary_count", "warnings", "output_file", "error")
        table_frame = ttk.Frame(root)
        table_frame.pack(fill=tk.BOTH, expand=True)
        self.audit_tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        y_scroll = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.audit_tree.yview)
        x_scroll = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.audit_tree.xview)
        self.audit_tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.audit_tree.grid(row=0, column=0, sticky=tk.NSEW)
        y_scroll.grid(row=0, column=1, sticky=tk.NS)
        x_scroll.grid(row=1, column=0, sticky=tk.EW)
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        headings = {
            "timestamp": "日時",
            "status": "状態",
            "period": "期間",
            "detail_count": "明細",
            "summary_count": "集計",
            "output_file": "出力",
            "error": "エラー",
        }
        headings["warnings"] = "警告"
        for column in columns:
            self.audit_tree.heading(column, text=headings[column])
            self.audit_tree.column(column, width=120, minwidth=70, stretch=True)
        self.audit_tree.column("output_file", width=260)
        self.audit_tree.column("error", width=260)
        self._refresh_audit_table()

    def _build_log_tab(self, root: ttk.Frame) -> None:
        ttk.Label(
            root,
            text="実行ログや監査ログの内容を確認できます。エラーが出た場合の原因確認に使用します。",
            justify=tk.LEFT,
            wraplength=900,
        ).pack(fill=tk.X, pady=(0, 8))
        buttons = ttk.Frame(root)
        buttons.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(buttons, text="ログ更新", command=self._refresh_log_text).pack(side=tk.LEFT)
        ttk.Button(buttons, text="監査ログ表示", command=self._refresh_audit_log_text).pack(side=tk.LEFT, padx=(8, 0))

        self.log_text = tk.Text(root, height=12)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self._refresh_log_text()

    def _setup_drop_target(self) -> None:
        if DND_FILES is None:
            self.drop_label.configure(text="ドラッグ＆ドロップには tkinterdnd2 が必要です。選択ボタンでも設定できます。")
            return
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind("<<Drop>>", self._handle_drop)

    def _handle_drop(self, event) -> None:
        paths = self.tk.splitlist(event.data)
        if not paths:
            return
        first_path = Path(paths[0])
        self.input_dir.set(str(first_path if first_path.is_dir() else first_path.parent))
        self.status.set("入力フォルダをドロップ内容から設定しました")

    def _choose_input_dir(self) -> None:
        choose_file = messagebox.askyesno("入力選択", "CSV/Excelファイルを選択しますか？\nいいえを選ぶとフォルダを選択します。")
        if choose_file:
            selected = filedialog.askopenfilename(
                initialdir=self.input_dir.get() or ".",
                filetypes=[
                    ("対応ファイル", "*.csv;*.xlsx;*.xls"),
                    ("CSVファイル", "*.csv"),
                    ("Excelファイル", "*.xlsx;*.xls"),
                    ("All files", "*.*"),
                ],
            )
        else:
            selected = filedialog.askdirectory(initialdir=self.input_dir.get() or ".")
        if selected:
            self.input_dir.set(selected)

    def _choose_config_file(self) -> None:
        selected = filedialog.askopenfilename(filetypes=[("JSON", "*.json"), ("All files", "*.*")])
        if selected:
            self.config_file.set(selected)
            self._load_config_from_file(Path(selected))

    def _choose_output_dir(self) -> None:
        selected = filedialog.askdirectory(initialdir=self.output_dir.get() or ".")
        if selected:
            self.output_dir.set(selected)

    def _choose_summary_csv_dir(self) -> None:
        selected = filedialog.askdirectory(initialdir=self.summary_csv_dir.get() or self.output_dir.get() or ".")
        if selected:
            self.summary_csv_dir.set(selected)

    def _config_dict(self) -> dict:
        return {
            "input_dir": self.input_dir.get(),
            "output_dir": self.output_dir.get(),
            "month": self._optional(self.month.get()),
            "start_date": self._optional(self.start_date.get()),
            "end_date": self._optional(self.end_date.get()),
            "product": self._optional(self.product.get()),
            "category": self._optional(self.category.get()),
            "pattern": self.pattern.get() or main.DEFAULT_PATTERN,
            "output_name": self._optional(self.output_name.get()),
            "group_by": self.group_by.get(),
            "all_summaries": self.all_summaries.get(),
            "monthly_trend": self.monthly_trend.get(),
            "charts": self.charts.get(),
            "dry_run": self.dry_run.get(),
            "notify": self.notify.get(),
            "notify_webhook_url": self._optional(self.notify_webhook_url.get()),
            "summary_csv_dir": self._optional(self.summary_csv_dir.get()),
            "column_aliases": self._column_aliases(),
            "warning_amount_threshold": self._warning_amount_threshold(),
            "style": self._style_config(),
            "audit_keep_count": self._optional_int(self.audit_keep_count.get()),
            "audit_keep_days": self._optional_int(self.audit_keep_days.get()),
        }

    def _state_dict(self) -> dict:
        state = self._config_dict()
        state.update(
            {
                "config_file": self._optional(self.config_file.get()),
                "preset": self._optional(self.preset.get()),
                "schedule_day": self.schedule_day.get(),
                "schedule_time": self.schedule_time.get(),
            }
        )
        return state

    def _apply_config(self, config: dict) -> None:
        self.input_dir.set(config.get("input_dir") or str(main.INPUT_DIR))
        self.output_dir.set(config.get("output_dir") or str(main.OUTPUT_DIR))
        self.month.set(config.get("month") or "")
        self.start_date.set(config.get("start_date") or "")
        self.end_date.set(config.get("end_date") or "")
        self.product.set(config.get("product") or "")
        self.category.set(config.get("category") or "")
        self.pattern.set(config.get("pattern") or main.DEFAULT_PATTERN)
        self.output_name.set(config.get("output_name") or "")
        self.group_by.set(config.get("group_by") or "product")
        self.all_summaries.set(bool(config.get("all_summaries", True)))
        self.monthly_trend.set(bool(config.get("monthly_trend", True)))
        self.charts.set(bool(config.get("charts", True)))
        self.dry_run.set(bool(config.get("dry_run", False)))
        self.notify.set(bool(config.get("notify", False)))
        self.notify_webhook_url.set(config.get("notify_webhook_url") or "")
        self.summary_csv_dir.set(config.get("summary_csv_dir") or "")
        self.column_aliases_json.set(json.dumps(config.get("column_aliases", {}), ensure_ascii=False, indent=2))
        self.warning_amount_threshold.set(str(config.get("warning_amount_threshold", 1_000_000)))
        style = config.get("style") or {}
        self.style_header_fill.set(str(style.get("header_fill", "D9EAF7")))
        self.style_total_fill.set(str(style.get("total_fill", "EAF4E2")))
        self.style_title_size.set(str(style.get("title_size", "14")))
        self.style_chart_height.set(str(style.get("chart_height", "7")))
        self.style_chart_width.set(str(style.get("chart_width", "14")))
        self.audit_keep_count.set("" if config.get("audit_keep_count") is None else str(config.get("audit_keep_count")))
        self.audit_keep_days.set("" if config.get("audit_keep_days") is None else str(config.get("audit_keep_days")))

    def _load_state(self) -> None:
        if not GUI_STATE_FILE.exists():
            self._refresh_history()
            self._refresh_audit_table()
            return
        try:
            state = json.loads(GUI_STATE_FILE.read_text(encoding="utf-8"))
            if isinstance(state, dict):
                state = main.normalize_legacy_value(state)
                self._apply_config(state)
                self.config_file.set(state.get("config_file") or "")
                self.schedule_day.set(state.get("schedule_day") or "1")
                self.schedule_time.set(state.get("schedule_time") or "09:00")
                if self.config_file.get():
                    self._load_presets_only(Path(self.config_file.get()))
                    self.preset.set(state.get("preset") or self.preset.get())
        except Exception as exc:
            self.status.set(f"前回状態の読み込みに失敗しました: {exc}")
        self._refresh_history()
        self._refresh_audit_table()

    def _save_state(self) -> None:
        GUI_STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
        GUI_STATE_FILE.write_text(json.dumps(self._state_dict(), ensure_ascii=False, indent=2), encoding="utf-8")

    def _on_close(self) -> None:
        try:
            self._save_state()
        finally:
            self.destroy()

    def _load_presets_only(self, config_file: Path) -> None:
        presets = main.list_presets(config_file)
        self.preset_combo.configure(values=presets)
        if presets and not self.preset.get():
            self.preset.set(presets[0])

    def _load_config_from_file(self, config_file: Path) -> None:
        config = main.read_config_document(config_file)
        base_config = {key: value for key, value in config.items() if key != "presets"}
        self._apply_config(base_config)
        self._load_presets_only(config_file)
        self.status.set("設定を読み込みました")
        self._save_state()

    def _apply_selected_preset(self) -> None:
        if not self.config_file.get().strip() or not self.preset.get().strip():
            self._show_error("設定ファイルとプリセットを選択してください。")
            return
        try:
            config = main.load_config(Path(self.config_file.get()), self.preset.get())
            self._apply_config(config)
            self.status.set("プリセットを適用しました")
            self._save_state()
        except Exception as exc:
            self._show_error(str(exc))

    def _load_config(self) -> None:
        selected = filedialog.askopenfilename(filetypes=[("JSON", "*.json"), ("All files", "*.*")])
        if not selected:
            return
        try:
            self.config_file.set(selected)
            self._load_config_from_file(Path(selected))
        except Exception as exc:
            self._show_error(str(exc))

    def _save_config(self) -> None:
        selected = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json"), ("All files", "*.*")])
        if not selected:
            return
        try:
            config = self._config_dict()
            main.validate_config_mapping(config, "GUI設定")
            Path(selected).write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
            self.status.set("設定を保存しました")
            self._save_state()
        except Exception as exc:
            self._show_error(str(exc))

    def _open_config_wizard(self) -> None:
        window = tk.Toplevel(self)
        window.title("設定ウィザード")
        window.geometry("640x560")

        input_var = tk.StringVar(value=self.input_dir.get())
        output_var = tk.StringVar(value=self.output_dir.get())
        month_var = tk.StringVar(value=self.month.get())
        pattern_var = tk.StringVar(value=self.pattern.get() or main.DEFAULT_PATTERN)
        output_name_var = tk.StringVar(value=self.output_name.get())

        fields = [
            ("入力フォルダ", input_var),
            ("出力フォルダ", output_var),
            ("対象月 YYYY-MM", month_var),
            ("読み込みファイル名パターン", pattern_var),
            ("出力ファイル名", output_name_var),
        ]
        for row, (label, variable) in enumerate(fields):
            ttk.Label(window, text=label).grid(row=row, column=0, sticky=tk.W, padx=12, pady=6)
            ttk.Entry(window, textvariable=variable).grid(row=row, column=1, sticky=tk.EW, padx=12, pady=6)

        ttk.Label(window, text="列名マッピング JSON").grid(row=len(fields), column=0, sticky=tk.NW, padx=12, pady=6)
        alias_text = tk.Text(window, height=10)
        alias_text.grid(row=len(fields), column=1, sticky=tk.NSEW, padx=12, pady=6)
        alias_text.insert(tk.END, self.column_aliases_json.get())

        def save_config() -> None:
            try:
                aliases = json.loads(alias_text.get("1.0", tk.END).strip() or "{}")
                if not isinstance(aliases, dict):
                    raise ValueError("列名マッピングはJSONオブジェクトで指定してください。")
                config = {
                    "input_dir": input_var.get().strip(),
                    "output_dir": output_var.get().strip(),
                    "month": self._optional(month_var.get()),
                    "pattern": pattern_var.get().strip() or main.DEFAULT_PATTERN,
                    "output_name": self._optional(output_name_var.get()),
                    "group_by": self.group_by.get(),
                    "all_summaries": self.all_summaries.get(),
                    "monthly_trend": self.monthly_trend.get(),
                    "charts": self.charts.get(),
                    "column_aliases": aliases,
                    "warning_amount_threshold": self._warning_amount_threshold(),
                    "style": self._style_config(),
                }
                main.validate_config_mapping(config, "設定ウィザード")
                selected = filedialog.asksaveasfilename(
                    defaultextension=".json",
                    initialfile="monthly_report_config.json",
                    filetypes=[("JSON", "*.json"), ("All files", "*.*")],
                )
                if not selected:
                    return
                Path(selected).write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
                self.config_file.set(selected)
                self._apply_config(config)
                self._load_presets_only(Path(selected))
                self._save_state()
                window.destroy()
            except Exception as exc:
                self._show_error(str(exc))

        buttons = ttk.Frame(window)
        buttons.grid(row=len(fields) + 1, column=0, columnspan=2, sticky=tk.EW, padx=12, pady=12)
        ttk.Button(buttons, text="保存して適用", command=save_config).pack(side=tk.RIGHT)
        ttk.Button(buttons, text="キャンセル", command=window.destroy).pack(side=tk.RIGHT, padx=(0, 8))
        window.columnconfigure(1, weight=1)
        window.rowconfigure(len(fields), weight=1)

    def _optional(self, value: str) -> str | None:
        value = value.strip()
        return value or None

    def _optional_int(self, value: str) -> int | None:
        value = value.strip()
        if not value:
            return None
        parsed = int(value)
        if parsed < 1:
            raise ValueError("監査履歴の保持設定は1以上で指定してください。")
        return parsed

    def _column_aliases(self) -> dict[str, str]:
        text = self.column_aliases_json.get().strip() or "{}"
        value = json.loads(text)
        if not isinstance(value, dict):
            raise ValueError("列名マッピングはJSONオブジェクトで指定してください。")
        return {str(source).strip(): str(target).strip() for source, target in value.items() if str(source).strip()}

    def _warning_amount_threshold(self) -> float:
        try:
            value = float(self.warning_amount_threshold.get())
        except ValueError as exc:
            raise ValueError("警告金額しきい値は数値で指定してください。") from exc
        if value < 0:
            raise ValueError("警告金額しきい値は0以上で指定してください。")
        return value

    def _style_config(self) -> dict:
        try:
            title_size = int(float(self.style_title_size.get()))
            chart_height = float(self.style_chart_height.get())
            chart_width = float(self.style_chart_width.get())
        except ValueError as exc:
            raise ValueError("Excel見た目設定の数値が不正です。") from exc
        style = {
            "header_fill": self.style_header_fill.get().strip() or "D9EAF7",
            "total_fill": self.style_total_fill.get().strip() or "EAF4E2",
            "title_size": title_size,
            "chart_height": chart_height,
            "chart_width": chart_width,
        }
        main.validate_style_config(style, "GUI設定")
        return style

    def _edit_column_aliases(self) -> None:
        window = tk.Toplevel(self)
        window.title("列名マッピング")
        window.geometry("620x420")
        text = tk.Text(window, height=16)
        text.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)
        text.insert(tk.END, self.column_aliases_json.get())

        def save() -> None:
            try:
                value = json.loads(text.get("1.0", tk.END))
                if not isinstance(value, dict):
                    raise ValueError("JSONオブジェクトで指定してください。")
                self.column_aliases_json.set(json.dumps(value, ensure_ascii=False, indent=2))
                self._save_state()
                window.destroy()
            except Exception as exc:
                self._show_error(str(exc))

        buttons = ttk.Frame(window)
        buttons.pack(fill=tk.X, padx=12, pady=(0, 12))
        ttk.Button(buttons, text="プリセット", command=self._edit_column_alias_presets).pack(side=tk.LEFT)
        ttk.Button(buttons, text="保存", command=save).pack(side=tk.RIGHT)
        ttk.Button(buttons, text="キャンセル", command=window.destroy).pack(side=tk.RIGHT, padx=(0, 8))

    def _edit_column_alias_presets(self) -> None:
        window = tk.Toplevel(self)
        window.title("列名マッピングプリセット")
        window.geometry("520x360")
        presets = main.read_column_alias_presets()

        listbox = tk.Listbox(window, height=10)
        listbox.pack(fill=tk.BOTH, expand=True, padx=12, pady=(12, 8))
        for name in sorted(presets):
            listbox.insert(tk.END, name)

        name_var = tk.StringVar()
        row = ttk.Frame(window)
        row.pack(fill=tk.X, padx=12, pady=(0, 8))
        ttk.Label(row, text="名前").pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=name_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))

        def selected_name() -> str | None:
            selection = listbox.curselection()
            if not selection:
                return None
            return str(listbox.get(selection[0]))

        def refresh() -> None:
            listbox.delete(0, tk.END)
            for preset_name in sorted(presets):
                listbox.insert(tk.END, preset_name)

        def apply_selected() -> None:
            preset_name = selected_name()
            if not preset_name:
                self._show_error("プリセットを選択してください。")
                return
            self.column_aliases_json.set(json.dumps(presets[preset_name], ensure_ascii=False, indent=2))
            self._save_state()
            window.destroy()

        def save_current() -> None:
            preset_name = name_var.get().strip()
            if not preset_name:
                self._show_error("プリセット名を入力してください。")
                return
            presets[preset_name] = self._column_aliases()
            main.write_column_alias_presets(presets)
            refresh()

        def delete_selected() -> None:
            preset_name = selected_name()
            if not preset_name:
                self._show_error("プリセットを選択してください。")
                return
            presets.pop(preset_name, None)
            main.write_column_alias_presets(presets)
            refresh()

        buttons = ttk.Frame(window)
        buttons.pack(fill=tk.X, padx=12, pady=(0, 12))
        ttk.Button(buttons, text="適用", command=apply_selected).pack(side=tk.LEFT)
        ttk.Button(buttons, text="現在の設定を保存", command=save_current).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(buttons, text="削除", command=delete_selected).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(buttons, text="閉じる", command=window.destroy).pack(side=tk.RIGHT)

    def _edit_style_config(self) -> None:
        window = tk.Toplevel(self)
        window.title("Excel見た目設定")
        window.geometry("440x260")
        fields = [
            ("ヘッダー色", self.style_header_fill),
            ("合計行の色", self.style_total_fill),
            ("タイトルサイズ", self.style_title_size),
            ("グラフ高さ", self.style_chart_height),
            ("グラフ幅", self.style_chart_width),
        ]
        for row, (label, variable) in enumerate(fields):
            ttk.Label(window, text=label).grid(row=row, column=0, sticky=tk.W, padx=12, pady=6)
            ttk.Entry(window, textvariable=variable).grid(row=row, column=1, sticky=tk.EW, padx=12, pady=6)

        def save() -> None:
            try:
                self._style_config()
                self._save_state()
                window.destroy()
            except Exception as exc:
                self._show_error(str(exc))

        buttons = ttk.Frame(window)
        buttons.grid(row=len(fields), column=0, columnspan=2, sticky=tk.EW, padx=12, pady=12)
        ttk.Button(buttons, text="保存", command=save).pack(side=tk.RIGHT)
        ttk.Button(buttons, text="キャンセル", command=window.destroy).pack(side=tk.RIGHT, padx=(0, 8))
        window.columnconfigure(1, weight=1)

    def _runtime_options(self) -> dict:
        return {
            "input_dir": Path(self.input_dir.get()),
            "output_dir": Path(self.output_dir.get()),
            "month": self._optional(self.month.get()),
            "start_date": self._optional(self.start_date.get()),
            "end_date": self._optional(self.end_date.get()),
            "product": self._optional(self.product.get()),
            "category": self._optional(self.category.get()),
            "group_by": self.group_by.get(),
            "pattern": self.pattern.get() or main.DEFAULT_PATTERN,
            "all_summaries": self.all_summaries.get(),
            "monthly_trend": self.monthly_trend.get(),
            "charts": self.charts.get(),
            "output_name": self._optional(self.output_name.get()),
            "summary_csv_dir": self._optional(self.summary_csv_dir.get()),
            "dry_run": self.dry_run.get(),
            "notify": self.notify.get(),
            "notify_webhook_url": self._optional(self.notify_webhook_url.get()),
            "column_aliases": self._column_aliases(),
            "warning_amount_threshold": self._warning_amount_threshold(),
            "style_config": self._style_config(),
            "audit_keep_count": self._optional_int(self.audit_keep_count.get()),
            "audit_keep_days": self._optional_int(self.audit_keep_days.get()),
        }

    def _validate_inputs(self) -> bool:
        errors: list[str] = []
        input_text = self.input_dir.get().strip()
        output_text = self.output_dir.get().strip()
        pattern = self.pattern.get().strip()
        group_by = self.group_by.get()

        if not input_text:
            errors.append("入力フォルダを指定してください。")
        else:
            input_path = Path(input_text)
            if not input_path.exists() or not input_path.is_dir():
                errors.append(f"入力フォルダが見つかりません: {input_path}")

        if not output_text:
            errors.append("出力フォルダを指定してください。")
        else:
            output_path = Path(output_text)
            if output_path.exists() and not output_path.is_dir():
                errors.append(f"出力パスがフォルダではありません: {output_path}")
            elif not output_path.exists() and not output_path.parent.exists():
                errors.append(f"出力フォルダの親フォルダが見つかりません: {output_path.parent}")

        if not pattern:
            errors.append("読み込みファイル名パターンを指定してください。")
        if group_by not in main.GROUP_BY_COLUMNS:
            errors.append("集計単位は product または category を選択してください。")

        try:
            self._column_aliases()
            self._warning_amount_threshold()
            self._style_config()
            self._optional_int(self.audit_keep_count.get())
            self._optional_int(self.audit_keep_days.get())
            main.validate_options(
                self._optional(self.month.get()),
                self._optional(self.start_date.get()),
                self._optional(self.end_date.get()),
                group_by,
            )
        except Exception as exc:
            errors.append(str(exc))

        if errors:
            self._show_error("\n".join(errors))
            return False
        self._save_state()
        return True

    def _run_async(self) -> None:
        if not self._validate_inputs():
            return
        self.status.set("Excelレポートを作成中です...")
        self.output.delete("1.0", tk.END)
        threading.Thread(target=self._run_report, daemon=True).start()

    def _preview(self) -> None:
        if not self._validate_inputs():
            return
        try:
            preview = main.build_preview(
                input_dir=Path(self.input_dir.get()),
                month=self._optional(self.month.get()),
                start_date=self._optional(self.start_date.get()),
                end_date=self._optional(self.end_date.get()),
                product=self._optional(self.product.get()),
                category=self._optional(self.category.get()),
                group_by=self.group_by.get(),
                pattern=self.pattern.get() or main.DEFAULT_PATTERN,
                limit=50,
                column_aliases=self._column_aliases(),
            )
            self._show_table("データプレビュー", preview)
        except Exception as exc:
            self._show_error(str(exc))

    def _summary_preview(self) -> None:
        if not self._validate_inputs():
            return
        try:
            preview = main.build_summary_preview(
                input_dir=Path(self.input_dir.get()),
                month=self._optional(self.month.get()),
                start_date=self._optional(self.start_date.get()),
                end_date=self._optional(self.end_date.get()),
                product=self._optional(self.product.get()),
                category=self._optional(self.category.get()),
                group_by=self.group_by.get(),
                pattern=self.pattern.get() or main.DEFAULT_PATTERN,
                all_summaries=self.all_summaries.get(),
                limit=50,
                column_aliases=self._column_aliases(),
            )
            self._show_summary_preview(preview)
        except Exception as exc:
            self._show_error(str(exc))

    def _review_current_settings(self) -> None:
        if not self._validate_inputs():
            return
        try:
            preview = main.build_summary_preview(
                input_dir=Path(self.input_dir.get()),
                month=self._optional(self.month.get()),
                start_date=self._optional(self.start_date.get()),
                end_date=self._optional(self.end_date.get()),
                product=self._optional(self.product.get()),
                category=self._optional(self.category.get()),
                group_by=self.group_by.get(),
                pattern=self.pattern.get() or main.DEFAULT_PATTERN,
                all_summaries=self.all_summaries.get(),
                limit=5,
                column_aliases=self._column_aliases(),
            )
            lines = [
                f"入力フォルダ: {self.input_dir.get()}",
                f"出力フォルダ: {self.output_dir.get()}",
                f"対象件数: {preview.detail_count}",
                f"作成される集計: {', '.join(preview.summaries)}",
                f"Excel出力: {'なし' if self.dry_run.get() else 'あり'}",
                f"集計CSV: {'あり' if self._optional(self.summary_csv_dir.get()) else 'なし'}",
                f"警告金額しきい値: {self._warning_amount_threshold():,.0f}",
                f"監査保持件数: {self._optional(self.audit_keep_count.get()) or '未設定'}",
                f"監査保持日数: {self._optional(self.audit_keep_days.get()) or '未設定'}",
            ]
            if preview.detail_count == 0:
                lines.append("警告: 出力対象の明細が0件です。")
            window = tk.Toplevel(self)
            window.title("設定内容チェック")
            window.geometry("720x420")
            text = tk.Text(window, height=16)
            text.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)
            text.insert(tk.END, "\n".join(lines))
            text.configure(state=tk.DISABLED)
        except Exception as exc:
            self._show_error(str(exc))

    def _run_report(self) -> None:
        try:
            result = main.run_report(
                input_dir=Path(self.input_dir.get()),
                output_dir=Path(self.output_dir.get()),
                month=self._optional(self.month.get()),
                start_date=self._optional(self.start_date.get()),
                end_date=self._optional(self.end_date.get()),
                product=self._optional(self.product.get()),
                category=self._optional(self.category.get()),
                group_by=self.group_by.get(),
                keep_reports=None,
                pattern=self.pattern.get() or main.DEFAULT_PATTERN,
                dry_run=self.dry_run.get(),
                all_summaries=self.all_summaries.get(),
                monthly_trend=self.monthly_trend.get(),
                charts=self.charts.get(),
                output_name=self._optional(self.output_name.get()),
                summary_csv_dir=Path(self.summary_csv_dir.get()) if self._optional(self.summary_csv_dir.get()) else None,
                column_aliases=self._column_aliases(),
                style_config=self._style_config(),
                warning_amount_threshold=self._warning_amount_threshold(),
            )
            main.append_audit_log(
                status="success",
                options=self._runtime_options(),
                result=result,
                keep_count=self._optional_int(self.audit_keep_count.get()),
                keep_days=self._optional_int(self.audit_keep_days.get()),
            )
            previous_rate = (
                f"{result.previous_month_change_rate:.1%}"
                if isinstance(result.previous_month_change_rate, (int, float))
                else str(result.previous_month_change_rate)
            )
            lines = [
                "レポート作成が完了しました。",
                "",
                "出力ファイル:",
                str(result.output_file) if result.output_file else "検証のみのためExcelファイルは作成していません。",
                "",
                "集計結果:",
                f"売上合計: {result.total_amount:,.0f}円",
                f"明細件数: {result.detail_count:,}件",
                f"数量合計: {result.total_quantity:,.0f}",
                f"平均単価: {result.average_unit_price:,.0f}円",
                f"対象日数: {result.target_days}日",
                f"商品数: {result.product_count}",
                f"カテゴリ数: {result.category_count}",
                f"前月売上: {result.previous_month_amount:,.0f}円",
                f"前月比: {previous_rate}",
                "",
                "確認事項:",
                f"未分類データ: {result.uncategorized_count:,}件",
                f"エラー行: {result.error_count:,}件",
                f"確認が必要なデータ: {result.uncategorized_count + result.error_count:,}件",
            ]
            for summary_csv_file in result.summary_csv_files:
                lines.append(f"集計CSV: {summary_csv_file}")
            for warning in result.warnings:
                lines.append(f"警告: {warning}")
            self.after(0, self._show_success, "\n".join(lines))
        except main.DataValidationError as exc:
            output_dir = Path(self.output_dir.get())
            report_file = main.default_error_report_path(output_dir)
            main.write_validation_error_report(exc, report_file)
            main.append_audit_log(
                status="validation_error",
                options=self._runtime_options(),
                error=str(exc),
                keep_count=self._optional_int(self.audit_keep_count.get()),
                keep_days=self._optional_int(self.audit_keep_days.get()),
            )
            if self.notify.get():
                main.notify_completion("レポート作成に失敗しました。", self._optional(self.notify_webhook_url.get()), "validation_error")
            self.after(0, self._show_validation_error, exc, report_file)
        except Exception as exc:
            main.append_audit_log(
                status="error",
                options=self._runtime_options(),
                error=str(exc),
                keep_count=self._optional_int(self.audit_keep_count.get()),
                keep_days=self._optional_int(self.audit_keep_days.get()),
            )
            if self.notify.get():
                main.notify_completion("レポート作成に失敗しました。", self._optional(self.notify_webhook_url.get()), "error")
            self.after(0, self._show_error, main.format_user_error_message(exc) + "\n\n詳細はログタブを確認してください。")

    def _show_table(self, title: str, preview: main.PreviewResult) -> None:
        window = tk.Toplevel(self)
        window.title(f"{title} ({len(preview.rows)} / {preview.total_count})")
        window.geometry("980x460")
        self._add_tree(window, preview)

    def _show_summary_preview(self, preview: main.SummaryPreviewResult) -> None:
        window = tk.Toplevel(self)
        window.title(f"集計プレビュー (明細 {preview.detail_count}件)")
        window.geometry("980x460")
        notebook = ttk.Notebook(window)
        notebook.pack(fill=tk.BOTH, expand=True)
        for name, table in preview.summaries.items():
            frame = ttk.Frame(notebook)
            notebook.add(frame, text=f"{name} ({table.total_count})")
            self._add_tree(frame, table)

    def _add_tree(self, parent: tk.Widget, preview: main.PreviewResult) -> None:
        tree = ttk.Treeview(parent, columns=preview.columns, show="headings")
        y_scroll = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=tree.yview)
        x_scroll = ttk.Scrollbar(parent, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        tree.grid(row=0, column=0, sticky=tk.NSEW)
        y_scroll.grid(row=0, column=1, sticky=tk.NS)
        x_scroll.grid(row=1, column=0, sticky=tk.EW)
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        for column in preview.columns:
            tree.heading(column, text=column)
            tree.column(column, width=120, minwidth=80, stretch=True)
        for row in preview.rows:
            tree.insert("", tk.END, values=row)

    def _show_success(self, message: str) -> None:
        self.status.set("レポート作成が完了しました")
        self.output.insert(tk.END, message)
        self._refresh_history()
        self._refresh_audit_table()
        self._refresh_audit_log_text()
        if self.notify.get():
            main.notify_completion("レポート作成が完了しました。", self._optional(self.notify_webhook_url.get()), "success")
        messagebox.showinfo("完了", message)

    def _show_error(self, message: str) -> None:
        self.status.set("エラー")
        if hasattr(self, "output"):
            self.output.insert(tk.END, message)
        messagebox.showerror("エラー", message)

    def _show_validation_error(self, error: main.DataValidationError, report_file: Path) -> None:
        message = "\n".join(
            [
                main.format_user_error_message(error),
                "",
                f"エラー一覧CSV: {report_file}",
                "",
                "詳細はログタブも確認してください。",
            ]
        )
        self._show_error(message)
        window = tk.Toplevel(self)
        window.title("エラー一覧と修正方法")
        window.geometry("1060x460")
        columns = ("message", "fix", "source_file", "source_row")
        tree = ttk.Treeview(window, columns=columns, show="headings")
        tree.pack(fill=tk.BOTH, expand=True)
        headings = {
            "message": "エラー内容",
            "fix": "修正方法",
            "source_file": "元ファイル",
            "source_row": "行番号",
        }
        for column in columns:
            tree.heading(column, text=headings[column])
            tree.column(column, width=320 if column in {"message", "fix"} else 120, stretch=True)
        for issue in error.issues:
            tree.insert("", tk.END, values=(issue.message, issue.fix, issue.source_file, issue.source_row))
        self._refresh_audit_table()
        self._refresh_audit_log_text()

    def _refresh_history(self) -> None:
        if not hasattr(self, "history_list"):
            return
        self.history_list.delete(0, tk.END)
        output_dir = Path(self.output_dir.get().strip() or main.OUTPUT_DIR)
        if not output_dir.exists() or not output_dir.is_dir():
            self.report_history = []
            self.history_list.insert(tk.END, "まだ作成済みレポートはありません。実行タブでExcelレポートを作成すると、ここに履歴が表示されます。")
            return
        self.report_history = sorted(output_dir.glob("*.xlsx"), key=lambda path: path.stat().st_mtime, reverse=True)
        if not self.report_history:
            self.history_list.insert(tk.END, "まだ作成済みレポートはありません。実行タブでExcelレポートを作成すると、ここに履歴が表示されます。")
            return
        for report_file in self.report_history:
            self.history_list.insert(tk.END, report_file.name)

    def _selected_report(self) -> Path | None:
        if not self.report_history:
            self._show_error("まだ作成済みレポートはありません。実行タブでExcelレポートを作成してください。")
            return None
        selection = self.history_list.curselection()
        if not selection:
            self._show_error("レポートを選択してください。")
            return None
        selected_index = selection[0]
        if selected_index >= len(self.report_history):
            self._show_error("レポートを選択してください。")
            return None
        return self.report_history[selected_index]

    def _open_selected_report(self) -> None:
        report_file = self._selected_report()
        if report_file is None:
            return
        try:
            subprocess.Popen(["cmd", "/c", "start", "", str(report_file)])
        except Exception as exc:
            self._show_error(str(exc))

    def _show_selected_report_detail(self) -> None:
        report_file = self._selected_report()
        if report_file is None:
            return
        stat = report_file.stat()
        window = tk.Toplevel(self)
        window.title("レポート詳細")
        window.geometry("640x220")
        text = tk.Text(window, height=8)
        text.pack(fill=tk.BOTH, expand=True)
        text.insert(
            tk.END,
            "\n".join(
                [
                    f"ファイル名: {report_file.name}",
                    f"パス: {report_file}",
                    f"サイズ: {stat.st_size:,} bytes",
                    f"更新日時: {report_file.stat().st_mtime}",
                ]
            ),
        )
        text.configure(state=tk.DISABLED)
        buttons = ttk.Frame(window)
        buttons.pack(fill=tk.X, pady=8)
        ttk.Button(buttons, text="開く", command=lambda: subprocess.Popen(["cmd", "/c", "start", "", str(report_file)])).pack(side=tk.LEFT, padx=8)
        ttk.Button(buttons, text="フォルダを開く", command=lambda: subprocess.Popen(["explorer", "/select,", str(report_file)])).pack(side=tk.LEFT)

    def _register_schedule(self) -> None:
        try:
            day = int(self.schedule_day.get())
            if day < 1 or day > 31:
                raise ValueError("実行日は1から31で指定してください。")
            script = main.BASE_DIR / "scripts" / "register_monthly_task.ps1"
            result = subprocess.run(
                [
                    "powershell.exe",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(script),
                    "-DayOfMonth",
                    str(day),
                    "-Time",
                    self.schedule_time.get(),
                ],
                cwd=main.BASE_DIR,
                capture_output=True,
                text=True,
                check=False,
            )
            if result.returncode != 0:
                raise RuntimeError(result.stderr or result.stdout)
            self.status.set("タスクを登録しました")
            messagebox.showinfo("完了", "タスクを登録しました。")
        except Exception as exc:
            self._show_error(str(exc))

    def _write_template(self) -> None:
        selected = filedialog.asksaveasfilename(
            defaultextension=".csv",
            initialfile="sales_template.csv",
            filetypes=[("CSV", "*.csv"), ("All files", "*.*")],
        )
        if not selected:
            return
        try:
            template_file = main.write_input_template(Path(selected))
            self.status.set("CSVテンプレートを作成しました")
            messagebox.showinfo("完了", f"CSVテンプレートを作成しました。\n{template_file}")
        except Exception as exc:
            self._show_error(str(exc))

    def _repair_legacy_text_files(self) -> None:
        try:
            repaired = main.repair_legacy_text_files(gui_state_file=GUI_STATE_FILE)
            changed_paths = [path for path, changed in repaired.items() if changed]
            if changed_paths:
                self.status.set("旧文字化け文言を修復しました")
                self._load_state()
                self._refresh_audit_log_text()
                self._refresh_log_text()
                messagebox.showinfo("完了", "旧文字化け文言を修復しました。\n" + "\n".join(changed_paths))
                return
            self.status.set("修復対象の旧文字化け文言はありませんでした")
            messagebox.showinfo("確認", "修復対象の旧文字化け文言はありませんでした。")
        except Exception as exc:
            self._show_error(str(exc))

    def _preview_legacy_text_repairs(self) -> None:
        try:
            changes_by_file = main.inspect_legacy_text_files(
                gui_state_file=GUI_STATE_FILE,
                audit_log_file=main.AUDIT_LOG_FILE,
                alias_preset_file=main.ALIAS_PRESET_FILE,
            )
            preview = main.format_legacy_text_preview(changes_by_file)
        except Exception as exc:
            self._show_error(str(exc))
            return

        window = tk.Toplevel(self)
        window.title("文字化け修復プレビュー")
        window.geometry("860x520")
        text = tk.Text(window, height=24)
        text.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)
        text.insert(tk.END, preview or "修復候補の旧文字化け文言はありませんでした。")
        text.configure(state=tk.DISABLED)

    def _refresh_log_text(self) -> None:
        if not hasattr(self, "log_text"):
            return
        if main.LOG_FILE.exists():
            content = main.normalize_legacy_text(main.LOG_FILE.read_text(encoding="utf-8", errors="replace"))
        else:
            content = "ログはまだありません。"
        self.log_text.delete("1.0", tk.END)
        self.log_text.insert(tk.END, content)

    def _refresh_audit_log_text(self) -> None:
        if not hasattr(self, "log_text"):
            return
        if main.AUDIT_LOG_FILE.exists():
            content = main.normalize_legacy_text(main.AUDIT_LOG_FILE.read_text(encoding="utf-8", errors="replace"))
        else:
            content = "監査ログはまだありません。"
        self.log_text.delete("1.0", tk.END)
        self.log_text.insert(tk.END, content)

    def _refresh_audit_table(self) -> None:
        if not hasattr(self, "audit_tree"):
            return
        self.audit_records = {}
        self.audit_tree.delete(*self.audit_tree.get_children())
        try:
            records = main.read_audit_log(limit=200)
        except Exception as exc:
            self.audit_tree.insert("", tk.END, values=("", "error", "", "", "", "", str(exc)))
            return
        try:
            filtered_records = main.filter_audit_records(
                records,
                filter_text=self.audit_filter_text.get(),
                filter_status=self.audit_filter_status.get(),
                date_from=self.audit_filter_date_from.get(),
                date_to=self.audit_filter_date_to.get(),
            )
        except Exception as exc:
            self.audit_tree.insert("", tk.END, values=("", "error", "", "", "", "", str(exc)))
            return
        for record in reversed(filtered_records):
            period = record.get("month") or f"{record.get('start_date') or ''} - {record.get('end_date') or ''}".strip()
            item_id = self.audit_tree.insert(
                "",
                tk.END,
                values=(
                    record.get("timestamp", ""),
                    record.get("status", ""),
                    period,
                    record.get("detail_count") or "",
                    record.get("summary_count") or "",
                    len(record.get("warnings") or []),
                    record.get("output_file") or "",
                    record.get("error") or "",
                ),
            )
            self.audit_records[item_id] = record
        if not filtered_records:
            self.audit_tree.insert(
                "",
                tk.END,
                values=("", "履歴なし", "", "", "", "", "実行タブでExcelレポートを作成すると、ここに履歴が表示されます。", ""),
            )

    def _filtered_audit_records(self) -> list[dict]:
        return [self.audit_records[item_id] for item_id in self.audit_tree.get_children() if item_id in self.audit_records]

    def _export_audit_csv(self) -> None:
        records = self._filtered_audit_records()
        if not records:
            self._show_error("出力対象の監査履歴がありません。")
            return
        selected = filedialog.asksaveasfilename(
            defaultextension=".csv",
            initialdir=str(main.LOG_DIR),
            initialfile=f"audit_history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            filetypes=[("CSV", "*.csv"), ("All files", "*.*")],
        )
        if not selected:
            return
        try:
            export_file = main.export_audit_log_csv(records, Path(selected))
            self.status.set("監査履歴CSVを出力しました")
            messagebox.showinfo("完了", f"監査履歴CSVを出力しました。\n{export_file}")
        except Exception as exc:
            self._show_error(str(exc))

    def _export_audit_summary_csv(self) -> None:
        records = self._filtered_audit_records()
        if not records:
            self._show_error("出力対象の監査履歴がありません。")
            return
        selected = filedialog.asksaveasfilename(
            defaultextension=".csv",
            initialdir=str(main.LOG_DIR),
            initialfile=f"audit_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            filetypes=[("CSV", "*.csv"), ("All files", "*.*")],
        )
        if not selected:
            return
        try:
            export_file = main.export_audit_summary_csv(records, Path(selected))
            self.status.set("監査履歴要約CSVを出力しました")
            messagebox.showinfo("完了", f"監査履歴要約CSVを出力しました。\n{export_file}")
        except Exception as exc:
            self._show_error(str(exc))

    def _export_audit_monthly_summary_csv(self) -> None:
        records = self._filtered_audit_records()
        if not records:
            self._show_error("出力対象の監査履歴がありません。")
            return
        selected = filedialog.asksaveasfilename(
            defaultextension=".csv",
            initialdir=str(main.LOG_DIR),
            initialfile=f"audit_monthly_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            filetypes=[("CSV", "*.csv"), ("All files", "*.*")],
        )
        if not selected:
            return
        try:
            export_file = main.export_audit_monthly_summary_csv(records, Path(selected))
            self.status.set("監査履歴月別要約CSVを出力しました")
            messagebox.showinfo("完了", f"監査履歴月別要約CSVを出力しました。\n{export_file}")
        except Exception as exc:
            self._show_error(str(exc))

    def _backup_audit_log(self) -> None:
        try:
            backup_file = main.backup_audit_log()
            if backup_file is None:
                self.status.set("バックアップ対象の監査履歴はありませんでした")
                messagebox.showinfo("確認", "バックアップ対象の監査履歴はありませんでした。")
                return
            self.status.set("監査履歴をバックアップしました")
            messagebox.showinfo("完了", f"監査履歴をバックアップしました。\n{backup_file}")
        except Exception as exc:
            self._show_error(str(exc))

    def _show_audit_anomalies(self) -> None:
        records = self._filtered_audit_records()
        if not records:
            self._show_error("確認対象の監査履歴がありません。")
            return
        anomaly = main.detect_audit_anomalies(records)
        lines = [
            f"実行件数: {anomaly.total_runs}",
            f"成功件数: {anomaly.success_count}",
            f"失敗件数: {anomaly.failure_count}",
            f"警告件数: {anomaly.warning_total}",
            f"最終実行日: {anomaly.last_run_date or 'なし'}",
            f"未実行日数: {anomaly.consecutive_missing_days}",
            f"失敗率: {anomaly.failure_rate:.1%}",
            f"アラート: {' / '.join(anomaly.alerts) if anomaly.alerts else 'なし'}",
        ]
        window = tk.Toplevel(self)
        window.title("監査履歴異常チェック")
        window.geometry("560x320")
        text = tk.Text(window, height=12)
        text.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)
        text.insert(tk.END, "\n".join(lines))
        text.configure(state=tk.DISABLED)

    def _selected_audit_record(self) -> dict | None:
        selection = self.audit_tree.selection()
        if not selection:
            self._show_error("監査履歴を選択してください。")
            return None
        return self.audit_records.get(selection[0])

    def _restore_selected_audit_record(self) -> None:
        record = self._selected_audit_record()
        if not record:
            return
        config = {
            "input_dir": record.get("input_dir") or str(main.INPUT_DIR),
            "output_dir": record.get("output_dir") or str(main.OUTPUT_DIR),
            "month": record.get("month"),
            "start_date": record.get("start_date"),
            "end_date": record.get("end_date"),
            "product": record.get("product"),
            "category": record.get("category"),
            "pattern": record.get("pattern") or main.DEFAULT_PATTERN,
            "output_name": record.get("output_name"),
            "group_by": record.get("group_by") or "product",
            "all_summaries": record.get("all_summaries", True),
            "monthly_trend": record.get("monthly_trend", True),
            "charts": record.get("charts", True),
            "dry_run": record.get("dry_run", False),
            "notify": record.get("notify", False),
            "notify_webhook_url": record.get("notify_webhook_url"),
            "summary_csv_dir": record.get("summary_csv_dir"),
            "column_aliases": record.get("column_aliases") or {},
            "warning_amount_threshold": record.get("warning_amount_threshold", 1_000_000),
            "style": record.get("style") or {},
        }
        self._apply_config(config)
        self._save_state()
        self.status.set("監査履歴から実行条件を復元しました")


def main_gui() -> None:
    main.setup_logging()
    app = MonthlyReportApp()
    app.mainloop()


if __name__ == "__main__":
    main_gui()

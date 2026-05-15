import json
import os
import subprocess
import sys
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
APP_BACKGROUND = "#F7F9FB"
APP_HEADING_COLOR = "#1F4E79"
APP_BORDER_COLOR = "#D9E2EC"
APP_SUCCESS_COLOR = "#2E7D32"
APP_ERROR_COLOR = "#C0392B"
PRIMARY_BUTTON_WIDTH = 20
SECONDARY_BUTTON_WIDTH = 18
HELP_WRAP_LENGTH = 420
PRESET_COMBO_WIDTH = 45
WINDOW_SCREEN_RATIO = 0.9
WINDOW_MAX_SIZE = (1400, 900)
WINDOW_MIN_SIZE = (1000, 650)
RUN_TAB_SCROLLBAR_WIDTH = 16
BUTTON_GRID_COLUMNS = 4
BUTTON_GRID_MIN_WIDTH = 180
PREFERRED_TTK_THEMES = ("vista", "xpnative", "default")
CHECKBOX_LABELS = {
    "all_summaries": "商品別・カテゴリ別集計を出力",
    "monthly_trend": "月別推移を出力",
    "charts": "グラフを出力",
    "dry_run": "検証のみ実行",
    "notify": "完了時に通知音を鳴らす",
}
RUN_TAB_FIELD_HELP_TEXTS = {
    "config_file": "保存済み設定を使う場合に指定します。",
    "input_dir": "CSVまたはExcelファイルを置くフォルダを指定します。\nファイル選択もできます。",
    "output_dir": "作成したExcelレポートの保存先です。",
    "month": "例: 2026-04\n月単位で集計する場合に指定します。",
    "start_date": "例: 2026-04-01\n任意期間で集計する場合に指定します。",
    "end_date": "例: 2026-04-30\n開始日とセットで使います。",
    "product": "空欄の場合はすべての商品を対象にします。",
    "category": "空欄の場合はすべてのカテゴリを対象にします。",
    "pattern": "例: sales_*.csv\nCSV・Excelの名前規則を指定します。",
    "output_name": "空欄の場合は日時付きのファイル名で自動作成します。",
    "summary_csv_dir": "集計結果CSVも出力したい場合だけ指定します。",
    "notify_webhook_url": "通知連携を使う場合だけ指定します。",
}
RUN_ACTION_BUTTONS = (
    ("Excelレポートを作成", "_run_async"),
    ("データをプレビュー", "_preview"),
    ("集計プレビュー", "_summary_preview"),
    ("プリセット適用", "_apply_selected_preset"),
    ("設定読込", "_load_config"),
    ("設定保存", "_save_config"),
    ("CSVテンプレート作成", "_write_template"),
    ("タスク登録", "_register_schedule"),
    ("設定ウィザード", "_open_config_wizard"),
    ("設定内容チェック", "_review_current_settings"),
    ("列名設定を開く", "_edit_column_aliases"),
    ("文字化け修復", "_repair_legacy_text_files"),
    ("列名候補を更新", "_refresh_column_aliases_from_input"),
    ("Excel見た目設定", "_edit_style_config"),
    ("Excelレポートを開く", "_open_latest_report"),
    ("出力フォルダを開く", "_open_output_folder"),
    ("エラーを確認", "_review_input_errors"),
)
ERROR_DIALOG_SIZE = "1200x700"
ERROR_DIALOG_MIN_SIZE = (900, 500)
ERROR_TABLE_DIALOG_SIZE = "1200x700"
ERROR_TABLE_DIALOG_MIN_SIZE = (900, 500)
TAB_LABELS = {
    "run": "実行",
    "error": "エラー確認",
    "history": "作成済みレポート",
    "audit": "実行履歴",
    "log": "詳細ログ",
}
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
MAPPING_STANDARD_LABELS = {
    "date": "日付",
    "product": "商品名",
    "category": "カテゴリ",
    "quantity": "数量",
    "unit_price": "単価",
    "amount": "金額",
}
MAPPING_REQUIRED_KEYS = ("date", "product", "quantity", "unit_price")
NO_REPORT_TO_OPEN_MESSAGE = "まだ開けるレポートがありません。\n先に「Excelレポートを作成」を実行してください。"
REPORT_FILE_MISSING_MESSAGE = "レポートファイルが見つかりません。\n削除または移動された可能性があります。"
OUTPUT_FOLDER_MISSING_MESSAGE = "出力フォルダが見つかりません。\n出力フォルダの指定を確認してください。"
REPORT_SELECTION_REQUIRED_MESSAGE = "開くレポートを一覧から選択してください。"
NO_VALIDATION_ERRORS_MESSAGE = "検証エラーはありません。\n現在の入力データは正常に読み込めます。"
NO_ERROR_CSV_MESSAGE = "まだエラーCSVは作成されていません。\n先に入力データを検証するか、レポートを作成してください。"
ERROR_REVIEW_COLUMNS = (
    ("source_file", "元ファイル名", 160),
    ("source_row", "行番号", 80),
    ("message", "エラー内容", 280),
    ("fix", "修正方法", 320),
    ("date", "日付", 120),
    ("product", "商品名", 140),
    ("category", "カテゴリ", 120),
    ("quantity", "数量", 90),
    ("unit_price", "単価", 90),
    ("amount", "金額", 100),
)


class MonthlyReportApp(BaseWindow):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self._apply_initial_window_size()
        self.resizable(True, True)
        self.report_history: list[Path] = []
        self.latest_report_path: Path | None = None
        self.latest_error_csv_path: Path | None = None
        self.audit_records: dict[str, dict] = {}
        self._init_variables()
        self._configure_style()
        self._build_widgets()
        self._load_state()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    @staticmethod
    def _calculate_window_size(screen_width: int, screen_height: int) -> tuple[int, int, int, int]:
        width = min(WINDOW_MAX_SIZE[0], int(screen_width * WINDOW_SCREEN_RATIO))
        height = min(WINDOW_MAX_SIZE[1], int(screen_height * WINDOW_SCREEN_RATIO))
        min_width = min(WINDOW_MIN_SIZE[0], width)
        min_height = min(WINDOW_MIN_SIZE[1], height)
        return width, height, min_width, min_height

    def _apply_initial_window_size(self) -> None:
        width, height, min_width, min_height = self._calculate_window_size(
            self.winfo_screenwidth(),
            self.winfo_screenheight(),
        )
        self.geometry(f"{width}x{height}")
        self.minsize(min_width, min_height)

    def _configure_style(self) -> None:
        self.configure(bg=APP_BACKGROUND)
        style = ttk.Style(self)
        self._apply_preferred_theme(style)
        style.configure(".", font=("Meiryo", 9))
        style.configure("TFrame", background=APP_BACKGROUND)
        style.configure("TLabel", background=APP_BACKGROUND, foreground="#1F2933")
        style.configure("Heading.TLabel", background=APP_BACKGROUND, foreground=APP_HEADING_COLOR, font=("Meiryo", 10, "bold"))
        style.configure("Help.TLabel", background=APP_BACKGROUND, foreground="#555555")
        style.configure("Success.TLabel", background=APP_BACKGROUND, foreground=APP_SUCCESS_COLOR)
        style.configure("Error.TLabel", background=APP_BACKGROUND, foreground=APP_ERROR_COLOR)
        style.configure("TNotebook", background=APP_BACKGROUND, borderwidth=0)
        style.configure("TNotebook.Tab", padding=(20, 10), font=("Meiryo", 10))
        style.map(
            "TNotebook.Tab",
            background=[("selected", "#EAF2FB"), ("active", "#F0F5FA")],
            foreground=[("selected", APP_HEADING_COLOR)],
        )
        style.configure("TButton", padding=(8, 6))
        style.configure("TCheckbutton", background=APP_BACKGROUND, foreground="#1F2933", padding=(2, 2))
        style.configure("Treeview.Heading", font=("Meiryo", 9, "bold"))

    @staticmethod
    def _apply_preferred_theme(style: ttk.Style) -> str:
        available_themes = {theme.lower(): theme for theme in style.theme_names()}
        for theme_name in PREFERRED_TTK_THEMES:
            matched = available_themes.get(theme_name)
            if matched:
                style.theme_use(matched)
                return matched
        current_theme = style.theme_use()
        if current_theme:
            return current_theme
        return style.theme_use(PREFERRED_TTK_THEMES[-1])

    def _button(self, parent: tk.Widget, text: str, command, *, width: int = PRIMARY_BUTTON_WIDTH, **kwargs) -> ttk.Button:
        return ttk.Button(parent, text=text, command=command, width=width, **kwargs)

    def _help_label(self, parent: tk.Widget, text: str) -> ttk.Label:
        return ttk.Label(
            parent,
            text=text,
            wraplength=HELP_WRAP_LENGTH,
            justify=tk.LEFT,
            anchor=tk.W,
            style="Help.TLabel",
        )

    def _create_scrollable_run_frame(self, parent: ttk.Frame) -> ttk.Frame:
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)
        canvas = tk.Canvas(parent, background=APP_BACKGROUND, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=canvas.yview)
        content = ttk.Frame(canvas, padding=(0, 0, RUN_TAB_SCROLLBAR_WIDTH, 0))
        window_id = canvas.create_window((0, 0), window=content, anchor=tk.NW)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky=tk.NSEW)
        scrollbar.grid(row=0, column=1, sticky=tk.NS)

        def update_scroll_region(_event: tk.Event) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        def update_content_width(event: tk.Event) -> None:
            canvas.itemconfigure(window_id, width=event.width)

        content.bind("<Configure>", update_scroll_region)
        canvas.bind("<Configure>", update_content_width)
        self._bind_canvas_mousewheel(canvas, content)
        return content

    def _bind_canvas_mousewheel(self, canvas: tk.Canvas, content: ttk.Frame) -> None:
        def on_mousewheel(event: tk.Event) -> str:
            if getattr(event, "num", None) == 4:
                delta = -1
            elif getattr(event, "num", None) == 5:
                delta = 1
            else:
                delta = int(-1 * (event.delta / 120))
            canvas.yview_scroll(delta, "units")
            return "break"

        def bind_scroll(_event: tk.Event) -> None:
            canvas.bind_all("<MouseWheel>", on_mousewheel)
            canvas.bind_all("<Button-4>", on_mousewheel)
            canvas.bind_all("<Button-5>", on_mousewheel)

        def unbind_scroll(_event: tk.Event) -> None:
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        content._run_tab_mousewheel_handler = on_mousewheel  # type: ignore[attr-defined]
        content.bind("<Enter>", bind_scroll)
        content.bind("<Leave>", unbind_scroll)
        canvas.bind("<Enter>", bind_scroll)
        canvas.bind("<Leave>", unbind_scroll)

    def _bind_run_tab_child_mousewheel(self, widget: tk.Widget, handler) -> None:
        for child in widget.winfo_children():
            child.bind("<MouseWheel>", handler, add="+")
            child.bind("<Button-4>", handler, add="+")
            child.bind("<Button-5>", handler, add="+")
            self._bind_run_tab_child_mousewheel(child, handler)

    def _build_run_action_buttons(self, parent: ttk.Frame) -> None:
        for column in range(BUTTON_GRID_COLUMNS):
            parent.columnconfigure(column, weight=1, minsize=BUTTON_GRID_MIN_WIDTH, uniform="run_actions")
        for index, (text, method_name) in enumerate(RUN_ACTION_BUTTONS):
            row, column = divmod(index, BUTTON_GRID_COLUMNS)
            self._button(parent, text=text, command=getattr(self, method_name)).grid(
                row=row,
                column=column,
                sticky=tk.EW,
                padx=(0 if column == 0 else 8, 0),
                pady=(0 if row == 0 else 8, 0),
            )

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
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        run_tab = ttk.Frame(self.notebook, padding=12)
        error_tab = ttk.Frame(self.notebook, padding=12)
        history_tab = ttk.Frame(self.notebook, padding=12)
        audit_tab = ttk.Frame(self.notebook, padding=12)
        log_tab = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(run_tab, text=TAB_LABELS["run"])
        self.notebook.add(error_tab, text=TAB_LABELS["error"])
        self.notebook.add(history_tab, text=TAB_LABELS["history"])
        self.notebook.add(audit_tab, text=TAB_LABELS["audit"])
        self.notebook.add(log_tab, text=TAB_LABELS["log"])

        self._build_run_tab(run_tab)
        self._build_error_tab(error_tab)
        self._build_history_tab(history_tab)
        self._build_audit_tab(audit_tab)
        self._build_log_tab(log_tab)

    def _build_run_tab(self, root: ttk.Frame) -> None:
        root = self._create_scrollable_run_frame(root)
        intro = ttk.Label(root, text=RUN_TAB_DESCRIPTION, justify=tk.LEFT, anchor=tk.W, wraplength=940, style="Heading.TLabel")
        intro.grid(row=0, column=0, columnspan=4, sticky=tk.EW, pady=(0, 8))
        steps = ttk.Label(root, text=RUN_TAB_STEPS, justify=tk.LEFT, anchor=tk.W, relief=tk.GROOVE, padding=10)
        steps.grid(row=1, column=0, columnspan=4, sticky=tk.EW, pady=(0, 12))

        fields = [
            ("設定ファイル", self.config_file, self._choose_config_file, RUN_TAB_FIELD_HELP_TEXTS["config_file"]),
            ("入力フォルダ", self.input_dir, self._choose_input_dir, RUN_TAB_FIELD_HELP_TEXTS["input_dir"]),
            ("出力フォルダ", self.output_dir, self._choose_output_dir, RUN_TAB_FIELD_HELP_TEXTS["output_dir"]),
            ("対象月", self.month, None, RUN_TAB_FIELD_HELP_TEXTS["month"]),
            ("開始日", self.start_date, None, RUN_TAB_FIELD_HELP_TEXTS["start_date"]),
            ("終了日", self.end_date, None, RUN_TAB_FIELD_HELP_TEXTS["end_date"]),
            ("商品名で絞り込み", self.product, None, RUN_TAB_FIELD_HELP_TEXTS["product"]),
            ("カテゴリで絞り込み", self.category, None, RUN_TAB_FIELD_HELP_TEXTS["category"]),
            ("読み込みファイル名パターン", self.pattern, None, RUN_TAB_FIELD_HELP_TEXTS["pattern"]),
            ("出力ファイル名", self.output_name, None, RUN_TAB_FIELD_HELP_TEXTS["output_name"]),
            ("集計CSVフォルダ", self.summary_csv_dir, self._choose_summary_csv_dir, RUN_TAB_FIELD_HELP_TEXTS["summary_csv_dir"]),
            ("通知Webhook URL", self.notify_webhook_url, None, RUN_TAB_FIELD_HELP_TEXTS["notify_webhook_url"]),
        ]
        field_start_row = 2
        for index, (label, variable, command, help_text) in enumerate(fields):
            row = field_start_row + index
            ttk.Label(root, text=label).grid(row=row, column=0, sticky=tk.W, pady=4)
            ttk.Entry(root, textvariable=variable).grid(row=row, column=1, sticky=tk.EW, pady=4)
            if command:
                self._button(root, text="選択", command=command, width=SECONDARY_BUTTON_WIDTH).grid(row=row, column=2, sticky=tk.EW, padx=(8, 0), pady=4)
            self._help_label(root, help_text).grid(row=row, column=3, sticky=tk.W, padx=(8, 0), pady=4)

        option_row = field_start_row + len(fields)
        ttk.Label(root, text="集計単位").grid(row=option_row, column=0, sticky=tk.W, pady=4)
        ttk.Combobox(root, textvariable=self.group_by, values=["product", "category"], state="readonly").grid(
            row=option_row,
            column=1,
            sticky=tk.W,
            pady=4,
        )
        ttk.Label(root, text="プリセット").grid(row=option_row, column=2, sticky=tk.W, padx=(8, 0), pady=4)
        self.preset_combo = ttk.Combobox(root, textvariable=self.preset, values=[], state="readonly", width=PRESET_COMBO_WIDTH)
        self.preset_combo.grid(row=option_row, column=3, sticky=tk.W, pady=4)

        checks = ttk.Frame(root)
        checks.grid(row=option_row + 1, column=0, columnspan=4, sticky=tk.W, pady=8)
        ttk.Checkbutton(checks, text=CHECKBOX_LABELS["all_summaries"], variable=self.all_summaries).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(checks, text=CHECKBOX_LABELS["monthly_trend"], variable=self.monthly_trend).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(checks, text=CHECKBOX_LABELS["charts"], variable=self.charts).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(checks, text=CHECKBOX_LABELS["dry_run"], variable=self.dry_run).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(checks, text=CHECKBOX_LABELS["notify"], variable=self.notify).pack(side=tk.LEFT)

        schedule_row = option_row + 2
        ttk.Label(root, text="定期実行 日/時刻").grid(row=schedule_row, column=0, sticky=tk.W, pady=(0, 8))
        schedule_frame = ttk.Frame(root)
        schedule_frame.grid(row=schedule_row, column=1, sticky=tk.W, pady=(0, 8))
        ttk.Entry(schedule_frame, textvariable=self.schedule_day, width=5).pack(side=tk.LEFT)
        ttk.Entry(schedule_frame, textvariable=self.schedule_time, width=8).pack(side=tk.LEFT, padx=(8, 0))

        audit_row = schedule_row + 1
        ttk.Label(root, text="監査保持 件数/日数").grid(row=audit_row, column=0, sticky=tk.W, pady=(0, 8))
        audit_frame = ttk.Frame(root)
        audit_frame.grid(row=audit_row, column=1, sticky=tk.W, pady=(0, 8))
        ttk.Entry(audit_frame, textvariable=self.audit_keep_count, width=8).pack(side=tk.LEFT)
        ttk.Entry(audit_frame, textvariable=self.audit_keep_days, width=8).pack(side=tk.LEFT, padx=(8, 0))

        action_row = audit_row + 1
        ttk.Label(root, text="操作", style="Heading.TLabel").grid(row=action_row, column=0, columnspan=4, sticky=tk.W, pady=(8, 4))
        action_frame = ttk.Frame(root)
        action_frame.grid(row=action_row + 1, column=0, columnspan=4, sticky=tk.EW, pady=(0, 8))
        self._build_run_action_buttons(action_frame)

        drop_row = action_row + 2
        self.drop_label = ttk.Label(root, text="CSV/Excelをここへドロップすると入力フォルダを設定します", relief=tk.GROOVE, anchor=tk.CENTER, padding=16)
        self.drop_label.grid(row=drop_row, column=0, columnspan=4, sticky=tk.EW, pady=(4, 8))
        self._setup_drop_target()

        status_row = drop_row + 1
        ttk.Label(root, text="結果・エラー表示", style="Heading.TLabel").grid(row=status_row, column=0, sticky=tk.W, pady=(0, 4))
        ttk.Label(root, textvariable=self.status).grid(row=status_row, column=1, columnspan=3, sticky=tk.W, pady=(0, 4))
        self.output = tk.Text(root, height=8, wrap=tk.WORD)
        self.output.grid(row=status_row + 1, column=0, columnspan=4, sticky=tk.NSEW, pady=(8, 0))

        root.columnconfigure(1, weight=1)
        root.columnconfigure(3, weight=0)
        root.rowconfigure(status_row + 1, weight=1)
        self._bind_run_tab_child_mousewheel(root, root._run_tab_mousewheel_handler)  # type: ignore[attr-defined]

    def _build_history_tab(self, root: ttk.Frame) -> None:
        ttk.Label(
            root,
            text="作成済みのExcelレポートを確認できます。レポート作成後に一覧へ表示されます。",
            justify=tk.LEFT,
            wraplength=900,
        ).pack(fill=tk.X, pady=(0, 8))
        buttons = ttk.Frame(root)
        buttons.pack(fill=tk.X, pady=(0, 8))
        self._button(buttons, text="履歴を更新", command=self._refresh_history, width=SECONDARY_BUTTON_WIDTH).pack(side=tk.LEFT)
        self._button(buttons, text="詳細", command=self._show_selected_report_detail, width=SECONDARY_BUTTON_WIDTH).pack(side=tk.LEFT, padx=(8, 0))
        self._button(buttons, text="選択したレポートを開く", command=self._open_selected_report).pack(side=tk.LEFT, padx=(8, 0))

        self.history_list = tk.Listbox(root, height=12)
        self.history_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.history_list.bind("<Double-Button-1>", lambda _event: self._show_selected_report_detail())
        history_scroll = ttk.Scrollbar(root, orient=tk.VERTICAL, command=self.history_list.yview)
        history_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.history_list.configure(yscrollcommand=history_scroll.set)
        self._refresh_history()

    def _build_error_tab(self, root: ttk.Frame) -> None:
        ttk.Label(
            root,
            text="入力データの検証エラーを確認できます。エラー内容と修正方法を確認し、CSV / Excelを修正してから再実行してください。",
            justify=tk.LEFT,
            wraplength=900,
        ).pack(fill=tk.X, pady=(0, 8))

        buttons = ttk.Frame(root)
        buttons.pack(fill=tk.X, pady=(0, 8))
        self._button(buttons, text="入力データを検証", command=self._review_input_errors).pack(side=tk.LEFT)
        self._button(buttons, text="エラー行を再読み込み", command=self._review_input_errors).pack(side=tk.LEFT, padx=(8, 0))
        self._button(buttons, text="エラーCSVを開く", command=self._open_error_csv).pack(side=tk.LEFT, padx=(8, 0))
        self._button(buttons, text="エラーCSVの保存先を開く", command=self._open_error_csv_folder).pack(side=tk.LEFT, padx=(8, 0))

        self.error_review_status = tk.StringVar(value=NO_VALIDATION_ERRORS_MESSAGE)
        ttk.Label(root, textvariable=self.error_review_status, justify=tk.LEFT, wraplength=900).pack(fill=tk.X, pady=(0, 8))

        table_frame = ttk.Frame(root)
        table_frame.pack(fill=tk.BOTH, expand=True)
        columns = tuple(column for column, _label, _width in ERROR_REVIEW_COLUMNS)
        self.error_tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        y_scroll = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.error_tree.yview)
        x_scroll = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.error_tree.xview)
        self.error_tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.error_tree.grid(row=0, column=0, sticky=tk.NSEW)
        y_scroll.grid(row=0, column=1, sticky=tk.NS)
        x_scroll.grid(row=1, column=0, sticky=tk.EW)
        self._bind_tree_mousewheel(self.error_tree)
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        for column, label, width in ERROR_REVIEW_COLUMNS:
            self.error_tree.heading(column, text=label)
            self.error_tree.column(column, width=width, minwidth=70, stretch=True)
        self._set_error_review_rows([])

    def _build_audit_tab(self, root: ttk.Frame) -> None:
        ttk.Label(
            root,
            text="実行履歴やエラー内容を確認できます。レポート作成時に問題が起きた場合は、この画面を確認してください。",
            justify=tk.LEFT,
            wraplength=900,
        ).pack(fill=tk.X, pady=(0, 8))
        buttons = ttk.Frame(root)
        buttons.pack(fill=tk.X, pady=(0, 8))
        self._button(buttons, text="履歴を更新", command=self._refresh_audit_table, width=SECONDARY_BUTTON_WIDTH).pack(side=tk.LEFT)
        self._button(buttons, text="CSVエクスポート", command=self._export_audit_csv, width=SECONDARY_BUTTON_WIDTH).pack(side=tk.LEFT, padx=(8, 0))
        self._button(buttons, text="要約CSV", command=self._export_audit_summary_csv, width=SECONDARY_BUTTON_WIDTH).pack(side=tk.LEFT, padx=(8, 0))
        self._button(buttons, text="月別要約CSV", command=self._export_audit_monthly_summary_csv, width=SECONDARY_BUTTON_WIDTH).pack(side=tk.LEFT, padx=(8, 0))
        self._button(buttons, text="バックアップ", command=self._backup_audit_log, width=SECONDARY_BUTTON_WIDTH).pack(side=tk.LEFT, padx=(8, 0))
        self._button(buttons, text="異常チェック", command=self._show_audit_anomalies, width=SECONDARY_BUTTON_WIDTH).pack(side=tk.LEFT, padx=(8, 0))

        self._button(buttons, text="履歴から条件復元", command=self._restore_selected_audit_record).pack(side=tk.LEFT, padx=(8, 0))

        ttk.Label(buttons, text="検索").pack(side=tk.LEFT, padx=(16, 4))
        ttk.Entry(buttons, textvariable=self.audit_filter_text, width=24).pack(side=tk.LEFT)
        ttk.Combobox(buttons, textvariable=self.audit_filter_status, values=["all", "success", "validation_error", "error"], state="readonly", width=16).pack(side=tk.LEFT, padx=(8, 0))
        filters = ttk.Frame(root)
        filters.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(filters, text="期間").pack(side=tk.LEFT)
        ttk.Entry(filters, textvariable=self.audit_filter_date_from, width=12).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Label(filters, text="〜").pack(side=tk.LEFT, padx=4)
        ttk.Entry(filters, textvariable=self.audit_filter_date_to, width=12).pack(side=tk.LEFT)
        self._button(filters, text="適用", command=self._refresh_audit_table, width=SECONDARY_BUTTON_WIDTH).pack(side=tk.LEFT, padx=(8, 0))
        self._button(filters, text="修復プレビュー", command=self._preview_legacy_text_repairs, width=SECONDARY_BUTTON_WIDTH).pack(side=tk.RIGHT)

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
        self._bind_tree_mousewheel(self.audit_tree)
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
        self._button(buttons, text="ログ更新", command=self._refresh_log_text, width=SECONDARY_BUTTON_WIDTH).pack(side=tk.LEFT)
        self._button(buttons, text="監査ログ表示", command=self._refresh_audit_log_text, width=SECONDARY_BUTTON_WIDTH).pack(side=tk.LEFT, padx=(8, 0))

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
            raise ValueError("実行履歴の保持設定は1以上で指定してください。")
        return parsed

    def _column_aliases(self) -> dict[str, str]:
        text = self.column_aliases_json.get().strip() or "{}"
        value = json.loads(text)
        if not isinstance(value, dict):
            raise ValueError("列名マッピングはJSONオブジェクトで指定してください。")
        return {str(source).strip(): str(target).strip() for source, target in value.items() if str(source).strip()}

    def _standard_to_source_mapping(self, aliases: dict[str, str] | None = None) -> dict[str, str]:
        standard_to_source = {key: "" for key in MAPPING_STANDARD_LABELS}
        for source, standard in (aliases or self._column_aliases()).items():
            if standard in standard_to_source:
                standard_to_source[standard] = source
        return standard_to_source

    def _aliases_from_standard_mapping(self, standard_to_source: dict[str, str]) -> dict[str, str]:
        return {
            source.strip(): standard
            for standard, source in standard_to_source.items()
            if standard in MAPPING_STANDARD_LABELS and source and source.strip()
        }

    def _read_input_columns(self) -> tuple[str, ...]:
        return main.read_sales_columns(
            Path(self.input_dir.get().strip() or main.INPUT_DIR),
            self.pattern.get().strip() or main.DEFAULT_PATTERN,
        )

    def _missing_required_mapping_labels(self, columns: tuple[str, ...] | None = None) -> list[str]:
        if columns is None:
            columns = self._read_input_columns()
        return main.missing_required_column_labels(columns, self._column_aliases())

    def _ensure_column_mapping_ready(self) -> bool:
        try:
            columns = self._read_input_columns()
            missing_labels = self._missing_required_mapping_labels(columns)
        except Exception:
            return True
        if not missing_labels:
            return True
        self._show_error(
            "\n".join(
                [
                    "列名設定が不足しています。",
                    "",
                    "未設定の項目:",
                    *[f"- {label}" for label in missing_labels],
                    "",
                    "入力ファイルの列名を確認し、列名マッピングを設定してください。",
                ]
            )
        )
        return False

    def _refresh_column_aliases_from_input(self) -> None:
        try:
            columns = self._read_input_columns()
            inferred = main.infer_column_aliases(columns)
            current = self._column_aliases()
            current.update(inferred)
            self.column_aliases_json.set(json.dumps(current, ensure_ascii=False, indent=2))
            self._save_state()
            missing = main.missing_required_column_labels(columns, current)
            if missing:
                self._show_error(
                    "列名候補を取得しましたが、未設定の項目があります。\n\n"
                    + "\n".join(f"- {label}" for label in missing)
                    + "\n\n列名設定を開いて、入力ファイルの列名を選択してください。"
                )
                return
            self.status.set("入力ファイルの列名候補を取得しました")
            messagebox.showinfo("完了", "入力ファイルの列名候補を取得し、自動推定した列名設定を反映しました。")
        except Exception as exc:
            self._show_error(main.format_user_error_message(exc))

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
        window.title("列名マッピング設定")
        window.geometry("760x440")

        ttk.Label(
            window,
            text="入力ファイルの列名を、レポート作成で使う標準項目に対応させます。列名が標準と違う場合に設定してください。",
            justify=tk.LEFT,
            wraplength=700,
        ).pack(fill=tk.X, padx=12, pady=(12, 8))

        columns_var = tk.StringVar(value="列名候補: 未取得")
        ttk.Label(window, textvariable=columns_var, wraplength=700, foreground="#555555").pack(fill=tk.X, padx=12, pady=(0, 8))

        mapping_frame = ttk.Frame(window)
        mapping_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)
        ttk.Label(mapping_frame, text="標準項目").grid(row=0, column=0, sticky=tk.W, pady=(0, 6))
        ttk.Label(mapping_frame, text="入力ファイルの列名").grid(row=0, column=1, sticky=tk.W, pady=(0, 6))

        current_mapping = self._standard_to_source_mapping()
        selected_columns = {key: tk.StringVar(value=current_mapping.get(key, "")) for key in MAPPING_STANDARD_LABELS}
        combos: dict[str, ttk.Combobox] = {}
        candidate_columns: list[str] = []

        def set_candidates(columns: tuple[str, ...]) -> None:
            nonlocal candidate_columns
            candidate_columns = list(columns)
            values = [""] + candidate_columns
            columns_var.set("列名候補: " + (", ".join(candidate_columns) if candidate_columns else "候補がありません"))
            for combo in combos.values():
                combo.configure(values=values)

        for row, key in enumerate(MAPPING_STANDARD_LABELS, start=1):
            required_mark = " *" if key in MAPPING_REQUIRED_KEYS else ""
            ttk.Label(mapping_frame, text=f"{MAPPING_STANDARD_LABELS[key]}{required_mark}").grid(row=row, column=0, sticky=tk.W, pady=4)
            combo = ttk.Combobox(mapping_frame, textvariable=selected_columns[key], values=[""], state="readonly")
            combo.grid(row=row, column=1, sticky=tk.EW, pady=4, padx=(8, 0))
            combos[key] = combo

        mapping_frame.columnconfigure(1, weight=1)

        def apply_inferred(columns: tuple[str, ...]) -> None:
            inferred = main.infer_column_aliases(columns)
            standard_to_source = self._standard_to_source_mapping(inferred)
            for key, value in standard_to_source.items():
                if value:
                    selected_columns[key].set(value)

        def read_columns() -> None:
            try:
                columns = self._read_input_columns()
                set_candidates(columns)
                apply_inferred(columns)
                self.status.set("入力ファイルの列名候補を取得しました")
            except Exception as exc:
                self._show_error(main.format_user_error_message(exc))

        def reset_to_standard() -> None:
            standard_names = tuple(MAPPING_STANDARD_LABELS)
            set_candidates(standard_names)
            for key in MAPPING_STANDARD_LABELS:
                selected_columns[key].set(key)

        def save() -> None:
            mapping = {key: selected_columns[key].get() for key in MAPPING_STANDARD_LABELS}
            aliases = self._aliases_from_standard_mapping(mapping)
            missing = [MAPPING_STANDARD_LABELS[key] for key in MAPPING_REQUIRED_KEYS if not mapping.get(key)]
            if missing:
                self._show_error(
                    "列名設定が不足しています。\n\n"
                    "未設定の項目:\n"
                    + "\n".join(f"- {label}" for label in missing)
                    + "\n\n入力ファイルの列名を確認し、列名マッピングを設定してください。"
                )
                return
            self.column_aliases_json.set(json.dumps(aliases, ensure_ascii=False, indent=2))
            self._save_state()
            self.status.set("列名設定を保存しました")
            messagebox.showinfo("完了", "列名設定を保存しました。次回のプレビュー・レポート作成に反映されます。")
            window.destroy()

        try:
            columns = self._read_input_columns()
            set_candidates(columns)
            if not any(current_mapping.values()):
                apply_inferred(columns)
        except Exception:
            set_candidates(())

        buttons = ttk.Frame(window)
        buttons.pack(fill=tk.X, padx=12, pady=(0, 12))
        ttk.Button(buttons, text="入力ファイルの列名を読み取る", command=read_columns).pack(side=tk.LEFT)
        ttk.Button(buttons, text="プリセット", command=self._edit_column_alias_presets).pack(side=tk.LEFT)
        ttk.Button(buttons, text="初期設定に戻す", command=reset_to_standard).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(buttons, text="列名設定を保存", command=save).pack(side=tk.RIGHT)
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
        if not self._ensure_column_mapping_ready():
            return
        self.status.set("Excelレポートを作成中です...")
        self.output.delete("1.0", tk.END)
        threading.Thread(target=self._run_report, daemon=True).start()

    def _preview(self) -> None:
        if not self._validate_inputs():
            return
        if not self._ensure_column_mapping_ready():
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
        if not self._ensure_column_mapping_ready():
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
                "",
                "次の操作:",
                "「Excelレポートを開く」ボタンで作成したレポートを確認できます。",
                "「出力フォルダを開く」ボタンで保存先フォルダを確認できます。",
            ]
            for summary_csv_file in result.summary_csv_files:
                lines.append(f"集計CSV: {summary_csv_file}")
            for warning in result.warnings:
                lines.append(f"警告: {warning}")
            self.after(0, self._show_success, "\n".join(lines), result.output_file)
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
            self.after(0, self._show_error, main.format_user_error_message(exc) + f"\n\n詳細は「{TAB_LABELS['log']}」タブを確認してください。")

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
        self._bind_tree_mousewheel(tree)

    def _bind_tree_mousewheel(self, tree: ttk.Treeview) -> None:
        def on_mousewheel(event: tk.Event) -> str:
            delta = int(-1 * (event.delta / 120)) if getattr(event, "delta", 0) else 0
            tree.yview_scroll(delta, "units")
            return "break"

        def on_linux_scroll_up(_event: tk.Event) -> str:
            tree.yview_scroll(-1, "units")
            return "break"

        def on_linux_scroll_down(_event: tk.Event) -> str:
            tree.yview_scroll(1, "units")
            return "break"

        tree.bind("<MouseWheel>", on_mousewheel)
        tree.bind("<Button-4>", on_linux_scroll_up)
        tree.bind("<Button-5>", on_linux_scroll_down)

    def _remember_latest_report(self, report_file: Path | None) -> None:
        if report_file is not None:
            self.latest_report_path = Path(report_file)

    def _show_success(self, message: str, report_file: Path | None = None) -> None:
        self.status.set("レポート作成が完了しました")
        self._remember_latest_report(report_file)
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
        window = tk.Toplevel(self)
        window.title("エラー内容と確認ポイント")
        window.geometry(ERROR_DIALOG_SIZE)
        window.minsize(*ERROR_DIALOG_MIN_SIZE)
        window.resizable(True, True)

        header = ttk.Label(window, text="エラーが発生しました", style="Error.TLabel", font=("Meiryo", 11, "bold"))
        header.pack(fill=tk.X, padx=12, pady=(12, 6))

        frame = ttk.Frame(window, padding=(12, 0, 12, 12))
        frame.pack(fill=tk.BOTH, expand=True)

        text = tk.Text(frame, wrap=tk.NONE, relief=tk.SOLID, borderwidth=1)
        y_scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=text.yview)
        x_scroll = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=text.xview)
        text.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        text.grid(row=0, column=0, sticky=tk.NSEW)
        y_scroll.grid(row=0, column=1, sticky=tk.NS)
        x_scroll.grid(row=1, column=0, sticky=tk.EW)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        text.insert(tk.END, message)
        text.configure(state=tk.DISABLED)

        def on_mousewheel(event: tk.Event) -> str:
            delta = int(-1 * (event.delta / 120)) if getattr(event, "delta", 0) else 0
            text.yview_scroll(delta, "units")
            return "break"

        text.bind("<MouseWheel>", on_mousewheel)
        text.bind("<Button-4>", lambda _event: (text.yview_scroll(-1, "units"), "break")[1])
        text.bind("<Button-5>", lambda _event: (text.yview_scroll(1, "units"), "break")[1])

        buttons = ttk.Frame(window, padding=(12, 0, 12, 12))
        buttons.pack(fill=tk.X)
        self._button(buttons, text="閉じる", command=window.destroy, width=SECONDARY_BUTTON_WIDTH).pack(side=tk.RIGHT)

    def _error_rows_from_validation_error(self, error: main.DataValidationError) -> list[dict[str, str]]:
        error_df = main.create_validation_error_rows(error)
        rows: list[dict[str, str]] = []
        for record in error_df.to_dict(orient="records"):
            rows.append({column: "" if record.get(column) is None else str(record.get(column, "")) for column, _label, _width in ERROR_REVIEW_COLUMNS})
        return rows

    def _set_error_review_rows(self, rows: list[dict[str, str]]) -> None:
        if not hasattr(self, "error_tree"):
            return
        self.error_tree.delete(*self.error_tree.get_children())
        if not rows:
            if hasattr(self, "error_review_status"):
                self.error_review_status.set(NO_VALIDATION_ERRORS_MESSAGE)
            self.error_tree.insert("", tk.END, values=("", "", "検証エラーはありません", "現在の入力データは正常に読み込めます。", "", "", "", "", "", ""))
            return
        if hasattr(self, "error_review_status"):
            self.error_review_status.set(f"検証エラーが{len(rows)}件あります。エラー内容と修正方法を確認してください。")
        for row in rows:
            self.error_tree.insert("", tk.END, values=tuple(row.get(column, "") for column, _label, _width in ERROR_REVIEW_COLUMNS))

    def _select_error_tab(self) -> None:
        if hasattr(self, "notebook") and hasattr(self, "error_tree"):
            parent = self.error_tree.winfo_toplevel()
            for tab_id in self.notebook.tabs():
                if self.notebook.tab(tab_id, "text") == TAB_LABELS["error"]:
                    self.notebook.select(tab_id)
                    break

    def _review_input_errors(self) -> None:
        if not self._validate_inputs():
            return
        try:
            merged_df = main.read_sales_files(
                Path(self.input_dir.get().strip() or main.INPUT_DIR),
                self.pattern.get().strip() or main.DEFAULT_PATTERN,
            )
            validated_df = main.validate_data(merged_df, self._column_aliases())
            main.filter_data(
                validated_df,
                month=self._optional(self.month.get()),
                start_date=self._optional(self.start_date.get()),
                end_date=self._optional(self.end_date.get()),
                product=self._optional(self.product.get()),
                category=self._optional(self.category.get()),
            )
            self.latest_error_csv_path = None
            self._set_error_review_rows([])
            self.status.set("入力データの検証が完了しました")
            self._select_error_tab()
        except main.DataValidationError as exc:
            output_dir = Path(self.output_dir.get().strip() or main.OUTPUT_DIR)
            report_file = main.default_error_report_path(output_dir)
            main.write_validation_error_report(exc, report_file)
            self.latest_error_csv_path = report_file
            self._set_error_review_rows(self._error_rows_from_validation_error(exc))
            self.status.set("入力データに検証エラーがあります")
            self._select_error_tab()
        except Exception as exc:
            self._show_error(main.format_user_error_message(exc))

    def _open_error_csv(self) -> None:
        if self.latest_error_csv_path is None:
            self._show_error(NO_ERROR_CSV_MESSAGE)
            return
        self._open_report_file(self.latest_error_csv_path)

    def _open_error_csv_folder(self) -> None:
        if self.latest_error_csv_path is None:
            self._show_error(NO_ERROR_CSV_MESSAGE)
            return
        if not self.latest_error_csv_path.exists():
            self._show_error("エラーCSVが見つかりません。\n削除または移動された可能性があります。")
            return
        self._open_path(self.latest_error_csv_path.parent)

    def _show_validation_error(self, error: main.DataValidationError, report_file: Path) -> None:
        message = "\n".join(
            [
                main.format_user_error_message(error),
                "",
                f"エラー一覧CSV: {report_file}",
                "",
                f"詳細は「{TAB_LABELS['log']}」タブも確認してください。",
            ]
        )
        self._show_error(message)
        self.latest_error_csv_path = report_file
        self._set_error_review_rows(self._error_rows_from_validation_error(error))
        self._select_error_tab()
        window = tk.Toplevel(self)
        window.title("エラー一覧と修正方法")
        window.geometry(ERROR_TABLE_DIALOG_SIZE)
        window.minsize(*ERROR_TABLE_DIALOG_MIN_SIZE)
        window.resizable(True, True)
        columns = ("message", "fix", "source_file", "source_row")
        frame = ttk.Frame(window, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)
        tree = ttk.Treeview(frame, columns=columns, show="headings")
        y_scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        x_scroll = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        tree.grid(row=0, column=0, sticky=tk.NSEW)
        y_scroll.grid(row=0, column=1, sticky=tk.NS)
        x_scroll.grid(row=1, column=0, sticky=tk.EW)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
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
        self._bind_tree_mousewheel(tree)
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
            self._show_error(REPORT_SELECTION_REQUIRED_MESSAGE)
            return None
        selected_index = selection[0]
        if selected_index >= len(self.report_history):
            self._show_error(REPORT_SELECTION_REQUIRED_MESSAGE)
            return None
        return self.report_history[selected_index]

    def _open_path(self, path: Path) -> None:
        if sys.platform == "win32":
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])

    def _open_report_file(self, report_file: Path | None) -> None:
        if report_file is None:
            self._show_error(NO_REPORT_TO_OPEN_MESSAGE)
            return
        report_file = Path(report_file)
        if not report_file.exists() or not report_file.is_file():
            self._show_error(REPORT_FILE_MISSING_MESSAGE)
            return
        try:
            self._open_path(report_file)
        except Exception as exc:
            self._show_error(f"レポートファイルを開けませんでした。\n{exc}")

    def _open_latest_report(self) -> None:
        self._open_report_file(self.latest_report_path)

    def _open_output_folder(self) -> None:
        output_dir = Path(self.output_dir.get().strip() or main.OUTPUT_DIR)
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            self._show_error(f"{OUTPUT_FOLDER_MISSING_MESSAGE}\n{exc}")
            return
        if not output_dir.exists() or not output_dir.is_dir():
            self._show_error(OUTPUT_FOLDER_MISSING_MESSAGE)
            return
        try:
            self._open_path(output_dir)
        except Exception as exc:
            self._show_error(f"出力フォルダを開けませんでした。\n{exc}")

    def _open_selected_report(self) -> None:
        report_file = self._selected_report()
        self._open_report_file(report_file)

    def _show_selected_report_detail(self) -> None:
        report_file = self._selected_report()
        if report_file is None:
            return
        if not report_file.exists() or not report_file.is_file():
            self._show_error(REPORT_FILE_MISSING_MESSAGE)
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
        ttk.Button(buttons, text="選択したレポートを開く", command=lambda: self._open_report_file(report_file)).pack(side=tk.LEFT, padx=8)
        ttk.Button(buttons, text="出力フォルダを開く", command=lambda: self._open_path(report_file.parent)).pack(side=tk.LEFT)

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
            self._show_error("出力対象の実行履歴がありません。")
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
            self.status.set("実行履歴CSVを出力しました")
            messagebox.showinfo("完了", f"実行履歴CSVを出力しました。\n{export_file}")
        except Exception as exc:
            self._show_error(str(exc))

    def _export_audit_summary_csv(self) -> None:
        records = self._filtered_audit_records()
        if not records:
            self._show_error("出力対象の実行履歴がありません。")
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
            self.status.set("実行履歴要約CSVを出力しました")
            messagebox.showinfo("完了", f"実行履歴要約CSVを出力しました。\n{export_file}")
        except Exception as exc:
            self._show_error(str(exc))

    def _export_audit_monthly_summary_csv(self) -> None:
        records = self._filtered_audit_records()
        if not records:
            self._show_error("出力対象の実行履歴がありません。")
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
            self.status.set("実行履歴月別要約CSVを出力しました")
            messagebox.showinfo("完了", f"実行履歴月別要約CSVを出力しました。\n{export_file}")
        except Exception as exc:
            self._show_error(str(exc))

    def _backup_audit_log(self) -> None:
        try:
            backup_file = main.backup_audit_log()
            if backup_file is None:
                self.status.set("バックアップ対象の実行履歴はありませんでした")
                messagebox.showinfo("確認", "バックアップ対象の実行履歴はありませんでした。")
                return
            self.status.set("実行履歴をバックアップしました")
            messagebox.showinfo("完了", f"実行履歴をバックアップしました。\n{backup_file}")
        except Exception as exc:
            self._show_error(str(exc))

    def _show_audit_anomalies(self) -> None:
        records = self._filtered_audit_records()
        if not records:
            self._show_error("確認対象の実行履歴がありません。")
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
        window.title("実行履歴異常チェック")
        window.geometry("560x320")
        text = tk.Text(window, height=12)
        text.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)
        text.insert(tk.END, "\n".join(lines))
        text.configure(state=tk.DISABLED)

    def _selected_audit_record(self) -> dict | None:
        selection = self.audit_tree.selection()
        if not selection:
            self._show_error("実行履歴を選択してください。")
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
        self.status.set("実行履歴から実行条件を復元しました")


def main_gui() -> None:
    main.setup_logging()
    app = MonthlyReportApp()
    app.mainloop()


if __name__ == "__main__":
    main_gui()

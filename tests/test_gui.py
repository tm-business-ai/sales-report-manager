import inspect
from pathlib import Path

import gui
import report


def test_gui_exposes_main_window_and_core_methods() -> None:
    app_class = gui.MonthlyReportApp

    assert gui.PRIMARY_BUTTON_WIDTH == 20
    assert gui.SECONDARY_BUTTON_WIDTH == 18
    assert gui.WINDOW_SCREEN_RATIO == 0.9
    assert gui.WINDOW_MAX_SIZE == (1400, 900)
    assert gui.WINDOW_MIN_SIZE == (1000, 650)
    assert gui.RUN_TAB_SCROLLBAR_WIDTH == 16
    assert gui.BUTTON_GRID_COLUMNS == 4
    assert gui.BUTTON_GRID_MIN_WIDTH == 180
    assert gui.HELP_WRAP_LENGTH == 420
    assert gui.PREFERRED_TTK_THEMES == ("vista", "xpnative", "default")
    assert gui.NO_REPORT_TO_OPEN_MESSAGE == "開くレポートを一覧から選択してください。"
    assert "削除または移動された可能性があります。" in gui.REPORT_FILE_MISSING_MESSAGE
    assert "出力" in gui.OUTPUT_FOLDER_MISSING_MESSAGE or "蜃ｺ蜉" in gui.OUTPUT_FOLDER_MISSING_MESSAGE

    for method_name in (
        "_edit_style_config",
        "_edit_column_alias_presets",
        "_open_config_wizard",
        "_review_current_settings",
        "_restore_selected_audit_record",
        "_refresh_audit_table",
        "_export_audit_csv",
        "_export_audit_summary_csv",
        "_export_audit_monthly_summary_csv",
        "_backup_audit_log",
        "_show_audit_anomalies",
        "_open_latest_report",
        "_open_output_folder",
        "_open_report_file",
        "_remember_latest_report",
        "_open_selected_report_folder",
        "_review_input_errors",
        "_open_error_csv",
        "_set_error_review_rows",
        "_error_rows_from_validation_error",
        "_configure_style",
        "_apply_preferred_theme",
        "_button",
        "_help_label",
        "_calculate_window_size",
        "_apply_initial_window_size",
        "_create_scrollable_run_frame",
        "_bind_canvas_mousewheel",
        "_bind_run_tab_child_mousewheel",
        "_bind_tree_mousewheel",
        "_report_row_from_path",
        "_is_error_log_row",
        "_filter_detail_log_rows",
        "_format_log_detail",
    ):
        assert hasattr(app_class, method_name)


def test_gui_window_size_is_based_on_screen_size() -> None:
    app_class = gui.MonthlyReportApp

    assert app_class._calculate_window_size(1920, 1080) == (1400, 900, 1000, 650)
    assert app_class._calculate_window_size(1024, 768) == (921, 691, 921, 650)


def test_gui_report_history_columns_and_row_conversion(tmp_path: Path) -> None:
    app_class = gui.MonthlyReportApp
    report_file = tmp_path / "monthly_report_202604.xlsx"
    report_file.write_bytes(b"dummy report")

    assert [label for _key, label, _width in gui.REPORT_HISTORY_COLUMNS] == [
        "作成日時",
        "対象月",
        "ファイル名",
        "種類",
        "サイズ",
        "保存先",
    ]
    row = app_class._report_row_from_path(report_file)

    assert row["target_month"] == "2026-04"
    assert row["file_name"] == "monthly_report_202604.xlsx"
    assert row["report_type"] == "月次レポート"
    assert row["size"].endswith("KB")
    assert row["path"] == str(report_file)


def test_gui_audit_buttons_and_scrollable_columns_are_defined() -> None:
    assert [label for label, _method_name in gui.AUDIT_ACTION_BUTTONS] == [
        "履歴を更新",
        "履歴をCSV出力",
        "要約CSVを作成",
        "月別要約CSVを作成",
        "履歴をバックアップ",
        "異常データを確認",
        "修復内容を確認",
    ]
    widths = {column: width for column, _label, width in gui.AUDIT_HISTORY_COLUMNS}

    assert widths["timestamp"] == 120
    assert widths["output_file"] == 260
    assert widths["error"] == 360
    assert "xscrollcommand=x_scroll.set" in inspect.getsource(gui.MonthlyReportApp._build_audit_tab)


def test_gui_detail_log_columns_error_detection_and_filtering() -> None:
    app_class = gui.MonthlyReportApp
    assert [label for _key, label, _width in gui.DETAIL_LOG_COLUMNS] == ["日時", "レベル", "処理", "内容", "ファイル"]

    rows = [
        {"level": "INFO", "message": "完了しました"},
        {"level": "ERROR", "message": "validation_error: 数量が不正です"},
        {"level": "INFO", "message": "処理に失敗しました"},
    ]

    assert app_class._is_error_log_row(rows[1])
    assert app_class._is_error_log_row(rows[2])
    assert app_class._filter_detail_log_rows(rows, errors_only=True) == rows[1:]
    assert app_class._filter_detail_log_rows(rows, errors_only=False) == rows


def test_gui_parse_log_line_and_detail_format() -> None:
    row = gui.MonthlyReportApp._parse_log_line(
        '{"timestamp":"2026-05-15 14:07","level":"error","process":"report","message":"失敗しました"}',
        "app.log",
    )

    assert row["timestamp"] == "2026-05-15 14:07"
    assert row["level"] == "ERROR"
    assert row["process"] == "report"
    assert row["message"] == "失敗しました"
    assert '"message": "失敗しました"' in gui.MonthlyReportApp._format_log_detail(row)


def test_gui_remembers_latest_report_path() -> None:
    app_class = gui.MonthlyReportApp
    app = object.__new__(app_class)
    app.latest_report_path = None

    app_class._remember_latest_report(app, Path("data/output/report.xlsx"))

    assert app.latest_report_path == Path("data/output/report.xlsx")


def test_gui_open_report_file_guidance_without_launching(tmp_path: Path) -> None:
    app_class = gui.MonthlyReportApp
    app = object.__new__(app_class)
    messages: list[str] = []
    opened: list[Path] = []
    app._show_error = messages.append
    app._open_path = opened.append

    app_class._open_report_file(app, None)
    assert gui.NO_REPORT_TO_OPEN_MESSAGE in messages[-1]

    missing = tmp_path / "missing.xlsx"
    app_class._open_report_file(app, missing)
    assert gui.REPORT_FILE_MISSING_MESSAGE in messages[-1]

    report_file = tmp_path / "report.xlsx"
    report_file.write_text("dummy", encoding="utf-8")
    app_class._open_report_file(app, report_file)
    assert opened == [report_file]


def test_gui_open_output_folder_creates_and_opens_folder(tmp_path: Path) -> None:
    app_class = gui.MonthlyReportApp
    app = object.__new__(app_class)
    opened: list[Path] = []
    messages: list[str] = []
    app.output_dir = type("DummyVar", (), {"get": lambda self: str(tmp_path / "output")})()
    app._open_path = opened.append
    app._show_error = messages.append

    app_class._open_output_folder(app)

    assert (tmp_path / "output").is_dir()
    assert opened == [tmp_path / "output"]
    assert messages == []


def test_gui_error_rows_include_fix_column() -> None:
    app_class = gui.MonthlyReportApp
    error = gui.main.DataValidationError(
        [
            report.ValidationIssue(
                "invalid_quantity",
                "数量に数値以外の値があります。不正な値: abc",
                "sales.csv",
                2,
                "数量には 1 以上の数値を入力してください。",
                {"quantity": "abc", "product": "りんご"},
            )
        ]
    )

    rows = app_class._error_rows_from_validation_error(object.__new__(app_class), error)

    assert rows[0]["source_file"] == "sales.csv"
    assert rows[0]["fix"].startswith("数量には")
    assert rows[0]["quantity"] == "abc"


def test_gui_open_error_csv_guidance_without_launching(tmp_path: Path) -> None:
    app_class = gui.MonthlyReportApp
    app = object.__new__(app_class)
    messages: list[str] = []
    opened: list[Path] = []
    app._show_error = messages.append
    app._open_report_file = opened.append
    app.latest_error_csv_path = None

    app_class._open_error_csv(app)
    assert gui.NO_ERROR_CSV_MESSAGE in messages[-1]

    error_csv = tmp_path / "errors.csv"
    app.latest_error_csv_path = error_csv
    app_class._open_error_csv(app)
    assert opened == [error_csv]

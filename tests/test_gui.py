import gui
import report
from pathlib import Path


def test_gui_exposes_main_window_and_new_editor_methods() -> None:
    app_class = gui.MonthlyReportApp

    assert gui.APP_TITLE == "売上データ自動集計・月末売上管理ツール"
    assert "月末確認用のExcelレポート" in gui.RUN_TAB_DESCRIPTION
    assert "Excelレポートを作成" in gui.RUN_TAB_STEPS
    assert gui.MAPPING_STANDARD_LABELS["date"] == "日付"
    assert "unit_price" in gui.MAPPING_REQUIRED_KEYS
    assert gui.PRIMARY_BUTTON_WIDTH == 20
    assert gui.SECONDARY_BUTTON_WIDTH == 18
    assert gui.ERROR_DIALOG_SIZE == "1200x700"
    assert gui.ERROR_DIALOG_MIN_SIZE == (900, 500)
    assert gui.ERROR_TABLE_DIALOG_SIZE == "1200x700"
    assert gui.TAB_LABELS == {
        "run": "実行",
        "error": "エラー確認",
        "history": "作成済みレポート",
        "audit": "実行履歴",
        "log": "詳細ログ",
    }
    assert "先に「Excelレポートを作成」" in gui.NO_REPORT_TO_OPEN_MESSAGE
    assert "削除または移動" in gui.REPORT_FILE_MISSING_MESSAGE
    assert "出力フォルダ" in gui.OUTPUT_FOLDER_MISSING_MESSAGE
    assert "検証エラーはありません" in gui.NO_VALIDATION_ERRORS_MESSAGE
    assert "まだエラーCSV" in gui.NO_ERROR_CSV_MESSAGE
    assert any(column == "fix" and label == "修正方法" for column, label, _width in gui.ERROR_REVIEW_COLUMNS)
    assert hasattr(gui.main, "format_user_error_message")
    assert hasattr(gui.main, "read_sales_columns")
    assert hasattr(gui.main, "infer_column_aliases")
    assert hasattr(app_class, "_edit_style_config")
    assert hasattr(app_class, "_edit_column_alias_presets")
    assert hasattr(app_class, "_open_config_wizard")
    assert hasattr(app_class, "_review_current_settings")
    assert hasattr(app_class, "_warning_amount_threshold")
    assert hasattr(app_class, "_restore_selected_audit_record")
    assert hasattr(app_class, "_refresh_audit_table")
    assert hasattr(app_class, "_export_audit_csv")
    assert hasattr(app_class, "_repair_legacy_text_files")
    assert hasattr(app_class, "_export_audit_summary_csv")
    assert hasattr(app_class, "_export_audit_monthly_summary_csv")
    assert hasattr(app_class, "_preview_legacy_text_repairs")
    assert hasattr(app_class, "_backup_audit_log")
    assert hasattr(app_class, "_show_audit_anomalies")
    assert hasattr(app_class, "_open_latest_report")
    assert hasattr(app_class, "_open_output_folder")
    assert hasattr(app_class, "_open_report_file")
    assert hasattr(app_class, "_remember_latest_report")
    assert hasattr(app_class, "_review_input_errors")
    assert hasattr(app_class, "_open_error_csv")
    assert hasattr(app_class, "_set_error_review_rows")
    assert hasattr(app_class, "_error_rows_from_validation_error")
    assert hasattr(app_class, "_configure_style")
    assert hasattr(app_class, "_button")
    assert hasattr(app_class, "_bind_tree_mousewheel")


def test_gui_column_mapping_helpers_convert_between_shapes() -> None:
    app_class = gui.MonthlyReportApp
    aliases = app_class._aliases_from_standard_mapping(
        object.__new__(app_class),
        {"date": "売上日", "product": "品名", "category": "", "quantity": "販売数", "unit_price": "販売単価", "amount": "売上金額"},
    )

    assert aliases == {"売上日": "date", "品名": "product", "販売数": "quantity", "販売単価": "unit_price", "売上金額": "amount"}


def test_gui_remembers_latest_report_path() -> None:
    app_class = gui.MonthlyReportApp
    app = object.__new__(app_class)
    app.latest_report_path = None

    app_class._remember_latest_report(app, Path("data/output/report.xlsx"))

    assert app.latest_report_path == Path("data/output/report.xlsx")


def test_gui_open_report_file_guidance_without_launching(tmp_path: Path, monkeypatch) -> None:
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
                "数量には 1、2、10 のような0以上の数値を入力してください。",
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

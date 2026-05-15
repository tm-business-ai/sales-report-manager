import gui


def test_gui_exposes_main_window_and_new_editor_methods() -> None:
    app_class = gui.MonthlyReportApp

    assert gui.APP_TITLE == "売上データ自動集計・月末売上管理ツール"
    assert "月末確認用のExcelレポート" in gui.RUN_TAB_DESCRIPTION
    assert "Excelレポートを作成" in gui.RUN_TAB_STEPS
    assert gui.MAPPING_STANDARD_LABELS["date"] == "日付"
    assert "unit_price" in gui.MAPPING_REQUIRED_KEYS
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


def test_gui_column_mapping_helpers_convert_between_shapes() -> None:
    app_class = gui.MonthlyReportApp
    aliases = app_class._aliases_from_standard_mapping(
        object.__new__(app_class),
        {"date": "売上日", "product": "品名", "category": "", "quantity": "販売数", "unit_price": "販売単価", "amount": "売上金額"},
    )

    assert aliases == {"売上日": "date", "品名": "product", "販売数": "quantity", "販売単価": "unit_price", "売上金額": "amount"}

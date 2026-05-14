import gui


def test_gui_exposes_main_window_and_new_editor_methods() -> None:
    app_class = gui.MonthlyReportApp

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

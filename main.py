import argparse
import csv
import importlib.util
import json
import logging
import sys
import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any
from urllib import request

from report import (
    DataValidationError,
    GROUP_BY_COLUMNS,
    REQUIRED_COLUMN_LABELS,
    REQUIRED_COLUMNS,
    STANDARD_COLUMN_ORDER,
    cleanup_old_reports,
    create_daily_trend,
    create_month_end_summary,
    create_monthly_trend,
    create_previous_month_comparison,
    create_uncategorized_rows,
    create_validation_error_rows,
    create_summaries,
    filter_data,
    format_validation_issues,
    infer_column_aliases,
    missing_required_column_labels,
    read_sales_columns,
    read_sales_files,
    save_to_excel,
    validate_data,
    validate_date_option,
    write_summary_csvs,
    write_validation_error_report,
)


BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "data" / "input"
OUTPUT_DIR = BASE_DIR / "data" / "output"
LOG_DIR = BASE_DIR / "logs"
LOG_FILE = LOG_DIR / "app.log"
AUDIT_LOG_FILE = LOG_DIR / "report_history.jsonl"
ALIAS_PRESET_FILE = LOG_DIR / "column_alias_presets.json"
LEGACY_BACKUP_SUFFIX = ".bak"
DEFAULT_PATTERN = "*.csv"
REQUIRED_MODULES = ("pandas", "openpyxl")
LEGACY_MOJIBAKE_MARKERS = (
    "繝",
    "縺",
    "螳",
    "逶",
    "險",
    "蜃",
    "譌",
    "蛻",
    "髮",
    "驕",
    "讀",
    "迥",
    "隴",
    "蜈",
    "蟇",
    "邨",
    "蝠",
    "騾",
    "豁ｴ",
    "繧",
    "遘",
    "邏",
)
CONFIG_KEYS = {
    "input_dir": (str,),
    "output_dir": (str,),
    "month": (str, type(None)),
    "start_date": (str, type(None)),
    "end_date": (str, type(None)),
    "product": (str, type(None)),
    "category": (str, type(None)),
    "group_by": (str,),
    "keep_reports": (int, type(None)),
    "pattern": (str,),
    "dry_run": (bool,),
    "all_summaries": (bool,),
    "monthly_trend": (bool,),
    "charts": (bool,),
    "output_name": (str, type(None)),
    "error_report": (str, type(None)),
    "summary_csv_dir": (str, type(None)),
    "summary_csv_prefix": (str,),
    "notify": (bool,),
    "notify_webhook_url": (str, type(None)),
    "column_aliases": (dict,),
    "warning_amount_threshold": (int, float),
    "style": (dict,),
    "audit_keep_count": (int, type(None)),
    "audit_keep_days": (int, type(None)),
}
STYLE_CONFIG_KEYS = {
    "header_fill": (str,),
    "total_fill": (str,),
    "title_size": (int, float),
    "chart_height": (int, float),
    "chart_width": (int, float),
}


@dataclass(frozen=True)
class ReportRunResult:
    output_file: Path | None
    detail_count: int
    summary_count: int
    dry_run: bool = False
    summary_csv_files: tuple[Path, ...] = ()
    warnings: tuple[str, ...] = ()
    total_amount: float = 0
    target_days: int = 0
    product_count: int = 0
    category_count: int = 0
    total_quantity: float = 0
    average_unit_price: float = 0
    previous_month_amount: float = 0
    previous_month_change_rate: float | str = "比較不可"
    uncategorized_count: int = 0
    error_count: int = 0


@dataclass(frozen=True)
class PreviewResult:
    columns: tuple[str, ...]
    rows: tuple[tuple[str, ...], ...]
    total_count: int
    source_files: tuple[str, ...] = ()


@dataclass(frozen=True)
class SummaryPreviewResult:
    summaries: dict[str, PreviewResult]
    detail_count: int


@dataclass(frozen=True)
class LegacyTextChange:
    location: str
    original: str
    normalized: str


@dataclass(frozen=True)
class AuditAnomalySummary:
    total_runs: int
    success_count: int
    failure_count: int
    warning_total: int
    last_run_date: str
    consecutive_missing_days: int
    failure_rate: float
    alerts: tuple[str, ...]


def _mojibake_score(text: str) -> int:
    return sum(text.count(marker) for marker in LEGACY_MOJIBAKE_MARKERS)


def normalize_legacy_text(text: str) -> str:
    if not isinstance(text, str) or _mojibake_score(text) == 0:
        return text

    candidates = {text}
    for encode_errors, decode_errors in (("ignore", "ignore"), ("replace", "replace"), ("replace", "ignore")):
        try:
            repaired = text.encode("cp932", errors=encode_errors).decode("utf-8", errors=decode_errors)
        except UnicodeError:
            continue
        if repaired:
            candidates.add(repaired)

    def quality(candidate: str) -> tuple[int, int, int, int]:
        readable_count = sum(
            1
            for char in candidate
            if char.isascii() or "\u3040" <= char <= "\u30ff" or "\u4e00" <= char <= "\u9fff"
        )
        return (_mojibake_score(candidate), -readable_count, candidate.count("\ufffd"), -len(candidate))

    return min(candidates, key=quality)


def normalize_legacy_value(value: Any) -> Any:
    if isinstance(value, str):
        return normalize_legacy_text(value)
    if isinstance(value, list):
        return [normalize_legacy_value(item) for item in value]
    if isinstance(value, tuple):
        return tuple(normalize_legacy_value(item) for item in value)
    if isinstance(value, dict):
        return {str(key): normalize_legacy_value(item) for key, item in value.items()}
    return value


def collect_legacy_text_changes(value: Any, location: str = "") -> list[LegacyTextChange]:
    changes: list[LegacyTextChange] = []
    if isinstance(value, str):
        normalized = normalize_legacy_text(value)
        if normalized != value:
            changes.append(LegacyTextChange(location or "<value>", value, normalized))
        return changes
    if isinstance(value, list):
        for index, item in enumerate(value):
            item_location = f"{location}[{index}]" if location else f"[{index}]"
            changes.extend(collect_legacy_text_changes(item, item_location))
        return changes
    if isinstance(value, tuple):
        for index, item in enumerate(value):
            item_location = f"{location}[{index}]" if location else f"[{index}]"
            changes.extend(collect_legacy_text_changes(item, item_location))
        return changes
    if isinstance(value, dict):
        for key, item in value.items():
            item_location = f"{location}.{key}" if location else str(key)
            changes.extend(collect_legacy_text_changes(item, item_location))
    return changes


def inspect_legacy_text_file(path: Path) -> list[LegacyTextChange]:
    if not path.exists():
        return []
    suffix = path.suffix.lower()
    content = path.read_text(encoding="utf-8", errors="replace")
    if suffix == ".jsonl":
        changes: list[LegacyTextChange] = []
        for line_number, line in enumerate(content.splitlines(), 1):
            if not line.strip():
                continue
            try:
                data = json.loads(line)
            except json.JSONDecodeError:
                continue
            changes.extend(collect_legacy_text_changes(data, f"line[{line_number}]"))
        return changes
    data = json.loads(content)
    return collect_legacy_text_changes(data)


def inspect_legacy_text_files(
    gui_state_file: Path | None = None,
    audit_log_file: Path | None = None,
    alias_preset_file: Path | None = None,
) -> dict[str, list[LegacyTextChange]]:
    gui_state_file = gui_state_file or LOG_DIR / "gui_state.json"
    audit_log_file = audit_log_file or AUDIT_LOG_FILE
    alias_preset_file = alias_preset_file or ALIAS_PRESET_FILE
    return {
        str(gui_state_file): inspect_legacy_text_file(gui_state_file),
        str(audit_log_file): inspect_legacy_text_file(audit_log_file),
        str(alias_preset_file): inspect_legacy_text_file(alias_preset_file),
    }


def setup_logging(max_bytes: int = 1_000_000, backup_count: int = 5) -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    if LOG_FILE.exists():
        try:
            LOG_FILE.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            LOG_FILE.replace(LOG_DIR / f"app_legacy_{timestamp}.log")

    handlers: list[logging.Handler]
    try:
        handlers = [RotatingFileHandler(LOG_FILE, maxBytes=max_bytes, backupCount=backup_count, encoding="utf-8")]
    except OSError:
        handlers = [logging.StreamHandler(sys.stderr)]
    logging.basicConfig(
        handlers=handlers,
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        force=True,
    )


def validate_keep_reports(value: str) -> int:
    try:
        keep_reports = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("--keep-reports は1以上の整数で指定してください。") from exc
    if keep_reports < 1:
        raise argparse.ArgumentTypeError("--keep-reports は1以上の整数で指定してください。")
    return keep_reports


def validate_positive_int_option(value: str, option_name: str) -> int:
    try:
        parsed = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"{option_name} は1以上の整数で指定してください。") from exc
    if parsed < 1:
        raise argparse.ArgumentTypeError(f"{option_name} は1以上の整数で指定してください。")
    return parsed


def validate_audit_keep_count(value: str) -> int:
    return validate_positive_int_option(value, "--audit-keep-count")


def validate_audit_keep_days(value: str) -> int:
    return validate_positive_int_option(value, "--audit-keep-days")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="CSVから月次売上レポートを作成します。")
    parser.add_argument("--config", type=Path, help="JSON設定ファイルを指定します。")
    parser.add_argument("--preset", help="設定ファイル内の presets から実行条件を選びます。")
    parser.add_argument("--month", default=None, help="集計対象月を YYYY-MM 形式で指定します。")
    parser.add_argument("--start-date", default=None, help="集計開始日を YYYY-MM-DD 形式で指定します。")
    parser.add_argument("--end-date", default=None, help="集計終了日を YYYY-MM-DD 形式で指定します。")
    parser.add_argument("--product", default=None, help="指定した商品名だけに絞り込みます。")
    parser.add_argument("--category", default=None, help="指定したカテゴリだけに絞り込みます。")
    parser.add_argument("--group-by", choices=sorted(GROUP_BY_COLUMNS), default=None, help="集計単位を指定します。")
    parser.add_argument("--all-summaries", action="store_true", default=None, help="商品別とカテゴリ別の集計シートを両方出力します。")
    parser.add_argument("--monthly-trend", action="store_true", default=None, help="月別推移シートを出力します。")
    parser.add_argument("--charts", action="store_true", default=None, help="Excelにグラフを追加します。")
    parser.add_argument("--input-dir", type=Path, default=None, help=f"入力CSV / Excelフォルダを指定します。既定値は {INPUT_DIR} です。")
    parser.add_argument("--output-dir", type=Path, default=None, help=f"Excel出力フォルダを指定します。既定値は {OUTPUT_DIR} です。")
    parser.add_argument("--output-name", default=None, help="出力Excelファイル名を指定します。")
    parser.add_argument("--pattern", default=None, help='読み込む入力ファイルのパターンを指定します。対応形式は .csv / .xlsx / .xls です。既定値は "*.csv" です。')
    parser.add_argument("--keep-reports", type=validate_keep_reports, default=None, help="出力フォルダに残す月次レポート数を1以上で指定します。")
    parser.add_argument("--dry-run", action="store_true", default=None, help="CSV検証と集計だけを実行し、Excelは出力しません。")
    parser.add_argument("--error-report", type=Path, default=None, help="検証エラーがある場合にエラー一覧CSVを出力します。")
    parser.add_argument("--summary-csv-dir", type=Path, default=None, help="集計結果CSVの出力フォルダを指定します。")
    parser.add_argument("--summary-csv-prefix", default=None, help="集計結果CSVのファイル名プレフィックスを指定します。")
    parser.add_argument("--check-setup", action="store_true", help="実行環境の初期設定を確認して終了します。")
    parser.add_argument("--preview", action="store_true", help="CSVを読み込み、出力せずに先頭行を表示します。")
    parser.add_argument("--preview-limit", type=int, default=10, help="--preview で表示する最大行数です。")
    parser.add_argument("--notify", action="store_true", default=None, help="処理完了時に通知音を鳴らします。")
    parser.add_argument("--notify-webhook-url", default=None, help="完了/失敗通知を送るWebhook URLを指定します。")
    parser.add_argument("--write-template", type=Path, default=None, help="入力CSVテンプレートを書き出して終了します。")
    parser.add_argument("--audit-date-from", default=None, help="監査履歴の開始日を YYYY-MM-DD 形式で指定します。")
    parser.add_argument("--audit-date-to", default=None, help="監査履歴の終了日を YYYY-MM-DD 形式で指定します。")
    parser.add_argument("--audit-keep-count", type=validate_audit_keep_count, default=None, help="監査履歴に残す件数を1以上で指定します。")
    parser.add_argument("--audit-keep-days", type=validate_audit_keep_days, default=None, help="監査履歴に残す日数を1以上で指定します。")
    parser.add_argument("--export-audit-summary", type=Path, default=None, help="監査履歴の要約CSVを書き出します。")
    parser.add_argument("--export-audit-monthly-summary", type=Path, default=None, help="監査履歴の月別要約CSVを書き出します。")
    parser.add_argument("--backup-audit-log", action="store_true", help="監査履歴JSONLをzip圧縮でバックアップします。")
    parser.add_argument("--prune-audit-log", action="store_true", help="監査履歴を保持設定に従って整理します。")
    parser.add_argument("--repair-legacy-text", action="store_true", help="GUI状態・監査履歴・プリセット内の旧文字化け文言を修復します。")
    parser.add_argument("--preview-repair-legacy-text", action="store_true", help="旧文字化け文言の修復候補を表示します。")
    return parser.parse_args()


def validate_config_mapping(config: dict[str, Any], context: str) -> None:
    for key, value in config.items():
        if key == "presets":
            continue
        normalized_key = key.replace("-", "_")
        if normalized_key not in CONFIG_KEYS:
            raise ValueError(f"{context} に未対応の設定キーがあります: {key}")
        if not isinstance(value, CONFIG_KEYS[normalized_key]):
            expected = ", ".join(type_.__name__ for type_ in CONFIG_KEYS[normalized_key])
            raise ValueError(f"{context} の設定値の型が不正です: {key}（期待: {expected}）")
        if normalized_key == "style":
            validate_style_config(value, context)


def validate_style_config(style: dict[str, Any], context: str) -> None:
    for key, value in style.items():
        normalized_key = key.replace("-", "_")
        if normalized_key not in STYLE_CONFIG_KEYS:
            raise ValueError(f"{context} の style に未対応の設定キーがあります: {key}")
        if not isinstance(value, STYLE_CONFIG_KEYS[normalized_key]):
            expected = ", ".join(type_.__name__ for type_ in STYLE_CONFIG_KEYS[normalized_key])
            raise ValueError(f"{context} の style.{key} の型が不正です（期待: {expected}）")


def read_config_document(config_file: Path) -> dict[str, Any]:
    try:
        config = json.loads(config_file.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"設定ファイルのJSON形式が不正です: {config_file}") from exc
    if not isinstance(config, dict):
        raise ValueError("設定ファイルのルートはJSONオブジェクトにしてください。")
    validate_config_mapping(config, "設定ファイル")
    return config


def list_presets(config_file: Path) -> tuple[str, ...]:
    config = read_config_document(config_file)
    presets = config.get("presets", {})
    if not isinstance(presets, dict):
        raise ValueError("presets はJSONオブジェクトにしてください。")
    return tuple(sorted(presets))


def load_config(config_file: Path | None, preset: str | None) -> dict[str, Any]:
    if config_file is None:
        return {}
    config = read_config_document(config_file)
    base_config = {key: value for key, value in config.items() if key != "presets"}
    if preset:
        presets = config.get("presets", {})
        if not isinstance(presets, dict) or preset not in presets:
            raise ValueError(f"指定された preset が見つかりません: {preset}")
        preset_config = presets[preset]
        if not isinstance(preset_config, dict):
            raise ValueError(f"preset はJSONオブジェクトにしてください: {preset}")
        validate_config_mapping(preset_config, f"preset {preset}")
        base_config.update(preset_config)
    return base_config


def load_config_legacy(config_file: Path | None, preset: str | None) -> dict[str, Any]:
    if config_file is None:
        return {}
    try:
        config = json.loads(config_file.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"設定ファイルのJSON形式が不正です: {config_file}") from exc
    if not isinstance(config, dict):
        raise ValueError("設定ファイルのルートはJSONオブジェクトにしてください。")

    validate_config_mapping(config, "設定ファイル")
    base_config = {key: value for key, value in config.items() if key != "presets"}
    if preset:
        presets = config.get("presets", {})
        if not isinstance(presets, dict) or preset not in presets:
            raise ValueError(f"指定された preset が見つかりません: {preset}")
        preset_config = presets[preset]
        if not isinstance(preset_config, dict):
            raise ValueError(f"preset はJSONオブジェクトにしてください: {preset}")
        validate_config_mapping(preset_config, f"preset {preset}")
        base_config.update(preset_config)
    return base_config


def get_option(args: argparse.Namespace, config: dict[str, Any], name: str, default: Any = None) -> Any:
    value = getattr(args, name, None)
    if value is not None:
        return value
    return config.get(name.replace("_", "-"), config.get(name, default))


def resolve_path_option(value: Any, default: Path) -> Path:
    if value is None:
        return default
    return value if isinstance(value, Path) else Path(value)


def build_options(args: argparse.Namespace) -> dict[str, Any]:
    config = load_config(args.config, args.preset)
    error_report = get_option(args, config, "error_report")
    summary_csv_dir = get_option(args, config, "summary_csv_dir")
    return {
        "input_dir": resolve_path_option(get_option(args, config, "input_dir"), INPUT_DIR),
        "output_dir": resolve_path_option(get_option(args, config, "output_dir"), OUTPUT_DIR),
        "month": get_option(args, config, "month"),
        "start_date": get_option(args, config, "start_date"),
        "end_date": get_option(args, config, "end_date"),
        "product": get_option(args, config, "product"),
        "category": get_option(args, config, "category"),
        "group_by": get_option(args, config, "group_by", "product"),
        "keep_reports": get_option(args, config, "keep_reports"),
        "pattern": get_option(args, config, "pattern", DEFAULT_PATTERN),
        "dry_run": bool(get_option(args, config, "dry_run", False)),
        "all_summaries": bool(get_option(args, config, "all_summaries", False)),
        "monthly_trend": bool(get_option(args, config, "monthly_trend", False)),
        "charts": bool(get_option(args, config, "charts", False)),
        "output_name": get_option(args, config, "output_name"),
        "error_report": resolve_path_option(error_report, Path("")) if error_report else None,
        "summary_csv_dir": resolve_path_option(summary_csv_dir, Path("")) if summary_csv_dir else None,
        "summary_csv_prefix": get_option(args, config, "summary_csv_prefix", "summary"),
        "notify": bool(get_option(args, config, "notify", False)),
        "notify_webhook_url": get_option(args, config, "notify_webhook_url"),
        "column_aliases": get_option(args, config, "column_aliases", {}),
        "warning_amount_threshold": get_option(args, config, "warning_amount_threshold", 1_000_000),
        "style_config": get_option(args, config, "style", {}),
        "audit_keep_count": get_option(args, config, "audit_keep_count"),
        "audit_keep_days": get_option(args, config, "audit_keep_days"),
    }


def check_environment(
    input_dir: Path = INPUT_DIR,
    output_dir: Path = OUTPUT_DIR,
    requirements_file: Path = BASE_DIR / "requirements.txt",
) -> list[str]:
    issues: list[str] = []
    if not requirements_file.exists():
        issues.append(f"requirements.txt が見つかりません: {requirements_file}")
    for module_name in REQUIRED_MODULES:
        if importlib.util.find_spec(module_name) is None:
            issues.append(f"必要なPythonモジュールが見つかりません: {module_name}")
    if not input_dir.exists():
        issues.append(f"入力パスが見つかりません: {input_dir}")
    elif not input_dir.is_dir() and input_dir.suffix.lower() not in {".csv", ".xlsx", ".xls"}:
        issues.append(f"入力パスが対応ファイルまたはフォルダではありません: {input_dir}")
    if not output_dir.exists():
        issues.append(f"出力フォルダが見つかりません: {output_dir}")
    elif not output_dir.is_dir():
        issues.append(f"出力パスがフォルダではありません: {output_dir}")
    return issues


def get_recovery_hint(exc: Exception) -> str:
    message = str(exc)
    if isinstance(exc, DataValidationError):
        return (
            "CSVまたはExcelの該当行を修正してください。"
            "--error-report を指定すると、エラー一覧CSVに修正方法を出力できます。"
        )
    if "入力フォルダが見つかりません" in message or "読み込み対象の売上データが見つかりません" in message:
        return (
            "入力フォルダ、data/input内のファイル、読み込みファイル名パターンを確認してください。"
            "例: sales_*.csv / sales_*.xlsx"
        )
    if "同時に指定できません" in message:
        return "対象月、または開始日・終了日のどちらか一方だけを指定してください。"
    if "開始日が終了日より後" in message:
        return "開始日を終了日以前にしてください。例: 開始日 2026-04-01 / 終了日 2026-04-30"
    if "対象月に一致する売上データ" in message or "指定条件に一致する売上データ" in message:
        return "対象月、日付列、商品名、カテゴリ名、読み込みファイル名パターンを確認してください。"
    if "古いExcel形式" in message or "xlrd" in message:
        return "Excelで .xlsx 形式に保存し直すか、requirements-optional.txt の内容をインストールしてください。"
    if "設定ファイル" in message or "preset" in message:
        return "config.example.json を参考に、設定キー、型、preset名を確認してください。"
    return "ログファイルの詳細を確認し、入力CSV/Excelと実行オプションを見直してください。"


def format_user_error_message(exc: Exception) -> str:
    if isinstance(exc, DataValidationError):
        lines = [
            "売上データの内容に修正が必要です。",
            "",
            "確認してください:",
            "- CSV / Excelの列名が正しいか",
            "- 日付、数量、単価、金額に不正な値がないか",
            "- エラー一覧CSVまたは画面の修正方法を確認してください",
            "",
            "詳細:",
        ]
        lines.extend(format_validation_issues(exc.issues, limit=5))
        return "\n".join(lines)
    return "\n".join(
        [
            "レポート作成中にエラーが発生しました。",
            "",
            "確認してください:",
            "- 入力フォルダに売上データがあるか",
            "- 対象月が正しいか",
            "- CSV / Excelの列名が正しいか",
            "- 日付、数量、単価、金額に不正な値がないか",
            "- 読み込みファイル名パターンが正しいか",
            "",
            f"詳細: {exc}",
            f"対処: {get_recovery_hint(exc)}",
        ]
    )


def validate_month(month: str | None) -> str | None:
    if month is None:
        return None
    try:
        datetime.strptime(month, "%Y-%m")
    except ValueError as exc:
        raise ValueError("--month は YYYY-MM 形式で指定してください。例: 2026-04") from exc
    return month


def validate_options(month: str | None, start_date: str | None, end_date: str | None, group_by: str) -> None:
    validate_month(month)
    start = validate_date_option(start_date, "--start-date")
    end = validate_date_option(end_date, "--end-date")
    if month and (start_date or end_date):
        raise ValueError("対象月と開始日・終了日は同時に指定できません。対象月または期間指定のどちらか一方を使ってください。")
    if start is not None and end is not None and start > end:
        raise ValueError(
            "開始日が終了日より後になっています。\n\n"
            "開始日と終了日を確認してください。\n\n"
            "修正例:\n"
            "開始日: 2026-04-01\n"
            "終了日: 2026-04-30"
        )
    if group_by not in GROUP_BY_COLUMNS:
        raise ValueError(f"--group-by は product または category を指定してください: {group_by}")


def previous_month_text(month: str) -> str:
    target = datetime.strptime(month, "%Y-%m")
    if target.month == 1:
        return f"{target.year - 1}-12"
    return f"{target.year}-{target.month - 1:02d}"


def build_preview(
    input_dir: Path,
    month: str | None,
    group_by: str,
    pattern: str = DEFAULT_PATTERN,
    *,
    start_date: str | None = None,
    end_date: str | None = None,
    product: str | None = None,
    category: str | None = None,
    limit: int = 10,
    column_aliases: dict | None = None,
) -> PreviewResult:
    validate_options(month, start_date, end_date, group_by)
    if limit < 1:
        raise ValueError("--preview-limit は1以上で指定してください。")
    merged_df = read_sales_files(input_dir, pattern)
    validated_df = validate_data(merged_df, column_aliases)
    target_df = filter_data(
        validated_df,
        month=month,
        start_date=start_date,
        end_date=end_date,
        product=product,
        category=category,
    )
    preview_df = target_df.head(limit).copy()
    columns = tuple(str(column) for column in preview_df.columns)
    rows = tuple(tuple("" if value is None else str(value) for value in row) for row in preview_df.to_numpy())
    return PreviewResult(
        columns=columns,
        rows=rows,
        total_count=len(target_df),
        source_files=tuple(str(name) for name in merged_df.attrs.get("source_files", ())),
    )


def build_summary_preview(
    input_dir: Path,
    month: str | None,
    group_by: str,
    pattern: str = DEFAULT_PATTERN,
    *,
    all_summaries: bool = False,
    start_date: str | None = None,
    end_date: str | None = None,
    product: str | None = None,
    category: str | None = None,
    limit: int = 20,
    column_aliases: dict | None = None,
) -> SummaryPreviewResult:
    validate_options(month, start_date, end_date, group_by)
    if limit < 1:
        raise ValueError("--preview-limit は1以上で指定してください。")
    merged_df = read_sales_files(input_dir, pattern)
    validated_df = validate_data(merged_df, column_aliases)
    target_df = filter_data(
        validated_df,
        month=month,
        start_date=start_date,
        end_date=end_date,
        product=product,
        category=category,
    )
    summary_dfs = create_summaries(target_df, group_by, all_summaries)
    summaries: dict[str, PreviewResult] = {}
    for name, summary_df in summary_dfs.items():
        preview_df = summary_df.head(limit).copy()
        columns = tuple(str(column) for column in preview_df.columns)
        rows = tuple(tuple("" if value is None else str(value) for value in row) for row in preview_df.to_numpy())
        summaries[name] = PreviewResult(columns=columns, rows=rows, total_count=len(summary_df))
    daily_df = create_daily_trend(target_df)
    preview_daily_df = daily_df.head(limit).copy()
    summaries["daily"] = PreviewResult(
        columns=tuple(str(column) for column in preview_daily_df.columns),
        rows=tuple(tuple("" if value is None else str(value) for value in row) for row in preview_daily_df.to_numpy()),
        total_count=len(daily_df),
    )
    return SummaryPreviewResult(summaries=summaries, detail_count=len(target_df))


def inspect_report_warnings(detail_df, *, high_amount_threshold: float = 1_000_000) -> tuple[str, ...]:
    warnings: list[str] = []
    if len(detail_df) == 0:
        warnings.append("出力対象の明細が0件です。条件または入力CSV / Excelを確認してください。")
        return tuple(warnings)

    if "category" in detail_df.columns:
        missing_category_count = int(detail_df["category"].fillna("").astype(str).str.strip().eq("").sum())
        if missing_category_count:
            warnings.append(f"カテゴリが空の明細が{missing_category_count}件あります。カテゴリ別集計を確認してください。")

    if "amount" in detail_df.columns:
        high_amount_count = int((detail_df["amount"] >= high_amount_threshold).sum())
        if high_amount_count:
            warnings.append(f"金額が{high_amount_threshold:,.0f}以上の明細が{high_amount_count}件あります。数量または単価を確認してください。")

    return tuple(warnings)


def print_preview(preview: PreviewResult) -> None:
    print(f"プレビュー件数: {len(preview.rows)} / {preview.total_count}")
    if preview.source_files:
        print("読み込みファイル:")
        for source_file in preview.source_files:
            print(f"- {source_file}")
    if not preview.columns:
        return
    print("\t".join(preview.columns))
    for row in preview.rows:
        print("\t".join(row))


def write_input_template(output_file: Path) -> Path:
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with output_file.open("w", newline="", encoding="utf-8-sig") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(["date", "product", "category", "quantity", "unit_price"])
        writer.writerow(["2026-04-01", "sample product", "sample category", "1", "100"])
    return output_file


def build_column_mapping_preview(input_dir: Path, pattern: str = DEFAULT_PATTERN) -> dict[str, Any]:
    columns = read_sales_columns(input_dir, pattern)
    aliases = infer_column_aliases(columns)
    missing_labels = missing_required_column_labels(columns, aliases)
    return {"columns": columns, "column_aliases": aliases, "missing_required_labels": tuple(missing_labels)}


def post_webhook_notification(webhook_url: str, message: str, status: str) -> None:
    payload = json.dumps({"text": message, "status": status}, ensure_ascii=False).encode("utf-8")
    webhook_request = request.Request(
        webhook_url,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with request.urlopen(webhook_request, timeout=10):
        pass


def notify_completion(message: str, webhook_url: str | None = None, status: str = "success") -> None:
    if sys.platform == "win32":
        try:
            import winsound

            winsound.MessageBeep(winsound.MB_ICONASTERISK)
        except Exception:
            logging.exception("通知音の再生に失敗しました。")
    if webhook_url:
        try:
            post_webhook_notification(webhook_url, message, status)
        except Exception:
            logging.exception("Webhook通知に失敗しました。")
    print(f"通知: {message}")


def default_error_report_path(output_dir: Path) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return output_dir / f"validation_errors_{timestamp}.csv"


def append_audit_log(
    *,
    status: str,
    options: dict[str, Any],
    result: ReportRunResult | None = None,
    error: str | None = None,
    audit_log_file: Path = AUDIT_LOG_FILE,
    keep_count: int | None = None,
    keep_days: int | None = None,
) -> Path:
    audit_log_file.parent.mkdir(parents=True, exist_ok=True)
    record = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "status": status,
        "input_dir": str(options.get("input_dir", "")),
        "output_dir": str(options.get("output_dir", "")),
        "month": options.get("month"),
        "start_date": options.get("start_date"),
        "end_date": options.get("end_date"),
        "product": options.get("product"),
        "category": options.get("category"),
        "group_by": options.get("group_by"),
        "pattern": options.get("pattern"),
        "all_summaries": options.get("all_summaries"),
        "monthly_trend": options.get("monthly_trend"),
        "charts": options.get("charts"),
        "output_name": options.get("output_name"),
        "summary_csv_dir": str(options.get("summary_csv_dir")) if options.get("summary_csv_dir") else None,
        "summary_csv_prefix": options.get("summary_csv_prefix"),
        "notify": options.get("notify"),
        "notify_webhook_url": options.get("notify_webhook_url"),
        "column_aliases": options.get("column_aliases", {}),
        "warning_amount_threshold": options.get("warning_amount_threshold"),
        "style": options.get("style_config", options.get("style", {})),
        "detail_count": result.detail_count if result else None,
        "summary_count": result.summary_count if result else None,
        "dry_run": result.dry_run if result else options.get("dry_run"),
        "output_file": str(result.output_file) if result and result.output_file else None,
        "summary_csv_files": [str(path) for path in result.summary_csv_files] if result else [],
        "warnings": list(result.warnings) if result else [],
        "error": error,
    }
    with audit_log_file.open("a", encoding="utf-8") as log_file:
        log_file.write(json.dumps(record, ensure_ascii=False) + "\n")
    prune_audit_log(audit_log_file, keep_count=keep_count, keep_days=keep_days)
    return audit_log_file


def read_audit_log(audit_log_file: Path = AUDIT_LOG_FILE, limit: int = 100) -> list[dict[str, Any]]:
    if limit < 1:
        raise ValueError("limit must be 1 or greater")
    if not audit_log_file.exists():
        return []

    records: list[dict[str, Any]] = []
    for line in audit_log_file.read_text(encoding="utf-8", errors="replace").splitlines():
        if not line.strip():
            continue
        try:
            record = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(record, dict):
            normalized = normalize_legacy_value(record)
            if isinstance(normalized, dict):
                records.append(normalized)
    return records[-limit:]


def _parse_audit_filter_date(value: str | None, option_name: str) -> date | None:
    if value is None:
        return None
    value = value.strip()
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError(f"{option_name} は YYYY-MM-DD 形式で指定してください。") from exc


def filter_audit_records(
    records: list[dict[str, Any]],
    *,
    filter_text: str = "",
    filter_status: str = "all",
    date_from: str | None = None,
    date_to: str | None = None,
) -> list[dict[str, Any]]:
    normalized_text = filter_text.strip().lower()
    from_date = _parse_audit_filter_date(date_from, "--audit-date-from")
    to_date = _parse_audit_filter_date(date_to, "--audit-date-to")
    if from_date and to_date and from_date > to_date:
        raise ValueError("--audit-date-from は --audit-date-to 以下で指定してください。")

    filtered: list[dict[str, Any]] = []
    for record in records:
        if filter_status and filter_status != "all" and record.get("status") != filter_status:
            continue
        timestamp = str(record.get("timestamp") or "")
        if from_date or to_date:
            try:
                record_date = datetime.fromisoformat(timestamp).date()
            except ValueError:
                continue
            if from_date and record_date < from_date:
                continue
            if to_date and record_date > to_date:
                continue
        searchable = " ".join(
            str(record.get(key, ""))
            for key in ("timestamp", "status", "month", "start_date", "end_date", "output_file", "error")
        ).lower()
        if normalized_text and normalized_text not in searchable:
            continue
        filtered.append(record)
    return filtered


def export_audit_log_csv(records: list[dict[str, Any]], output_file: Path) -> Path:
    fieldnames = [
        "timestamp",
        "status",
        "period",
        "month",
        "start_date",
        "end_date",
        "detail_count",
        "summary_count",
        "dry_run",
        "warning_count",
        "warnings",
        "input_dir",
        "output_dir",
        "product",
        "category",
        "group_by",
        "pattern",
        "all_summaries",
        "monthly_trend",
        "charts",
        "output_name",
        "summary_csv_dir",
        "summary_csv_prefix",
        "notify",
        "notify_webhook_url",
        "output_file",
        "summary_csv_files",
        "column_aliases",
        "warning_amount_threshold",
        "style",
        "error",
    ]
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with output_file.open("w", newline="", encoding="utf-8-sig") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
        writer.writeheader()
        for raw_record in records:
            record = normalize_legacy_value(raw_record)
            if not isinstance(record, dict):
                continue
            warnings = [str(item) for item in record.get("warnings") or []]
            summary_csv_files = [str(item) for item in record.get("summary_csv_files") or []]
            writer.writerow(
                {
                    "timestamp": record.get("timestamp") or "",
                    "status": record.get("status") or "",
                    "period": record.get("month") or f"{record.get('start_date') or ''} - {record.get('end_date') or ''}".strip(),
                    "month": record.get("month") or "",
                    "start_date": record.get("start_date") or "",
                    "end_date": record.get("end_date") or "",
                    "detail_count": record.get("detail_count") or "",
                    "summary_count": record.get("summary_count") or "",
                    "dry_run": record.get("dry_run"),
                    "warning_count": len(warnings),
                    "warnings": " | ".join(warnings),
                    "input_dir": record.get("input_dir") or "",
                    "output_dir": record.get("output_dir") or "",
                    "product": record.get("product") or "",
                    "category": record.get("category") or "",
                    "group_by": record.get("group_by") or "",
                    "pattern": record.get("pattern") or "",
                    "all_summaries": record.get("all_summaries"),
                    "monthly_trend": record.get("monthly_trend"),
                    "charts": record.get("charts"),
                    "output_name": record.get("output_name") or "",
                    "summary_csv_dir": record.get("summary_csv_dir") or "",
                    "summary_csv_prefix": record.get("summary_csv_prefix") or "",
                    "notify": record.get("notify"),
                    "notify_webhook_url": record.get("notify_webhook_url") or "",
                    "output_file": record.get("output_file") or "",
                    "summary_csv_files": " | ".join(summary_csv_files),
                    "column_aliases": json.dumps(record.get("column_aliases") or {}, ensure_ascii=False, sort_keys=True),
                    "warning_amount_threshold": record.get("warning_amount_threshold") or "",
                    "style": json.dumps(record.get("style") or {}, ensure_ascii=False, sort_keys=True),
                    "error": record.get("error") or "",
                }
            )
    return output_file


def export_audit_summary_csv(records: list[dict[str, Any]], output_file: Path) -> Path:
    normalized_records = [record for record in (normalize_legacy_value(item) for item in records) if isinstance(record, dict)]
    timestamps = [str(record.get("timestamp") or "") for record in normalized_records if record.get("timestamp")]
    warning_total = sum(len(record.get("warnings") or []) for record in normalized_records)
    anomaly = detect_audit_anomalies(normalized_records)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with output_file.open("w", newline="", encoding="utf-8-sig") as csv_file:
        writer = csv.DictWriter(
            csv_file,
            fieldnames=[
                "total_runs",
                "success_count",
                "validation_error_count",
                "error_count",
                "other_status_count",
                "warning_total",
                "dry_run_count",
                "earliest_timestamp",
                "latest_timestamp",
                "failure_rate",
                "consecutive_missing_days",
                "last_run_date",
                "alerts",
            ],
        )
        writer.writeheader()
        writer.writerow(
            {
                "total_runs": len(normalized_records),
                "success_count": sum(1 for record in normalized_records if record.get("status") == "success"),
                "validation_error_count": sum(1 for record in normalized_records if record.get("status") == "validation_error"),
                "error_count": sum(1 for record in normalized_records if record.get("status") == "error"),
                "other_status_count": sum(
                    1 for record in normalized_records if record.get("status") not in {"success", "validation_error", "error"}
                ),
                "warning_total": warning_total,
                "dry_run_count": sum(1 for record in normalized_records if record.get("dry_run")),
                "earliest_timestamp": min(timestamps) if timestamps else "",
                "latest_timestamp": max(timestamps) if timestamps else "",
                "failure_rate": f"{anomaly.failure_rate:.3f}",
                "consecutive_missing_days": anomaly.consecutive_missing_days,
                "last_run_date": anomaly.last_run_date,
                "alerts": " | ".join(anomaly.alerts),
            }
        )
    return output_file


def export_audit_monthly_summary_csv(records: list[dict[str, Any]], output_file: Path) -> Path:
    monthly: dict[str, dict[str, int]] = {}
    for raw_record in records:
        record = normalize_legacy_value(raw_record)
        if not isinstance(record, dict):
            continue
        timestamp = str(record.get("timestamp") or "")
        month = timestamp[:7] if len(timestamp) >= 7 else "unknown"
        bucket = monthly.setdefault(
            month,
            {
                "total_runs": 0,
                "success_count": 0,
                "validation_error_count": 0,
                "error_count": 0,
                "other_status_count": 0,
                "warning_total": 0,
                "dry_run_count": 0,
            },
        )
        bucket["total_runs"] += 1
        status = str(record.get("status") or "")
        status_key = f"{status}_count"
        if status_key in bucket:
            bucket[status_key] += 1
        else:
            bucket["other_status_count"] += 1
        bucket["warning_total"] += len(record.get("warnings") or [])
        if record.get("dry_run"):
            bucket["dry_run_count"] += 1

    output_file.parent.mkdir(parents=True, exist_ok=True)
    with output_file.open("w", newline="", encoding="utf-8-sig") as csv_file:
        writer = csv.DictWriter(
            csv_file,
            fieldnames=[
                "month",
                "total_runs",
                "success_count",
                "validation_error_count",
                "error_count",
                "other_status_count",
                "warning_total",
                "dry_run_count",
                "alerts",
            ],
        )
        writer.writeheader()
        for month in sorted(monthly):
            month_records = [normalize_legacy_value(record) for record in records if str(record.get("timestamp") or "").startswith(month)]
            anomaly = detect_audit_anomalies([record for record in month_records if isinstance(record, dict)], today=_month_end_date(month))
            writer.writerow({"month": month, **monthly[month], "alerts": " | ".join(anomaly.alerts)})
    return output_file


def detect_audit_anomalies(records: list[dict[str, Any]], *, today: date | None = None) -> AuditAnomalySummary:
    normalized_records = [record for record in (normalize_legacy_value(item) for item in records) if isinstance(record, dict)]
    total_runs = len(normalized_records)
    success_count = sum(1 for record in normalized_records if record.get("status") == "success")
    failure_count = sum(1 for record in normalized_records if record.get("status") in {"validation_error", "error"})
    warning_total = sum(len(record.get("warnings") or []) for record in normalized_records)
    run_dates = [_safe_timestamp_date(str(record.get("timestamp") or "")) for record in normalized_records]
    run_dates = [item for item in run_dates if item is not None]
    reference_date = today or datetime.now().date()
    last_run = max(run_dates) if run_dates else None
    missing_days = (reference_date - last_run).days if last_run else 0
    failure_rate = (failure_count / total_runs) if total_runs else 0.0
    alerts: list[str] = []
    if total_runs and failure_rate >= 0.3:
        alerts.append(f"失敗率が高めです ({failure_rate:.0%})")
    if missing_days >= 3:
        alerts.append(f"{missing_days}日間実行がありません")
    if warning_total >= max(3, total_runs):
        alerts.append(f"警告件数が多めです ({warning_total}件)")
    return AuditAnomalySummary(
        total_runs=total_runs,
        success_count=success_count,
        failure_count=failure_count,
        warning_total=warning_total,
        last_run_date=last_run.isoformat() if last_run else "",
        consecutive_missing_days=missing_days,
        failure_rate=failure_rate,
        alerts=tuple(alerts),
    )


def prune_audit_log(
    audit_log_file: Path = AUDIT_LOG_FILE,
    *,
    keep_count: int | None = None,
    keep_days: int | None = None,
) -> int:
    if keep_count is None and keep_days is None:
        return 0
    if keep_count is not None and keep_count < 1:
        raise ValueError("keep_count は1以上で指定してください。")
    if keep_days is not None and keep_days < 1:
        raise ValueError("keep_days は1以上で指定してください。")
    if not audit_log_file.exists():
        return 0

    entries: list[tuple[str, dict[str, Any]]] = []
    for line in audit_log_file.read_text(encoding="utf-8", errors="replace").splitlines():
        if not line.strip():
            continue
        try:
            record = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(record, dict):
            entries.append((line, record))

    filtered = entries
    if keep_days is not None:
        cutoff = datetime.now().date().toordinal() - (keep_days - 1)
        filtered = [
            (line, record)
            for line, record in filtered
            if (
                isinstance(record.get("timestamp"), str)
                and _safe_timestamp_date(record["timestamp"]) is not None
                and _safe_timestamp_date(record["timestamp"]).toordinal() >= cutoff
            )
        ]
    if keep_count is not None and len(filtered) > keep_count:
        filtered = filtered[-keep_count:]

    if len(filtered) == len(entries):
        return 0
    audit_log_file.write_text("\n".join(line for line, _record in filtered) + ("\n" if filtered else ""), encoding="utf-8")
    return len(entries) - len(filtered)


def _safe_timestamp_date(timestamp: str) -> date | None:
    try:
        return datetime.fromisoformat(timestamp).date()
    except ValueError:
        return None


def _month_end_date(month_text: str) -> date:
    year = int(month_text[:4])
    month = int(month_text[5:7])
    if month == 12:
        return date(year, month, 31)
    next_month = date(year + (month // 12), (month % 12) + 1, 1)
    return date.fromordinal(next_month.toordinal() - 1)


def backup_audit_log(audit_log_file: Path = AUDIT_LOG_FILE, backup_dir: Path | None = None) -> Path | None:
    if not audit_log_file.exists():
        return None
    target_dir = backup_dir or audit_log_file.parent
    target_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_path = target_dir / f"{audit_log_file.stem}_{timestamp}.zip"
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.write(audit_log_file, arcname=audit_log_file.name)
    return zip_path


def _backup_path(path: Path) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return path.with_name(f"{path.name}.{timestamp}{LEGACY_BACKUP_SUFFIX}")


def _normalize_json_file(path: Path) -> bool:
    if not path.exists():
        return False
    original = path.read_text(encoding="utf-8", errors="replace")
    data = json.loads(original)
    normalized = normalize_legacy_value(data)
    if normalized == data:
        return False
    normalized_text = json.dumps(normalized, ensure_ascii=False, indent=2)
    backup_path = _backup_path(path)
    path.replace(backup_path)
    path.write_text(normalized_text + "\n", encoding="utf-8")
    return True


def _normalize_jsonl_file(path: Path) -> bool:
    if not path.exists():
        return False
    original = path.read_text(encoding="utf-8", errors="replace")
    normalized_lines: list[str] = []
    changed = False
    for line in original.splitlines():
        if not line.strip():
            continue
        try:
            data = json.loads(line)
        except json.JSONDecodeError:
            normalized_lines.append(line)
            continue
        normalized = normalize_legacy_value(data)
        normalized_line = json.dumps(normalized, ensure_ascii=False)
        if normalized != data:
            changed = True
        normalized_lines.append(normalized_line)
    if not changed:
        return False
    backup_path = _backup_path(path)
    path.replace(backup_path)
    path.write_text("\n".join(normalized_lines) + "\n", encoding="utf-8")
    return True


def repair_legacy_text_files(
    gui_state_file: Path | None = None,
    audit_log_file: Path | None = None,
    alias_preset_file: Path | None = None,
) -> dict[str, bool]:
    gui_state_file = gui_state_file or LOG_DIR / "gui_state.json"
    audit_log_file = audit_log_file or AUDIT_LOG_FILE
    alias_preset_file = alias_preset_file or ALIAS_PRESET_FILE
    results = {
        str(gui_state_file): False,
        str(audit_log_file): False,
        str(alias_preset_file): False,
    }
    try:
        results[str(gui_state_file)] = _normalize_json_file(gui_state_file)
    except FileNotFoundError:
        pass
    try:
        results[str(audit_log_file)] = _normalize_jsonl_file(audit_log_file)
    except FileNotFoundError:
        pass
    try:
        results[str(alias_preset_file)] = _normalize_json_file(alias_preset_file)
    except FileNotFoundError:
        pass
    return results


def format_legacy_text_preview(changes_by_file: dict[str, list[LegacyTextChange]]) -> str:
    lines: list[str] = []
    for file_path, changes in changes_by_file.items():
        if not changes:
            continue
        lines.append(file_path)
        for change in changes[:20]:
            lines.append(f"- {change.location}: {change.original} -> {change.normalized}")
        if len(changes) > 20:
            lines.append(f"- ... {len(changes) - 20}件省略")
    return "\n".join(lines)


def read_column_alias_presets(preset_file: Path = ALIAS_PRESET_FILE) -> dict[str, dict[str, str]]:
    if not preset_file.exists():
        return {}
    try:
        presets = json.loads(preset_file.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"列名マッピングプリセットのJSON形式が不正です: {preset_file}") from exc
    if not isinstance(presets, dict):
        raise ValueError("列名マッピングプリセットはJSONオブジェクトで指定してください。")

    normalized: dict[str, dict[str, str]] = {}
    for name, aliases in presets.items():
        if not isinstance(aliases, dict):
            raise ValueError(f"列名マッピングプリセットはJSONオブジェクトで指定してください: {name}")
        normalized[str(name)] = {str(source): str(target) for source, target in aliases.items()}
    return normalized


def write_column_alias_presets(presets: dict[str, dict[str, str]], preset_file: Path = ALIAS_PRESET_FILE) -> Path:
    normalized: dict[str, dict[str, str]] = {}
    for name, aliases in presets.items():
        if not isinstance(aliases, dict):
            raise ValueError(f"列名マッピングプリセットはJSONオブジェクトで指定してください: {name}")
        normalized[str(name)] = {str(source): str(target) for source, target in aliases.items()}
    preset_file.parent.mkdir(parents=True, exist_ok=True)
    preset_file.write_text(json.dumps(normalized, ensure_ascii=False, indent=2), encoding="utf-8")
    return preset_file


def run_report(
    input_dir: Path,
    output_dir: Path,
    month: str | None,
    group_by: str,
    keep_reports: int | None,
    pattern: str = DEFAULT_PATTERN,
    *,
    dry_run: bool = False,
    all_summaries: bool = False,
    start_date: str | None = None,
    end_date: str | None = None,
    product: str | None = None,
    category: str | None = None,
    monthly_trend: bool = False,
    charts: bool = False,
    output_name: str | None = None,
    summary_csv_dir: Path | None = None,
    summary_csv_prefix: str = "summary",
    style_config: dict | None = None,
    column_aliases: dict | None = None,
    warning_amount_threshold: float = 1_000_000,
) -> ReportRunResult:
    validate_options(month, start_date, end_date, group_by)
    merged_df = read_sales_files(input_dir, pattern)
    validated_df = validate_data(merged_df, column_aliases)
    target_df = filter_data(
        validated_df,
        month=month,
        start_date=start_date,
        end_date=end_date,
        product=product,
        category=category,
    )
    summary_dfs = create_summaries(target_df, group_by, all_summaries)
    daily_trend_df = create_daily_trend(target_df)
    monthly_trend_df = create_monthly_trend(target_df) if monthly_trend else None
    uncategorized_df = create_uncategorized_rows(target_df)
    validation_error_df = create_validation_error_rows()
    uncategorized_count = len(uncategorized_df)
    error_count = len(validation_error_df)
    summary_month = month or target_df["date"].min().strftime("%Y-%m")
    previous_month = previous_month_text(summary_month)
    previous_month_df = validated_df[validated_df["date"].dt.strftime("%Y-%m") == previous_month].copy()
    if product:
        previous_month_df = previous_month_df[previous_month_df["product"] == product].copy()
    if category and "category" in previous_month_df.columns:
        previous_month_df = previous_month_df[previous_month_df["category"] == category].copy()
    month_end_summary_df = create_month_end_summary(
        target_df,
        summary_month,
        uncategorized_count=uncategorized_count,
        error_count=error_count,
    )
    previous_month_comparison_df = create_previous_month_comparison(target_df, previous_month_df, summary_month)
    summary_count = sum(len(summary_df) for summary_df in summary_dfs.values())
    warnings = inspect_report_warnings(target_df, high_amount_threshold=warning_amount_threshold)
    total_amount = float(target_df["amount"].sum())
    total_quantity = float(target_df["quantity"].sum())
    average_unit_price = total_amount / total_quantity if total_quantity else 0
    target_days = int(target_df["date"].dt.normalize().nunique())
    product_count = int(target_df["product"].nunique()) if "product" in target_df.columns else 0
    category_count = int(target_df["category"].nunique()) if "category" in target_df.columns else 0
    previous_month_amount = float(previous_month_df["amount"].sum()) if not previous_month_df.empty else 0
    previous_month_change_rate: float | str = (
        (total_amount - previous_month_amount) / previous_month_amount if previous_month_amount else "比較不可"
    )

    if dry_run:
        logging.info("dry-run完了: 明細%s件、集計%s件", len(target_df), summary_count)
        return ReportRunResult(
            None,
            len(target_df),
            summary_count,
            dry_run=True,
            warnings=warnings,
            total_amount=total_amount,
            target_days=target_days,
            product_count=product_count,
            category_count=category_count,
            total_quantity=total_quantity,
            average_unit_price=average_unit_price,
            previous_month_amount=previous_month_amount,
            previous_month_change_rate=previous_month_change_rate,
            uncategorized_count=uncategorized_count,
            error_count=error_count,
        )

    summary_csv_files = tuple(write_summary_csvs(summary_dfs, summary_csv_dir, summary_csv_prefix)) if summary_csv_dir else ()
    output_file = save_to_excel(
        target_df,
        summary_dfs,
        output_dir,
        month,
        group_by,
        input_dir=input_dir,
        pattern=pattern,
        all_summaries=all_summaries,
        start_date=start_date,
        end_date=end_date,
        product=product,
        category=category,
        daily_trend_df=daily_trend_df,
        monthly_trend_df=monthly_trend_df,
        month_end_summary_df=month_end_summary_df,
        previous_month_comparison_df=previous_month_comparison_df,
        uncategorized_df=uncategorized_df,
        validation_error_df=validation_error_df,
        output_name=output_name,
        charts=charts,
        summary_csv_dir=summary_csv_dir,
        style_config=style_config,
    )
    deleted_reports = cleanup_old_reports(output_dir, keep_reports)
    logging.info("Excel出力完了: %s", output_file.name)
    for deleted_report in deleted_reports:
        logging.info("古いレポートを削除: %s", deleted_report.name)
    return ReportRunResult(
        output_file,
        len(target_df),
        summary_count,
        summary_csv_files=summary_csv_files,
        warnings=warnings,
        total_amount=total_amount,
        target_days=target_days,
        product_count=product_count,
        category_count=category_count,
        total_quantity=total_quantity,
        average_unit_price=average_unit_price,
        previous_month_amount=previous_month_amount,
        previous_month_change_rate=previous_month_change_rate,
        uncategorized_count=uncategorized_count,
        error_count=error_count,
    )


def main() -> int:
    try:
        args = parse_args()
        if args.preview_repair_legacy_text:
            changes_by_file = inspect_legacy_text_files()
            preview = format_legacy_text_preview(changes_by_file)
            if preview:
                print("旧文字化け文言の修復候補:")
                print(preview)
            else:
                print("修復候補の旧文字化け文言はありませんでした。")
            return 0
        if args.repair_legacy_text:
            repaired = repair_legacy_text_files()
            changed_paths = [path for path, changed in repaired.items() if changed]
            if changed_paths:
                print("旧文字化け文言を修復しました。")
                for path in changed_paths:
                    print(f"- {path}")
            else:
                print("修復対象の旧文字化け文言はありませんでした。")
            return 0
        if args.audit_date_from:
            _parse_audit_filter_date(args.audit_date_from, "--audit-date-from")
        if args.audit_date_to:
            _parse_audit_filter_date(args.audit_date_to, "--audit-date-to")
        options = build_options(args)
        if args.backup_audit_log:
            backup_file = backup_audit_log(AUDIT_LOG_FILE)
            if backup_file is None:
                print("バックアップ対象の監査履歴はありませんでした。")
            else:
                print(f"監査履歴バックアップ: {backup_file}")
            return 0
        if args.prune_audit_log:
            removed = prune_audit_log(
                AUDIT_LOG_FILE,
                keep_count=options.get("audit_keep_count"),
                keep_days=options.get("audit_keep_days"),
            )
            print(f"監査履歴を整理しました: {removed}件削除")
            return 0
        if args.export_audit_summary or args.export_audit_monthly_summary:
            records = filter_audit_records(
                read_audit_log(limit=1_000_000),
                date_from=args.audit_date_from,
                date_to=args.audit_date_to,
            )
            if args.export_audit_summary:
                export_file = export_audit_summary_csv(records, args.export_audit_summary)
                print(f"監査履歴要約CSV: {export_file}")
            if args.export_audit_monthly_summary:
                export_file = export_audit_monthly_summary_csv(records, args.export_audit_monthly_summary)
                print(f"監査履歴月別要約CSV: {export_file}")
            return 0
        if args.write_template:
            template_file = write_input_template(args.write_template)
            print(f"入力CSVテンプレートを作成しました: {template_file}")
            return 0
        if args.check_setup:
            issues = check_environment(input_dir=options["input_dir"], output_dir=options["output_dir"])
            if issues:
                print("初期設定チェック: NG")
                for issue in issues:
                    print(f"- {issue}")
                print("対処: requirements.txt を使って依存ライブラリをインストールし、data/input と data/output を確認してください。")
                return 1
            print("初期設定チェック: OK")
            return 0
        if args.preview:
            preview = build_preview(
                input_dir=options["input_dir"],
                month=options["month"],
                group_by=options["group_by"],
                pattern=options["pattern"],
                start_date=options["start_date"],
                end_date=options["end_date"],
                product=options["product"],
                category=options["category"],
                limit=args.preview_limit,
                column_aliases=options["column_aliases"],
            )
            print_preview(preview)
            return 0
        setup_logging()
        logging.info("処理開始")
        result = run_report(
            input_dir=options["input_dir"],
            output_dir=options["output_dir"],
            month=options["month"],
            group_by=options["group_by"],
            keep_reports=options["keep_reports"],
            pattern=options["pattern"],
            dry_run=options["dry_run"],
            all_summaries=options["all_summaries"],
            start_date=options["start_date"],
            end_date=options["end_date"],
            product=options["product"],
            category=options["category"],
            monthly_trend=options["monthly_trend"],
            charts=options["charts"],
            output_name=options["output_name"],
            summary_csv_dir=options["summary_csv_dir"],
            summary_csv_prefix=options["summary_csv_prefix"],
            style_config=options["style_config"],
            column_aliases=options["column_aliases"],
            warning_amount_threshold=options["warning_amount_threshold"],
        )
        logging.info("処理終了")
        append_audit_log(
            status="success",
            options=options,
            result=result,
            keep_count=options.get("audit_keep_count"),
            keep_days=options.get("audit_keep_days"),
        )

        if result.dry_run:
            print("検証が完了しました。Excelは出力していません。")
        else:
            print("処理が完了しました。")
        print(f"明細{result.detail_count}件、集計{result.summary_count}件を確認しました。")
        print(f"売上合計: {result.total_amount:,.0f}")
        print(f"数量合計: {result.total_quantity:,.0f}")
        print(f"平均単価: {result.average_unit_price:,.0f}")
        print(f"対象日数: {result.target_days}日")
        print(f"商品数: {result.product_count}")
        print(f"カテゴリ数: {result.category_count}")
        print(f"前月売上: {result.previous_month_amount:,.0f}")
        previous_rate = result.previous_month_change_rate
        if isinstance(previous_rate, (int, float)):
            print(f"前月比: {previous_rate:.1%}")
        else:
            print(f"前月比: {previous_rate}")
        print(f"未分類データ: {result.uncategorized_count}件")
        print(f"エラー行: {result.error_count}件")
        print(f"確認が必要なデータ: {result.uncategorized_count + result.error_count}件")
        if result.output_file is not None:
            print(f"出力ファイル: {result.output_file}")
        for summary_csv_file in result.summary_csv_files:
            print(f"集計CSV: {summary_csv_file}")
        for warning in result.warnings:
            print(f"警告: {warning}")
        if options["notify"]:
            notify_completion("レポート作成が完了しました。", options.get("notify_webhook_url"), "success")
        return 0

    except DataValidationError as exc:
        logging.exception("CSV検証エラーが発生しました。")
        print(format_user_error_message(exc))
        print(f"ログファイル: {LOG_FILE.resolve()}")
        error_report = options.get("error_report") if "options" in locals() else None
        if error_report:
            report_file = write_validation_error_report(exc, error_report)
            print(f"エラー一覧CSV: {report_file}")
        if "options" in locals():
            append_audit_log(
                status="validation_error",
                options=options,
                error=str(exc),
                keep_count=options.get("audit_keep_count"),
                keep_days=options.get("audit_keep_days"),
            )
            if options.get("notify"):
                notify_completion("レポート作成に失敗しました。", options.get("notify_webhook_url"), "validation_error")
        return 1
    except Exception as exc:
        logging.exception("エラーが発生しました。")
        if "options" in locals():
            append_audit_log(
                status="error",
                options=options,
                error=str(exc),
                keep_count=options.get("audit_keep_count"),
                keep_days=options.get("audit_keep_days"),
            )
            if options.get("notify"):
                notify_completion("レポート作成に失敗しました。", options.get("notify_webhook_url"), "error")
        print(format_user_error_message(exc))
        print(f"ログファイル: {LOG_FILE.resolve()}")
        return 1


if __name__ == "__main__":
    sys.exit(main())

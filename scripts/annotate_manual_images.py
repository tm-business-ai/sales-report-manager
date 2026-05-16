from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SOURCE_DIR = PROJECT_ROOT / "_private_materials" / "images" / "manual_source"
OUTPUT_DIR = PROJECT_ROOT / "_private_materials" / "images" / "annotated"

RED = (220, 38, 38)
WHITE = (255, 255, 255)
BLACK = (24, 24, 27)
YELLOW = (254, 249, 195)


def _font(size: int, bold: bool = False) -> ImageFont.ImageFont:
    candidates = [
        Path(r"C:\Windows\Fonts\meiryob.ttc" if bold else r"C:\Windows\Fonts\meiryo.ttc"),
        Path(r"C:\Windows\Fonts\YuGothM.ttc"),
        Path(r"C:\Windows\Fonts\msgothic.ttc"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return ImageFont.truetype(str(candidate), size)
    return ImageFont.load_default()


def _marker(draw: ImageDraw.ImageDraw, xy: tuple[int, int], number: int, label: str | None = None) -> None:
    x, y = xy
    radius = 18
    draw.ellipse((x - radius, y - radius, x + radius, y + radius), fill=RED, outline=WHITE, width=3)
    text = str(number)
    font = _font(20, bold=True)
    bbox = draw.textbbox((0, 0), text, font=font)
    draw.text((x - (bbox[2] - bbox[0]) / 2, y - (bbox[3] - bbox[1]) / 2 - 1), text, fill=WHITE, font=font)

    if label:
        label_font = _font(15)
        pad = 6
        lb = draw.textbbox((0, 0), label, font=label_font)
        lx = min(x + radius + 6, max(0, draw.im.size[0] - (lb[2] - lb[0]) - pad * 2 - 4))
        ly = max(4, y - 14)
        draw.rounded_rectangle(
            (lx, ly, lx + (lb[2] - lb[0]) + pad * 2, ly + (lb[3] - lb[1]) + pad * 2),
            radius=5,
            fill=YELLOW,
            outline=RED,
            width=2,
        )
        draw.text((lx + pad, ly + pad - 1), label, fill=BLACK, font=label_font)


def _box(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int]) -> None:
    draw.rounded_rectangle(box, radius=6, outline=RED, width=3)


ANNOTATIONS: dict[str, list[dict[str, object]]] = {
    "gui_main.png": [
        {"n": 1, "p": (0.09, 0.10), "label": "設定ファイル", "box": (0.03, 0.08, 0.45, 0.14)},
        {"n": 2, "p": (0.09, 0.18), "label": "入力フォルダ", "box": (0.03, 0.16, 0.76, 0.22)},
        {"n": 3, "p": (0.09, 0.25), "label": "出力フォルダ", "box": (0.03, 0.23, 0.76, 0.29)},
        {"n": 4, "p": (0.08, 0.34), "label": "対象月", "box": (0.03, 0.31, 0.25, 0.38)},
        {"n": 13, "p": (0.60, 0.34), "label": "集計単位", "box": (0.50, 0.31, 0.76, 0.38)},
        {"n": 14, "p": (0.08, 0.47), "label": "出力オプション", "box": (0.03, 0.42, 0.76, 0.55)},
        {"n": 21, "p": (0.16, 0.86), "label": "Excelレポート作成", "box": (0.04, 0.82, 0.30, 0.91)},
        {"n": 22, "p": (0.43, 0.86), "label": "データプレビュー", "box": (0.31, 0.82, 0.55, 0.91)},
        {"n": 23, "p": (0.68, 0.86), "label": "集計プレビュー", "box": (0.56, 0.82, 0.80, 0.91)},
    ],
    "gui_error_review.png": [
        {"n": 1, "p": (0.11, 0.13), "label": "入力データを検証", "box": (0.03, 0.09, 0.24, 0.17)},
        {"n": 3, "p": (0.40, 0.13), "label": "エラーCSV", "box": (0.27, 0.09, 0.54, 0.17)},
        {"n": 5, "p": (0.08, 0.33), "label": "エラー一覧", "box": (0.03, 0.23, 0.97, 0.74)},
        {"n": 6, "p": (0.56, 0.80), "label": "修正方法", "box": (0.03, 0.74, 0.97, 0.94)},
    ],
    "gui_report_history.png": [
        {"n": 1, "p": (0.10, 0.14), "label": "一覧を更新", "box": (0.03, 0.09, 0.21, 0.18)},
        {"n": 2, "p": (0.33, 0.14), "label": "レポートを開く", "box": (0.22, 0.09, 0.45, 0.18)},
        {"n": 4, "p": (0.73, 0.14), "label": "詳細を表示", "box": (0.63, 0.09, 0.84, 0.18)},
        {"n": 5, "p": (0.10, 0.34), "label": "作成済みレポート", "box": (0.03, 0.24, 0.97, 0.83)},
    ],
    "gui_execution_history.png": [
        {"n": 1, "p": (0.08, 0.11), "label": "履歴を更新", "box": (0.03, 0.07, 0.18, 0.15)},
        {"n": 2, "p": (0.27, 0.11), "label": "CSV出力", "box": (0.19, 0.07, 0.36, 0.15)},
        {"n": 8, "p": (0.10, 0.25), "label": "検索・絞り込み", "box": (0.03, 0.19, 0.96, 0.31)},
        {"n": 12, "p": (0.78, 0.37), "label": "条件を復元", "box": (0.64, 0.32, 0.93, 0.41)},
        {"n": 13, "p": (0.08, 0.54), "label": "実行履歴一覧", "box": (0.03, 0.43, 0.97, 0.88)},
    ],
    "gui_detail_log.png": [
        {"n": 1, "p": (0.09, 0.13), "label": "ログを更新", "box": (0.03, 0.08, 0.21, 0.17)},
        {"n": 2, "p": (0.30, 0.13), "label": "エラーのみ", "box": (0.22, 0.08, 0.41, 0.17)},
        {"n": 5, "p": (0.08, 0.37), "label": "ログ一覧", "box": (0.03, 0.23, 0.97, 0.67)},
        {"n": 11, "p": (0.11, 0.78), "label": "ログ詳細", "box": (0.03, 0.69, 0.97, 0.94)},
    ],
}


def _scale_box(size: tuple[int, int], box: tuple[float, float, float, float]) -> tuple[int, int, int, int]:
    width, height = size
    return tuple(int(value * (width if index % 2 == 0 else height)) for index, value in enumerate(box))  # type: ignore[return-value]


def annotate_image(source: Path, output: Path, annotations: list[dict[str, object]]) -> None:
    image = Image.open(source).convert("RGBA")
    draw = ImageDraw.Draw(image)
    width, height = image.size

    for item in annotations:
        if "box" in item:
            _box(draw, _scale_box((width, height), item["box"]))  # type: ignore[arg-type]

    for item in annotations:
        px, py = item["p"]  # type: ignore[misc]
        _marker(draw, (int(px * width), int(py * height)), int(item["n"]), str(item.get("label", "")))

    output.parent.mkdir(parents=True, exist_ok=True)
    image.convert("RGB").save(output, "PNG")


def main() -> int:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    missing: list[str] = []
    created: list[Path] = []

    for filename, annotations in ANNOTATIONS.items():
        source = SOURCE_DIR / filename
        output = OUTPUT_DIR / filename.replace(".png", "_annotated.png")
        if not source.exists():
            missing.append(str(source))
            continue
        annotate_image(source, output, annotations)
        created.append(output)

    if created:
        print("Created annotated images:")
        for path in created:
            print(f"  - {path}")
    else:
        print("No annotated images were created.")

    if missing:
        print("Missing source images:")
        for path in missing:
            print(f"  - {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

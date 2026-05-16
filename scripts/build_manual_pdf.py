from __future__ import annotations

import re
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer


PROJECT_ROOT = Path(__file__).resolve().parents[1]
MANUAL_MD = PROJECT_ROOT / "_private_materials" / "operation_manual.md"
OUTPUT_PDF = PROJECT_ROOT / "_private_materials" / "operation_manual.pdf"
FONT_NAME = "ManualJapanese"


def register_font() -> str:
    candidates = [
        Path(r"C:\Windows\Fonts\meiryo.ttc"),
        Path(r"C:\Windows\Fonts\YuGothR.ttc"),
        Path(r"C:\Windows\Fonts\YuGothM.ttc"),
        Path(r"C:\Windows\Fonts\msgothic.ttc"),
    ]
    for font_path in candidates:
        if font_path.exists():
            try:
                pdfmetrics.registerFont(TTFont(FONT_NAME, str(font_path)))
                return FONT_NAME
            except Exception:
                continue
    return "Helvetica"


def _styles(font_name: str) -> dict[str, ParagraphStyle]:
    base = getSampleStyleSheet()
    return {
        "h1": ParagraphStyle(
            "h1",
            parent=base["Title"],
            fontName=font_name,
            fontSize=18,
            leading=24,
            spaceAfter=8,
            wordWrap="CJK",
        ),
        "h2": ParagraphStyle(
            "h2",
            parent=base["Heading2"],
            fontName=font_name,
            fontSize=13,
            leading=18,
            spaceBefore=8,
            spaceAfter=5,
            wordWrap="CJK",
            textColor=colors.HexColor("#111827"),
        ),
        "body": ParagraphStyle(
            "body",
            parent=base["BodyText"],
            fontName=font_name,
            fontSize=9.5,
            leading=14,
            spaceAfter=4,
            alignment=TA_LEFT,
            wordWrap="CJK",
        ),
        "bullet": ParagraphStyle(
            "bullet",
            parent=base["BodyText"],
            fontName=font_name,
            fontSize=9,
            leading=13,
            leftIndent=10,
            firstLineIndent=-7,
            spaceAfter=3,
            wordWrap="CJK",
        ),
    }


def _escape(text: str) -> str:
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    text = re.sub(r"`([^`]+)`", r"<font color='#374151'>\1</font>", text)
    text = re.sub(r"\*\*([^*]+)\*\*", r"<b>\1</b>", text)
    return text


def build_story(markdown: str, styles: dict[str, ParagraphStyle]) -> list[object]:
    story: list[object] = []
    image_re = re.compile(r"!\[[^\]]*\]\(([^)]+)\)")
    max_width = A4[0] - 36 * mm
    max_height = 105 * mm

    for raw_line in markdown.splitlines():
        line = raw_line.strip()
        if not line:
            story.append(Spacer(1, 2))
            continue
        image_match = image_re.match(line)
        if image_match:
            image_path = (PROJECT_ROOT / image_match.group(1)).resolve()
            if image_path.exists():
                img = Image(str(image_path))
                ratio = min(max_width / img.imageWidth, max_height / img.imageHeight, 1)
                img.drawWidth = img.imageWidth * ratio
                img.drawHeight = img.imageHeight * ratio
                story.append(Spacer(1, 4))
                story.append(img)
                story.append(Spacer(1, 6))
            else:
                story.append(Paragraph(f"画像が見つかりません: {_escape(str(image_path))}", styles["body"]))
            continue
        if line.startswith("# "):
            if story:
                story.append(PageBreak())
            story.append(Paragraph(_escape(line[2:]), styles["h1"]))
        elif line.startswith("## "):
            story.append(Paragraph(_escape(line[3:]), styles["h2"]))
        elif line.startswith("- "):
            story.append(Paragraph("• " + _escape(line[2:]), styles["bullet"]))
        elif re.match(r"^\d+\. ", line):
            story.append(Paragraph(_escape(line), styles["body"]))
        else:
            story.append(Paragraph(_escape(line), styles["body"]))
    return story


def main() -> int:
    if not MANUAL_MD.exists():
        print(f"Manual markdown was not found: {MANUAL_MD}")
        return 1

    OUTPUT_PDF.parent.mkdir(parents=True, exist_ok=True)
    font_name = register_font()
    styles = _styles(font_name)
    story = build_story(MANUAL_MD.read_text(encoding="utf-8"), styles)
    doc = SimpleDocTemplate(
        str(OUTPUT_PDF),
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=16 * mm,
        bottomMargin=16 * mm,
        title="売上データ自動集計・月末売上管理ツール 操作マニュアル",
    )
    doc.build(story)
    print(f"Created PDF: {OUTPUT_PDF}")
    print(f"Font: {font_name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def read_text(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


def test_private_material_scripts_exist() -> None:
    assert (ROOT / "scripts" / "annotate_manual_images.py").exists()
    assert (ROOT / "scripts" / "build_manual_pdf.py").exists()


def test_optional_requirements_include_manual_dependencies() -> None:
    requirements = read_text("requirements-optional.txt").lower().splitlines()
    assert "reportlab" in requirements
    assert "pillow" in requirements


def test_gitignore_excludes_private_and_runtime_materials() -> None:
    gitignore = read_text(".gitignore")
    for pattern in [
        "_private_materials/",
        "release/",
        "dist/",
        "build/",
        "*.exe",
        "*.spec",
        "data/output/",
        "logs/",
        "tmp_release_check/",
        "__pycache__/",
        ".venv/",
    ]:
        assert pattern in gitignore


def test_release_package_includes_private_manual_when_available() -> None:
    script = read_text("scripts/build_release_package.ps1")
    assert "_private_materials" in script
    assert "operation_manual.pdf" in script
    assert "images\\annotated" in script


def test_readme_mentions_private_materials_without_sales_copy() -> None:
    readme = read_text("README.md")
    assert "_private_materials" in readme
    assert "## 商品説明文" not in readme
    assert "## 短めの商品紹介文" not in readme


def test_distribution_docs_reference_private_manual_workflow() -> None:
    quick_start = read_text("docs/README_QUICK_START.txt")
    distribution = read_text("docs/distribution_guide.md")

    assert "operation_manual.pdf" in quick_start
    assert "_private_materials" in distribution
    assert "operation_manual.pdf" in distribution

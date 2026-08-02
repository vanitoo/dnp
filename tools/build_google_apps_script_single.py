from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SOURCE_DIR = ROOT / "google-apps-script"
OUTPUT_DIR = ROOT / "google-apps-script-single"
OUTPUT_FILE = OUTPUT_DIR / "Code.gs"

FILES = [
    "Code.gs",
    "Utils.gs",
    "Menu.gs",
    "Setup.gs",
    "YearSheets.gs",
    "Services.gs",
    "Template.gs",
    "PdfGenerator.gs",
    "PdfCleanup.gs",
    "Mail.gs",
    "Journal.gs",
]


def main() -> None:
    missing = [name for name in FILES if not (SOURCE_DIR / name).exists()]
    if missing:
        raise SystemExit("Missing source files: " + ", ".join(missing))

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    parts = [
        "/**\n"
        " * DNP Receipts — единый файл Google Apps Script.\n"
        " * Файл создан автоматически из модулей каталога google-apps-script.\n"
        " * Для тестирования достаточно вставить только этот Code.gs.\n"
        " */\n"
    ]

    for name in FILES:
        source = (SOURCE_DIR / name).read_text(encoding="utf-8").strip()
        parts.append(f"\n\n// ============================================================\n")
        parts.append(f"// MODULE: {name}\n")
        parts.append("// ============================================================\n\n")
        parts.append(source)
        parts.append("\n")

    OUTPUT_FILE.write_text("".join(parts), encoding="utf-8")
    print(f"Created {OUTPUT_FILE}")


if __name__ == "__main__":
    main()

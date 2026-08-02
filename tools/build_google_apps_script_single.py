from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SOURCE_DIR = ROOT / "google-apps-script"
OUTPUT_DIR = ROOT / "google-apps-script-single"
OUTPUT_FILE = OUTPUT_DIR / "Code.gs"

# Единственный источник истины — модульная версия в google-apps-script.
# Единый Code.gs всегда собирается из этого списка в указанном порядке.
FILES = [
    "Code.gs",
    "Utils.gs",
    "Menu.gs",
    "Setup.gs",
    "YearSheets.gs",
    "Services.gs",
    "Template.gs",
    "PdfGenerator.gs",
    "PdfLog.gs",
    "PdfCleanup.gs",
    "Mail.gs",
    "Journal.gs",
]

REQUIRED_SYMBOLS = [
    "function generatePdfsForMonth(",
    "function generatePdfsForMonthWithLog(",
    "function openPdfLogSheet(",
    "function clearPdfLog(",
    "function startPdfGenerationFromDialog(",
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
        " * Не редактируйте его вручную: изменения будут перезаписаны сборкой.\n"
        " * Для тестирования достаточно вставить только этот Code.gs.\n"
        " */\n"
    ]

    source_contents: dict[str, str] = {}

    for name in FILES:
        source = (SOURCE_DIR / name).read_text(encoding="utf-8").strip()
        source_contents[name] = source
        parts.append("\n\n// ============================================================\n")
        parts.append(f"// MODULE: {name}\n")
        parts.append("// ============================================================\n\n")
        parts.append(source)
        parts.append("\n")

    generated = "".join(parts)

    # Проверяем, что каждый модуль действительно попал в единый файл без изменений.
    for name, source in source_contents.items():
        marker = f"// MODULE: {name}"
        if marker not in generated:
            raise SystemExit(f"Generated file does not contain module marker: {name}")
        if source not in generated:
            raise SystemExit(f"Generated file does not contain exact source content: {name}")

    # Проверяем ключевые функции, чтобы не повторилась ошибка с отсутствующим лог-генератором.
    missing_symbols = [symbol for symbol in REQUIRED_SYMBOLS if symbol not in generated]
    if missing_symbols:
        raise SystemExit("Generated file is missing required symbols: " + ", ".join(missing_symbols))

    OUTPUT_FILE.write_text(generated, encoding="utf-8")
    print(f"Created {OUTPUT_FILE}")
    print(f"Modules included: {len(FILES)}")
    print("Parity check passed: modular and single-file versions are identical.")


if __name__ == "__main__":
    main()

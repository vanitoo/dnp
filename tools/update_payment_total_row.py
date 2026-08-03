from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
GENERATOR = ROOT / "google-apps-script" / "PdfGenerator.gs"
CORE = ROOT / "google-apps-script" / "Code.gs"

OLD_BLOCK = """  table.getRow(0).editAsText().setBold(true);\n  table.getRow(table.getNumRows() - 1).editAsText().setBold(true);\n  applyPaymentTableColumnWidths_(table);\n"""

NEW_BLOCK = """  table.getRow(0).editAsText().setBold(true);\n  applyPaymentTableColumnWidths_(table);\n\n  // Последняя строка: первые пять колонок объединяются под надпись\n  // «ИТОГО К ОПЛАТЕ», сумма остаётся в отдельной последней колонке.\n  const totalRow = table.getRow(table.getNumRows() - 1);\n  for (let mergeIndex = 0; mergeIndex < 4; mergeIndex++) {\n    totalRow.getCell(1).merge();\n  }\n  totalRow.getCell(0).setText('ИТОГО К ОПЛАТЕ');\n  totalRow.getCell(1).setText(formatReceiptMoney_(total) + ' руб.');\n  totalRow.getCell(0).editAsText().setBold(true);\n  totalRow.getCell(1).editAsText().setBold(true);\n  totalRow.getCell(0).getChild(0).asParagraph()\n    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);\n"""


def replace_once(path: Path, old: str, new: str) -> bool:
    text = path.read_text(encoding="utf-8")
    if new in text:
        return False
    if old not in text:
        raise SystemExit(f"Expected block not found in {path}")
    path.write_text(text.replace(old, new, 1), encoding="utf-8")
    return True


def main() -> None:
    changed = replace_once(GENERATOR, OLD_BLOCK, NEW_BLOCK)

    core = CORE.read_text(encoding="utf-8")
    if "const DNP_VERSION = '3.7.2';" in core:
        CORE.write_text(
            core.replace("const DNP_VERSION = '3.7.2';", "const DNP_VERSION = '3.7.3';", 1),
            encoding="utf-8",
        )
        changed = True
    elif "const DNP_VERSION = '3.7.3';" not in core:
        raise SystemExit("Unexpected DNP_VERSION in Code.gs")

    print("Payment total row migration applied." if changed else "Already up to date.")


if __name__ == "__main__":
    main()

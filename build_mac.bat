@echo off
chcp 65001 >nul

echo === DNP Receipts: сборка EXE ===
python -m pip install --upgrade pip
python -m pip install pyinstaller openpyxl python-docx num2words

pyinstaller           --windowed --name DNP_Receipts receipt_generator.py

echo.
echo Готово: dist\DNP_Receipts.exe
pause



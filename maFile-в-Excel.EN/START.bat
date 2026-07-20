@echo off
chcp 65001 > nul
color 0A
title maFile-to-Excel Converter [v1.0]
cls

echo ===================================================
echo             maFile-to-Excel Converter
echo ===================================================
echo   [*] Creator: ChuCha
echo   [*] Status: Initializing configuration files...
echo   [*] Template: "template.xlsx" is ready.
echo ---------------------------------------------------
echo   [*] Running account table generation...
echo.

python "Under the hood\generate_sheet.py"

echo.
echo ===================================================
echo   [+] Job finished successfully!
echo   [+] Output tables saved to:
echo       "Your Ready Accounts"
echo ===================================================
echo.
echo Press any key to exit...
pause > nul

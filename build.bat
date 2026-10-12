@echo off
chcp 65001 >nul
rem Сборка: dist\GIB.exe и (если установлен Inno Setup 6) установщик dist\GIB-Setup.exe
if not exist .venv (
    python -m venv .venv
)
call .venv\Scripts\activate.bat
pip install -r requirements.txt
pyinstaller --noconfirm --onefile --windowed --name GIB ^
    --icon assets\icon.ico ^
    --add-data "assets\icon.png;assets" ^
    --add-data "assets\gib_template.xlsm;assets" ^
    --paths gib_app ^
    gib_app\main.py

set ISCC=
if exist "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" set ISCC="%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set ISCC="%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set ISCC="%ProgramFiles%\Inno Setup 6\ISCC.exe"
if defined ISCC (
    %ISCC% installer.iss
    echo.
    echo Готово: dist\GIB.exe и dist\GIB-Setup.exe
) else (
    echo.
    echo Готово: dist\GIB.exe
    echo Для установщика поставьте Inno Setup 6 и запустите build.bat ещё раз.
)
pause

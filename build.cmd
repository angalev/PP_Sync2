@echo off
rem Сборка dist\ПП_синхронизация.exe. Запускать двойным щелчком или из консоли.
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set "PY=.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"

"%PY%" -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo Ставлю PyInstaller...
    "%PY%" -m pip install pyinstaller || goto :fail
)

"%PY%" -m PyInstaller --noconfirm --clean build_exe.spec || goto :fail

echo.
echo Готово: %CD%\dist\ПП_синхронизация.exe
goto :eof

:fail
echo.
echo Сборка не удалась, см. сообщения выше.
exit /b 1

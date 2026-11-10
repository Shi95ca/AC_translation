@echo off
chcp 65001 >nul
setlocal EnableExtensions
cd /d "%~dp0"

set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

if exist "venv\Scripts\python.exe" goto :check_deps

echo [1/2] Первичная настройка: поиск Python в системе...
set "SYS_PY="

py -3 --version >nul 2>&1 && set "SYS_PY=py -3"
if not defined SYS_PY python.exe --version >nul 2>&1 && set "SYS_PY=python.exe"

if not defined SYS_PY (
    echo [ОШИБКА] Python 3 не найден на этом компьютере!
    echo Скачайте и установите Python с https://www.python.org/downloads/
    echo ВАЖНО: при установке обязательно поставьте галочку "Add python.exe to PATH".
    pause
    exit /b 1
)

echo [1/2] Создание локального виртуального окружения venv...
%SYS_PY% -m venv venv
if errorlevel 1 (
    echo [ОШИБКА] Не удалось создать виртуальное окружение venv.
    pause
    exit /b 1
)

:check_deps
"venv\Scripts\python.exe" -c "import docx, xlrd, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo [2/2] Установка необходимых библиотек ^(python-docx, xlrd, openpyxl^)...
    "venv\Scripts\python.exe" -m pip install --upgrade pip >nul
    "venv\Scripts\python.exe" -m pip install python-docx xlrd openpyxl
    if errorlevel 1 (
        echo [ОШИБКА] Не удалось установить пакеты. Проверьте подключение к интернету.
        pause
        exit /b 1
    )
)

"venv\Scripts\python.exe" translate_card.py %*
echo.
pause
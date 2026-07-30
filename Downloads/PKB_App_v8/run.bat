@echo off
chcp 65001 >nul
echo ==========================================
echo   PKB App v11 - Запуск GUI
echo ==========================================
echo.

REM Проверяем наличие Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ОШИБКА] Python не найден. Установите Python 3.10+ и добавьте в PATH.
    pause
    exit /b 1
)

REM Проверяем зависимости
echo [INFO] Проверка зависимостей...
python -c "import pdfplumber, openpyxl, pypdf" >nul 2>&1
if errorlevel 1 (
    echo [INFO] Установка зависимостей...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [ОШИБКА] Не удалось установить зависимости.
        pause
        exit /b 1
    )
)

REM Проверяем шаблон
if not exist "assets\template.xlsx" (
    echo [ОШИБКА] Шаблон не найден: assets\template.xlsx
    echo [INFO] Поместите template.xlsx в папку assets\
    pause
    exit /b 1
)

REM Создаём папку history если нет
if not exist "history" mkdir history

echo [INFO] Запуск GUI...
echo.
pythonw app.py

if errorlevel 1 (
    echo.
    echo [ОШИБКА] При запуске произошла ошибка. Запускаем с консолью...
    python app.py
    pause
)

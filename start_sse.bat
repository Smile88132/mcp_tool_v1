@echo off
rem Venn MCP Server (SSE Daemon)

:: ===============================================================================
:: PYTHON YOLU AYARI (Gerekiyorsa asagidaki yolu kendi sanal PC'nize gore duzenleyin)
:: ===============================================================================
set "PYTHON_CMD=C:\Users\ismailyilmaz\AppData\Local\Programs\Python\Python314\python.exe"

echo ===============================================================================
echo   INTRANET VENN MCP SUNUCUSU - SSE/HTTP AG SERVISI BASLATILIYOR
echo ===============================================================================
echo.

if exist "%PYTHON_CMD%" (
    echo [*] Python Yolu: "%PYTHON_CMD%"
    goto python_ok
)

where python >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    set "PYTHON_CMD=python"
    goto python_ok
)

echo [HATA] Python bulunamadi! Lutfen start_sse.bat icindeki PYTHON_CMD yolunu guncelleyin.
pause
exit /b 1

:python_ok
"%PYTHON_CMD%" setup_environment.py
if %ERRORLEVEL% NEQ 0 (
    echo [HATA] Ortam hazirlanamadi!
    pause
    exit /b 1
)

echo [*] Sunucu http://0.0.0.0:8000/sse adresinde baslatiliyor...
"%PYTHON_CMD%" app.py --transport sse --host 0.0.0.0 --port 8000
pause

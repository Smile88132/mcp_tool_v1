@echo off
rem Venn Server - Direct Network Launcher

set "PYTHON_CMD=C:\Users\ismailyilmaz\AppData\Local\Programs\Python\Python314\python.exe"

echo ===============================================================================
echo   INTRANET VENN SUNUCUSU - AGDAKI HERKESE ACIK WEB SERVISI BASLATILIYOR
echo ===============================================================================
echo.

if exist "%PYTHON_CMD%" goto python_ok
where python >nul 2>nul && set "PYTHON_CMD=python" && goto python_ok

echo [HATA] Python bulunamadi!
pause
exit /b 1

:python_ok
"%PYTHON_CMD%" setup_environment.py
if %ERRORLEVEL% NEQ 0 (
    echo [HATA] Ortam hazirlanamadi!
    pause
    exit /b 1
)

echo [*] Sunucu http://0.0.0.0:8000 adresinde baslatiliyor...
echo [*] Agdaki tum cihazlardan http://<SANAL_PC_IP>:8000/docs adresine erisebilirsiniz.
"%PYTHON_CMD%" app.py --transport http --host 0.0.0.0 --port 8000
pause

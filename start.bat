@echo off
rem Venn Nonconformity Server - Intranet Launcher

:: ===============================================================================
:: PYTHON YOLU AYARI
:: ===============================================================================
set "PYTHON_CMD=C:\Users\ismailyilmaz\AppData\Local\Programs\Python\Python314\python.exe"

echo ===============================================================================
echo     INTRANET VENN DIAGRAM RAG SUNUCUSU (AGDAKI HERKESE ACIK WEB SERVISI)
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

echo [HATA] Python bulunamadi! Lutfen start.bat icindeki PYTHON_CMD yolunu guncelleyin.
pause
exit /b 1

:python_ok
echo.
echo [*] Bagimliliklar kontrol ediliyor...
"%PYTHON_CMD%" setup_environment.py
if %ERRORLEVEL% NEQ 0 (
    echo [HATA] Paket kurulumu basarisiz oldu!
    pause
    exit /b 1
)

echo.
if "%1"=="--http" goto run_http
if "%1"=="--stdio" goto cli_stdio
if "%1"=="--test" goto cli_test
if "%1"=="--demo" goto cli_demo

:menu
echo -------------------------------------------------------------------------------
echo  Lutfen calistirmak istediginiz modu secin:
echo -------------------------------------------------------------------------------
echo  [1] Agdaki Herkese Acik Baslat (http://0.0.0.0:8000 - ONERILEN)
echo  [2] stdio Modu    (Ayni sanal PC'deki LLM istemcileri / Claude Desktop / Cursor)
echo  [3] Hizli Demo    (example_usage.py ile gercek veri analizini calistir)
echo  [4] Testleri Kos  (test_app.py ile birim testlerini calistir)
echo  [5] Cikis
echo -------------------------------------------------------------------------------
set /p secim="Seciminiz [1-5]: "

if "%secim%"=="1" goto run_http
if "%secim%"=="2" goto run_stdio
if "%secim%"=="3" goto run_demo
if "%secim%"=="4" goto run_test
if "%secim%"=="5" exit /b 0

echo Gecersiz secim, lutfen 1 ile 5 arasinda bir deger girin.
goto menu

:cli_stdio
"%PYTHON_CMD%" app.py --transport stdio
exit /b %ERRORLEVEL%

:cli_demo
"%PYTHON_CMD%" example_usage.py
exit /b %ERRORLEVEL%

:cli_test
"%PYTHON_CMD%" test_app.py
exit /b %ERRORLEVEL%

:run_http
echo.
echo ===============================================================================
echo [WEB SERVISI] Agdaki Herkese Acik Olarak Baslatiliyor...
echo Adres: http://0.0.0.0:8000
echo Tarayicidan veya baska PC'den erisim: http://^<SANAL_PC_IP^>:8000/docs
echo Durdurmak icin CTRL+C tuslarina basin.
echo ===============================================================================
"%PYTHON_CMD%" app.py --transport http --host 0.0.0.0 --port 8000
pause
exit /b 0

:run_stdio
echo.
echo ===============================================================================
echo [MCP Sunucusu] stdio modunda baslatiliyor...
echo ===============================================================================
"%PYTHON_CMD%" app.py --transport stdio
pause
exit /b 0

:run_demo
echo.
echo ===============================================================================
echo [Demo] Uygunsuzluk.xlsx uzerinde canli analiz calistiriliyor...
echo ===============================================================================
"%PYTHON_CMD%" example_usage.py
echo.
pause
goto menu

:run_test
echo.
echo ===============================================================================
echo [Test] Birim testleri calistiriliyor...
echo ===============================================================================
"%PYTHON_CMD%" test_app.py
echo.
pause
goto menu

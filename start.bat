@echo off
rem Venn Nonconformity MCP Server - Intranet Launcher

:: ===============================================================================
:: PYTHON YOLU AYARI (Gerekiyorsa asagidaki yolu kendi sanal PC'nize gore duzenleyin)
:: Ornek 1: set "PYTHON_CMD=C:\Users\ismailyilmaz\AppData\Local\Programs\Python\Python314\python.exe"
:: Ornek 2: set "PYTHON_CMD=C:\Python311\python.exe"
:: Ornek 3: set "PYTHON_CMD=C:\ProgramData\anaconda3\python.exe"
:: Ornek 4: set "PYTHON_CMD=python"  (Sistem PATH'inde tanimliysa)
:: ===============================================================================
set "PYTHON_CMD=C:\Users\ismailyilmaz\AppData\Local\Programs\Python\Python314\python.exe"

echo ===============================================================================
echo     INTRANET VENN DIAGRAM RAG MCP SUNUCUSU (OTOMATIK KURULUM VE BASLATICI)
echo ===============================================================================
echo.

:: 1. Belirtilen Python Yolunun Kontrolü
if exist "%PYTHON_CMD%" (
    echo [*] Tanimlanan Python yolu kullaniliyor:
    echo     "%PYTHON_CMD%"
    goto python_ok
)

:: Belirtilen yol bulunamazsa sistem PATH'inden 'python' ara
where python >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    set "PYTHON_CMD=python"
    echo [*] Belirtilen ozel yol bulunamadi, sistem PATH uzerindeki Python kullaniliyor.
    goto python_ok
)

:: Ikisi de yoksa hata ver
echo [HATA] Python calistirilabilir dosyasi bulunamadi!
echo Aranan yol: "%PYTHON_CMD%"
echo.
echo Lutfen start.bat dosyasini metin duzenleyici ile acip 11. satirdaki
echo "PYTHON_CMD" degiskenine sanal PC'nizdeki dogru python.exe yolunu yazin.
echo.
pause
exit /b 1

:python_ok
echo.
echo [*] Bagimliliklar kontrol ediliyor ve gerekiyorsa internetsiz (wheels/) kuruluyor...
"%PYTHON_CMD%" setup_environment.py
if %ERRORLEVEL% NEQ 0 (
    echo [HATA] Paket kurulumu basarisiz oldu!
    pause
    exit /b 1
)

echo.
if "%1"=="--stdio" goto cli_stdio
if "%1"=="--sse" goto cli_sse
if "%1"=="--test" goto cli_test
if "%1"=="--demo" goto cli_demo

:menu
echo -------------------------------------------------------------------------------
echo  Lutfen calistirmak istediginiz modu secin:
echo -------------------------------------------------------------------------------
echo  [1] stdio Modu    (Ayni sanal PC'deki LLM istemcileri / Claude Desktop / Cursor)
echo  [2] SSE/HTTP Modu (Intranet agindan erisim icin: http://0.0.0.0:8000/sse)
echo  [3] Hizli Demo    (example_usage.py ile gercek veri analizini calistir)
echo  [4] Testleri Kos  (test_app.py ile birim testlerini calistir)
echo  [5] Cikis
echo -------------------------------------------------------------------------------
set /p secim="Seciminiz [1-5]: "

if "%secim%"=="1" goto run_stdio
if "%secim%"=="2" goto run_sse
if "%secim%"=="3" goto run_demo
if "%secim%"=="4" goto run_test
if "%secim%"=="5" exit /b 0

echo Gecersiz secim, lutfen 1 ile 5 arasinda bir deger girin.
goto menu

:cli_stdio
"%PYTHON_CMD%" app.py --transport stdio
exit /b %ERRORLEVEL%

:cli_sse
"%PYTHON_CMD%" app.py --transport sse --host 0.0.0.0 --port 8000
exit /b %ERRORLEVEL%

:cli_demo
"%PYTHON_CMD%" example_usage.py
exit /b %ERRORLEVEL%

:cli_test
"%PYTHON_CMD%" test_app.py
exit /b %ERRORLEVEL%

:run_stdio
echo.
echo ===============================================================================
echo [MCP Sunucusu] stdio modunda baslatiliyor...
echo JSON-RPC iletisimi baslatildi. LLM istemcinizden komut bekleniyor.
echo ===============================================================================
"%PYTHON_CMD%" app.py --transport stdio
pause
exit /b 0

:run_sse
echo.
echo ===============================================================================
echo [MCP Sunucusu] SSE/HTTP modunda baslatiliyor...
echo Ag Dinleme Adresi: http://0.0.0.0:8000/sse
echo Intranet istemcileri bu adrese baglanabilir.
echo Sunucuyu durdurmak icin CTRL+C tuslarina basin.
echo ===============================================================================
"%PYTHON_CMD%" app.py --transport sse --host 0.0.0.0 --port 8000
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

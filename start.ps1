# Venn Nonconformity MCP Server - PowerShell Launcher
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ===============================================================================
# PYTHON YOLU AYARI (Gerekiyorsa burayı güncelleyin)
# ===============================================================================
$PYTHON_CMD = "C:\Users\ismailyilmaz\AppData\Local\Programs\Python\Python314\python.exe"

if (-not (Test-Path $PYTHON_CMD)) {
    if (Get-Command python -ErrorAction SilentlyContinue) {
        $PYTHON_CMD = "python"
    } else {
        Write-Host "[HATA] Python bulunamadı! Lütfen start.ps1 içindeki `$PYTHON_CMD yolunu güncelleyin." -ForegroundColor Red
        exit 1
    }
}

Write-Host "===============================================================================" -ForegroundColor Cyan
Write-Host "    INTRANET VENN DIAGRAM RAG MCP SUNUCUSU (POWERSHELL BASLATICI)" -ForegroundColor Cyan
Write-Host "===============================================================================" -ForegroundColor Cyan
Write-Host "[*] Python Yolu: $PYTHON_CMD" -ForegroundColor Yellow

# 1. Ortam ve Paket Kontrolü
Write-Host "`n[*] Bağımlılıklar kontrol ediliyor..." -ForegroundColor Yellow
& $PYTHON_CMD setup_environment.py
if ($LASTEXITCODE -ne 0) {
    Write-Host "[HATA] Paket kurulumu başarısız!" -ForegroundColor Red
    exit 1
}

# 2. Seçenek Menüsü
Write-Host "`nLütfen çalıştırmak istediğiniz modu seçin:" -ForegroundColor Green
Write-Host " [1] stdio Modu    (Aynı PC'deki LLM / Claude Desktop / Cursor)"
Write-Host " [2] SSE/HTTP Modu (Intranet ağı üzerinden erişim: http://0.0.0.0:8000/sse)"
Write-Host " [3] Canlı Demo    (example_usage.py ile analiz)"
Write-Host " [4] Testleri Koş  (test_app.py ile birim testleri)"
Write-Host " [5] Çıkış"

$secim = Read-Host "`nSeçiminiz [1-5]"

switch ($secim) {
    "1" {
        Write-Host "`n[MCP Sunucusu] stdio modunda başlatılıyor..." -ForegroundColor Cyan
        & $PYTHON_CMD app.py --transport stdio
    }
    "2" {
        Write-Host "`n[MCP Sunucusu] SSE/HTTP modunda dinleniyor: http://0.0.0.0:8000/sse" -ForegroundColor Cyan
        & $PYTHON_CMD app.py --transport sse --host 0.0.0.0 --port 8000
    }
    "3" {
        & $PYTHON_CMD example_usage.py
    }
    "4" {
        & $PYTHON_CMD test_app.py
    }
    "5" {
        exit 0
    }
    Default {
        Write-Host "Geçersiz seçim!" -ForegroundColor Red
    }
}

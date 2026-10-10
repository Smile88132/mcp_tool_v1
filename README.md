# Intranet Venn Şeması Uygunsuzluk Analizi MCP Sunucusu

Bu proje, kurum içi kısıtlı ve güvenli intranet ortamlarında çalışan Sanal Bilgisayarlar (Virtual PC) için geliştirilmiş, **Model Context Protocol (MCP)** standartlarına uygun bir Python sunucusu ve RAG (Retrieval-Augmented Generation) analiz aracıdır.

Sistem, Excel formatındaki faaliyet verilerini ve test aşamalarını (**Dry-run**, **SKT - Saha Kabul Testi**, **FKT - Fabrika Kabul Testi**) okur ve uygunsuzlukları **Venn Şeması (küme matematiği)** ile analiz ederek LLM'e yapılandırılmış JSON çıktısı sağlar.

---

## 🚀 Özellikler

1. **Venn Şeması Küme Analizi (`analyze_venn_nonconformities`):**
   - **Fazlar Arası (`compare_by='phases'`):** Belirli bir faaliyet altındaki veya tüm dosyadaki Dry-run, SKT ve FKT uygunsuzluklarını kümelere (`set`) ayırır.
     - Sadece Dry-run'da olanlar (çözülenler)
     - Sadece SKT'de ortaya çıkanlar
     - Sadece FKT'de ortaya çıkanlar
     - İkili kesişimler (Örn: SKT & FKT ortak)
     - Üç aşamada da devam eden kronik uygunsuzluklar (Dry-run ∩ SKT ∩ FKT)
     - Jaccard benzerlik oranları ve LLM için hazır RAG özet içgörüsü.
   - **Faaliyetler Arası (`compare_by='activities'`):** Farklı faaliyetler arasındaki benzer ve ortak uygunsuzlukları karşılaştırır.

2. **Veri Keşif Araçları:**
   - `list_activities_and_phases`: Excel'deki faaliyet adlarını, mevcut fazları ve kayıt sayılarını listeler (LLM'in faaliyet adlarını doğru bilmesini sağlar).
   - `preview_excel_data`: Excel sütunlarını ve ilk birkaç satırı inceler.

3. **İntranet ve Güvenlik:**
   - Sıfır telemetri ve sıfır dış ağ bağlantısı (tamamen offline/yerel çalışır).
   - Akıllı sütun eşleme (Türkçe karakter ve farklı sütun başlıklarına toleranslı).
   - Hata durumunda mevcut sütun ve faaliyetleri listeleyen şeffaf hata yönetimi.

4. **Çift Taşıma Protokolü (Dual Transport):**
   - **`stdio`:** LLM istemcisi (Claude Desktop, Cursor, yerel ajan) aynı sanal makinede çalıştığında.
   - **`sse` (HTTP/SSE):** LLM istemcisi ağdaki başka bir intranet sunucusundan bağlandığında (örn: `http://0.0.0.0:8000/sse`).

---

## 📁 Proje Dosya Yapısı

```
mcp tool v1/
├── app.py                      # Ana MCP Sunucusu ve Araçları
├── setup_environment.py        # Eksik paketleri internetsiz/otomatik kuran betik
├── start.bat                   # Tek tıkla paketleri kurup başlatan Windows başlatıcısı
├── start_sse.bat               # Ağ servisi olarak tek tıkla başlatan servis başlatıcı
├── start.ps1                   # PowerShell başlatıcısı
├── wheels/                     # Tamamen internetsiz kurulum için 36 adet hazır .whl paketi
├── test_app.py                 # Birim testleri (8 test)
├── example_usage.py            # Canlı veri analiz demosu
├── Uygunsuzluk.xlsx            # Gerçek veri kaynağı (7.123 satır)
├── ornek_uygunsuzluklar.xlsx   # Sentetik test verisi
├── requirements.txt            # Python bağımlılıkları listesi
├── environment.yml             # Kurum içi Conda ortam şablonu
├── mcp_tool.md                 # Gereksinim belgesi
└── README.md                   # Bu kullanım kılavuzu
```

---

## ⚡ TEK TIKLA BAŞLATMA (ÖNERİLEN)

Proje klasöründeki **`start.bat`** dosyasına çift tıklayarak çalıştırabilirsiniz:
1. Python'un varlığını kontrol eder.
2. Gerekli paketleri (`mcp`, `pandas`, `openpyxl`, `uvicorn` vb.) eksikse **`wheels/`** klasöründen **tamamen internetsiz (çevrimdışı)** olarak otomatik kurar.
3. Karşınıza etkileşimli bir menü çıkarır:
   - `[1]` stdio Modu (Yerel Claude / Cursor / LLM istemcileri için)
   - `[2]` SSE / HTTP Modu (Intranet ağından erişim için: `http://0.0.0.0:8000/sse`)
   - `[3]` Hızlı Demo (`example_usage.py`)
   - `[4]` Testleri Koş (`test_app.py`)

> [!TIP]
> **Python Yolu Ayarı:**  
> Sanal PC'nizde Python sistem PATH'ine ekli değilse veya özel bir dizindeyse (örn: Conda / Program Files), [`start.bat`](file:///c:/Users/ismailyilmaz/Desktop/mcp%20tool%20v1/start.bat) ve [`start_sse.bat`](file:///c:/Users/ismailyilmaz/Desktop/mcp%20tool%20v1/start_sse.bat) dosyalarının 10. satırındaki `PYTHON_CMD` yolunu dilediğiniz gibi güncelleyebilirsiniz:
> ```bat
> set "PYTHON_CMD=C:\Users\ismailyilmaz\AppData\Local\Programs\Python\Python314\python.exe"
> ```
---

## 🛠️ Manuel Kurulum Adımları (Intranet / Çevrimdışı Ortam)

### Yöntem A: Hazır `wheels/` Klasöründen Çevrimdışı Kurulum

Proje klasöründe 36 adet `.whl` dosyası hazır olarak indirilmiştir. İnternetsiz sanal makinede şu komutu çalıştırmanız yeterlidir:

```bash
python setup_environment.py
```
Veya doğrudan pip ile:
```bash
python -m pip install --no-index --find-links=./wheels -r requirements.txt
```

### Yöntem B: Kurum İçi Conda Kanalı ile Kurulum

Eğer kurumunuzda yerel bir Conda deposu varsa:

```bash
# 1. Conda ortamını oluşturun
conda env create -f environment.yml

# 2. Ortamı aktif edin
conda activate mcp_venn_env
```

---

## 🖥️ Çalıştırma Seçenekleri

### 1. `stdio` Modu (Yerel LLM İstemcisi)

LLM istemcisi (Cursor, Claude Desktop, yerel Python betiği vb.) aynı Sanal PC üzerinde çalışıyorsa:

```bash
python app.py --transport stdio
```

#### Örnek Claude Desktop / Cline / Cursor Yapılandırması (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "venn_analiz": {
      "command": "python",
      "args": [
        "C:\\Users\\ismailyilmaz\\Desktop\\mcp tool v1\\app.py",
        "--transport",
        "stdio"
      ]
    }
  }
}
```

---

### 2. `sse` Modu (Intranet Ağ Üzerinden Erişim)

LLM istemcisi veya kurumsal RAG portalı ağdaki farklı bir sunucuda çalışıyorsa, bu makinede bir HTTP/SSE servisi ayağa kaldırılır:

```bash
python app.py --transport sse --host 0.0.0.0 --port 8000
```

Bu durumda sunucu `http://<SANAL_PC_IP>:8000/sse` adresinde dinlemeye başlar.

#### Uzak İstemci JSON Yapılandırması:
```json
{
  "mcpServers": {
    "venn_analiz_remote": {
      "url": "http://192.168.1.50:8000/sse"
    }
  }
}
```

---

## 🧪 Testlerin Çalıştırılması

Tüm fonksiyonların ve Excel okuma mantığının sorunsuz çalıştığını doğrulamak için:

```bash
python test_app.py
```

Çıktı:
```text
Ran 7 tests in 0.38s
OK
```

---

## 📊 Örnek Çıktı Formatı

`analyze_venn_nonconformities(file_path="...", activity_name="Elektrik Montaj", compare_by="phases")` çağrıldığında dönen JSON:

```json
{
  "durum": "basarili",
  "analiz_tipi": "fazlar_arasi_venn_semasi",
  "kapsam": "Faaliyet: 'Elektrik Montaj'",
  "rag_ozet_yorumu": "Faaliyet: 'Elektrik Montaj' kapsamında incelenen 5 benzersiz uygunsuzluktan 2 tanesi yalnızca Dry-run aşamasında görülüp sonraki aşamalarda çözülmüştür. 1 adet uygunsuzluk ise Dry-run, SKT ve FKT aşamalarının üçünde de tekrarlamış olup kronik kök neden analizi gerektirmektedir.",
  "venn_analizi": {
    "ozet_istatistikler": {
      "toplam_benzersiz_Dry-run": 3,
      "toplam_benzersiz_SKT": 2,
      "toplam_benzersiz_FKT": 2,
      "toplam_benzersiz_uygunsuzluk_sayisi": 5
    },
    "venn_bolgeleri": {
      "sadece_Dry-run": {
        "adet": 2,
        "ogeler": ["Kablo kesiti standart dışı", "Topraklama direnci yüksek"]
      },
      "sadece_SKT": {
        "adet": 1,
        "ogeler": ["Klemens bağlantısı gevşek"]
      },
      "sadece_FKT": {
        "adet": 1,
        "ogeler": ["Kablo bağı kopuk"]
      },
      "ucunde_ortak_kesisim": {
        "adet": 1,
        "ogeler": ["Pano içi etiketleme eksik"]
      }
    }
  }
}
```

# ROL VE BAĞLAM
Sen uzman bir Python geliştiricisi ve Yapay Zeka Entegrasyon mühendisisin. Görevin, standart Model Context Protocol (MCP) Python SDK'sını kullanarak bir MCP Sunucusu (Server) ve yapay zekanın kullanabileceği bir MCP Aracı (Tool) yazmaktır.

**Kritik Ortam Kısıtlaması:** Bu sistem, internet erişimi kısıtlı olan bir Intranet portalındaki Sanal Bilgisayarda (Virtual PC) çalışacaktır. Sistemde güvenlik önlemleri yüksektir, dış dünyaya API çağrısı yapılamaz ve paket kurulumları kurum içi (internal) Conda kanalları veya çevrimdışı `.whl` dosyaları üzerinden yapılmaktadır.

# VERİ YAPISI VE HEDEF
Elimizde bir Excel (`.xlsx`) dosyası bulunuyor. Bu dosyada çeşitli "Faaliyetler" ve bu faaliyetlerde ortaya çıkan "Uygunsuzluklar" yer alıyor.
Faaliyetler ayrıca kendi içinde 3 farklı test/faz aşamasına ayrılıyor:
1. Dry-run
2. SKT (Saha Kabul Testi)
3. FKT (Fabrika Kabul Testi)

**Amacımız:** Yapay zekanın bu verileri okuyup, farklı fazlar (dry-run, skt, fkt) veya farklı faaliyetler arasındaki uygunsuzlukları "Venn Şeması" mantığıyla analiz edebilmesini sağlamak. Yani hangi uygunsuzluklar sadece Dry-run'da var, hangileri hem SKT hem FKT'de ortak, hangileri tamamen benzersiz gibi küme mantığıyla çalışan bir RAG aracı oluşturmak.

# İSTENEN MCP ARAÇLARI (TOOLS)
Lütfen aşağıdaki özelliklere sahip tek bir Python MCP Sunucusu betiği (örn: `app.py`) yaz.

**Araç 1: `analyze_venn_nonconformities`**
* **Girdi Parametreleri:**
  * `file_path` (string): Excel dosyasının sanal makinedeki mutlak (absolute) yolu.
  * `activity_name` (string, opsiyonel): Eğer verilirse sadece o faaliyete ait fazları karşılaştırır.
  * `compare_by` (string): 'phases' (Dry-run, SKT, FKT arası) veya 'activities' (Faaliyetler arası).
* **İşlev:** 
  * `pandas` kütüphanesini kullanarak Excel dosyasını okumalı. (Örnek sütun isimleri: `Faaliyet`, `Faz`, `Uygunsuzluk`).
  * `compare_by == 'phases'` ise: Belirtilen faaliyet altındaki Dry-run, SKT ve FKT uygunsuzluklarını Python `set()` (küme) objelerine dönüştürmeli.
  * Kümeler arası Venn şeması matematiğini hesaplamalı (Sadece Dry-run, Ortak Kesişim, Farklar vb.).
* **Çıktı (Return):** Küme analizlerinin sonucunu, LLM'in kolayca okuyup yorumlayabileceği yapılandırılmış bir JSON string olarak döndürmeli.

# TEKNİK GEREKSİNİMLER VE INTRANET UYARLARI
1. Resmi `@modelcontextprotocol/sdk` (Python) kütüphanesini kullan.
2. **İletişim Protokolü:** LLM istemcisi (Client) aynı sanal makinede çalışıyorsa `stdio` kullanılabilir. Ancak LLM istemcisi Intranet üzerindeki başka bir sunucudaysa, aracın ağ üzerinden dinleyebilmesi için kodu **SSE (Server-Sent Events) / HTTP** üzerinden çalışacak şekilde yapılandır (FastAPI veya Starlette kullanarak). Lütfen her iki taşıma (transport) yöntemini de koda ekle veya nasıl değiştirileceğini yorum satırlarıyla açıkla.
3. **Güvenlik ve Hata Ayıklama:** Telemetri veya dış internet bağlantısı kurmaya çalışan hiçbir eklenti kullanma. Dosya yolu (path) veya Excel sütun isimleri eşleşmezse, hatayı sessizce yutmak yerine mevcut sütunları listeleyen güvenli bir hata yakalama (error handling) bloğu yaz.
4. **Kurulum Adımları:** Çıktının sonuna, bu kodun kurum içi bir Conda ortamında (internal Conda channel) nasıl ayağa kaldırılacağına dair kısa bir `README` veya komut listesi ekle.
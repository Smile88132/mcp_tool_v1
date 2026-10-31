# 📘 Dış Uygulama Kimlik Aktarımı (SSO / Iframe Push) Entegrasyon Rehberi

> **Mimarisi:** Iframe Push Mimarisi (`postMessage` tabanlı)  
> **Hedef Kitle:** Ana platform (Kokpit/Portal) içerisine iframe olarak gömülecek dış web uygulamalarının yazılımcıları  
> **Gereksinim:** Dış uygulamanın harici bir Keycloak veya LDAP bağlantısı kurmasına **gerek yoktur**. Ana portal kimlik doğrulamasını üstlenir ve oturum verilerini otomatik olarak dış siteye iletir.

---

## 1. Mimari ve Çalışma Mantığı

Ana platform, dış uygulamayı bir `<iframe>` içerisinde açtığında; kullanıcı oturum açmışsa tarayıcı seviyesinde güvenli `window.postMessage` protokolünü kullanarak kullanıcının JWT erişim belirtecini (`token`) ve profil bilgilerini (`user`) doğrudan iframe penceresine basar.

```
┌──────────────────────────────────────────────────────────────┐
│                    ANA SİSTEM (KOKPİT)                       │
│                                                              │
│  1. Kullanıcı oturum açtı (Keycloak JWT Token mevcut)        │
│  2. Dış site iframe olarak yüklenir (DirectAppViewer)        │
│  3. iframe yüklendiği anda tarayıcı içi mesaj tetiklenir:    │
│     postMessage({ type: 'kokpit-identity', token, user })    │
└───────────────────────────────┬──────────────────────────────┘
                                │ window.postMessage
                                ▼
┌──────────────────────────────────────────────────────────────┐
│             DIŞ UYGULAMA (iframe İçerisindeki Site)          │
│                                                              │
│  1. Frontend: window.addEventListener('message', ...)        │
│  2. Token ve User objesi doğrulanarak yakalanır              │
│  3. Oturum başlatılır (SessionStorage / React State)         │
│  4. (Opsiyonel) Dış sitenin kendi backend'ine POST edilir:   │
│     POST /api/identity { token, user }                       │
└──────────────────────────────────────────────────────────────┘
```

---

## 2. Ana Portal Tarafındaki Ayar (Admin)

Ana sistem yöneticisi dış uygulamayı eklerken:
1. **Araçlar** -> **Yeni Araç Ekle** adımı açılır.
2. Tür: **"Dış Uygulama (Web App)"** seçilir.
3. Adres (URL): Dış sitenizin adresi yazılır (örn: `http://192.168.1.50:8080/` veya `https://uygulamaniz.sirket.com`).
4. **"Kokpit Kimliği Gönder"** seçeneği **AÇIK** duruma getirilir.
5. Kaydedilir.

---

## 3. Dış Sitede Yapılması Zorunlu Ön Hazırlık (Iframe İzinleri)

Dış sitenizin iframe içinde görüntülenebilmesi için sunucunuzun HTTP başlıklarında iframe kısıtlaması bulunmamalıdır:

* ❌ **Engeller:** `X-Frame-Options: DENY` veya `SAMEORIGIN` (Sayfanın açılmasını engeller).
* ✅ **Doğru Yapılandırma (Nginx / Apache / IIS / Express):**
  İlgili başlığı kaldırın veya yalnızca ana portal origin'ine izin verin:
  ```http
  Content-Security-Policy: frame-ancestors 'self' http://localhost:3004 https://kokpit.sirketiniz.com;
  ```

---

## 4. İletilen Kimlik Paketi Formatı

Ana sistem tarafından iframe penceresine gönderilen `event.data` JSON nesnesinin şeması:

```json
{
  "type": "kokpit-identity",
  "status": "success",
  "source": "kokpit-iframe-bridge",
  "timestamp": "2026-10-29T10:00:00.000Z",
  "token": "eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9...",
  "user": {
    "id": 145,
    "sicil": "12345",
    "username": "ahmet.yilmaz",
    "isim": "Ahmet Yılmaz",
    "mail": "ahmet.yilmaz@sirket.com",
    "picture": "/uploads/photos/12345.jpg",
    "role": "Admin",
    "mudurluk": "UGES Teslimat Yönetim Müdürlüğü",
    "mudurluk_id": 191148,
    "direktorluk": "Program Direktörlüğü",
    "birim": "Planlama Birimi",
    "unvan": "Uzman",
    "hierarchicLevel": 2
  }
}
```

### Alan Açıklamaları

| Alan | Tip | Açıklama |
|---|---|---|
| `type` | `string` | **Zorunlu kontrol:** Her zaman `"kokpit-identity"` değerindedir. |
| `source` | `string` | **Güvenlik kontrolü:** Her zaman `"kokpit-iframe-bridge"` değerindedir. |
| `status` | `string` | `"success"` (Başarılı aktarım). |
| `token` | `string` | Keycloak JWT Bearer Token. Ana sistem backend servislerine veya ortak API'lere yapılacak isteklerde `Authorization: Bearer <token>` olarak kullanılır. |
| `user.sicil` | `string` | Kullanıcının personel sicil numarası. |
| `user.isim` | `string` | Ad Soyad. |
| `user.mail` | `string` | Kurumsal e-posta adresi. |
| `user.role` | `string` | Rol (`Admin`, `User`, vb.). |
| `user.mudurluk` | `string` | Bağlı olduğu müdürlük adı. |
| `user.mudurluk_id` | `number` | Müdürlük kurumsal HR ID değeri. |

---

## 5. Dış Uygulama Kod Örnekleri

### A) Saf HTML / Vanilla JavaScript (Hızlı Entegrasyon)

Dış sitenizin ana `index.html` dosyasına eklemeniz yeterlidir:

```html
<!DOCTYPE html>
<html lang="tr">
<head>
  <meta charset="UTF-8">
  <title>Dış Uygulama</title>
</head>
<body>

  <div id="auth-status">Kimlik bilgisi bekleniyor...</div>

  <script>
    (function() {
      // 1. Güvenli kabul edilen ana sistem origin'leri
      const ALLOWED_ORIGINS = [
        'http://localhost:3004',
        'https://kokpit.sirketiniz.com' // Kendi ana platform adresinizle güncelleyin
      ];

      window.addEventListener('message', function(event) {
        // Güvenlik filtreleri
        if (!ALLOWED_ORIGINS.includes(event.origin) && !event.origin.includes(':3004')) {
          return;
        }
        if (!event.data || event.data.type !== 'kokpit-identity') return;
        if (event.data.source !== 'kokpit-iframe-bridge') return;
        if (event.data.status !== 'success') return;

        const { token, user } = event.data;

        // 2. Token ve profil bilgisini oturum süresince sakla
        sessionStorage.setItem('auth_token', token);
        sessionStorage.setItem('current_user', JSON.stringify(user));

        console.log('✅ Kimlik başarıyla alındı:', user.isim, `(Sicil: ${user.sicil})`);

        // 3. UI'ı güncelle ve uygulamanı başlat
        document.getElementById('auth-status').innerText = 
          `Hoş geldiniz, ${user.isim} (${user.mudurluk || user.role})`;

        // Kendi uygulamanızın başlatma fonksiyonu
        initializeApp(user, token);
      });

      function initializeApp(user, token) {
        // Uygulama render / dashboard yükleme kodlarınız...
      }
    })();
  </script>
</body>
</html>
```

---

### B) Modern React (Hook & Dinleyici Mimarisi)

Dış siteniz bir React (SPA) uygulaması ise, sayfa açılışında kimliği bekleyen ve timeout sonrasında fallback veren hook:

```javascript
// useKokpitIdentity.js
import { useState, useEffect } from 'react';

const ALLOWED_ORIGINS = [
  'http://localhost:3004',
  'https://kokpit.sirketiniz.com'
];

export function useKokpitIdentity() {
  const [identity, setIdentity] = useState({
    user: null,
    token: null,
    loading: true
  });

  useEffect(() => {
    let resolved = false;

    const handleMessage = (event) => {
      // Origin doğrulaması
      const isAllowed = ALLOWED_ORIGINS.includes(event.origin) || event.origin.endsWith(':3004');
      if (!isAllowed) return;

      if (event.data?.type === 'kokpit-identity' && event.data?.source === 'kokpit-iframe-bridge') {
        const { token, user } = event.data;
        resolved = true;
        
        sessionStorage.setItem('kokpit_token', token);
        sessionStorage.setItem('kokpit_user', JSON.stringify(user));

        setIdentity({
          user,
          token,
          loading: false
        });
      }
    };

    window.addEventListener('message', handleMessage);

    // Eğer iframe dışında bağımsız açıldıysa veya mesaj 4 saniye içinde gelmediyse:
    const timer = setTimeout(() => {
      if (!resolved) {
        const cachedUser = sessionStorage.getItem('kokpit_user');
        const cachedToken = sessionStorage.getItem('kokpit_token');
        if (cachedUser) {
          setIdentity({
            user: JSON.parse(cachedUser),
            token: cachedToken,
            loading: false
          });
        } else {
          setIdentity((prev) => ({ ...prev, loading: false }));
        }
      }
    }, 4000);

    return () => {
      window.removeEventListener('message', handleMessage);
      clearTimeout(timer);
    };
  }, []);

  return identity;
}
```

---

### C) Dış Sitenin Kendi Backend'ine Kimliği Aktarması (Tier-2 Senaryosu)

Dış sitenizin kendine ait bir Backend servisi varsa (Node.js, Python, Java, .NET) ve backend oturumu kurmak istiyorsanız:

1. **Frontend'den Kendi Backend'inize Push Edin:**  
   ```javascript
   await fetch('/api/identity', {
     method: 'POST',
     headers: { 'Content-Type': 'application/json' },
     body: JSON.stringify({
       user: payload.user,
       token: payload.token,
       source: 'kokpit-iframe-bridge'
     })
   });
   ```

2. **Backend Örneği (Python / Django):**
   ```python
   # views.py
   import json
   from django.http import JsonResponse
   from django.views.decorators.csrf import csrf_exempt

   @csrf_exempt
   def receive_identity(request):
       if request.method != 'POST':
           return JsonResponse({'error': 'Sadece POST desteklenir.'}, status=405)
       
       data = json.loads(request.body.decode('utf-8'))
       user = data.get('user')
       token = data.get('token')
       
       if not user or not user.get('sicil'):
           return JsonResponse({'error': 'Geçersiz kullanıcı paketi'}, status=400)
       
       # Oturum açma veya belleğe alma
       request.session['current_user'] = user
       request.session['jwt_token'] = token
       
       return JsonResponse({
           'status': 'success',
           'message': f"{user.get('isim')} oturumu açıldı."
       })
   ```

---

## 6. Güvenlik Kontrol Listesi

- [x] **Origin Kontrolü Yapın:** Gelen mesajın `event.origin` adresinin ana portal adresi olduğundan emin olmadan oturum açmayın.
- [x] **Source Kontrolü Yapın:** `event.data.source === 'kokpit-iframe-bridge'` kontrolü sahte mesajları engeller.
- [x] **Token Saklama:** Token'ı `localStorage` yerine `sessionStorage` veya bellekte tutun (sekme kapandığında otomatik silinir).
- [x] **HTTPS Kullanımı:** Canlı ortamlarda hem ana portalın hem dış uygulamanın HTTPS üzerinden çalışması gerekmektedir.

---

## 7. Sorun Giderme (Troubleshooting)

| Belirti / Hata | Olası Neden | Çözüm |
|---|---|---|
| **Mesaj hiç gelmiyor** | "Kokpit Kimliği Gönder" seçeneği kapalıdır. | Ana sistem araç ayarlarından toggle'ın mavi/aktif olduğunu doğrulayın. |
| **Iframe boş/gri ekran kalıyor** | Dış sitede `X-Frame-Options` engeli vardır. | Dış sitenizin sunucusunda `X-Frame-Options: SAMEORIGIN` başlığını kaldırın. |
| **Origin hatası veriyor** | Port veya protokol uyuşmazlığı vardır. | `console.log(event.origin)` ile gelen gerçek origin'i yazdırıp izin verilenler listesine ekleyin. |
| **Token ile API'ye istek atıldığında 401 dönüyor** | Keycloak token süresi dolmuştur (ortalama 5-15 dk). | Sayfayı yenileyin; iframe yenilendiğinde ana sistem güncel geçerli token'ı otomatik tekrar iletir. |

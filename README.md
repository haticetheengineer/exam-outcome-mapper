# Sınav ÖÇ Eşleştirme Sistemi

Akreditasyon raporu için sınav sorularını öğrenim çıktılarıyla eşleştiren web uygulaması.

---

## 🚀 Streamlit Cloud'a Yükleme (Adım Adım)

### 1. GitHub Hesabı Aç
- github.com → Sign Up

### 2. Yeni Repository Oluştur
- GitHub'da sağ üst → "+" → "New repository"
- İsim: `sinav-oc-sistemi`
- Public seç
- "Create repository" tıkla

### 3. Dosyaları Yükle
"uploading an existing file" linkine tıkla, şu dosyaları yükle:
- `app.py`
- `requirements.txt`

### 4. Streamlit Cloud Hesabı
- share.streamlit.io → Google ile giriş yap

### 5. Uygulamayı Deploy Et
- "New app" tıkla
- Repository: `sinav-oc-sistemi`
- Main file: `app.py`
- "Deploy!" tıkla

### 6. API Anahtarı Ekle (ÖNEMLİ)
- App ayarları → "Secrets" sekmesi
- Şunu yapıştır:
```
ANTHROPIC_API_KEY = "sk-ant-..."
```
- console.anthropic.com'dan API anahtarı alın

---

## 📁 Proje Yapısı

```
sinav-oc-sistemi/
├── app.py              # Ana uygulama
├── requirements.txt    # Kütüphaneler
└── .streamlit/
    └── secrets.toml    # API anahtarı (local için)
```

---

## 🎯 Kullanım

1. **Sınav Yükle** → TXT, DOCX veya PDF
2. **ÖÇ Tanımla** → Elle yaz veya fotoğraf yükle
3. **Eşleştir** → Her soruya ÖÇ ata
4. **Excel İndir** → 3 sayfalık akreditasyon raporu

---

## 📊 Excel Çıktısı

- **Sayfa 1:** Soru-ÖÇ Eşleştirme (soru metni, ÖÇ, zorluk, başarı %)
- **Sayfa 2:** ÖÇ Özet (ortalama başarı, değerlendirme)
- **Sayfa 3:** Öğrenci Sonuçları (öğrenci bazında puan)

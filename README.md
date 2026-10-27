# İlan Çalışma Alanı — Kurulum

Revy'den aldığın ilan listeleriyle müşteri dosyası bazında çalışan, Karma App'ten tamamen ayrı özel uygulama.

## Akış
1. **Veri**: müşteri dosyası aç, Revy'den satılık (ve kiralık) Excel'i yükle.
2. **Bölge özeti**: ilçe/mahalle ort. m² fiyat, ort. m² kira, kira çarpanı, brüt getiri (olağan dışı fiyatlılar ortalamaya girmez).
3. **1. Eleme**: ilanlara bak, ✔ beğen / ✖ ele, not yaz.
4. **Danışman görüşmesi**: beğendiklerin için net fiyat, kira, aidat, tapu vb. kaydet, 2. elemeye al.
5. **2. Eleme**: net rakamlarla yeniden hesaplanmış karşılaştırma; en fazla 3 ilanı Final'e al.
6. **Final 3**: müşteriye gönderilecek HTML sunum sayfası + tüm çalışmanın Excel'i.

## Kurulum (bir kez)
1. **Yeni Supabase projesi** aç (mevcut Karma App projesinden ayrı). SQL Editor'de `schema.sql` dosyasının tamamını çalıştır.
2. **Yeni GitHub reposu** aç (ör. `ilan-calisma`, private) ve bu klasördeki dosyaları (app.py, db.py, excel_oku.py, hesap.py, sunum.py, requirements.txt) reponun ana dizinine yükle.
3. **Streamlit Cloud**: New app → repo → `app.py`. Settings → Secrets:
```
SUPABASE_URL = "https://xxxx.supabase.co"
SUPABASE_KEY = "service_role anahtarı"
APP_PASSWORD = "seçtiğin şifre"
```
Anahtarları asla repoya yazma, sadece Secrets'a gir.

## Hesap yöntemi
- m² fiyatı = fiyat ÷ m². Referans: mahallede yeterli ilan varsa mahalle, yoksa ilçe (eşikler soldaki Hesap ayarları'nda).
- Aykırı ilan: referans grubunun Q1 − k·IQR ... Q3 + k·IQR aralığı dışı (k=1,5). Listede ⚠ ile işaretlenir, ortalamaya girmez.
- Tahmini kira = bölge ort. m² kira × m². Kira çarpanı = fiyat ÷ aylık kira (ay). Brüt getiri % = 12 ÷ çarpan × 100.
- Danışmandan net fiyat/kira girilirse 2. Eleme ve Final bu rakamlarla hesaplanır.

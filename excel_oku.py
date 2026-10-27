"""Revy ilan listesi (Excel) -> sisteme yazılacak standart tablo."""
import pandas as pd

# Revy sütun adı -> iç ad
SUTUNLAR = {
    "İlan Url": "url", "İlan tarihi": "ilan_tarihi", "Mülk türü": "mulk_turu", "İşlem tipi": "tur",
    "M2": "m2", "Fiyat": "fiyat", "İlk Fiyat": "ilk_fiyat", "İlçe": "ilce", "Mahalle": "mahalle",
    "Oda sayısı": "oda", "Bulunduğu kat": "kat", "Toplam kat sayısı": "toplam_kat", "Bina Yaşı": "bina_yasi",
    "Site içerisinde": "site_icinde", "Site Adı": "site_adi", "Kullanım Durumu": "kullanim",
    "İlan sahibi türü": "sahip_turu", "İlan sahibi": "ilan_sahibi", "Ofis": "ofis",
    "İlan Yayın Süresi": "yayin_suresi", "İlan Kaynağı": "kaynak", "İlan Başlığı": "baslik",
}
ZORUNLU = ["url", "tur", "m2", "fiyat", "ilce"]
METIN = ["mulk_turu", "ilce", "mahalle", "oda", "kat", "bina_yasi", "site_icinde", "site_adi", "kullanim",
         "sahip_turu", "ilan_sahibi", "ofis", "kaynak", "baslik"]
SAYI = ["m2", "fiyat", "ilk_fiyat", "toplam_kat", "yayin_suresi"]
TUR = {"satılık": "satilik", "satilik": "satilik", "kiralık": "kiralik", "kiralik": "kiralik"}


def _tr_kucuk(s):
    return str(s or "").replace("İ", "i").replace("I", "ı").lower().strip()


def oku(dosya):
    """Excel dosyasını okur. Dönüş: (DataFrame, rapor sözlüğü). Hata durumunda ValueError."""
    ham = pd.read_excel(dosya)
    ham.columns = [str(c).strip() for c in ham.columns]
    eksik = [rev for rev, ic in SUTUNLAR.items() if ic in ZORUNLU and rev not in ham.columns]
    if eksik:
        raise ValueError("Excel'de şu sütunlar yok: " + ", ".join(eksik) + ". Revy'den 'Excel'e Aktar' ile alınan ilan listesi olmalı.")
    df = ham[[c for c in SUTUNLAR if c in ham.columns]].rename(columns=SUTUNLAR).copy()
    n0 = len(df)
    for c in SAYI:
        if c in df:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in METIN:
        if c in df:
            df[c] = df[c].astype("string").str.strip()
    df["tur"] = df["tur"].map(lambda v: TUR.get(_tr_kucuk(v)))
    if "ilan_tarihi" in df:
        df["ilan_tarihi"] = pd.to_datetime(df["ilan_tarihi"], dayfirst=True, errors="coerce").dt.strftime("%Y-%m-%d")
    df["url"] = df["url"].astype("string").str.strip()
    gecerli = df["url"].notna() & df["url"].str.startswith("http", na=False) & df["tur"].notna() \
        & (df["m2"] > 0) & (df["fiyat"] > 0) & df["ilce"].notna()
    atlanan = int((~gecerli).sum())
    df = df[gecerli].drop_duplicates(subset=["tur", "url"]).reset_index(drop=True)
    for c in SUTUNLAR.values():
        if c not in df:
            df[c] = None
    rapor = {
        "okunan": n0, "atlanan": atlanan, "tekrar": int(gecerli.sum()) - len(df), "kayit": len(df),
        "satilik": int((df["tur"] == "satilik").sum()), "kiralik": int((df["tur"] == "kiralik").sum()),
    }
    return df, rapor

"""Supabase erişimi. Anahtarlar yalnız Streamlit secrets'ta durur (SUPABASE_URL, SUPABASE_KEY = service_role)."""
import json

import pandas as pd
import streamlit as st

SAYFA = 1000
PARCA = 300


@st.cache_resource
def _istemci():
    from supabase import create_client
    url = str(st.secrets.get("SUPABASE_URL", "")).strip().strip('"').strip("'")
    key = str(st.secrets.get("SUPABASE_KEY", "")).strip().strip('"').strip("'")
    if not url.startswith("https://") or not url.isascii():
        st.error("SUPABASE_URL hatalı. Secrets'ta https://....supabase.co biçiminde, düz tırnak içinde olmalı.")
        st.stop()
    if len(key) < 40 or not key.isascii() or " " in key:
        st.error("SUPABASE_KEY hatalı: gerçek anahtar yerine örnek metin, boşluk ya da Türkçe/özel karakter var. "
                 "Supabase → Project Settings → API Keys → Legacy → service_role → Reveal → Copy ile kopyalayıp Secrets'a yeniden yapıştırın.")
        st.stop()
    return create_client(url, key)


def dosyalar():
    return _istemci().table("musteri_dosyalari").select("*").order("olusturma", desc=True).execute().data or []


def dosya_olustur(alanlar):
    r = _istemci().table("musteri_dosyalari").insert(alanlar).execute().data
    return r[0] if r else None


def dosya_guncelle(dosya_id, alanlar):
    _istemci().table("musteri_dosyalari").update(alanlar).eq("id", dosya_id).execute()


def dosya_sil(dosya_id):
    _istemci().table("musteri_dosyalari").delete().eq("id", dosya_id).execute()


def ilanlar_getir(dosya_id):
    cikti, bas = [], 0
    while True:
        r = (_istemci().table("ilanlar").select("*").eq("dosya_id", dosya_id)
             .order("id").range(bas, bas + SAYFA - 1).execute().data or [])
        cikti += r
        if len(r) < SAYFA:
            return cikti
        bas += SAYFA


def ilanlar_yaz(dosya_id, df):
    """Excel'den gelen ilanları ekler/günceller (aynı dosya + tür + url → fiyat vb. güncellenir; eleme aşaması, not ve
    görüşme bilgisi korunur)."""
    kayitlar = json.loads(df.assign(dosya_id=str(dosya_id)).to_json(orient="records", force_ascii=False))
    for i in range(0, len(kayitlar), PARCA):
        _istemci().table("ilanlar").upsert(kayitlar[i:i + PARCA], on_conflict="dosya_id,tur,url").execute()
    return len(kayitlar)


def ilan_guncelle(ilan_id, alanlar):
    _istemci().table("ilanlar").update(alanlar).eq("id", int(ilan_id)).execute()


def adaylari_sil(dosya_id, tur=None):
    """Yalnız henüz işlem görmemiş ('aday') ilanları siler."""
    q = _istemci().table("ilanlar").delete().eq("dosya_id", dosya_id).eq("asama", "aday")
    if tur:
        q = q.eq("tur", tur)
    q.execute()


def ilan_toplu_guncelle(ilan_idleri, alanlar):
    ids = [int(i) for i in ilan_idleri]
    for i in range(0, len(ids), 200):
        _istemci().table("ilanlar").update(alanlar).in_("id", ids[i:i + 200]).execute()

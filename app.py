"""İlan Çalışma Alanı — müşteri dosyalarıyla ilan eleme (Revy Excel → 1. eleme → danışman görüşmesi → 2. eleme → final 3)."""
import hmac
import re
from datetime import date, datetime, timezone

import numpy as np
import pandas as pd
import streamlit as st

import db
import excel_oku
import hesap
import sunum

st.set_page_config(page_title="İlan Çalışma Alanı", page_icon="🗂️", layout="wide")

ASAMA_AD = {"aday": "Aday", "birinci": "1. eleme geçti", "ikinci": "2. eleme", "final": "Final", "elendi": "Elendi"}
TAPU = ["", "Kat mülkiyeti", "Kat irtifakı", "Hisseli", "Arsa / müstakil", "Diğer"]
KREDI = ["", "Uygun", "Uygun değil", "Bilinmiyor"]
KIRACI = ["", "Boş", "Kiracılı", "Mülk sahibi oturuyor"]
FINAL_SAYISI = 3
KALKTI = "kalkti"   # elenen_asama değeri: ilan yayından kalkmış

SIRALAMA = {
    "Bölge ortalamasına göre en ucuz": ("fark_pct", True),
    "Kira çarpanı (düşükten yükseğe)": ("carpan", True),
    "Brüt getiri (yüksekten düşüğe)": ("getiri_pct", False),
    "Fiyat (düşükten yükseğe)": ("fiyat", True),
    "Fiyat düşüşü (çoktan aza)": ("dusus_pct", False),
    "Yayın süresi (uzundan kısaya)": ("yayin_suresi", False),
}


# ── Giriş ──────────────────────────────────────────────────────────────
def _giris():
    sifre = st.secrets.get("APP_PASSWORD", "")
    if not sifre:
        st.error("APP_PASSWORD tanımlı değil. Streamlit → Settings → Secrets bölümüne ekleyin.")
        st.stop()
    if st.session_state.get("ok"):
        return
    st.title("🗂️ İlan Çalışma Alanı")
    with st.form("giris"):
        p = st.text_input("Şifre", type="password")
        if st.form_submit_button("Giriş"):
            if hmac.compare_digest(str(p), str(sifre)):
                st.session_state["ok"] = True
                st.rerun()
            st.error("Şifre yanlış.")
    st.stop()


# ── Veri ───────────────────────────────────────────────────────────────
def _yenile():
    st.session_state["v"] = st.session_state.get("v", 0) + 1
    _yukle.clear()
    _hesap.clear()


@st.cache_data(ttl=300, show_spinner="İlanlar yükleniyor...")
def _yukle(dosya_id, v):
    return pd.DataFrame(db.ilanlar_getir(dosya_id))


@st.cache_data(ttl=300, show_spinner="Bölge ortalamaları hesaplanıyor...")
def _hesap(dosya_id, v, mn_s, mn_k, k):
    df = _yukle(dosya_id, v)
    if df.empty or "tur" not in df:
        bos = pd.DataFrame()
        return bos, bos, bos
    sat, kira, oz = hesap.hesapla(df, mn_s, mn_k, k)
    return hesap.net_hesapla(sat), kira, oz


def _simdi():
    return datetime.now(timezone.utc).isoformat()


def _ilan_guncelle(ilan_id, **alanlar):
    alanlar["guncelleme"] = _simdi()
    db.ilan_guncelle(ilan_id, alanlar)


def _mesaj():
    m = st.session_state.pop("_mesaj", None)
    if m:
        st.success(m)


def _kisa(x, n=0):
    return None if x is None or (isinstance(x, float) and np.isnan(x)) else round(float(x), n)


# ── Müşteri dosyası ────────────────────────────────────────────────────
def _yeni_dosya_kaydet():
    ad = (st.session_state.get("yd_ad") or "").strip()
    if not ad:
        st.session_state["_hata"] = "Müşteri adı gerekli."
        return
    r = db.dosya_olustur({
        "ad": ad, "telefon": st.session_state.get("yd_tel") or None, "hedef": st.session_state.get("yd_hedef") or None,
        "butce_min": st.session_state.get("yd_bmin") or None, "butce_max": st.session_state.get("yd_bmax") or None,
        "bolgeler": st.session_state.get("yd_bolge") or None, "notlar": st.session_state.get("yd_not") or None,
    })
    if r:
        st.session_state["dosya_sec"] = r["id"]
        st.session_state["_mesaj"] = f"'{ad}' dosyası oluşturuldu. Veri sekmesinden Revy Excel'ini yükleyebilirsin."


def _kenar_cubugu():
    dos = db.dosyalar()
    st.sidebar.title("🗂️ İlan Çalışma Alanı")
    with st.sidebar.expander("➕ Yeni müşteri dosyası", expanded=not dos):
        with st.form("yeni_dosya", clear_on_submit=True):
            st.text_input("Müşteri adı *", key="yd_ad")
            st.text_input("Telefon", key="yd_tel")
            st.text_input("Hedef (ör. yatırım, 1+1, kira getirisi)", key="yd_hedef")
            a, b = st.columns(2)
            a.number_input("Bütçe min (TL)", min_value=0, step=100000, key="yd_bmin")
            b.number_input("Bütçe max (TL)", min_value=0, step=100000, key="yd_bmax")
            st.text_input("Bölgeler", key="yd_bolge")
            st.text_area("Notlar", key="yd_not", height=70)
            st.form_submit_button("Dosyayı oluştur", on_click=_yeni_dosya_kaydet)
    if st.session_state.pop("_hata", None):
        st.sidebar.error("Müşteri adı gerekli.")
    if not dos:
        st.info("Başlamak için soldan bir müşteri dosyası oluştur.")
        st.stop()
    ids = [d["id"] for d in dos]
    if st.session_state.get("dosya_sec") not in ids:
        st.session_state["dosya_sec"] = ids[0]
    etiket = {d["id"]: d["ad"] for d in dos}
    secim = st.sidebar.selectbox("Müşteri dosyası", ids, format_func=lambda i: etiket[i], key="dosya_sec")
    dosya = next(d for d in dos if d["id"] == secim)
    with st.sidebar.expander("⚙️ Hesap ayarları"):
        mn_s = st.number_input("Min. satılık ilan (mahalle referansı için)", 3, 30, hesap.MIN_N_SATILIK, key="ay_ms")
        mn_k = st.number_input("Min. kiralık ilan (mahalle referansı için)", 2, 30, hesap.MIN_N_KIRALIK, key="ay_mk")
        k = st.number_input("Aykırı sınırı (IQR çarpanı)", 0.5, 4.0, hesap.IQR_K, 0.25, key="ay_k",
                            help="Küçük değer = daha fazla ilan aykırı sayılır. 1,5 standart yöntemdir.")
    if st.sidebar.button("Çıkış"):
        st.session_state.clear()
        st.rerun()
    return dosya, int(mn_s), int(mn_k), float(k)


# ── Sekme: Veri ────────────────────────────────────────────────────────
def _sekme_veri(dosya, df):
    st.subheader(f"{dosya['ad']}")
    bilgi = [x for x in (dosya.get("telefon"), dosya.get("hedef"), dosya.get("bolgeler")) if x]
    if dosya.get("butce_min") or dosya.get("butce_max"):
        bilgi.append(f"Bütçe {sunum.tl(dosya.get('butce_min') or 0)} – {sunum.tl(dosya.get('butce_max') or 0)}")
    if bilgi:
        st.caption(" · ".join(bilgi))
    if dosya.get("notlar"):
        st.write(dosya["notlar"])
    if not df.empty:
        t = df.pivot_table(index="tur", columns="asama", values="id", aggfunc="count", fill_value=0)
        t.index = [("Satılık" if i == "satilik" else "Kiralık") for i in t.index]
        t.columns = [ASAMA_AD.get(c, c) for c in t.columns]
        t["Toplam"] = t.sum(axis=1)
        st.dataframe(t, use_container_width=True)
    st.markdown("**Revy ilan listesi yükle**")
    st.caption("Revy'de kriterini ara, 'Excel'e Aktar' de, dosyayı buraya bırak. Satılık ve kiralık dosyaları ayrı ayrı yükle (tür, "
               "dosyadaki 'İşlem tipi' sütunundan otomatik anlaşılır). Aynı ilan tekrar yüklenirse fiyatı güncellenir, eleme durumun korunur.")
    yuk = st.file_uploader("Excel dosyaları", type=["xlsx"], accept_multiple_files=True,
                           key=f"up_{dosya['id']}_{st.session_state.get('up_n', 0)}")
    parcalar = []
    for f in yuk or []:
        try:
            d, r = excel_oku.oku(f)
        except Exception as e:      # noqa: BLE001
            st.error(f"{f.name}: {e}")
            continue
        st.info(f"{f.name}: {r['kayit']} ilan ({r['satilik']} satılık, {r['kiralik']} kiralık)"
                + (f" · {r['atlanan']} satır eksik veri nedeniyle atlandı" if r["atlanan"] else "")
                + (f" · {r['tekrar']} tekrar" if r["tekrar"] else ""))
        parcalar.append(d)
    kalkan = pd.DataFrame()
    if parcalar and not df.empty:
        yeni_url = pd.concat(parcalar).groupby("tur")["url"].apply(set).to_dict()
        kalkan = pd.concat([df[(df["tur"] == t) & (df["asama"] != "elendi") & (~df["url"].isin(u))] for t, u in yeni_url.items()])
    kalkan_sil = False
    if len(kalkan):
        say = kalkan["asama"].map(ASAMA_AD).value_counts().to_dict()
        st.warning(f"Dosyada olup bu yeni listede olmayan {len(kalkan)} ilan var, yayından kalkmış olabilir "
                   f"({', '.join(f'{k}: {v}' for k, v in say.items())}).")
        kalkan_sil = st.checkbox("Bunları 'yayından kalktı' say ve listelerden çıkar (Elenenler bölümünden geri alınabilir)",
                                 value=True, key="kalkan_sil",
                                 help="Yeni liste farklı bir kriterle alındıysa bu kutuyu kapat.")
    if parcalar and st.button("Müşteri dosyasına ekle", type="primary", key="ekle"):
        tum = pd.concat(parcalar).drop_duplicates(subset=["tur", "url"], keep="last")
        n = db.ilanlar_yaz(dosya["id"], tum)
        if kalkan_sil and len(kalkan):
            db.ilan_toplu_guncelle(list(kalkan["id"]), {"asama": "elendi", "elenen_asama": KALKTI, "guncelleme": _simdi()})
        st.session_state["up_n"] = st.session_state.get("up_n", 0) + 1
        st.session_state["_mesaj"] = f"{n} ilan eklendi/güncellendi." + (f" {len(kalkan)} ilan yayından kalktı olarak çıkarıldı." if kalkan_sil and len(kalkan) else "")
        _yenile()
        st.rerun()

    with st.expander("Müşteri bilgilerini düzenle"):
        with st.form(f"dz_{dosya['id']}"):
            ad = st.text_input("Müşteri adı", dosya["ad"])
            tel = st.text_input("Telefon", dosya.get("telefon") or "")
            hedef = st.text_input("Hedef", dosya.get("hedef") or "")
            a, b = st.columns(2)
            bmin = a.number_input("Bütçe min (TL)", 0, value=int(dosya.get("butce_min") or 0), step=100000)
            bmax = b.number_input("Bütçe max (TL)", 0, value=int(dosya.get("butce_max") or 0), step=100000)
            bol = st.text_input("Bölgeler", dosya.get("bolgeler") or "")
            nt = st.text_area("Notlar", dosya.get("notlar") or "")
            if st.form_submit_button("Kaydet"):
                db.dosya_guncelle(dosya["id"], {"ad": ad.strip() or dosya["ad"], "telefon": tel or None, "hedef": hedef or None,
                                                "butce_min": bmin or None, "butce_max": bmax or None, "bolgeler": bol or None, "notlar": nt or None})
                st.session_state["_mesaj"] = "Müşteri bilgileri kaydedildi."
                st.rerun()
    with st.expander("Listeyi temizle / dosyayı sil"):
        st.caption("Temizleme yalnızca henüz işlem görmemiş 'Aday' ilanları siler; 1. eleme, görüşme, 2. eleme ve final ilanları korunur.")
        tur = st.radio("Silinecek tür", ["Satılık", "Kiralık", "İkisi"], horizontal=True, key="tmz_tur")
        if st.button("Aday ilanları temizle", key="tmz"):
            db.adaylari_sil(dosya["id"], {"Satılık": "satilik", "Kiralık": "kiralik"}.get(tur))
            st.session_state["_mesaj"] = "Aday ilanlar temizlendi."
            _yenile()
            st.rerun()
        st.divider()
        onay = st.checkbox("Bu müşteri dosyasını ve tüm ilanlarını kalıcı olarak silmek istiyorum", key="sil_onay")
        if st.button("Dosyayı sil", disabled=not onay, key="sil"):
            db.dosya_sil(dosya["id"])
            st.session_state.pop("dosya_sec", None)
            _yenile()
            st.rerun()


# ── Sekme: Bölge özeti ─────────────────────────────────────────────────
def _sekme_ozet(oz, mn_s, mn_k):
    if oz.empty:
        st.info("Önce satılık ve kiralık Revy listesini yükle.")
        return
    st.caption(f"Ortalamalar olağan dışı yüksek/düşük fiyatlı ilanlar hariç hesaplanır. Mahallede en az {mn_s} satılık ve {mn_k} kiralık ilan "
               "yoksa 'Yeterli veri' işaretsizdir; ilan satırlarında bu durumda ilçe ortalaması kullanılır.")
    a, b = st.columns([1, 2])
    sev = a.radio("Düzey", ["İlçe", "Mahalle"], horizontal=True, key="oz_sev")
    if sev == "İlçe":
        d = oz[oz["Seviye"] == "ilçe"]
    else:
        d = oz[oz["Seviye"] == "mahalle"]
        ilceler = sorted(d["İlçe"].dropna().unique())
        sec = b.multiselect("İlçe", ilceler, key="oz_ilce")
        if sec:
            d = d[d["İlçe"].isin(sec)]
    d = d.sort_values("Satılık ilan", ascending=False).drop(columns=["Seviye"] + (["Mahalle"] if sev == "İlçe" else []))
    st.dataframe(d, use_container_width=True, hide_index=True, column_config={
        "Ort. m² fiyat": st.column_config.NumberColumn(format="%.0f"), "Medyan m² fiyat": st.column_config.NumberColumn(format="%.0f"),
        "Ort. m² kira": st.column_config.NumberColumn(format="%.0f"), "Kira çarpanı (ay)": st.column_config.NumberColumn(format="%.0f"),
        "Brüt getiri %": st.column_config.NumberColumn(format="%.1f"),
    })


# ── Sekme: 1. Eleme ────────────────────────────────────────────────────
def _yas_aralik(v):
    """Revy bina yaşı metni → (en az, en çok). '0', '5', '6-10', '11-15', '21 ve üzeri' ..."""
    s = hesap.tr_kucuk(v)
    m = re.findall(r"\d+", s)
    if not m:
        return None, None
    lo = int(m[0])
    if len(m) >= 2:
        return lo, int(m[1])
    return lo, (None if ("ve üzeri" in s or "+" in s or "üzeri" in s) else lo)


def _kat_tipi(v):
    s = hesap.tr_kucuk(v)
    if not s:
        return ""
    if "çatı" in s or "cati" in s or "teras" in s:
        return "cati"
    if any(x in s for x in ("zemin", "giriş", "bahçe", "bodrum", "kot", "yüksek giriş")):
        return "zemin"
    return "normal"


def _filtre_df(sat, anahtar):
    with st.expander("🔎 Filtreler", expanded=True):
        c = st.columns([2, 2, 1, 1])
        ilce = c[0].multiselect("İlçe", sorted(sat["ilce"].dropna().unique()), key=f"{anahtar}_ilce")
        d = sat[sat["ilce"].isin(ilce)] if ilce else sat
        mah = c[1].multiselect("Mahalle", sorted(d["mahalle"].dropna().unique()), key=f"{anahtar}_mah")
        fmin = c[2].number_input("Fiyat min (TL)", 0, step=250000, key=f"{anahtar}_fmin")
        fmax = c[3].number_input("Fiyat max (TL)", 0, step=250000, key=f"{anahtar}_fmax")

        c = st.columns([2, 2, 2, 2])
        bina = c[0].multiselect("Bina yaşı grubu", ["Yeni (0–10 yaş)", "Eski (11 yaş ve üzeri)"], key=f"{anahtar}_bina",
                                help="Boş bırakırsan hepsi gelir. 11-15 yaş eski binaya dahildir.")
        yaslar = sorted(sat["bina_yasi"].dropna().astype(str).unique(), key=lambda x: (_yas_aralik(x)[0] is None, _yas_aralik(x)[0] or 0))
        yas = c[1].multiselect("Bina yaşı (tam seçim)", yaslar, key=f"{anahtar}_yas")
        site = c[2].radio("Site içinde", ["Hepsi", "Evet", "Hayır"], horizontal=True, key=f"{anahtar}_site")
        kul = c[3].multiselect("Kullanım durumu", sorted(sat["kullanim"].dropna().unique()), key=f"{anahtar}_kul")

        c = st.columns([2, 2, 3])
        cati = c[0].checkbox("Çatı / en üst katları gizle", value=True, key=f"{anahtar}_cati",
                             help="'Çatı' yazanlar ve kat numarası binanın kat sayısına eşit olanlar (3 katlı binanın 3. katı gibi)")
        zemin = c[1].checkbox("Zemin / giriş / bahçe / bodrum katlarını gizle", value=True, key=f"{anahtar}_zemin")
        katlar = sorted(sat["kat"].dropna().astype(str).unique(), key=lambda x: (not x.split(".")[0].isdigit(), int(x.split(".")[0]) if x.split(".")[0].isdigit() else 0, x))
        kat_gizle = c[2].multiselect("Başka kat(lar)ı gizle", katlar, key=f"{anahtar}_katg")

        c = st.columns([2, 2, 2, 2])
        gizle = c[0].checkbox("Aykırı fiyatlıları gizle", key=f"{anahtar}_gizle")
        alti = c[1].checkbox("Yalnız bölge ort. altındakiler", key=f"{anahtar}_alt")
        sira = c[2].selectbox("Sıralama", list(SIRALAMA), key=f"{anahtar}_sira")
        ara = c[3].text_input("Ara (başlık, site, ofis)", key=f"{anahtar}_ara")

    n0 = len(sat)
    if ilce:
        sat = sat[sat["ilce"].isin(ilce)]
    if mah:
        sat = sat[sat["mahalle"].isin(mah)]
    if fmin:
        sat = sat[sat["fiyat"] >= fmin]
    if fmax:
        sat = sat[sat["fiyat"] <= fmax]
    if kul:
        sat = sat[sat["kullanim"].isin(kul)]
    if gizle:
        sat = sat[~sat["aykiri"]]
    if alti:
        sat = sat[sat["fark_pct"] <= 0]
    if site != "Hepsi":
        sat = sat[sat["site_icinde"].fillna("").map(hesap.tr_kucuk) == site.lower()]
    if yas:
        sat = sat[sat["bina_yasi"].astype(str).isin(yas)]
    if bina and len(bina) == 1:
        ar = sat["bina_yasi"].map(_yas_aralik)
        yeni = ar.map(lambda t: t[1] is not None and t[1] <= 10)
        eski = ar.map(lambda t: t[0] is not None and t[0] >= 11)
        sat = sat[yeni] if bina[0].startswith("Yeni") else sat[eski]
    tip = sat["kat"].map(_kat_tipi)
    if cati:
        kat_no = pd.to_numeric(sat["kat"].astype(str).str.extract(r"^(\d+)\.")[0], errors="coerce")
        ust = (kat_no == sat["toplam_kat"]) & (sat["toplam_kat"] >= 2)      # 3 katlı binanın 3. katı gibi
        sat = sat[(tip != "cati") & ~ust]
        tip = tip[(tip != "cati") & ~ust]
    if zemin:
        sat = sat[tip != "zemin"]
    if kat_gizle:
        sat = sat[~sat["kat"].astype(str).isin(kat_gizle)]
    if ara:
        a = ara.lower()
        sat = sat[sat[["baslik", "site_adi", "ofis", "ilan_sahibi"]].fillna("").agg(" ".join, axis=1).str.lower().str.contains(a, regex=False)]
    kol, art = SIRALAMA[sira]
    sig = (tuple(ilce), tuple(mah), fmin, fmax, tuple(kul), gizle, alti, sira, ara, site, tuple(yas), tuple(bina), cati, zemin, tuple(kat_gizle))
    return sat.sort_values(kol, ascending=art, na_position="last"), sig


def _durum(fark, aykiri):
    if aykiri:
        return "⚠ aykırı"
    if pd.isna(fark):
        return "—"
    return "🟢 ucuz" if fark <= -10 else ("🔴 pahalı" if fark >= 10 else "🟡 normal")


def _kat_metni(r):
    k = str(r.get("kat") or "").replace(" Kat", "").replace(".", "").strip()
    t = r.get("toplam_kat")
    return f"{k} / {int(t)}" if k and pd.notna(t) else k


def _liste_df(d, secenek=None):
    df = pd.DataFrame({
        "Durum": [_durum(f, a) for f, a in zip(d["fark_pct"], d["aykiri"])],
        "İlan": d["url"].values, "İlçe": d["ilce"].values,
        "Mahalle": d["mahalle"].fillna("").str.replace(r"\s*Mah\.?$", "", regex=True).values,
        "m²": d["m2"].values, "Kat": [_kat_metni(r) for _, r in d.iterrows()], "Bina yaşı": d["bina_yasi"].astype(str).replace("nan", "").values,
        "Fiyat": d["fiyat"].values, "m² fiyat": d["m2_fiyat"].values, "Bölge ort. m²": d["ref_ort"].values, "Fark %": d["fark_pct"].values,
        "Düşüş %": d["dusus_pct"].values, "Tahmini kira": d["tahmini_kira"].values, "Çarpan (ay)": d["carpan"].values, "Getiri %": d["getiri_pct"].values,
        "Kullanım": d["kullanim"].values, "Yayın (gün)": d["yayin_suresi"].values,
        "Ofis / ilan sahibi": d["ofis"].fillna(d["ilan_sahibi"]).values, "Kaynak": d["kaynak"].values,
    }, index=d["id"].values)
    # veri yoksa boş kolonları gizle (ör. kiralık listesi yüklenmemişse kira/çarpan/getiri)
    for c in ("Tahmini kira", "Çarpan (ay)", "Getiri %", "Düşüş %", "Yayın (gün)"):
        if df[c].isna().all() or (c == "Düşüş %" and (df[c].fillna(0) == 0).all()):
            df = df.drop(columns=c)
    return df


def _kolon_ayar():
    n = st.column_config.NumberColumn
    return {
        "Beğen": st.column_config.CheckboxColumn("✔", help="Beğen", width=50, pinned=True),
        "Ele": st.column_config.CheckboxColumn("✖", help="Ele", width=50, pinned=True),
        "Durum": st.column_config.TextColumn("Durum", width=95, help="Bölge ortalamasına göre m² fiyatı: 🟢 %10+ ucuz, 🔴 %10+ pahalı"),
        "İlan": st.column_config.LinkColumn("İlan", display_text="aç ↗", width=60),
        "İlçe": st.column_config.TextColumn(width=85), "Mahalle": st.column_config.TextColumn(width=130),
        "m²": n(format="%.0f", width=55), "Kat": st.column_config.TextColumn(width=65), "Bina yaşı": st.column_config.TextColumn("Yaş", width=55),
        "Fiyat": n("Fiyat (TL)", format="localized", width=100), "m² fiyat": n("m² fiyat", format="localized", width=85),
        "Bölge ort. m²": n("Bölge ort.", format="localized", width=85), "Fark %": n("Fark", format="%+.1f%%", width=70),
        "Düşüş %": n("Düşüş", format="%.1f%%", width=65), "Tahmini kira": n("Tah. kira", format="localized", width=80),
        "Çarpan (ay)": n("Çarpan", format="%.0f ay", width=75), "Getiri %": n("Getiri", format="%.1f%%", width=70),
        "Kullanım": st.column_config.TextColumn(width=85), "Yayın (gün)": n("Gün", format="%.0f", width=50),
        "Ofis / ilan sahibi": st.column_config.TextColumn("Ofis", width=140), "Kaynak": st.column_config.TextColumn(width=90),
        "Not": st.column_config.TextColumn("Not", width=200),
        "Geri al": st.column_config.CheckboxColumn("↩", help="Geri al", width=50),
        "Kalktı": st.column_config.CheckboxColumn("🗑", help="İlan yayından kalkmış: listeden çıkar", width=50, pinned=True),
        "Neden": st.column_config.TextColumn(width=110),
    }


def _sekme_birinci(sat, dosya):
    aday = sat[sat["asama"] == "aday"]
    if aday.empty:
        st.info("Aday ilan yok. Veri sekmesinden Revy listesini yükle ya da elenenleri geri al.")
    else:
        d, sig = _filtre_df(aday, "b1")
        st.caption(f"{len(d)} / {len(aday)} ilan · {int(d['aykiri'].sum())} aykırı fiyatlı (bölge ortalamasına girmez). "
                   "Beğendiklerini ✔, istemediklerini ✖ ile işaretle, sonra Kaydet.")
        ed = _liste_df(d, None)
        ed["Not"] = d["not_metni"].fillna("").values
        ed.insert(0, "Kalktı", False)
        ed.insert(0, "Ele", False)
        ed.insert(0, "Beğen", False)
        disabled = [c for c in ed.columns if c not in ("Beğen", "Ele", "Kalktı", "Not")]
        son = st.data_editor(ed, column_config=_kolon_ayar(), disabled=disabled, hide_index=True, use_container_width=True,
                             height=560, key=f"ed_b1_{hash(sig)}_{len(ed)}")
        c1, c2 = st.columns([1, 4])
        if c1.button("Kaydet", type="primary", key="b1_kaydet"):
            cift = son.index[(son[["Beğen", "Ele", "Kalktı"]].sum(axis=1) > 1)]
            if len(cift):
                st.error(f"{len(cift)} ilanda birden fazla işaret var (Beğen / Ele / Kalktı), birini bırak.")
            else:
                n = 0
                kalkti = list(son.index[son["Kalktı"]])
                if kalkti:
                    db.ilan_toplu_guncelle(kalkti, {"asama": "elendi", "elenen_asama": KALKTI, "guncelleme": _simdi()})
                    n += len(kalkti)
                for i in son.index[son["Beğen"]]:
                    _ilan_guncelle(i, asama="birinci", not_metni=son.at[i, "Not"] or None)
                    n += 1
                for i in son.index[son["Ele"]]:
                    _ilan_guncelle(i, asama="elendi", elenen_asama="aday", not_metni=son.at[i, "Not"] or None)
                    n += 1
                for i in son.index[~son["Beğen"] & ~son["Ele"] & ~son["Kalktı"]]:
                    if (son.at[i, "Not"] or "") != (ed.at[i, "Not"] or ""):
                        _ilan_guncelle(i, not_metni=son.at[i, "Not"] or None)
                        n += 1
                st.session_state["_mesaj"] = f"{n} ilan güncellendi."
                _yenile()
                st.rerun()
    elen = sat[sat["asama"] == "elendi"]
    if not elen.empty:
        with st.expander(f"Elenenler ({len(elen)})"):
            ed2 = _liste_df(elen, None)
            ed2["Neden"] = elen["elenen_asama"].map(lambda x: "Yayından kalktı" if x == KALKTI else "Elendi").values
            ed2["Not"] = elen["not_metni"].fillna("").values
            ed2.insert(0, "Geri al", False)
            s2 = st.data_editor(ed2, column_config=_kolon_ayar(),
                                disabled=[c for c in ed2.columns if c != "Geri al"], hide_index=True, use_container_width=True,
                                key=f"ed_elen_{len(ed2)}")
            if st.button("Seçilenleri geri al", key="geri_al"):
                for i in s2.index[s2["Geri al"]]:
                    onceki = elen.loc[elen["id"] == i, "elenen_asama"].iloc[0]
                    _ilan_guncelle(i, asama=onceki if onceki in ("aday", "birinci") else "aday", elenen_asama=None)
                _yenile()
                st.rerun()


# ── Sekme: Danışman görüşmesi ──────────────────────────────────────────
def _fl(v):
    try:
        return float(v) if v not in (None, "") else 0.0
    except (TypeError, ValueError):
        return 0.0


def _sekme_gorusme(sat):
    b = sat[sat["asama"] == "birinci"].sort_values(["ilce", "mahalle", "fiyat"])
    if b.empty:
        st.info("1. elemeyi geçen ilan yok. 1. Eleme sekmesinde beğendiğin ilanları işaretle.")
        return
    st.caption(f"{len(b)} ilan danışman görüşmesi bekliyor. Görüşmede öğrendiklerini kaydet; hazır olanı 2. elemeye al.")
    for _, r in b.iterrows():
        g = r["gorusme"] if isinstance(r["gorusme"], dict) else {}
        ozet = f"{r['ilce']} · {r['mahalle']} · {sunum.tl(r['fiyat'])} · {r['m2']:.0f} m²"
        if g:
            ozet += " · ✔ görüşme var"
        with st.expander(ozet):
            cc = st.columns(4)
            cc[0].metric("m² fiyat", sunum.tl(r["m2_fiyat"]), None if pd.isna(r["fark_pct"]) else f"{r['fark_pct']:+.1f}% bölge ort.", delta_color="inverse")
            cc[1].metric("Tahmini kira", sunum.tl(r["tahmini_kira"]))
            cc[2].metric("Çarpan", "—" if pd.isna(r["carpan"]) else f"{r['carpan']:.0f} ay")
            cc[3].metric("Brüt getiri", "—" if pd.isna(r["getiri_pct"]) else f"%{r['getiri_pct']:.1f}")
            st.caption(f"{r.get('baslik') or ''} · {r.get('ofis') or r.get('ilan_sahibi') or ''} · {r.get('kaynak') or ''} · {r.get('kullanim') or ''}")
            st.link_button("İlanı aç", r["url"])
            with st.form(f"g_{r['id']}"):
                a, b2, c = st.columns(3)
                dan = a.text_input("Danışman", g.get("danisman_adi") or (r.get("ilan_sahibi") or ""))
                tel = b2.text_input("Telefon", g.get("danisman_tel") or "")
                tarih = c.date_input("Görüşme tarihi", value=date.fromisoformat(g["tarih"]) if g.get("tarih") else date.today())
                a, b2, c = st.columns(3)
                net = a.number_input("Net (pazarlık sonrası) fiyat, TL", 0.0, step=50000.0, value=_fl(g.get("net_fiyat")), help="0 = bilgi yok, ilan fiyatı kullanılır")
                kira = b2.number_input("Aylık kira (mevcut / beklenen), TL", 0.0, step=500.0, value=_fl(g.get("kira")), help="0 = bilgi yok, bölge tahmini kullanılır")
                aidat = c.number_input("Aidat, TL", 0.0, step=100.0, value=_fl(g.get("aidat")))
                a, b2, c = st.columns(3)
                tapu = a.selectbox("Tapu", TAPU, index=TAPU.index(g.get("tapu")) if g.get("tapu") in TAPU else 0)
                kredi = b2.selectbox("Krediye uygunluk", KREDI, index=KREDI.index(g.get("kredi")) if g.get("kredi") in KREDI else 0)
                kiraci = c.selectbox("Kiracı durumu", KIRACI, index=KIRACI.index(g.get("kiraci")) if g.get("kiraci") in KIRACI else 0)
                neden = st.text_input("Satış nedeni", g.get("neden") or "")
                notlar = st.text_area("Notlar", g.get("notlar") or "", height=80)
                b1, b2, b3, b4 = st.columns(4)
                kaydet = b1.form_submit_button("Kaydet")
                ikinci = b2.form_submit_button("Kaydet ve 2. elemeye al", type="primary")
                ele = b3.form_submit_button("Ele")
                kalkti = b4.form_submit_button("🗑 Yayından kalktı")
            if kalkti:
                _ilan_guncelle(r["id"], asama="elendi", elenen_asama=KALKTI)
                st.session_state["_mesaj"] = "İlan yayından kalktı olarak listeden çıkarıldı."
                _yenile()
                st.rerun()
            if kaydet or ikinci or ele:
                yeni = {"danisman_adi": dan or None, "danisman_tel": tel or None, "tarih": tarih.isoformat(), "net_fiyat": net or None,
                        "kira": kira or None, "aidat": aidat or None, "tapu": tapu or None, "kredi": kredi or None,
                        "kiraci": kiraci or None, "neden": neden or None, "notlar": notlar or None}
                alan = {"gorusme": yeni}
                if ikinci:
                    alan["asama"] = "ikinci"
                if ele:
                    alan.update(asama="elendi", elenen_asama="birinci")
                _ilan_guncelle(r["id"], **alan)
                st.session_state["_mesaj"] = "Görüşme kaydedildi." + (" İlan 2. elemeye alındı." if ikinci else " İlan elendi." if ele else "")
                _yenile()
                st.rerun()


# ── Sekme: 2. Eleme ────────────────────────────────────────────────────
def _net_liste(d):
    df = pd.DataFrame({
        "İlan": d["url"].values, "İlçe": d["ilce"].values, "Mahalle": d["mahalle"].values, "m²": d["m2"].values,
        "İlan fiyatı": d["fiyat"].values, "Net fiyat": d["fiyat_net"].values, "m² fiyat": d["m2_fiyat_net"].values,
        "Bölge ort. m²": d["ref_ort"].values, "Fark %": d["fark_net"].values, "Aylık kira": d["kira_x"].values,
        "Kira kaynağı": d["kira_kaynak"].values, "Çarpan (ay)": d["carpan_net"].values, "Getiri %": d["getiri_net"].values,
        "Aidat": d["gorusme"].map(lambda g: (g or {}).get("aidat")).values,
        "Tapu": d["gorusme"].map(lambda g: (g or {}).get("tapu")).values,
        "Danışman": d["gorusme"].map(lambda g: (g or {}).get("danisman_adi")).values,
    }, index=d["id"].values)
    return df


def _net_ayar():
    n0 = st.column_config.NumberColumn
    return {**_kolon_ayar(), "İlan fiyatı": n0(format="localized"), "Net fiyat": n0(format="localized"), "m² fiyat": n0(format="localized"),
            "Aylık kira": n0(format="localized"), "Aidat": n0(format="localized"),
            "Final": st.column_config.CheckboxColumn("Final'e al ★", width="small"),
            "Ele": st.column_config.CheckboxColumn("Ele ✖", width="small"),
            "Geri": st.column_config.CheckboxColumn("1. elemeye geri ↩", width="small")}


def _sekme_ikinci(sat):
    d = sat[sat["asama"] == "ikinci"].sort_values("carpan_net", na_position="last")
    final_var = int((sat["asama"] == "final").sum())
    if d.empty:
        st.info("2. elemede ilan yok. Danışman görüşmesi sekmesinde 'Kaydet ve 2. elemeye al' ile ilan ekle.")
        return
    st.caption(f"{len(d)} ilan · Rakamlar danışmandan öğrenilen net fiyat ve kira ile yeniden hesaplandı (yoksa ilan fiyatı ve bölge tahmini). "
               f"Final listesi en fazla {FINAL_SAYISI} ilan alır (şu an {final_var}).")
    ed = _net_liste(d)
    ed.insert(0, "Geri", False)
    ed.insert(0, "Ele", False)
    ed.insert(0, "Final", False)
    son = st.data_editor(ed, column_config=_net_ayar(), disabled=[c for c in ed.columns if c not in ("Final", "Ele", "Geri")],
                         hide_index=True, use_container_width=True, key=f"ed_b2_{len(ed)}")
    if st.button("Kaydet", type="primary", key="b2_kaydet"):
        secilen = son.index[son["Final"]]
        if final_var + len(secilen) > FINAL_SAYISI:
            st.error(f"Final listesi en fazla {FINAL_SAYISI} ilan alır. Önce Final sekmesinden birini çıkar ya da daha az seç.")
        elif (son["Final"] & (son["Ele"] | son["Geri"])).any() or (son["Ele"] & son["Geri"]).any():
            st.error("Bir ilanda birden fazla işlem işaretli.")
        else:
            for i in secilen:
                _ilan_guncelle(i, asama="final")
            for i in son.index[son["Ele"]]:
                _ilan_guncelle(i, asama="elendi", elenen_asama="ikinci")
            for i in son.index[son["Geri"]]:
                _ilan_guncelle(i, asama="birinci")
            _yenile()
            st.rerun()


# ── Sekme: Final ───────────────────────────────────────────────────────
def _sekme_final(dosya, sat, oz):
    f = sat[sat["asama"] == "final"].sort_values("carpan_net", na_position="last")
    if f.empty:
        st.info("Final listesi boş. 2. Eleme sekmesinde en fazla 3 ilanı Final'e al.")
    else:
        cols = st.columns(min(len(f), 3))
        for col, (_, r) in zip(cols, f.iterrows()):
            with col:
                with st.container(border=True):
                    st.markdown(f"**{r['ilce']} · {r['mahalle']}**")
                    st.caption(str(r.get("baslik") or "")[:100])
                    st.markdown(f"### {sunum.tl(r['fiyat_net'])}")
                    for k, v in sunum._satirlar(r, danisman=True):
                        st.markdown(f"<div style='display:flex;justify-content:space-between;gap:8px;border-top:1px solid rgba(128,128,128,.25);padding:4px 0;font-size:13.5px'>"
                                    f"<span style='opacity:.7'>{k}</span><span style='text-align:right'>{v}</span></div>", unsafe_allow_html=True)
                    st.link_button("İlanı aç", r["url"])
                    if st.button("Final'den çıkar", key=f"fc_{r['id']}"):
                        _ilan_guncelle(r["id"], asama="ikinci")
                        _yenile()
                        st.rerun()
        st.divider()
        dn = st.checkbox("Sunum sayfasına danışman bilgisi ve notları ekle", key="sunum_dan")
        a, b = st.columns(2)
        a.download_button("⬇ Müşteri sunum sayfası (HTML)", sunum.sunum_html(dosya["ad"], f, dn),
                          file_name=f"{dosya['ad']}_secilen_ilanlar.html".replace(" ", "_"), mime="text/html", key="dl_html")
    parcalar = {"Final": f, "2. Eleme": sat[sat["asama"] == "ikinci"], "1. Eleme": sat[sat["asama"] == "birinci"]}
    st.download_button("⬇ Tüm çalışma (Excel)", sunum.excel(dosya["ad"], parcalar, oz),
                       file_name=f"{dosya['ad']}_calisma.xlsx".replace(" ", "_"),
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_xlsx")


# ── Ana akış ───────────────────────────────────────────────────────────
def main():
    _giris()
    dosya, mn_s, mn_k, k = _kenar_cubugu()
    _mesaj()
    v = st.session_state.get("v", 0)
    df = _yukle(dosya["id"], v)
    sat, kira, oz = _hesap(dosya["id"], v, mn_s, mn_k, k)
    t = st.tabs(["📥 Veri", "📊 Bölge özeti", "1️⃣ 1. Eleme", "📞 Danışman görüşmesi", "2️⃣ 2. Eleme", "⭐ Final 3"])
    with t[0]:
        _sekme_veri(dosya, df)
    if sat.empty:
        for i in range(1, 6):
            with t[i]:
                st.info("Önce Veri sekmesinden Revy ilan listesini yükle.")
        return
    with t[1]:
        _sekme_ozet(oz, mn_s, mn_k)
    with t[2]:
        _sekme_birinci(sat, dosya)
    with t[3]:
        _sekme_gorusme(sat)
    with t[4]:
        _sekme_ikinci(sat)
    with t[5]:
        _sekme_final(dosya, sat, oz)


main()

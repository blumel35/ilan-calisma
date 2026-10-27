"""İlan Çalışma Alanı — müşteri dosyalarıyla ilan eleme (Revy Excel → 1. eleme → danışman görüşmesi → 2. eleme → final 3)."""
import hmac
import html
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
    """Veri değişti (yükleme/temizleme/silme): her şeyi yeniden oku ve hesapla."""
    st.session_state["v"] = st.session_state.get("v", 0) + 1
    st.session_state.pop("durum", None)
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
    return hesap.hesapla(df, mn_s, mn_k, k)


@st.cache_data(ttl=300)
def _dosyalar(n):
    return db.dosyalar()


def _dosyalar_yenile():
    st.session_state["dn"] = st.session_state.get("dn", 0) + 1


DURUM_KOL = ["asama", "elenen_asama", "not_metni", "gorusme"]


def _durum_hazir(dosya_id, v, df):
    """Eleme durumu (aşama, not, görüşme) oturumda tutulur: tıklamada veritabanı yazılır, ekran anında güncellenir."""
    d = st.session_state.get("durum")
    if d and d["dosya"] == dosya_id and d["v"] == v:
        return
    m = {}
    if not df.empty:
        m = df.set_index("id")[DURUM_KOL].to_dict("index")
    st.session_state["durum"] = {"dosya": dosya_id, "v": v, "map": m}


def _uygula(d):
    """Hesap tablosuna güncel eleme durumunu işler."""
    if d is None or d.empty:
        return d
    m = st.session_state["durum"]["map"]
    d = d.copy()
    for c in DURUM_KOL:
        d[c] = pd.Series([m.get(i, {}).get(c) for i in d["id"]], index=d.index, dtype=object)
    return d


def _simdi():
    return datetime.now(timezone.utc).isoformat()


def _guncelle(ids, **alan):
    ids = [int(i) for i in ids]
    if not ids:
        return
    yaz = {**alan, "guncelleme": _simdi()}
    if len(ids) == 1:
        db.ilan_guncelle(ids[0], yaz)
    else:
        db.ilan_toplu_guncelle(ids, yaz)
    m = st.session_state["durum"]["map"]
    for i in ids:
        m.setdefault(i, {}).update({k: v for k, v in alan.items() if k in DURUM_KOL})


def _ilan_guncelle(ilan_id, **alanlar):
    _guncelle([ilan_id], **alanlar)


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
    _dosyalar_yenile()
    if r:
        st.session_state["dosya_sec"] = r["id"]
        st.session_state["_mesaj"] = f"'{ad}' dosyası oluşturuldu. Veri sekmesinden Revy Excel'ini yükleyebilirsin."


def _kenar_cubugu():
    dos = _dosyalar(st.session_state.get('dn', 0))
    st.sidebar.markdown("### 🗂️ İlan Çalışma Alanı")
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
    bilgi = [x for x in (dosya.get("telefon"), dosya.get("hedef"), dosya.get("bolgeler")) if x]
    if dosya.get("butce_min") or dosya.get("butce_max"):
        bilgi.append(f"Bütçe {sunum.tl(dosya.get('butce_min') or 0)} – {sunum.tl(dosya.get('butce_max') or 0)}")
    st.markdown(f"<div class='baslik'><b>{html.escape(dosya['ad'])}</b><span>{html.escape(' · '.join(bilgi))}</span></div>", unsafe_allow_html=True)
    if dosya.get("notlar"):
        st.markdown(f"<div class='notlar'>{html.escape(dosya['notlar'])}</div>", unsafe_allow_html=True)
    if not df.empty:
        chips = ""
        for tur, ad in (("satilik", "Satılık"), ("kiralik", "Kiralık")):
            d = df[df["tur"] == tur]
            if d.empty:
                continue
            c = d["asama"].value_counts().to_dict()
            chips += (f"<div class='chipsatir'><b>{ad}</b><span class='chip'>Toplam <b>{len(d)}</b></span>"
                      + "".join(f"<span class='chip'>{ASAMA_AD[k]} <b>{c[k]}</b></span>" for k in ASAMA_AD if c.get(k)) + "</div>")
        st.markdown(chips, unsafe_allow_html=True)
    with st.expander("📥 Revy ilan listesi yükle", expanded=df.empty):
        _yukleme(dosya, df)
    _veri_ayarlari(dosya, df)


def _yukleme(dosya, df):
    st.caption("Revy'de kriterini ara, 'Excel'e Aktar' de, dosyayı buraya bırak. Satılık ve kiralık dosyaları ayrı ayrı yükle (tür, "
               "'İşlem tipi' sütunundan otomatik anlaşılır). Aynı ilan tekrar yüklenirse fiyatı güncellenir, eleme durumun korunur.")
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
    degistir = False
    if parcalar and not df.empty:
        degistir = st.checkbox("Mevcut listeyi tamamen silip bunu baştan yükle (eski ilanlar, notlar ve eleme durumları gider)",
                               key="degistir")
    kalkan = pd.DataFrame()
    if parcalar and not df.empty and not degistir:
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
        if degistir:
            for t in tum["tur"].unique():
                db.ilanlar_sil(dosya["id"], t)
        n = db.ilanlar_yaz(dosya["id"], tum)
        if kalkan_sil and len(kalkan):
            db.ilan_toplu_guncelle(list(kalkan["id"]), {"asama": "elendi", "elenen_asama": KALKTI, "guncelleme": _simdi()})
        st.session_state["up_n"] = st.session_state.get("up_n", 0) + 1
        st.session_state["_mesaj"] = f"{n} ilan eklendi/güncellendi." + (f" {len(kalkan)} ilan yayından kalktı olarak çıkarıldı." if kalkan_sil and len(kalkan) else "")
        _yenile()
        st.rerun()


def _veri_ayarlari(dosya, df):
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
                _dosyalar_yenile()
                st.session_state["_mesaj"] = "Müşteri bilgileri kaydedildi."
                st.rerun()
    with st.expander("🗑 Listeyi temizle / dosyayı sil"):
        tur = st.radio("Tür", ["Satılık", "Kiralık", "İkisi"], horizontal=True, key="tmz_tur")
        tur_kod = {"Satılık": "satilik", "Kiralık": "kiralik"}.get(tur)
        c1, c2 = st.columns(2)
        with c1:
            st.caption("Yalnızca henüz işlem görmemiş **Aday** ilanları siler; beğendiklerin, görüşmeler ve final korunur.")
            if st.button("Aday ilanları temizle", key="tmz"):
                db.adaylari_sil(dosya["id"], tur_kod)
                st.session_state["_mesaj"] = "Aday ilanlar temizlendi."
                _yenile()
                st.rerun()
        with c2:
            st.caption("Seçilen türdeki **bütün** ilanları siler (beğendiklerin, görüşmeler ve notlar dahil). Müşteri dosyası kalır.")
            onay0 = st.checkbox("Listeyi tümden silmek istiyorum", key="tum_onay")
            if st.button("Tüm ilanları sil", disabled=not onay0, key="tum_sil", type="primary"):
                db.ilanlar_sil(dosya["id"], tur_kod)
                st.session_state["_mesaj"] = "Liste tümden silindi. Yeni Revy listesini yükleyebilirsin."
                _yenile()
                st.rerun()
        st.divider()
        onay = st.checkbox("Bu müşteri dosyasını ve tüm ilanlarını kalıcı olarak silmek istiyorum", key="sil_onay")
        if st.button("Dosyayı sil", disabled=not onay, key="sil"):
            db.dosya_sil(dosya["id"])
            _dosyalar_yenile()
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
        c = st.columns([2, 2, 1.1, 1.1, 2])
        ilce = c[0].multiselect("İlçe", sorted(sat["ilce"].dropna().unique()), key=f"{anahtar}_ilce", placeholder="Tümü")
        d = sat[sat["ilce"].isin(ilce)] if ilce else sat
        mah = c[1].multiselect("Mahalle", sorted(d["mahalle"].dropna().unique()), key=f"{anahtar}_mah", placeholder="Tümü")
        fmin = c[2].number_input("Fiyat min (TL)", 0, step=250000, key=f"{anahtar}_fmin")
        fmax = c[3].number_input("Fiyat max (TL)", 0, step=250000, key=f"{anahtar}_fmax")
        kul = c[4].multiselect("Kullanım durumu", sorted(sat["kullanim"].dropna().unique()), key=f"{anahtar}_kul", placeholder="Tümü")

        c = st.columns([2, 1.8, 2.2, 2.4])
        bina = c[0].multiselect("Bina yaşı grubu", ["Yeni (0–10 yaş)", "Eski (11 yaş ve üzeri)"], key=f"{anahtar}_bina", placeholder="Tümü",
                                help="Boş bırakırsan hepsi gelir. 11-15 yaş eski binaya dahildir.")
        yaslar = sorted(sat["bina_yasi"].dropna().astype(str).unique(), key=lambda x: (_yas_aralik(x)[0] is None, _yas_aralik(x)[0] or 0))
        yas = c[1].multiselect("Bina yaşı (tam)", yaslar, key=f"{anahtar}_yas", placeholder="Tümü")
        site = c[2].radio("Site içinde", ["Hepsi", "Evet", "Hayır"], horizontal=True, key=f"{anahtar}_site")
        katlar = sorted(sat["kat"].dropna().astype(str).unique(), key=lambda x: (not x.split(".")[0].isdigit(), int(x.split(".")[0]) if x.split(".")[0].isdigit() else 0, x))
        kat_gizle = c[3].multiselect("Başka kat(lar)ı gizle", katlar, key=f"{anahtar}_katg", placeholder="Tümü")

        c = st.columns([1.5, 1.7, 1.3, 1.3, 2.2, 2])
        cati = c[0].checkbox("Çatı / en üst kat gizle", value=True, key=f"{anahtar}_cati",
                             help="'Çatı' yazanlar ve kat numarası binanın kat sayısına eşit olanlar (3 katlı binanın 3. katı gibi)")
        zemin = c[1].checkbox("Zemin / giriş / bahçe / bodrum gizle", value=True, key=f"{anahtar}_zemin")
        gizle = c[2].checkbox("Aykırıları gizle", key=f"{anahtar}_gizle")
        alti = c[3].checkbox("Yalnız bölge altı", key=f"{anahtar}_alt")
        sira = c[4].selectbox("Sıralama", list(SIRALAMA), key=f"{anahtar}_sira")
        ara = c[5].text_input("Ara (başlık, site, ofis)", key=f"{anahtar}_ara")

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
    sat = sat.assign(_ay=sat["aykiri"].astype(int)).sort_values(["_ay", kol], ascending=[True, art], na_position="last").drop(columns="_ay")
    return sat, sig


def _durum(fark, aykiri):
    if aykiri:
        return "⚠ aykırı"
    if pd.isna(fark):
        return "—"
    em = "🟢" if fark <= -10 else ("🔴" if fark >= 10 else "🟡")
    return f"{em} {fark:+.0f}%".replace("-", "−")


def _kat_metni(r):
    k = str(r.get("kat") or "").replace(" Kat", "").replace(".", "").strip()
    t = r.get("toplam_kat")
    return f"{k}/{int(t)}" if k and pd.notna(t) else k


def _liste_df(d, secenek=None):
    mah = d["mahalle"].fillna("").str.replace(r"\s*Mah\.?$", "", regex=True)
    df = pd.DataFrame({
        "Durum": [_durum(f, a) for f, a in zip(d["fark_pct"], d["aykiri"])],
        "Konum": (d["ilce"].fillna("") + " · " + mah).values,
        "Fiyat": d["fiyat"].values, "m²": d["m2"].values, "m² fiyat": d["m2_fiyat"].values,
        "Kat": [_kat_metni(r) for _, r in d.iterrows()], "Yaş": d["bina_yasi"].astype(str).replace("nan", "").values,
        "Çarpan (ay)": d["carpan"].values, "Getiri %": d["getiri_pct"].values,
        "Kullanım": d["kullanim"].values, "İlan": d["url"].values,
        "Bölge ort. m²": d["ref_ort"].values, "Fark %": d["fark_pct"].values, "Düşüş %": d["dusus_pct"].values,
        "Tahmini kira": d["tahmini_kira"].values, "Yayın (gün)": d["yayin_suresi"].values,
        "Ofis / ilan sahibi": d["ofis"].fillna(d["ilan_sahibi"]).values, "Kaynak": d["kaynak"].values,
    }, index=d["id"].values)
    # veri yoksa boş kolonları gizle (ör. kiralık listesi yüklenmemişse kira/çarpan/getiri)
    for c in ("Tahmini kira", "Çarpan (ay)", "Getiri %", "Düşüş %", "Yayın (gün)"):
        if df[c].isna().all() or (c == "Düşüş %" and (df[c].fillna(0) == 0).all()):
            df = df.drop(columns=c)
    return df


KOMPAKT = ["Beğen", "Ele", "Kalktı", "Geri al", "Durum", "Konum", "Fiyat", "m²", "m² fiyat", "Kat", "Yaş", "Çarpan (ay)", "Getiri %",
           "Kullanım", "İlan", "Neden", "Not"]


def _sutun_sirasi(cols, detay):
    return list(cols) if detay else [c for c in KOMPAKT if c in cols]


def _kolon_ayar():
    n = st.column_config.NumberColumn
    t = st.column_config.TextColumn
    return {
        "Beğen": st.column_config.CheckboxColumn("✔", help="Beğen: 1. elemeyi geçir", width=44, pinned=True),
        "Ele": st.column_config.CheckboxColumn("✖", help="Ele: istemiyorum", width=44, pinned=True),
        "Kalktı": st.column_config.CheckboxColumn("🗑", help="İlan yayından kalkmış: listeden çıkar", width=44, pinned=True),
        "Geri al": st.column_config.CheckboxColumn("↩", help="Geri al", width=44, pinned=True),
        "Durum": t("Durum", width=82, help="Bölge ortalamasına göre m² fiyat farkı: 🟢 %10+ ucuz · 🟡 normal · 🔴 %10+ pahalı"),
        "Konum": t("Konum", width=170), "Fiyat": n("Fiyat (TL)", format="localized", width=95),
        "m²": n("m²", format="%.0f", width=48), "m² fiyat": n("m² fiyat", format="localized", width=82),
        "Kat": t("Kat", width=52), "Yaş": t("Yaş", width=52),
        "Çarpan (ay)": n("Çarpan", format="%.0f ay", width=70), "Getiri %": n("Getiri", format="%.1f%%", width=62),
        "Kullanım": t("Kullanım", width=80), "İlan": st.column_config.LinkColumn("İlan", display_text="aç ↗", width=52),
        "Bölge ort. m²": n("Bölge ort.", format="localized", width=82), "Fark %": n("Fark", format="%+.1f%%", width=62),
        "Düşüş %": n("Düşüş", format="%.1f%%", width=58), "Tahmini kira": n("Tah. kira", format="localized", width=78),
        "Yayın (gün)": n("Gün", format="%.0f", width=44),
        "Ofis / ilan sahibi": t("Ofis", width=140), "Kaynak": t("Kaynak", width=85),
        "Neden": t("Neden", width=105), "Not": t("Not", width=240),
    }


def _b1_degisti(anahtar):
    """Tablodaki işaretleme/not değişikliği: veritabanına yazar, satır anında listeden düşer."""
    ids = st.session_state.get("_b1_ids", [])
    ed = st.session_state.get(anahtar, {}).get("edited_rows", {})
    kalkti, ele, begen, notlar = [], [], [], {}
    for pos, ch in ed.items():
        if int(pos) >= len(ids):
            continue
        i = ids[int(pos)]
        if ch.get("Kalktı"):
            kalkti.append(i)
        elif ch.get("Ele"):
            ele.append(i)
        elif ch.get("Beğen"):
            begen.append(i)
        if "Not" in ch and not (ch.get("Kalktı") or ch.get("Ele")):
            notlar[i] = (ch["Not"] or "").strip() or None
    try:
        _guncelle(kalkti, asama="elendi", elenen_asama=KALKTI)
        _guncelle(ele, asama="elendi", elenen_asama="aday")
        _guncelle(begen, asama="birinci")
        for i, n in notlar.items():
            _guncelle([i], not_metni=n)
    except Exception as e:      # noqa: BLE001
        st.session_state["_b1_hata"] = str(e)
    st.session_state["_b1_ver"] = st.session_state.get("_b1_ver", 0) + 1


def _b1_geri(anahtar):
    ids = st.session_state.get("_b1_elen_ids", [])
    ed = st.session_state.get(anahtar, {}).get("edited_rows", {})
    m = st.session_state["durum"]["map"]
    try:
        for pos, ch in ed.items():
            if ch.get("Geri al") and int(pos) < len(ids):
                i = ids[int(pos)]
                onceki = m.get(i, {}).get("elenen_asama")
                _guncelle([i], asama=onceki if onceki in ("aday", "birinci") else "aday", elenen_asama=None)
    except Exception as e:      # noqa: BLE001
        st.session_state["_b1_hata"] = str(e)
    st.session_state["_b1_ver"] = st.session_state.get("_b1_ver", 0) + 1


@st.fragment
def _sekme_birinci(sat, dosya):
    """Fragment: bu bölümdeki her tıklama yalnız bu bölümü yeniler (tüm uygulamayı değil)."""
    sat = _uygula(sat)
    hata = st.session_state.pop("_b1_hata", None)
    if hata:
        st.error(f"Kaydedilemedi: {hata}")
    ver = st.session_state.get("_b1_ver", 0)
    aday = sat[sat["asama"] == "aday"]
    if aday.empty:
        st.info("Aday ilan yok. Veri sekmesinden Revy listesini yükle ya da elenenleri geri al.")
    else:
        d, sig = _filtre_df(aday, "b1")
        st.caption(f"{len(d)} / {len(aday)} ilan · {int(d['aykiri'].sum())} aykırı fiyatlı (bölge ortalamasına girmez). "
                   "✔ beğen, ✖ ele, 🗑 yayından kalkmış: işaretlediğin anda kaydedilir ve satır listeden çıkar.")
        ed = _liste_df(d, None)
        ed["Not"] = d["not_metni"].fillna("").values
        ed.insert(0, "Kalktı", False)
        ed.insert(0, "Ele", False)
        ed.insert(0, "Beğen", False)
        st.session_state["_b1_ids"] = [int(i) for i in ed.index]
        anahtar = f"ed_b1_{dosya['id']}_{ver}"
        detay = st.toggle("Ayrıntılı sütunlar", key="b1_detay")
        st.data_editor(ed, column_config=_kolon_ayar(), column_order=_sutun_sirasi(ed.columns, detay),
                       disabled=[c for c in ed.columns if c not in ("Beğen", "Ele", "Kalktı", "Not")],
                       hide_index=True, use_container_width=True, height=min(560, 44 + 35 * len(ed)), row_height=35, key=anahtar,
                       on_change=_b1_degisti, args=(anahtar,))
    elen = sat[sat["asama"] == "elendi"]
    if not elen.empty:
        with st.expander(f"Elenenler ({len(elen)})"):
            ed2 = _liste_df(elen, None)
            ed2["Neden"] = elen["elenen_asama"].map(lambda x: "Yayından kalktı" if x == KALKTI else "Elendi").values
            ed2["Not"] = elen["not_metni"].fillna("").values
            ed2.insert(0, "Geri al", False)
            st.session_state["_b1_elen_ids"] = [int(i) for i in ed2.index]
            anahtar2 = f"ed_elen_{dosya['id']}_{ver}"
            st.data_editor(ed2, column_config=_kolon_ayar(), column_order=_sutun_sirasi(ed2.columns, False),
                           disabled=[c for c in ed2.columns if c != "Geri al"], hide_index=True, use_container_width=True,
                           height=min(360, 44 + 35 * len(ed2)), row_height=35, key=anahtar2, on_change=_b1_geri, args=(anahtar2,))


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
                st.rerun()


# ── Sekme: 2. Eleme ────────────────────────────────────────────────────
def _net_liste(d):
    df = pd.DataFrame({
        "Konum": (d["ilce"].fillna("") + " · " + d["mahalle"].fillna("").str.replace(r"\s*Mah\.?$", "", regex=True)).values,
        "m²": d["m2"].values, "İlan": d["url"].values,
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
            "Final": st.column_config.CheckboxColumn("★ Final", help="Final listesine al", width=62, pinned=True),
            "Ele": st.column_config.CheckboxColumn("✖", help="Ele", width=44, pinned=True),
            "Geri": st.column_config.CheckboxColumn("↩", help="1. elemeye geri gönder", width=44, pinned=True)}


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
                         hide_index=True, use_container_width=True, row_height=35, height=min(420, 44 + 35 * len(ed)),
                         key=f"ed_b2_{len(ed)}")
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
CSS = """
<style>
.block-container{padding:3.4rem 1.6rem 1rem;max-width:1600px}
[data-testid="stVerticalBlock"]{gap:.5rem}
[data-testid="stWidgetLabel"] p{font-size:.76rem;opacity:.75}
[data-testid="stExpander"] summary{padding:.3rem .7rem}
[data-testid="stExpander"] summary p{font-size:.9rem;font-weight:600}
[data-testid="stSidebar"] [data-testid="stVerticalBlock"]{gap:.4rem}
[data-testid="stSidebarUserContent"]{padding-top:1.2rem}
hr{margin:.6rem 0!important}
.baslik{display:flex;gap:14px;align-items:baseline;flex-wrap:wrap;margin-bottom:.1rem}
.baslik b{font-size:1.3rem}.baslik span{opacity:.65;font-size:.85rem}
.notlar{background:rgba(128,128,128,.10);border-left:3px solid #1f3a5f;padding:.35rem .7rem;border-radius:4px;font-size:.86rem;margin:.2rem 0 .4rem}
.chipsatir{display:flex;gap:6px;align-items:center;flex-wrap:wrap;margin:.1rem 0 .3rem}
.chipsatir>b{min-width:58px;font-size:.85rem}
.chip{background:rgba(128,128,128,.14);border-radius:999px;padding:2px 11px;font-size:.78rem}
.chip b{margin-left:3px}
.st-key-bolum [role="radiogroup"]{gap:.4rem;flex-wrap:wrap}
.st-key-bolum label{background:rgba(128,128,128,.13);padding:.32rem .85rem;border-radius:999px;margin:0;cursor:pointer}
.st-key-bolum label>div:first-child{display:none}
.st-key-bolum label p{font-size:.88rem;font-weight:600}
.st-key-bolum label:has(input:checked){background:#1f3a5f}
.st-key-bolum label:has(input:checked) p{color:#fff}
</style>
"""

BOLUMLER = ["veri", "ozet", "b1", "gor", "b2", "final"]


def main():
    _giris()
    dosya, mn_s, mn_k, k = _kenar_cubugu()
    _mesaj()
    v = st.session_state.get("v", 0)
    df = _yukle(dosya["id"], v)
    sat, kira, oz = _hesap(dosya["id"], v, mn_s, mn_k, k)
    _durum_hazir(dosya["id"], v, df)
    dfd = _uygula(df)
    say = dfd[dfd["tur"] == "satilik"]["asama"].value_counts().to_dict() if not dfd.empty else {}
    ad = {"veri": "📥 Veri", "ozet": "📊 Bölge özeti", "b1": f"1️⃣ 1. Eleme ({say.get('aday', 0)})",
          "gor": f"📞 Danışman görüşmesi ({say.get('birinci', 0)})", "b2": f"2️⃣ 2. Eleme ({say.get('ikinci', 0)})",
          "final": f"⭐ Final ({say.get('final', 0)})"}
    st.markdown(CSS, unsafe_allow_html=True)
    bolum = st.radio("Bölüm", BOLUMLER, format_func=ad.get, horizontal=True, key="bolum", label_visibility="collapsed")
    if bolum == "veri":
        _sekme_veri(dosya, dfd)
        return
    if sat.empty:
        st.info("Önce Veri bölümünden Revy ilan listesini yükle.")
        return
    if bolum == "ozet":
        _sekme_ozet(oz, mn_s, mn_k)
    elif bolum == "b1":
        _sekme_birinci(sat, dosya)
    else:
        satd = hesap.net_hesapla(_uygula(sat))
        if bolum == "gor":
            _sekme_gorusme(satd)
        elif bolum == "b2":
            _sekme_ikinci(satd)
        else:
            _sekme_final(dosya, satd, oz)


main()

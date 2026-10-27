"""Bölge ortalamaları (aykırı ilanlar hariç), kira tahmini, kira çarpanı ve brüt getiri.

Yöntem
- m² fiyatı = fiyat / m². Kiralık için m² kira = aylık kira / m².
- Referans grup: ilanın mahallesinde yeterli ilan varsa mahalle, yoksa ilçe, o da yetersizse referans yok.
- Aykırı ilan: referans grubunun değer dağılımında alt/üst çeyrek ± k × çeyrekler açıklığı (IQR) dışında kalan ilan.
  Aykırılar ortalamaya ve medyana GİRMEZ; listede işaretli kalır.
- Tahmini aylık kira = referans m² kira ortalaması × ilanın m²'si.
- Kira çarpanı = fiyat / aylık kira (ay). Brüt getiri % = 12 / çarpan × 100.
"""
import re

import numpy as np
import pandas as pd

MIN_N_SATILIK = 6
MIN_N_KIRALIK = 4
IQR_K = 1.5


def tr_kucuk(s):
    s = str(s if s is not None and not (isinstance(s, float) and np.isnan(s)) else "")
    s = s.replace("İ", "i").replace("I", "ı").lower().strip()
    return re.sub(r"\s+", " ", s)


def mahalle_anahtar(m):
    s = tr_kucuk(m)
    s = re.sub(r"\b(mahallesi|mah\.?|mh\.?)(?=\s|$)", "", s)
    return re.sub(r"\s+", " ", s).strip()


def _grup_stat(df, anahtar, deger, k):
    """{anahtar: {n, n_ok, lo, hi, ort, medyan}} — ort/medyan aykırılar hariç."""
    sonuc = {}
    if df.empty:
        return sonuc
    for key, g in df.groupby(anahtar):
        v = g[deger].to_numpy(float)
        if len(v) >= 4:
            q1, q3 = np.percentile(v, [25, 75])
            iqr = q3 - q1
            lo, hi = q1 - k * iqr, q3 + k * iqr
        else:
            lo, hi = -np.inf, np.inf
        ok = v[(v >= lo) & (v <= hi)]
        sonuc[key] = {"n": len(v), "n_ok": len(ok), "lo": lo, "hi": hi,
                      "ort": float(ok.mean()) if len(ok) else np.nan, "medyan": float(np.median(ok)) if len(ok) else np.nan}
    return sonuc


def _sec(ilce, mah, st_m, st_i, min_n):
    s = st_m.get((ilce, mah))
    if s and s["n"] >= min_n:
        return "mahalle", s
    s = st_i.get(ilce)
    if s and s["n"] >= min_n:
        return "ilçe", s
    return "", None


def _hazirla(df, tur, deger_adi):
    d = df[df["tur"] == tur].copy()
    for c in ("m2", "fiyat", "ilk_fiyat", "yayin_suresi"):
        d[c] = pd.to_numeric(d.get(c), errors="coerce")
    d = d[(d["m2"] > 0) & (d["fiyat"] > 0)].copy()
    d["ilce_k"] = d["ilce"].map(tr_kucuk)
    d["mah_k"] = d["mahalle"].map(mahalle_anahtar)
    d[deger_adi] = d["fiyat"] / d["m2"]
    return d


def hesapla(df, min_n_sat=MIN_N_SATILIK, min_n_kira=MIN_N_KIRALIK, k=IQR_K):
    """df: ilanlar tablosu (tur, ilce, mahalle, m2, fiyat, ilk_fiyat, ...). Dönüş: (satılık_df, kiralık_df, özet_df)."""
    sat = _hazirla(df, "satilik", "m2_fiyat")
    kira = _hazirla(df, "kiralik", "m2_kira")
    st_sm = _grup_stat(sat, ["ilce_k", "mah_k"], "m2_fiyat", k)
    st_si = _grup_stat(sat, "ilce_k", "m2_fiyat", k)
    st_km = _grup_stat(kira, ["ilce_k", "mah_k"], "m2_kira", k)
    st_ki = _grup_stat(kira, "ilce_k", "m2_kira", k)

    def _ref(d, deger, st_m, st_i, mn, on):
        duzey, n, ort, med, aykiri, yon = [], [], [], [], [], []
        for ilce, mah, v in zip(d["ilce_k"], d["mah_k"], d[deger]):
            dz, s = _sec(ilce, mah, st_m, st_i, mn)
            if s is None:
                duzey.append(""); n.append(0); ort.append(np.nan); med.append(np.nan); aykiri.append(False); yon.append("")
                continue
            duzey.append(dz); n.append(s["n_ok"]); ort.append(s["ort"]); med.append(s["medyan"])
            if v < s["lo"]:
                aykiri.append(True); yon.append("düşük")
            elif v > s["hi"]:
                aykiri.append(True); yon.append("yüksek")
            else:
                aykiri.append(False); yon.append("")
        d[on + "duzey"], d[on + "n"], d[on + "ort"], d[on + "medyan"] = duzey, n, ort, med
        d["aykiri"], d["aykiri_yon"] = aykiri, yon
        return d

    sat = _ref(sat, "m2_fiyat", st_sm, st_si, min_n_sat, "ref_")
    if not kira.empty:
        kira = _ref(kira, "m2_kira", st_km, st_ki, min_n_kira, "ref_")
    else:
        for c in ("ref_duzey", "ref_n", "ref_ort", "ref_medyan", "aykiri", "aykiri_yon"):
            kira[c] = []

    # Satılık ilanlara kira tahmini (referans: aynı mahalle/ilçe kiralık ortalaması, aykırılar hariç)
    kd, kn, ko = [], [], []
    for ilce, mah in zip(sat["ilce_k"], sat["mah_k"]):
        dz, s = _sec(ilce, mah, st_km, st_ki, min_n_kira)
        kd.append(dz); kn.append(s["n_ok"] if s else 0); ko.append(s["ort"] if s else np.nan)
    sat["kira_duzey"], sat["kira_n"], sat["kira_ort_m2"] = kd, kn, ko
    sat["fark_pct"] = (sat["m2_fiyat"] / sat["ref_ort"] - 1) * 100
    sat["tahmini_kira"] = sat["kira_ort_m2"] * sat["m2"]
    sat["carpan"] = sat["fiyat"] / sat["tahmini_kira"]
    sat["getiri_pct"] = 12 / sat["carpan"] * 100
    ilk = sat["ilk_fiyat"]
    sat["dusus_pct"] = np.where((ilk > 0) & (sat["fiyat"] < ilk), (ilk - sat["fiyat"]) / ilk * 100, 0.0)

    return sat, kira, bolge_ozet(sat, kira, min_n_sat, min_n_kira, k)


def bolge_ozet(sat, kira, min_n_sat=MIN_N_SATILIK, min_n_kira=MIN_N_KIRALIK, k=IQR_K):
    """İlçe ve mahalle bazında özet (aykırılar hariç ortalama)."""
    satir = []
    for seviye, anahtar in (("ilçe", "ilce_k"), ("mahalle", ["ilce_k", "mah_k"])):
        ss = _grup_stat(sat, anahtar, "m2_fiyat", k)
        ks = _grup_stat(kira, anahtar, "m2_kira", k) if not kira.empty else {}
        adlar = {}
        for _, r in sat.iterrows():
            key = r["ilce_k"] if seviye == "ilçe" else (r["ilce_k"], r["mah_k"])
            adlar.setdefault(key, (r["ilce"], r["mahalle"]))
        for key, s in ss.items():
            kk = ks.get(key)
            ilce, mah = adlar.get(key, ("", ""))
            ort_m2 = s["ort"]
            kira_m2 = kk["ort"] if kk and kk["n"] >= 1 else np.nan
            satir.append({
                "Seviye": seviye, "İlçe": ilce, "Mahalle": "" if seviye == "ilçe" else mah,
                "Satılık ilan": s["n"], "Aykırı": s["n"] - s["n_ok"], "Ort. m² fiyat": ort_m2, "Medyan m² fiyat": s["medyan"],
                "Kiralık ilan": kk["n"] if kk else 0, "Ort. m² kira": kira_m2,
                "Kira çarpanı (ay)": ort_m2 / kira_m2 if kira_m2 and kira_m2 > 0 else np.nan,
                "Brüt getiri %": 12 * kira_m2 / ort_m2 * 100 if kira_m2 and kira_m2 > 0 and ort_m2 > 0 else np.nan,
                "Yeterli veri": s["n"] >= min_n_sat and (kk["n"] if kk else 0) >= min_n_kira,
            })
    return pd.DataFrame(satir)


def _sayi(v):
    try:
        f = float(v)
        return f if f > 0 else np.nan
    except (TypeError, ValueError):
        return np.nan


def net_hesapla(sat):
    """Danışman görüşmesinden gelen net fiyat ve kira varsa onlarla yeniden hesaplar (yoksa ilan fiyatı ve tahmini kira)."""
    d = sat.copy()
    g = d["gorusme"].map(lambda x: x if isinstance(x, dict) else {})
    net_f = g.map(lambda x: _sayi(x.get("net_fiyat")))
    kira_g = g.map(lambda x: _sayi(x.get("kira")))
    d["fiyat_net"] = net_f.fillna(d["fiyat"])
    d["kira_x"] = kira_g.fillna(d["tahmini_kira"])
    d["kira_kaynak"] = np.where(kira_g.notna(), "danışman", np.where(d["tahmini_kira"].notna(), "tahmin", ""))
    d["m2_fiyat_net"] = d["fiyat_net"] / d["m2"]
    d["fark_net"] = (d["m2_fiyat_net"] / d["ref_ort"] - 1) * 100
    d["carpan_net"] = d["fiyat_net"] / d["kira_x"]
    d["getiri_net"] = 12 / d["carpan_net"] * 100
    return d

"""Çıktılar: müşteriye sunulacak sayfa (HTML) ve Excel."""
import html as _h
import io

import numpy as np
import pandas as pd


def tl(x):
    return "—" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x:,.0f}".replace(",", ".") + " TL"


def sayi(x, nd=1):
    return "—" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x:,.{nd}f}".replace(",", "_").replace(".", ",").replace("_", ".")


def _g(r):
    g = r.get("gorusme")
    return g if isinstance(g, dict) else {}


def _satirlar(r, danisman):
    g = _g(r)
    s = [
        ("İlan fiyatı", tl(r["fiyat"])),
        ("Danışmandan öğrenilen net fiyat", tl(r["fiyat_net"]) if r["fiyat_net"] != r["fiyat"] else None),
        ("Alan", f"{sayi(r['m2'], 0)} m² · {r.get('oda') or ''}".strip(" ·")),
        ("m² fiyatı", f"{tl(r['m2_fiyat_net'])} (bölge ort. {('%+.1f' % r['fark_net']).replace('.', ',')}%)" if pd.notna(r["fark_net"]) else tl(r["m2_fiyat_net"])),
        ("Bölge ort. m² fiyatı", tl(r["ref_ort"]) if pd.notna(r["ref_ort"]) else None),
        ("Aylık kira", (tl(r["kira_x"]) + (" (tahmini)" if r["kira_kaynak"] == "tahmin" else " (danışman)")) if pd.notna(r["kira_x"]) else None),
        ("Kira çarpanı", f"{sayi(r['carpan_net'], 0)} ay ({sayi(r['carpan_net'] / 12, 1)} yıl)" if pd.notna(r["carpan_net"]) else None),
        ("Brüt kira getirisi", f"%{sayi(r['getiri_net'], 1)}" if pd.notna(r["getiri_net"]) else None),
        ("Kullanım durumu", r.get("kullanim")),
        ("Bina yaşı / kat", " / ".join(x for x in [str(r.get("bina_yasi") or ""), str(r.get("kat") or "")] if x)),
        ("Aidat", tl(_aidat(g))), ("Tapu", g.get("tapu")), ("Krediye uygunluk", g.get("kredi")),
    ]
    if danisman:
        s += [("Danışman", " · ".join(x for x in [g.get("danisman_adi") or r.get("ilan_sahibi"), g.get("danisman_tel")] if x)),
              ("Satış nedeni", g.get("neden")), ("Notlar", g.get("notlar"))]
    return [(k, v) for k, v in s if v not in (None, "", "—", "nan")]


def _aidat(g):
    try:
        f = float(g.get("aidat"))
        return f if f > 0 else None
    except (TypeError, ValueError):
        return None


_CSS = """
:root{--bg:#f5f6f8;--kart:#fff;--yazi:#1f2937;--soluk:#6b7280;--cizgi:#e5e7eb;--vurgu:#1f3a5f}
@media(prefers-color-scheme:dark){:root{--bg:#111827;--kart:#1f2937;--yazi:#f3f4f6;--soluk:#9ca3af;--cizgi:#374151;--vurgu:#93c5fd}}
body{margin:0;background:var(--bg);color:var(--yazi);font-family:Segoe UI,Arial,sans-serif;padding:16px}
h1{font-size:20px;margin:0 0 4px}.alt{color:var(--soluk);font-size:13px;margin-bottom:14px}
.izgara{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:14px}
.kart{background:var(--kart);border:1px solid var(--cizgi);border-radius:14px;padding:16px}
.kart h2{font-size:15px;margin:0 0 2px;color:var(--vurgu)}.kart .bas{color:var(--soluk);font-size:12.5px;margin-bottom:10px}
.kart .fiyat{font-size:22px;font-weight:800;margin:6px 0 10px}
table{width:100%;border-collapse:collapse;font-size:13.5px}td{padding:6px 0;border-top:1px solid var(--cizgi);vertical-align:top}
td:first-child{color:var(--soluk);padding-right:10px;width:46%}
a{color:var(--vurgu)}.not{color:var(--soluk);font-size:12px;line-height:1.5;margin-top:14px}
"""


def sunum_html(musteri, satirlar_df, danisman=False):
    kartlar = ""
    for _, r in satirlar_df.iterrows():
        g = _g(r)
        baslik = f"{r['ilce']} · {r['mahalle']}"
        kartlar += (
            f"<div class='kart'><h2>{_h.escape(baslik)}</h2><div class='bas'>{_h.escape(str(r.get('baslik') or ''))[:110]}</div>"
            f"<div class='fiyat'>{_h.escape(tl(r['fiyat_net']))}</div><table>"
            + "".join(f"<tr><td>{_h.escape(k)}</td><td>{_h.escape(str(v))}</td></tr>" for k, v in _satirlar(r, danisman))
            + f"</table><p><a href='{_h.escape(str(r['url']))}' target='_blank' rel='noopener'>İlanı aç</a></p></div>"
        )
    return (
        "<!doctype html><html lang='tr'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'>"
        f"<title>{_h.escape(musteri)} — seçilen ilanlar</title><style>{_CSS}</style></head><body>"
        f"<h1>{_h.escape(musteri)} için seçilen ilanlar</h1><div class='alt'>{len(satirlar_df)} ilan karşılaştırması</div>"
        f"<div class='izgara'>{kartlar}</div>"
        "<div class='not'>Bölge ortalamaları, aynı kriterde Revy'den alınan ilanlardan hesaplanmıştır; olağan dışı yüksek ve düşük fiyatlı ilanlar "
        "ortalamaya dahil edilmemiştir. Tahmini kira, aynı bölgedeki kiralık ilanların m² ortalamasından türetilir. "
        "Kira çarpanı = fiyat ÷ aylık kira; brüt getiri = 12 ÷ çarpan. Rakamlar bilgilendirme amaçlıdır.</div></body></html>"
    ).encode("utf-8")


KOLONLAR = [
    ("ilce", "İlçe"), ("mahalle", "Mahalle"), ("oda", "Oda"), ("m2", "m²"), ("fiyat", "İlan fiyatı"), ("fiyat_net", "Net fiyat"),
    ("m2_fiyat_net", "m² fiyat"), ("ref_ort", "Bölge ort. m²"), ("fark_net", "Fark %"), ("kira_x", "Aylık kira"),
    ("kira_kaynak", "Kira kaynağı"), ("carpan_net", "Kira çarpanı (ay)"), ("getiri_net", "Brüt getiri %"),
    ("kullanim", "Kullanım"), ("yayin_suresi", "Yayın (gün)"), ("ofis", "Ofis"), ("ilan_sahibi", "İlan sahibi"), ("kaynak", "Kaynak"),
    ("not_metni", "Not"), ("url", "Link"),
]


def excel(musteri, parcalar, ozet):
    """parcalar: {'Final 3': df, '2. Eleme': df, '1. Eleme': df} (net_hesapla uygulanmış)."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for ad, d in parcalar.items():
            if d is None or d.empty:
                continue
            out = d[[c for c, _ in KOLONLAR if c in d.columns]].rename(columns=dict(KOLONLAR))
            g = d["gorusme"].map(lambda x: x if isinstance(x, dict) else {})
            for anahtar, etiket in (("danisman_adi", "Danışman"), ("danisman_tel", "Danışman tel"), ("aidat", "Aidat"),
                                    ("tapu", "Tapu"), ("kredi", "Kredi"), ("neden", "Satış nedeni"), ("notlar", "Görüşme notu")):
                out[etiket] = g.map(lambda x, a=anahtar: x.get(a))
            out.to_excel(w, sheet_name=ad[:31], index=False)
            ws = w.sheets[ad[:31]]
            for i, col in enumerate(out.columns, 1):
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = 16
        if ozet is not None and not ozet.empty:
            ozet.drop(columns=["Seviye"], errors="ignore").to_excel(w, sheet_name="Bölge özeti", index=False)
    return buf.getvalue()

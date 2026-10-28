# -*- coding: utf-8 -*-
"""StockFlow : stocks de chantier et simulation des premières ruptures.

Unité de temps : jour calendaire. Une réception devient disponible en fin de
jour, après la demande. La demande non satisfaite n'est pas reportée au lendemain.
"""
from __future__ import annotations
import copy
from datetime import date, timedelta
import hashlib
import html
import io
import json
import math
import os
from pathlib import Path
import sqlite3

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from openpyxl import Workbook
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

MATERIALS = [
    ("ciment", "Ciment", "sacs", "#B9A6DF", "inventory_2"),
    ("sable", "Sable", "m³", "#DFB466", "landscape"),
    ("gravier", "Gravier", "m³", "#92B7B2", "grain"),
    ("bois", "Bois", "m³", "#C58C64", "forest"),
    ("acier", "Acier", "kg", "#8696C8", "construction"),
]
PINK, INK, GREEN = "#D46B56", "#30243E", "#328574"


def defaults(example=False):
    values = [(120, 12, 90, 7), (30, 2.5, 20, 6), (40, 3, 25, 8), (12, .6, 8, 10), (1800, 100, 1200, 9)]
    rows = []
    for (key, name, unit, color, icon), (stock, daily, delivery, lead) in zip(MATERIALS, values):
        rows.append(dict(id=key, stock=stock if example else 0., daily=daily if example else 0.,
                         delivery=delivery if example else 0., lead=lead, variation=30., late=2,
                         periodic=False, interval=14, enabled=True))
    return dict(schema=1, start=date.today().isoformat(), horizon=30, count=2000, seed=42, materials=rows)


def numeric(value, name, low, high, integer=False):
    try:
        result = float(value)
    except (TypeError, ValueError):
        raise ValueError(f"{name} : nombre valide attendu.") from None
    if not math.isfinite(result) or not low <= result <= high or (integer and not result.is_integer()):
        raise ValueError(f"{name} : valeur {'entière ' if integer else ''}entre {low} et {high} attendue.")
    return int(result) if integer else result


def validate(config):
    if not isinstance(config, dict) or config.get("schema") != 1:
        raise ValueError("Configuration StockFlow non reconnue.")
    try:
        start = date.fromisoformat(str(config.get("start")))
    except ValueError:
        raise ValueError("Date de départ invalide.") from None
    result = dict(schema=1, start=start.isoformat(),
                  horizon=numeric(config.get("horizon"), "Horizon", 7, 180, True),
                  count=numeric(config.get("count"), "Simulations", 500, 10000, True),
                  seed=numeric(config.get("seed"), "Graine", 0, 2_000_000_000, True), materials=[])
    rows = config.get("materials")
    if not isinstance(rows, list) or len(rows) != 5 or not all(isinstance(r, dict) for r in rows):
        raise ValueError("Les cinq matériaux doivent être présents une seule fois.")
    if {r.get("id") for r in rows} != {x[0] for x in MATERIALS}:
        raise ValueError("Identifiants des matériaux invalides.")
    for key, name, *_ in MATERIALS:
        raw = next(r for r in rows if r["id"] == key)
        row = {"id": key}
        for field, lo, hi, integer in [("stock", 0, 1e9, False), ("daily", 0, 1e7, False),
                                       ("delivery", 0, 1e9, False), ("lead", 1, 90, True),
                                       ("variation", 0, 100, False), ("late", 0, 30, True), ("interval", 1, 180, True)]:
            row[field] = numeric(raw.get(field), f"{name} / {field}", lo, hi, integer)
        for field in ("enabled", "periodic"):
            if type(raw.get(field)) is not bool:
                raise ValueError(f"{name} / {field} : case booléenne attendue.")
            row[field] = raw[field]
        result["materials"].append(row)
    return result


def fingerprint(config):
    return hashlib.sha256(json.dumps(validate(config), sort_keys=True).encode()).hexdigest()


def store_path():
    root = Path(os.environ.get("STOCKFLOW_DATA_DIR", str(Path.home() / "Documents" / "Codex" / "data" / "stockflow") if os.name == "nt" else str(Path.home() / ".local" / "share" / "stockflow"))).expanduser().resolve()
    app = Path(__file__).resolve().parent
    if root == app or app in root.parents:
        raise ValueError("Le stockage doit se trouver hors du dossier StockFlow.")
    root.mkdir(parents=True, exist_ok=True)
    return root / "stockflow.sqlite3"


def storage(config=None, expected=None):
    conn = sqlite3.connect(store_path(), timeout=10)
    try:
        with conn:
            conn.execute("CREATE TABLE IF NOT EXISTS configuration (id INTEGER PRIMARY KEY CHECK(id=1),revision INTEGER,payload TEXT)")
            conn.execute("INSERT OR IGNORE INTO configuration VALUES (1,0,?)", (json.dumps(defaults()),))
            if config is not None:
                payload = json.dumps(validate(config), allow_nan=False)
                cursor = conn.execute("UPDATE configuration SET payload=?,revision=revision+1 WHERE id=1 AND revision=?", (payload, expected))
                if cursor.rowcount != 1:
                    raise ValueError("Les données ont changé dans une autre session. Actualiser puis réessayer.")
            revision, payload = conn.execute("SELECT revision,payload FROM configuration WHERE id=1").fetchone()
        return validate(json.loads(payload)), revision
    finally:
        conn.close()


def wilson(successes, count):
    p = successes / count
    z = 1.959963984540054
    den = 1 + z*z/count
    mid = (p + z*z/(2*count)) / den
    half = z*math.sqrt(p*(1-p)/count + z*z/(4*count*count))/den
    return max(0., mid-half), min(1., mid+half)


def deterministic(row, horizon):
    """Bilan sans variation ni retard : réception à la FIN du jour prévu."""
    stock, first = row["stock"], horizon + 1
    points = [stock]
    for day in range(1, horizon + 1):
        if stock + 1e-10 < row["daily"] and first > horizon:
            first = day
        stock = max(0., stock - row["daily"])
        arrival = row["delivery"] > 0 and (day == row["lead"] or
                  (row["periodic"] and day > row["lead"] and (day-row["lead"]) % row["interval"] == 0))
        if arrival:
            stock += row["delivery"]
        points.append(stock)
    return first, np.asarray(points)


@st.cache_data(ttl=1800, max_entries=8, show_spinner=False)
def simulate(config):
    config = validate(config)
    n, h = config["count"], config["horizon"]
    result = {"signature": fingerprint(config), "config": config, "materials": {}}
    global_first = np.full(n, h+1, dtype=np.int16)
    for index, row in enumerate(config["materials"]):
        if not row["enabled"]:
            continue
        # Flux indépendants pour conserver la même demande lors des comparaisons
        # de stocks ou de livraisons, à graine et variation identiques.
        demand_rng = np.random.default_rng(np.random.SeedSequence([config["seed"], index, 1]))
        delivery_rng = np.random.default_rng(np.random.SeedSequence([config["seed"], index, 2]))
        lead = row["lead"] + delivery_rng.integers(0, row["late"]+1, size=n)
        receipts = np.zeros((n, h), dtype=np.float64)
        nominal = row["lead"]
        if row["delivery"] > 0:
            k = 0
            while nominal <= h:
                due = lead if k == 0 else nominal + delivery_rng.integers(0, row["late"]+1, size=n)
                within = due <= h
                np.add.at(receipts, (np.flatnonzero(within), due[within]-1), row["delivery"])
                if not row["periodic"]:
                    break
                nominal += row["interval"]
                k += 1
        stock_paths = np.zeros((n, h+1))
        demand_paths = np.zeros((n, h+1))
        stock_paths[:, 0] = row["stock"]
        stock = np.full(n, row["stock"], dtype=float)
        first = np.full(n, h+1, dtype=np.int16)
        unmet_total = np.zeros(n)
        lead_need = np.zeros(n)
        mean_demand, mean_served, mean_unmet = [], [], []
        for day in range(1, max(h, row["lead"] + row["late"]) + 1):
            mean, spread = row["daily"], row["variation"] / 100
            demand = np.full(n, mean) if mean == 0 or spread == 0 else demand_rng.triangular(mean*(1-spread), mean, mean*(1+spread), n)
            lead_need += demand * (day <= lead)
            if day > h:
                continue
            # La tolérance évite les fausses ruptures dues aux arrondis binaires.
            insufficient = demand > stock + np.maximum(1e-10, np.abs(demand)*1e-12)
            first[(first == h+1) & insufficient] = day
            served = np.minimum(stock, demand)
            unmet = demand - served
            unmet_total += unmet
            stock = np.maximum(stock - demand, 0.) + receipts[:, day-1]
            stock_paths[:, day] = stock
            demand_paths[:, day] = demand_paths[:, day-1] + demand
            mean_demand.append(float(demand.mean()))
            mean_served.append(float(served.mean()))
            mean_unmet.append(float(unmet.mean()))
        hits = int(np.count_nonzero(first <= h))
        global_first = np.minimum(global_first, first)
        result["materials"][row["id"]] = dict(
            first=first, hits=hits, probability=hits/n, interval=wilson(hits, n),
            day_quantiles=np.quantile(first, [.1, .5, .9], method="inverted_cdf").astype(int),
            stock_quantiles=np.quantile(stock_paths, [.1, .5, .9], axis=0),
            mean_stock=stock_paths.mean(axis=0),
            demand_quantiles=np.quantile(demand_paths, [.1, .5, .9], axis=0),
            examples=stock_paths[:5].copy(), cumulative_probability=np.array([np.mean(first <= d) for d in range(h+1)]),
            mean_demand=np.array(mean_demand), mean_served=np.array(mean_served), mean_unmet=np.array(mean_unmet),
            mean_receipts=receipts.mean(axis=0), unmet_total=unmet_total,
            lead_need95=float(np.quantile(lead_need, .95)), first_delivery=lead,
            deterministic_first=deterministic(row, h)[0], deterministic_stock=deterministic(row, h)[1])
    if not result["materials"]:
        raise ValueError("Activer au moins un matériau avant de simuler.")
    hits = int(np.count_nonzero(global_first <= h))
    result.update(global_first=global_first, global_probability=hits/n, global_interval=wilson(hits, n))
    return result


def fmt(value, digits=1):
    return f"{value:,.{digits}f}".replace(",", " ").replace(".", ",")


def day_label(value, horizon):
    return f"> {horizon} j" if int(value) > horizon else f"J{int(value)}"


def summary_frame(result):
    rows = []
    cfg = result["config"]
    for key, name, unit, *_ in MATERIALS:
        if key not in result["materials"]:
            continue
        r = result["materials"][key]
        rows.append({"Matériau": name, "Unité": unit, "Ruptures simulées": r["hits"], "Scénarios": cfg["count"],
                     "Probabilité (%)": 100*r["probability"], "IC fréquence bas (%)": 100*r["interval"][0],
                     "IC fréquence haut (%)": 100*r["interval"][1],
                     "Délai q10": day_label(r["day_quantiles"][0], cfg["horizon"]),
                     "Délai médian": day_label(r["day_quantiles"][1], cfg["horizon"]),
                     "Délai q90": day_label(r["day_quantiles"][2], cfg["horizon"]),
                     "Stock de protection q95": r["lead_need95"]})
    return pd.DataFrame(rows)


def daily_frame(result, key):
    r, cfg = result["materials"][key], result["config"]
    h = cfg["horizon"]
    return pd.DataFrame({"Jour": range(1, h+1),
        "Date": [(date.fromisoformat(cfg["start"])+timedelta(days=d-1)).isoformat() for d in range(1, h+1)],
        "Demande moyenne simulée": r["mean_demand"], "Utilisation moyenne servie": r["mean_served"],
        "Demande moyenne non servie": r["mean_unmet"], "Réceptions moyennes": r["mean_receipts"],
        "Stock fin de jour q10": r["stock_quantiles"][0, 1:], "Stock fin de jour médian": r["stock_quantiles"][1, 1:],
        "Stock fin de jour q90": r["stock_quantiles"][2, 1:], "Probabilité de première rupture cumulée (%)": r["cumulative_probability"][1:]*100})


def csv_bytes(frame):
    return frame.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")


def excel_bytes(result):
    book = Workbook()
    book.remove(book.active)
    frames = [("Synthese", summary_frame(result))] + [(name, daily_frame(result, key)) for key, name, *_ in MATERIALS if key in result["materials"]]
    for name, frame in frames:
        sheet = book.create_sheet(name)
        sheet.append(list(frame.columns))
        for row in frame.itertuples(index=False, name=None):
            sheet.append(list(row))
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column in sheet.columns:
            sheet.column_dimensions[column[0].column_letter].width = min(35, max(16, len(str(column[0].value))))
    buf = io.BytesIO(); book.save(buf)
    return buf.getvalue()


def report_pdf(result, fictitious):
    import reportlab
    fonts = Path(reportlab.__file__).resolve().parent / "fonts"
    for name, filename in [("SF", "Vera.ttf"), ("SF-Bold", "VeraBd.ttf")]:
        if name not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont(name, str(fonts / filename)))
    styles = getSampleStyleSheet()
    for style in styles.byName.values():
        style.fontName = "SF-Bold" if style.name in ("Title", "Heading2") else "SF"
    styles["Normal"].fontSize = 9; styles["Normal"].leading = 13
    paragraph = lambda text: Paragraph(html.escape(str(text)), styles["Normal"])
    cfg = result["config"]
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, rightMargin=1.3*cm, leftMargin=1.3*cm, topMargin=1.5*cm, bottomMargin=1.5*cm,
                            title="StockFlow - Analyse des stocks")
    elements = [Paragraph("StockFlow | Analyse des stocks", styles["Title"]),
        paragraph(f"{'EXEMPLE FICTIF' if fictitious else 'PARAMETRES UTILISATEUR'} - Départ J1 : {cfg['start']} - Horizon {cfg['horizon']} jours - {cfg['count']} scénarios - Graine {cfg['seed']}"), Spacer(1, .4*cm)]
    rows = [[paragraph(x) for x in ["Matériau", "Rupture avant fin", "Médiane du délai", "Protection q95"]]]
    for key, name, unit, *_ in MATERIALS:
        if key in result["materials"]:
            r = result["materials"][key]
            rows.append([paragraph(name), paragraph(f"{fmt(r['probability']*100)} % ({r['hits']}/{cfg['count']})"),
                         paragraph(day_label(r["day_quantiles"][1], cfg["horizon"])), paragraph(f"{fmt(r['lead_need95'])} {unit}")])
    t = Table(rows, colWidths=[3.1*cm, 5*cm, 4.7*cm, 5.1*cm], repeatRows=1)
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EADFEB")),
                          ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.white]),
                          ("VALIGN", (0, 0), (-1, -1), "TOP"), ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8)]))
    elements += [t, Spacer(1, .4*cm), paragraph("Une rupture est la première journée où le stock disponible ne couvre pas toute la demande. Les livraisons deviennent disponibles en fin de jour. Les demandes non servies ne sont pas reportées. '> H j' signifie que ce quantile n'est pas atteint dans l'horizon, pas une date connue au-delà."), Spacer(1, .3*cm)]
    for row in cfg["materials"]:
        if not row["enabled"]:
            continue
        _, name, unit, *_ = next(m for m in MATERIALS if m[0] == row["id"])
        elements.append(paragraph(f"{name} : stock {row['stock']:g} {unit}, demande {row['daily']:g} {unit}/jour, livraison {row['delivery']:g} {unit}, délai {row['lead']} j, retard uniforme de 0 à {row['late']} j, variation demande ±{row['variation']:g} %, {'livraisons répétées tous les '+str(row['interval'])+' j' if row['periodic'] else 'une seule livraison prévue'}."))
        elements.append(Spacer(1, .15*cm))
    elements += [Spacer(1, .3*cm), paragraph("Hypothèses : demande triangulaire symétrique, indépendance des jours et des matériaux, tous les jours calendaires consommés. Le stock de protection q95 est le quantile de demande cumulée pendant le délai d'une nouvelle commande. Ce repère ne garantit pas l'absence de rupture sur tout le projet. Les fréquences Monte-Carlo sont conditionnelles à ces choix, qui doivent être adaptés au chantier.")]
    def footer(canvas, document):
        canvas.setFont("SF", 8)
        canvas.drawString(1.3*cm, .8*cm, "StockFlow - Simulation pédagogique des approvisionnements")
        canvas.drawRightString(19.5*cm, .8*cm, str(document.page))
    doc.build(elements, onFirstPage=footer, onLaterPages=footer)
    return buf.getvalue()


CSS = """<style>
.stApp{background:#F7F4F0;color:#30243E}[data-testid="stAppViewContainer"]{background:#F7F4F0}
[data-testid="stHeader"]{background:#F7F4F0F0}.block-container{max-width:1350px;padding-top:4.5rem;padding-bottom:3rem}
h1,h2,h3{color:#30243E!important;letter-spacing:-.04em}p,label,h1,h2,h3{overflow-wrap:anywhere}
[data-testid="stSidebar"]{background:#30243E}[data-testid="stSidebar"] p,[data-testid="stSidebar"] label{color:#F2EAF7!important}
[data-testid="stSidebarNav"] a{color:#F2EAF7!important;border-radius:10px}
[data-testid="stSidebarNav"] a[aria-current="page"]{background:#E8B1A3!important}
[data-testid="stSidebarNav"] a[aria-current="page"] span,[data-testid="stSidebarNav"] a[aria-current="page"] p{color:#30243E!important}
[data-testid="stSidebar"] button[kind="secondary"]{background:#E7B0A1;color:#30243E;border:0}
[data-testid="stMain"] [data-testid="stCaptionContainer"],[data-testid="stMain"] [data-testid="stCaptionContainer"] p{color:#665B6F!important}
[data-testid="stMain"] button[kind="primary"]{background:#B85446;border-color:#B85446;color:white}
.stButton button,.stFormSubmitButton button,.stDownloadButton button{min-height:44px;border-radius:11px}
button p{color:inherit!important}.sf-context{display:inline-block;padding:7px 12px;border-radius:20px;background:#EAE0EF;color:#594366;font-size:12px;margin-bottom:16px}
.sf-brand{font-size:28px;font-weight:850;letter-spacing:-1.2px;color:#F2EAF7;padding:14px 0}.sf-brand span{color:#E7AD9D}
.sf-hero{display:grid;grid-template-columns:1.4fr 1fr;gap:30px;background:#30243E;color:#F5EFF6;border-radius:24px;overflow:hidden;padding:35px 38px;margin-bottom:24px;min-height:245px;position:relative}
.sf-hero h1{font-size:clamp(30px,3.3vw,44px);line-height:1.15;color:#F7EFF6!important;padding:0;margin:12px 0;letter-spacing:-1.5px}
.sf-hero p{color:#D4C7DB;font-size:14px;line-height:1.7;max-width:540px}.sf-eyebrow{font-size:11px;letter-spacing:2px;color:#EAB5A7;font-weight:750}
.sf-art{position:relative;display:flex;align-items:flex-end;justify-content:center;gap:15px;min-height:220px;padding-bottom:30px}
.sf-tower{width:54px;border-radius:9px 9px 4px 4px;background:linear-gradient(90deg,#B9A6DF,#8C6FB4);position:relative;box-shadow:10px 12px 0 #1D17242D}
.sf-tower:before{content:'';position:absolute;top:10px;left:12px;right:12px;height:3px;background:#ffffff66;box-shadow:0 18px #ffffff33,0 36px #ffffff33}
.sf-tower:nth-child(1){height:115px}.sf-tower:nth-child(2){height:155px;background:linear-gradient(90deg,#ECBA99,#C98775)}.sf-tower:nth-child(3){height:85px;background:linear-gradient(90deg,#A2D4BF,#65A38F)}
.sf-art-label{position:absolute;bottom:0;color:#CEBFD6;font-size:10px;letter-spacing:2px}
.sf-kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin-bottom:24px}
.sf-kpi{border:1px solid #E3D9E5;background:#FFFDFB;border-radius:17px;padding:18px;min-width:0}
.sf-kpi small{display:block;color:#706079;font-size:12px}.sf-kpi strong{display:block;color:#30243E;font-size:29px;line-height:1.5;overflow-wrap:anywhere}.sf-kpi span{font-size:11px;color:#706079}
.sf-materials{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:12px;margin:8px 0 22px}
.sf-material{border:1px solid #E3D9E5;border-radius:17px;background:#FFFDFC;padding:18px 15px;min-width:0;border-top:5px solid var(--accent)}
.sf-material b{font-size:16px;color:#30243E}.sf-material strong{display:block;margin:9px 0;font-size:25px;line-height:1.3;color:#30243E;overflow-wrap:anywhere}.sf-material small{font-size:11px;color:#75637B;display:block;line-height:1.8}
.sf-note{border-left:4px solid #D9826D;background:#F3E7E1;padding:14px 17px;border-radius:0 12px 12px 0;font-size:13px;line-height:1.7;color:#654A43;margin:10px 0 20px}
@media(max-width:1000px){.sf-materials{grid-template-columns:repeat(3,minmax(0,1fr))}.sf-kpis{grid-template-columns:repeat(2,minmax(0,1fr))}.sf-hero{grid-template-columns:1fr .7fr;padding:28px}}
@media(max-width:640px){.block-container{padding:4rem 1rem 3rem}.sf-hero{grid-template-columns:1fr;padding:26px}.sf-art{display:none}.sf-hero h1{font-size:32px}.sf-materials{grid-template-columns:repeat(2,minmax(0,1fr))}.sf-material{padding:14px}.sf-material strong{font-size:22px}.sf-kpi{padding:14px}.sf-kpi strong{font-size:24px}}
</style>"""


def current():
    return st.session_state["sf_demo"] if st.session_state["sf_workspace"] == "Exemple fictif" else st.session_state["sf_real"]


def save(config):
    config = validate(config)
    if st.session_state["sf_workspace"] == "Exemple fictif":
        st.session_state["sf_demo"] = config
    else:
        storage(config, st.session_state["sf_revision"])
    st.session_state.pop("sf_result", None)
    st.session_state.pop("sf_comparison", None)
    st.session_state["sf_notice"] = "Paramètres enregistrés. Relancer les simulations pour actualiser les résultats."


def load_demo():
    st.session_state["sf_demo"] = defaults(True)
    st.session_state["sf_workspace"] = "Exemple fictif"
    st.session_state.pop("sf_result", None)
    st.session_state.pop("sf_comparison", None)


def change_workspace():
    st.session_state["sf_workspace"] = st.session_state["_sf_workspace"]
    st.session_state.pop("sf_result", None)
    st.session_state.pop("sf_comparison", None)


def run_current():
    with st.spinner("Simulation des consommations et des approvisionnements…"):
        try:
            st.session_state["sf_result"] = simulate(current())
        except ValueError as exc:
            st.error(str(exc))


def valid_result():
    result = st.session_state.get("sf_result")
    if result and result["signature"] == fingerprint(current()):
        return result
    st.info("Renseigne les stocks, puis clique sur Lancer les simulations. L'exemple fictif permet de découvrir les résultats.")
    return None


def title(name, description):
    st.title(name); st.caption(description)


def kpis(items):
    st.html('<div class="sf-kpis">'+''.join(f'<div class="sf-kpi"><small>{html.escape(label)}</small><strong>{html.escape(str(value))}</strong><span>{html.escape(note)}</span></div>' for label,value,note in items)+'</div>')


def plot(fig, alt, height=340):
    fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=height,
                      font=dict(family="Segoe UI,Arial", color=INK, size=12), margin=dict(l=16,r=16,t=30,b=22),
                      legend=dict(orientation="h", y=-.2), hovermode="x unified")
    fig.update_xaxes(automargin=True, gridcolor="#E7DFE6"); fig.update_yaxes(automargin=True, gridcolor="#E7DFE6")
    st.plotly_chart(fig, width="stretch", theme=None, config={"displaylogo":False}, alt=alt)


def material_picker(result, key):
    ids = list(result["materials"])
    lookup = {m[0]:m[1] for m in MATERIALS}
    return st.selectbox("Matériau à examiner", ids, format_func=lambda x:lookup[x], key=key)


def stock_chart(result, key, examples=False):
    r = result["materials"][key]
    unit = next(m[2] for m in MATERIALS if m[0] == key)
    days = np.arange(result["config"]["horizon"]+1)
    q = r["stock_quantiles"]
    fig = go.Figure()
    fig.add_scatter(x=days,y=q[2],mode="lines",line=dict(width=0),name="Quantile 90 %",hoverinfo="skip",showlegend=False)
    fig.add_scatter(x=days,y=q[0],mode="lines",line=dict(width=0),fill="tonexty",fillcolor="rgba(153,122,181,.20)",name="Bande 10-90 %")
    fig.add_scatter(x=days,y=q[1],name="Stock médian",line=dict(color="#815DA2",width=3))
    fig.add_scatter(x=days,y=r["deterministic_stock"],name="Sans variation ni retard",line=dict(color=PINK,width=2,dash="dot"))
    if examples:
        for i, series in enumerate(r["examples"]):
            fig.add_scatter(x=days,y=series,name=f"Trajectoire {i+1}",line=dict(width=1),opacity=.45,showlegend=False)
    fig.update_layout(xaxis_title="Jour (J0 = stock initial)",yaxis_title=f"Stock de fin de jour ({unit})")
    plot(fig, f"Évolution simulée du stock en {unit}, médiane et quantiles 10 et 90 pour cent",420 if examples else 340)
    if examples:
        st.caption("Les cinq lignes fines sont cinq trajectoires individuelles simulées. La bande représente les quantiles 10-90 % de tous les scénarios.")


def dashboard():
    st.html('''<div class="sf-hero"><div><div class="sf-eyebrow">STOCKFLOW / APPROVISIONNEMENTS DE CHANTIER</div><h1>Anticiper le manque.<br>Garder le chantier en mouvement.</h1><p>Ciment, sable, gravier, bois et acier : visualise tes réserves et explore les dates possibles de rupture.</p></div><div class="sf-art" aria-hidden="true"><div class="sf-tower"></div><div class="sf-tower"></div><div class="sf-tower"></div><div class="sf-art-label">MESURER · SIMULER · ANTICIPER</div></div></div>''')
    config = current()
    st.html('<div class="sf-materials">'+''.join(
        f'<div class="sf-material" style="--accent:{color}"><b>{name}</b><strong>{fmt(row["stock"])} {unit}</strong><small>{fmt(row["daily"])} {unit}/jour</small><small>{"Activé" if row["enabled"] else "Exclu des simulations"}</small></div>'
        for row,(key,name,unit,color,icon) in zip(config["materials"],MATERIALS))+'</div>')
    with st.container(horizontal=True):
        if st.button("Lancer les simulations", type="primary", icon=":material/play_arrow:"):
            run_current()
        st.page_link(st.session_state["sf_pages"]["stocks"],label="Modifier mes stocks",icon=":material/edit:")
    result = st.session_state.get("sf_result")
    if not result or result["signature"] != fingerprint(config):
        st.html('<div class="sf-note">Commence par <b>Stocks et livraisons</b>, ou charge <b>l’exemple fictif</b> dans le menu. La simulation estime une première demande non satisfaite ; les paramètres restent sous ton contrôle.</div>')
        return
    active = result["materials"]
    medians = [int(r["day_quantiles"][1]) for r in active.values()]
    kpis([("Au moins un matériau en rupture",fmt(result["global_probability"]*100)+" %",f"sur {config['horizon']} jours"),
          ("Première médiane par matériau",day_label(min(medians),config["horizon"]),"parmi les matériaux activés"),
          ("Scénarios simulés",fmt(config["count"],0),"pour chaque matériau"),
          ("Matériaux étudiés",str(len(active)),"unités conservées séparément")])
    st.subheader("Quels matériaux risquent de manquer ?")
    frame = summary_frame(result)
    st.dataframe(frame[["Matériau","Unité","Probabilité (%)","Délai q10","Délai médian","Délai q90"]],hide_index=True,width="stretch",alt="Probabilités et délais de première rupture par matériau")
    st.caption(f"Un délai « > {config['horizon']} j » signifie que le quantile n'est pas atteint dans la période étudiée. Il ne donne pas une date précise après cette période.")
    key = material_picker(result,"dash_material")
    stock_chart(result,key)


def stocks_page():
    title("Stocks et livraisons","Saisis toutes les quantités dans l'unité indiquée. Les changements sont appliqués après enregistrement.")
    cfg = current()
    with st.form("global_parameters"):
        a,b,c = st.columns(3)
        with a: start = st.date_input("Premier jour de consommation (J1)", date.fromisoformat(cfg["start"]))
        with b: horizon = st.number_input("Période étudiée (jours)",7,180,cfg["horizon"])
        with c: count = st.selectbox("Nombre de simulations",[500,1000,2000,5000,10000],index=[500,1000,2000,5000,10000].index(cfg["count"]) if cfg["count"] in [500,1000,2000,5000,10000] else 2)
        seed = st.number_input("Graine aléatoire",0,2_000_000_000,cfg["seed"])
        if st.form_submit_button("Enregistrer la période",type="primary"):
            try:
                new=copy.deepcopy(cfg);new.update(start=start.isoformat(),horizon=horizon,count=count,seed=seed);save(new);st.rerun()
            except (ValueError,OSError,sqlite3.Error) as exc:st.error(str(exc))
    st.html('<div class="sf-note"><b>Convention :</b> la consommation a lieu pendant la journée. Une livraison reçue à J7 devient utilisable à la fin de J7, donc pour J8. Tous les jours calendaires consomment. Une livraison unique et des livraisons répétées sont deux choix distincts.</div>')
    for index,(key,name,unit,color,icon) in enumerate(MATERIALS):
        row=cfg["materials"][index]
        with st.expander(f"{name} · stock {fmt(row['stock'])} {unit}",expanded=index==0):
            with st.form("material_"+key):
                enabled=st.checkbox("Inclure ce matériau",row["enabled"])
                a,b=st.columns(2)
                with a:
                    stock=st.number_input(f"Stock actuel ({unit})",0.,1e9,float(row["stock"]),key="stock_"+key)
                    delivery=st.number_input(f"Quantité par livraison ({unit})",0.,1e9,float(row["delivery"]),key="delivery_"+key,help="0 signifie qu'aucune livraison n'est planifiée.")
                    variation=st.number_input("Variation de consommation (± %)",0.,100.,float(row["variation"]),key="var_"+key)
                with b:
                    daily=st.number_input(f"Consommation moyenne ({unit}/jour)",0.,1e7,float(row["daily"]),key="daily_"+key)
                    lead=st.number_input("Délai de la première livraison (jours)",1,90,row["lead"],key="lead_"+key)
                    late=st.number_input("Retard supplémentaire maximum (jours)",0,30,row["late"],key="late_"+key)
                periodic=st.checkbox("Répéter les livraisons",row["periodic"],key="periodic_"+key)
                interval=st.number_input("Intervalle entre livraisons prévues (jours)",1,180,row["interval"],key="interval_"+key,help="Utilisé uniquement si Répéter les livraisons est coché. Ce n'est pas le délai initial.")
                st.caption("La variation correspond à une loi triangulaire symétrique. Le retard est uniforme entre 0 et le maximum, en jours entiers. Mettre ces deux valeurs à 0 pour un test exact.")
                if st.form_submit_button("Enregistrer "+name,type="primary",width="stretch"):
                    try:
                        new=copy.deepcopy(cfg);new["materials"][index].update(stock=stock,daily=daily,delivery=delivery,lead=lead,late=late,variation=variation,periodic=periodic,interval=interval,enabled=enabled);save(new);st.rerun()
                    except (ValueError,OSError,sqlite3.Error) as exc:st.error(str(exc))


def simulations_page():
    title("Durées probables avant rupture","Une rupture signifie que le stock disponible ne suffit pas à satisfaire toute la demande d'une journée.")
    if st.button("Lancer les simulations",type="primary",icon=":material/play_arrow:"):run_current()
    result=valid_result()
    if not result:return
    key=material_picker(result,"sim_material")
    r=result["materials"][key];cfg=result["config"];h=cfg["horizon"]
    unit=next(m[2] for m in MATERIALS if m[0]==key)
    q=r["day_quantiles"]
    kpis([("Rupture sur la période",fmt(100*r["probability"])+" %",f"{r['hits']} cas sur {cfg['count']}"),
          ("Délai q10",day_label(q[0],h),"10 % des scénarios au plus tard"),
          ("Délai médian",day_label(q[1],h),"50 % des scénarios au plus tard"),
          ("Délai q90",day_label(q[2],h),"90 % des scénarios au plus tard")])
    lo,hi=r["interval"]
    st.caption(f"Intervalle de Wilson à 95 % pour la fréquence : {fmt(lo*100,2)} à {fmt(hi*100,2)} %. Il décrit la précision Monte-Carlo, sans valider les hypothèses du modèle.")
    fig=go.Figure(go.Scatter(x=np.arange(h+1),y=r["cumulative_probability"]*100,mode="lines",fill="tozeroy",line=dict(color=PINK,width=3),fillcolor="rgba(212,107,86,.15)"))
    fig.update_layout(xaxis_title="Jour",yaxis_title="P(première rupture déjà survenue) (%)",yaxis_range=[0,102]);plot(fig,"Probabilité cumulée de première rupture selon le jour")
    st.caption("Cette courbe ne diminue pas : une livraison ultérieure ne supprime pas une rupture déjà survenue.")
    days,counts=np.unique(r["first"][r["first"]<=h],return_counts=True)
    fig=go.Figure(go.Bar(x=days,y=counts,marker_color="#9471AF"))
    fig.update_layout(xaxis_title="Jour de première rupture",yaxis_title="Nombre de scénarios")
    plot(fig,"Distribution des premières ruptures observées dans la période")
    st.caption(f"{cfg['count']-r['hits']} scénario(s) sans rupture observée avant J{h} sont exclus de ces barres mais inclus dans les probabilités et quantiles. Leur date future reste inconnue.")
    stock_chart(result,key,examples=True)
    st.subheader("Un repère pour préparer une nouvelle commande")
    row=next(x for x in cfg["materials"] if x["id"]==key)
    target=r["lead_need95"]
    st.markdown(f"La demande cumulée pendant le délai d'une nouvelle commande a un quantile 95 % de **{fmt(target)} {unit}**. Stock actuel : **{fmt(row['stock'])} {unit}** ; écart à ce repère : **{fmt(max(0,target-row['stock']))} {unit}**.")
    st.caption("Ce calcul ignore les livraisons déjà prévues et les coûts. C'est une protection ponctuelle du délai sous le modèle, pas une quantité optimale à commander ni une garantie sur tout le chantier.")


def usage_page():
    title("Utilisation des matériaux","Comparer la demande prévue à ce que le stock permet effectivement de servir.")
    result=valid_result()
    if not result:return
    key=material_picker(result,"usage_material");r=result["materials"][key]
    unit=next(m[2] for m in MATERIALS if m[0]==key);h=result["config"]["horizon"]
    kpis([("Demande moyenne totale",fmt(r["mean_demand"].sum()),unit),
          ("Utilisation servie moyenne",fmt(r["mean_served"].sum()),unit),
          ("Demande non servie moyenne",fmt(r["mean_unmet"].sum()),unit),
          ("Livraisons moyennes",fmt(r["mean_receipts"].sum()),unit)])
    days=np.arange(1,h+1)
    fig=go.Figure()
    for field,label,color in [("mean_served","Utilisation servie",GREEN),("mean_unmet","Demande non servie",PINK)]:
        fig.add_bar(x=days,y=r[field],name=label,marker_color=color)
    fig.add_scatter(x=days,y=r["mean_demand"],name="Demande totale",line=dict(color=INK,width=2))
    fig.update_layout(barmode="stack",xaxis_title="Jour",yaxis_title=f"Quantité quotidienne ({unit})");plot(fig,"Demande, utilisation servie et demande non satisfaite par jour")
    fig=go.Figure()
    fig.add_scatter(x=days,y=np.cumsum(r["mean_demand"]),name="Demande cumulée",line=dict(color=INK,width=3))
    fig.add_scatter(x=days,y=np.cumsum(r["mean_served"]),name="Utilisation servie cumulée",line=dict(color=GREEN,width=3))
    fig.update_layout(xaxis_title="Jour",yaxis_title=f"Quantité cumulée ({unit})");plot(fig,"Demande et utilisation servie cumulées")
    st.caption("Ces courbes sont simulées, pas des consommations relevées sur le chantier. Les demandes manquées ne sont pas reportées. Les réceptions apparaissent dans le stock en fin de journée.")
    st.dataframe(daily_frame(result,key),hide_index=True,width="stretch",alt="Détail journalier simulé pour le matériau sélectionné")
    st.download_button("Télécharger le suivi journalier CSV",csv_bytes(daily_frame(result,key)),f"StockFlow_{key}_journalier.csv","text/csv")


def comparison_page():
    title("Comparer une solution","Tester davantage de stock, une livraison plus tôt ou une autre fréquence, sans modifier le scénario enregistré.")
    baseline=valid_result()
    if not baseline:return
    key=material_picker(baseline,"compare_material");cfg=baseline["config"];index=next(i for i,r in enumerate(cfg["materials"]) if r["id"]==key);row=cfg["materials"][index]
    unit=next(m[2] for m in MATERIALS if m[0]==key)
    with st.form("alternative_"+key):
        a,b=st.columns(2)
        with a:
            extra=st.number_input(f"Stock supplémentaire ({unit})",0.,1e9,0.)
            qty=st.number_input(f"Nouvelle quantité par livraison ({unit})",0.,1e9,float(row["delivery"]))
            periodic=st.checkbox("Livraisons répétées dans la solution",row["periodic"])
        with b:
            lead=st.number_input("Nouveau délai initial (jours)",1,90,row["lead"])
            late=st.number_input("Nouveau retard maximum (jours)",0,30,row["late"])
            interval=st.number_input("Nouvel intervalle de livraison (jours)",1,180,row["interval"])
        go_run=st.form_submit_button("Comparer les deux scénarios",type="primary",width="stretch")
    if go_run:
        try:
            alternative=copy.deepcopy(cfg);alternative["materials"][index].update(stock=row["stock"]+extra,delivery=qty,lead=lead,late=late,periodic=periodic,interval=interval)
            with st.spinner("Comparaison avec les mêmes tirages de demande…"):
                alt=simulate(alternative)
            st.session_state["sf_comparison"]=dict(signature=baseline["signature"],key=key,result=alt)
        except ValueError as exc:st.error(str(exc))
    stored=st.session_state.get("sf_comparison")
    if not stored or stored["signature"]!=baseline["signature"] or stored["key"]!=key:return
    alternative=stored["result"];base=baseline["materials"][key];alt=alternative["materials"][key]
    frame=pd.DataFrame({"Indicateur":["Risque de rupture (%)","Délai médian","Demande totale non servie moyenne ("+unit+")"],
                        "Situation enregistrée":[fmt(base["probability"]*100),day_label(base["day_quantiles"][1],cfg["horizon"]),fmt(base["unmet_total"].mean())],
                        "Solution testée":[fmt(alt["probability"]*100),day_label(alt["day_quantiles"][1],cfg["horizon"]),fmt(alt["unmet_total"].mean())]})
    st.dataframe(frame,hide_index=True,width="stretch",alt="Comparaison des probabilités et des pénuries des deux scénarios")
    fig=go.Figure()
    for r,name,col in [(base,"Situation enregistrée",PINK),(alt,"Solution testée",GREEN)]:
        fig.add_scatter(x=np.arange(cfg["horizon"]+1),y=r["cumulative_probability"]*100,name=name,line=dict(color=col,width=3))
    fig.update_layout(xaxis_title="Jour",yaxis_title="Probabilité cumulée de rupture (%)",yaxis_range=[0,102]);plot(fig,"Comparaison des probabilités cumulées de première rupture")
    tested=alternative["config"]["materials"][index]
    st.caption(f"Solution calculée : stock {fmt(tested['stock'])} {unit}, livraison {fmt(tested['delivery'])} {unit}, délai {tested['lead']} j, retard max {tested['late']} j, {'répétition tous les '+str(tested['interval'])+' j' if tested['periodic'] else 'une seule livraison'}. Resoumettre après toute modification du formulaire. Aucune solution n'est enregistrée automatiquement.")
    st.download_button("Exporter la comparaison CSV",csv_bytes(frame),"StockFlow_comparaison.csv","text/csv")


def guide_page():
    title("Guide et exports","Comprendre le modèle, tester les calculs et sauvegarder les paramètres.")
    st.markdown("**Parcours simple :** charger l'exemple fictif → Stocks et livraisons → Lancer les simulations → Durées avant rupture → Utilisation des matériaux → Comparer une solution.")
    st.markdown("**Définition :** rupture = première journée avec une demande supérieure au stock disponible. Un stock qui finit exactement à zéro ne manque que lors de la prochaine demande positive. Une livraison est ajoutée à la fin du jour prévu. La date de départ est J1, le stock initial est J0.")
    st.markdown("**Hypothèses :** consommation triangulaire symétrique autour de la moyenne, bornes moyenne × (1 ± variation %). Les demandes sont indépendantes entre jours et matériaux. Le retard de chaque livraison suit une loi uniforme entière de 0 au maximum choisi. Les livraisons périodiques ont des dates prévues D, D+intervalle, D+2×intervalle… ; les retards s'ajoutent à ces dates sans changer les suivantes.")
    st.markdown("**Interprétation :** les quantiles du délai incluent les scénarios sans rupture avant l'horizon. Une médiane > H jours signifie que moins de 50 % des scénarios ont rompu dans la période ; la date future n'est pas connue. Les bandes de stock 10-90 % décrivent les trajectoires. Les probabilités ne garantissent pas le chantier réel. Les consommations fractionnaires (sacs compris) sont une approximation continue.")
    st.subheader("Tests exacts à réaliser sur un matériau")
    st.caption("Mettre variation = 0 %, retard = 0 et horizon = 20 jours pour ces tests. Les autres matériaux peuvent être exclus.")
    cases=pd.DataFrame([
        [100,10,0,5,"Non","Première rupture J11"],
        [100,10,50,5,"Non","Première rupture J16"],
        [100,10,100,10,"Non","Aucune rupture avant J20, quantiles > 20 j"],
        [100,10,100,11,"Non","Rupture J11 ; livraison en fin de J11"],
        [100,10,50,5,"Oui, tous les 5 j","Aucune rupture avant J20"],
        [0,0,0,5,"Non","Aucune demande, aucune rupture"],
    ],columns=["Stock","Demande/j","Livraison","Délai j","Répétition","Résultat attendu"])
    st.dataframe(cases,hide_index=True,width="stretch",alt="Cas de test déterministes et résultats de première rupture attendus")
    result=st.session_state.get("sf_result")
    if result and result["signature"]==fingerprint(current()):
        st.subheader("Télécharger les résultats simulés")
        st.download_button("Synthèse CSV",csv_bytes(summary_frame(result)),"StockFlow_synthese.csv","text/csv")
        st.download_button("Classeur Excel",excel_bytes(result),"StockFlow_resultats.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.download_button("Rapport PDF",report_pdf(result,st.session_state["sf_workspace"]=="Exemple fictif"),"StockFlow_rapport.pdf","application/pdf")
        scenario_rows=[]
        for key,name,unit,*_ in MATERIALS:
            if key in result["materials"]:
                for i,first in enumerate(result["materials"][key]["first"]):
                    scenario_rows.append((i+1,name,int(first) if first<=result["config"]["horizon"] else None,first<=result["config"]["horizon"]))
        trials=pd.DataFrame(scenario_rows,columns=["Scénario","Matériau","Jour de première rupture observée","Rupture observée dans l'horizon"])
        st.download_button("Premières ruptures de tous les scénarios CSV",csv_bytes(trials),"StockFlow_scenarios.csv","text/csv")
    backup=dict(application="StockFlow",fictitious=st.session_state["sf_workspace"]=="Exemple fictif",config=current())
    st.download_button("Sauvegarder les paramètres JSON",json.dumps(backup,ensure_ascii=False,indent=2).encode("utf-8"),"StockFlow_parametres.json","application/json")
    with st.expander("Importer des paramètres sauvegardés"):
        with st.form("import_config"):
            uploaded=st.file_uploader("Fichier JSON StockFlow",type=["json"],max_upload_size=1)
            confirmed=st.checkbox("Remplacer les paramètres de l'espace actif")
            imported=st.form_submit_button("Importer les paramètres")
        if imported:
            try:
                if not uploaded or not confirmed:raise ValueError("Choisir le fichier et confirmer le remplacement.")
                if uploaded.size>1024*1024:raise ValueError("Fichier limité à 1 Mo.")
                payload=json.loads(uploaded.getvalue())
                if not isinstance(payload,dict) or payload.get("application")!="StockFlow":raise ValueError("Sauvegarde StockFlow non reconnue.")
                if payload.get("fictitious") and st.session_state["sf_workspace"]=="Mes données":raise ValueError("Importer cette sauvegarde fictive dans l'espace Exemple fictif.")
                save(payload.get("config"));st.rerun()
            except (ValueError,OSError,sqlite3.Error) as exc:st.error(str(exc))
    st.caption("Les paramètres Mes données sont enregistrés sur le serveur. Sur un hébergement à disque temporaire, conserver le JSON. Les simulations sont conservées dans la session et doivent être relancées après changement des paramètres.")


def main():
    st.set_page_config(page_title="StockFlow · Gestion des matériaux",page_icon=":material/inventory_2:",layout="wide")
    st.html(CSS)
    st.session_state.setdefault("sf_workspace","Mes données")
    st.session_state.setdefault("sf_demo",defaults(True))
    try:st.session_state["sf_real"],st.session_state["sf_revision"]=storage()
    except (ValueError,OSError,sqlite3.Error) as exc:st.error(str(exc));st.stop()
    definitions=[("accueil",dashboard,"Vue d’ensemble","dashboard"),("stocks",stocks_page,"Stocks et livraisons","inventory_2"),
                 ("simulations",simulations_page,"Durées avant rupture","query_stats"),("utilisation",usage_page,"Utilisation des matériaux","bar_chart"),
                 ("comparaison",comparison_page,"Comparer une solution","compare_arrows"),("guide",guide_page,"Guide et exports","menu_book")]
    pages={slug:st.Page(function,title=name,icon=f":material/{icon}:",url_path=slug,default=slug=="accueil") for slug,function,name,icon in definitions}
    st.session_state["sf_pages"]=pages
    navigation=st.navigation(list(pages.values()))
    with st.sidebar:
        st.html('<div class="sf-brand">Stock<span>Flow</span></div>')
        st.session_state["_sf_workspace"]=st.session_state["sf_workspace"]
        st.selectbox("Espace de travail",["Mes données","Exemple fictif"],key="_sf_workspace",on_change=change_workspace)
        st.button("Charger l'exemple fictif",icon=":material/science:",on_click=load_demo,key="load_demo",width="stretch")
        st.caption("5 matériaux · stocks et approvisionnements")
        st.caption("StockFlow 1.0 · Génie civil")
    st.html(f'<div class="sf-context">{st.session_state["sf_workspace"].upper()} · paramètres et résultats simulés identifiés</div>')
    if message:=st.session_state.pop("sf_notice",None):st.success(message)
    navigation.run()


if __name__=="__main__":main()

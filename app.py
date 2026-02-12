# app.py
# Dashboard moderne (palette ministère) – lecture d'une feuille Excel : "PLANNIF"
# Fichier: "Portefeuille 24-35_v8_opt_src_v9.xlsx"

import pandas as pd
import streamlit as st
import plotly.express as px
from datetime import datetime

# ================== CONFIG STREAMLIT ==================
st.set_page_config(page_title="Planification des investissements", layout="wide")

# ================== PALETTE MINISTÈRE (HEX) ==================
COL = {
    "bleu_profond": "#223B78",
    "vert_dynamique": "#8DBA47",
    "orange_flash": "#E28332",
    "bleu_ciel": "#4B9AD1",
    "bleu_marine": "#223651",
    "gris_acier": "#BCC3CF",
    "bg": "#F6F7FA",
    "card": "#FFFFFF",
    "border": "#E5E7EB",
    "muted": "#6B7280",
}

# ================== STYLE ==================
st.markdown(
    f"""
<style>
  .appview-container {{ background: {COL["bg"]}; }}
  .header {{
    background: {COL["bleu_profond"]};
    color: white;
    padding: 16px 18px;
    border-radius: 16px;
    border: 1px solid rgba(255,255,255,.15);
  }}
  .header-title {{ font-size: 18px; font-weight: 800; margin: 0; }}
  .header-sub {{ font-size: 12px; opacity: .9; margin-top: 4px; }}

  .kpi {{
    background: {COL["card"]};
    border: 1px solid {COL["border"]};
    border-radius: 16px;
    padding: 14px 16px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.06);
  }}
  .kpi .label {{
    color: {COL["muted"]};
    font-size: 12px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: .04em;
  }}
  .kpi .value {{
    color: {COL["bleu_marine"]};
    font-size: 26px;
    font-weight: 900;
    margin-top: 4px;
  }}
  .kpi .sub {{
    color: {COL["muted"]};
    font-size: 12px;
    margin-top: 4px;
  }}

  .section-title {{
    color: {COL["bleu_marine"]};
    font-weight: 900;
  }}
</style>
""",
    unsafe_allow_html=True,
)

# ================== PARAMS FICHIER ==================
FILE_PATH = r"Portefeuille 24-35_v8_opt_src_v9.xlsx"
SHEET_NAME = "PLANNIF"
REFRESH_SECONDS = 30  # actualisation auto

# ================== LOAD DATA ==================
@st.cache_data(ttl=REFRESH_SECONDS)
def load_data():
    df = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME)

    # Harmoniser colonnes (trim)
    df.columns = [str(c).strip() for c in df.columns]

    # Convertir numériques si présents
    for c in ["AC", "TR", "Budget année"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    # Normaliser Oui/Non si présent
    if "Projet annoncé" in df.columns:
        df["Projet annoncé"] = df["Projet annoncé"].astype(str).str.strip()

    return df


df = load_data()

# ================== EN-TÊTE ==================
today = datetime.now().strftime("%d/%m/%Y %H:%M")
st.markdown(
    f"""
<div class="header">
  <div class="header-title">PLANIFICATION DES INVESTISSEMENTS AÉROPORTUAIRES (000$)</div>
  <div class="header-sub">Données en date du {today} — mise à jour auto ~ {REFRESH_SECONDS}s</div>
</div>
""",
    unsafe_allow_html=True,
)
st.write("")

# ================== FILTRES ==================
f = df.copy()

colf1, colf2, colf3 = st.columns([1, 1, 1])

if "Année" in f.columns:
    years = sorted([y for y in f["Année"].dropna().unique().tolist()])
    default_years = years[-5:] if len(years) > 5 else years
    sel_years = colf1.multiselect("Années", years, default=default_years)
    if sel_years:
        f = f[f["Année"].isin(sel_years)]
else:
    colf1.info("Colonne 'Année' non trouvée.")

if "Type investissement" in f.columns:
    types = sorted([t for t in f["Type investissement"].dropna().unique().tolist()])
    sel_types = colf2.multiselect("Type investissement", types, default=types)
    if sel_types:
        f = f[f["Type investissement"].isin(sel_types)]
else:
    colf2.info("Colonne 'Type investissement' non trouvée (optionnelle).")

if "Projet annoncé" in f.columns:
    sel_ann = colf3.selectbox("Projet annoncé", ["Tous", "Oui", "Non"], index=0)
    if sel_ann != "Tous":
        f = f[f["Projet annoncé"].str.contains(sel_ann, na=False)]
else:
    colf3.info("Colonne 'Projet annoncé' non trouvée (optionnelle).")

st.write("")

# ================== KPI ==================
def kpi_card(label, value, sub=""):
    st.markdown(
        f"""
    <div class="kpi">
      <div class="label">{label}</div>
      <div class="value">{value}</div>
      <div class="sub">{sub}</div>
    </div>
    """,
        unsafe_allow_html=True,
    )

total_budget = f["Budget année"].sum() if "Budget année" in f.columns else 0
total_projects = f["Numéro projet"].nunique() if "Numéro projet" in f.columns else len(f)
annonces = (
    f["Projet annoncé"].str.contains("Oui", na=False).sum()
    if "Projet annoncé" in f.columns
    else None
)

# Heuristique "en travaux" (tu peux ajuster)
projects_travaux = (
    f["État avancement"]
    .astype(str)
    .str.contains("Plans et devis|AO Publié|Travaux", case=False, regex=True, na=False)
    .sum()
    if "État avancement" in f.columns
    else None
)

risk_val = f["TR"].sum() if "TR" in f.columns else None

k1, k2, k3, k4 = st.columns(4)
with k1:
    kpi_card("Total planification", f"{total_budget:,.0f} $", "Somme Budget année")
with k2:
    kpi_card("Projets (distincts)", f"{total_projects:,}", "Nb Numéro projet")
with k3:
    if annonces is None:
        kpi_card("Projets annoncés", "—", "Colonne absente")
    else:
        kpi_card("Projets annoncés", f"{annonces:,}", "Oui")
with k4:
    if risk_val is None:
        kpi_card("Valeurs à risques (TR)", "—", "Colonne absente")
    else:
        kpi_card("Valeurs à risques (TR)", f"{risk_val:,.0f} $", "Somme TR")

st.write("")

# ================== TABLEAU CENTRAL (MATRIX) ==================
st.markdown('<div class="section-title">Planification (tableau)</div>', unsafe_allow_html=True)

has_matrix = all(c in f.columns for c in ["Année", "Budget année"])
if has_matrix:
    if "Type investissement" in f.columns:
        idx = "Type investissement"
    elif "État avancement" in f.columns:
        idx = "État avancement"
    else:
        idx = None

    if idx is None:
        st.info("Ajoute 'Type investissement' (ou 'État avancement') pour des lignes plus utiles.")
    else:
        bud = f.pivot_table(
            index=idx,
            columns="Année",
            values="Budget année",
            aggfunc="sum",
            fill_value=0,
        )
        bud["Total"] = bud.sum(axis=1)
        bud.loc["Total actuel"] = bud.sum(axis=0)

        st.dataframe(bud.style.format("{:,.0f}"), use_container_width=True)

        # Bloc "Nombre de projets en travaux" si possible
        if "État avancement" in f.columns and "Numéro projet" in f.columns:
            work = f.copy()
            work["En travaux?"] = work["État avancement"].astype(str).str.contains(
                "Plans et devis|AO Publié|Travaux", case=False, regex=True, na=False
            )
            nb = work[work["En travaux?"]].pivot_table(
                index=idx,
                columns="Année",
                values="Numéro projet",
                aggfunc=pd.Series.nunique,
                fill_value=0,
            )
            nb["Total"] = nb.sum(axis=1)
            nb.loc["Total actuel"] = nb.sum(axis=0)

            st.write("")
            st.markdown(
                '<div class="section-title">Nombre de projets en travaux</div>',
                unsafe_allow_html=True,
            )
            st.dataframe(nb, use_container_width=True)
else:
    st.warning("Colonnes minimales requises: 'Année' + 'Budget année'.")

st.write("")

# ================== GRAPHIQUES ==================
st.markdown('<div class="section-title">Graphiques</div>', unsafe_allow_html=True)

g1, g2 = st.columns([1, 1])

# 1) Répartition par état d'avancement
with g1:
    if "État avancement" in f.columns:
        counts = f["État avancement"].astype(str).value_counts().reset_index()
        counts.columns = ["État", "Nombre"]
        fig = px.bar(counts, x="État", y="Nombre", title="Répartition par état d’avancement")
        fig.update_traces(marker_color=COL["bleu_profond"])
        fig.update_layout(
            plot_bgcolor="white",
            paper_bgcolor="white",
            font=dict(color=COL["bleu_marine"]),
            xaxis_tickangle=-35,
            margin=dict(l=10, r=10, t=50, b=90),
            showlegend=False,
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Colonne 'État avancement' non trouvée.")

# 2) Budget par année
with g2:
    if "Année" in f.columns and "Budget année" in f.columns:
        by = f.groupby("Année")["Budget année"].sum().reset_index()
        fig = px.line(by, x="Année", y="Budget année", markers=True, title="Budget par année")
        fig.update_traces(line_color=COL["vert_dynamique"])
        fig.update_layout(
            plot_bgcolor="white",
            paper_bgcolor="white",
            font=dict(color=COL["bleu_marine"]),
            margin=dict(l=10, r=10, t=50, b=10),
            showlegend=False,
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Colonnes 'Année' et/ou 'Budget année' non trouvées.")

st.write("")

# 3) Budget par site (si une colonne Site/Aéroport existe)
site_col = None
for candidate in ["Aéroport", "Site", "Localisation"]:
    if candidate in f.columns:
        site_col = candidate
        break

if site_col and "Budget année" in f.columns:
    st.markdown(
        f'<div class="section-title">Budget par {site_col} (tri décroissant)</div>',
        unsafe_allow_html=True,
    )
    site = f.groupby(site_col)["Budget année"].sum().sort_values(ascending=False).reset_index()
    fig = px.bar(site, x=site_col, y="Budget année", title="")
    fig.update_traces(marker_color=COL["bleu_profond"])
    fig.update_layout(
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(color=COL["bleu_marine"]),
        xaxis_title="",
        yaxis_title="Budget (000$)",
        xaxis_tickangle=-60,
        margin=dict(l=10, r=10, t=10, b=140),
        showlegend=False,
    )
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("Si tu veux le graphique “Budget par aéroport/site”, ajoute une colonne 'Aéroport' (ou 'Site').")

# ================== TABLE DETAIL (optionnelle) ==================
with st.expander("Voir la table (données filtrées)"):
    st.dataframe(f, use_container_width=True)

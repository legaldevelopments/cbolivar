"""
Tablero Predial Ciudad Bolívar 2026 — Streamlit
Análisis de prefacturado 2026 y límite máximo de incremento (80% vs 2025)
Fuentes: prefacturado.xlsx + PREDIOS EXCEDEN LIMITE 80%.xlsx

Ejecutar: streamlit run tablero_ciudad_bolivar.py
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import numpy as np
import os
import io

# ── CONFIG ────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Predial Ciudad Bolívar 2026",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

RUTA_PREF = os.path.join(os.path.dirname(__file__), "prefacturado.xlsx")
RUTA_EXC  = os.path.join(os.path.dirname(__file__), "PREDIOS EXCEDEN LIMITE 80%.xlsx")

AZUL_OSC  = "#1F4E79"
AZUL_MED  = "#2E75B6"
AZUL_CLAR = "#BDD7EE"
VERDE     = "#70AD47"
AMBAR     = "#FFD966"
ROJO      = "#FF7C80"
NARANJA   = "#F4B183"
GRIS      = "#D9D9D9"

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  .stApp { background-color: #F0F4F8; }
  [data-testid="stSidebar"] {
      background: linear-gradient(180deg, #1F4E79 0%, #2E75B6 100%);
  }
  [data-testid="stSidebar"] * { color: white !important; }
  [data-testid="stSidebar"] .stSelectbox label,
  [data-testid="stSidebar"] .stMultiSelect label,
  [data-testid="stSidebar"] .stSlider label { color: #BDD7EE !important; font-weight: 600; }

  .kpi-card {
      background: white; border-radius: 12px; padding: 18px 14px;
      text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.10);
      border-top: 4px solid #2E75B6; min-height: 120px;
      display: flex; flex-direction: column; justify-content: center;
  }
  .kpi-card.verde  { border-top-color: #70AD47; }
  .kpi-card.ambar  { border-top-color: #FFD966; }
  .kpi-card.rojo   { border-top-color: #FF7C80; }
  .kpi-card.oscuro { border-top-color: #1F4E79; }
  .kpi-card.naranja{ border-top-color: #F4B183; }
  .kpi-valor { font-size: 1.65rem; font-weight: 800; color: #1F4E79; line-height:1.1; }
  .kpi-label { font-size: 0.73rem; color: #666; margin-top:5px; font-weight:500;
               text-transform:uppercase; letter-spacing:0.5px; }
  .kpi-sub   { font-size: 0.78rem; color: #2E75B6; font-weight:600; margin-top:3px; }

  .sec-tit {
      background: linear-gradient(90deg, #1F4E79, #2E75B6);
      color: white; padding: 7px 16px; border-radius: 8px;
      font-size: 0.92rem; font-weight: 700; letter-spacing: 0.4px;
      margin: 20px 0 10px 0;
  }
  .header-main {
      background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 55%, #3A9AD9 100%);
      border-radius: 14px; padding: 26px 30px; color: white;
      margin-bottom: 22px; box-shadow: 0 4px 15px rgba(31,78,121,0.3);
  }
  .header-main h1 { margin:0; font-size:1.75rem; font-weight:800; }
  .header-main p  { margin:4px 0 0; opacity:0.85; font-size:0.9rem; }
  .alerta-rojo {
      background:#FFF0F0; border-left:4px solid #FF7C80;
      border-radius:6px; padding:10px 14px; font-size:0.82rem; color:#555; margin-top:6px;
  }
  .alerta-verde {
      background:#F0FFF0; border-left:4px solid #70AD47;
      border-radius:6px; padding:10px 14px; font-size:0.82rem; color:#555; margin-top:6px;
  }
  .nota-legal {
      background:#FFF9E6; border-left:4px solid #FFD966;
      border-radius:6px; padding:10px 14px; font-size:0.80rem; color:#555; margin-top:8px;
  }
</style>
""", unsafe_allow_html=True)


# ── CARGA DE DATOS ────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Cargando datos Ciudad Bolívar 2026...")
def cargar():
    # ── 1. PREFACTURADO ──────────────────────────────────────────────────────
    df_raw = pd.read_excel(RUTA_PREF, sheet_name="prefacturado")

    # Normalizar nombres de columnas
    df_raw.columns = [
        c.strip().lower()
         .replace(" ", "_")
         .replace("(", "")
         .replace(")", "")
        for c in df_raw.columns
    ]
    # Columna resultante: 'predial_periodo_trimestre'

    for col in ["base_gravable", "derecho", "tarifa", "predial_periodo_trimestre"]:
        df_raw[col] = pd.to_numeric(df_raw[col], errors="coerce").fillna(0)

    df_raw["matricula"] = df_raw["matricula"].astype(str).str.strip()

    # Agregar por matrícula (predio)
    # base_gravable es el mismo valor para todos los propietarios de un predio → "first"
    # predial_periodo_trimestre es proporcional al derecho de cada propietario → "sum"
    df_pref = df_raw.groupby("matricula").agg(
        ZONA          = ("sector",                  "first"),
        DEST_COD      = ("dest_igac",               "first"),
        DEST_NOM      = ("descripigac",             "first"),
        DIRECCION     = ("direccion",               "first"),
        N_PROPIETARIOS= ("cod_nit",                 "count"),
        BASE_GRAVABLE = ("base_gravable",           "first"),   # misma para todos los propietarios
        TARIFA        = ("tarifa",                  "first"),
        PREDIAL_TRIM  = ("predial_periodo_trimestre","sum"),
    ).reset_index()

    df_pref["PREDIAL_ANUAL_2026"] = df_pref["PREDIAL_TRIM"] * 4
    df_pref["ZONA"]     = df_pref["ZONA"].str.upper()
    df_pref["TARIFA_MIL"] = df_pref["TARIFA"] * 1000   # decimal → por mil

    # ── 2. ARCHIVO PREDIOS EXCEDEN LÍMITE 80% ────────────────────────────────
    df_exc = pd.read_excel(RUTA_EXC, sheet_name="Hoja4", header=0)

    # Renombrar por posición (evita problemas de codificación)
    cols_exc = [
        "llave_cruce", "nit", "nro_ficha", "matricula",
        "area_t_2025", "area_t_2026", "val_area_t",
        "area_c_2025", "area_c_2026", "val_area_c",
        "base_2025", "base_2026",
        "derecho_2025", "derecho_2026",
        "tarifa_orig", "tarifa_2025", "tarifa_2026",
        "dest_2025", "dest_2026", "val_dest",
        "exencion",
        "facturado_trim_2025", "total_anual_2025",
        "facturado_trim_2026", "total_anual_2026",
        "tope_80pct", "debido_liquidar_80pct", "regla_1",
    ]
    df_exc.columns = cols_exc[: len(df_exc.columns)]
    df_exc["matricula"] = df_exc["matricula"].astype(str).str.strip()

    for col in ["total_anual_2025", "total_anual_2026", "tope_80pct", "debido_liquidar_80pct"]:
        df_exc[col] = pd.to_numeric(df_exc[col], errors="coerce")

    # Agregar EXCEDEN por matrícula (suma de todos los propietarios)
    df_exc_agg = df_exc.groupby("matricula").agg(
        total_anual_2025      = ("total_anual_2025",      "sum"),
        total_anual_2026      = ("total_anual_2026",      "sum"),
        tope_80pct            = ("tope_80pct",            "sum"),
        debido_liquidar_80pct = ("debido_liquidar_80pct", "sum"),
        regla_1               = ("regla_1",               "first"),
    ).reset_index()

    # ── 3. JOIN ───────────────────────────────────────────────────────────────
    df = df_pref.merge(df_exc_agg, on="matricula", how="left")

    # ── 4. MÉTRICAS DERIVADAS ─────────────────────────────────────────────────
    df["_excede_80pct"]  = (df["regla_1"].fillna("").str.strip().str.upper() == "SE PASA")
    df["_tiene_hist"]    = df["total_anual_2025"].notna() & (df["total_anual_2025"] > 0)

    # Impuesto correcto: si excede → aplicar límite 80%, si no → el prefacturado
    df["IMPTO_CORRECTO_2026"] = np.where(
        df["_excede_80pct"],
        df["debido_liquidar_80pct"].fillna(df["PREDIAL_ANUAL_2026"]),
        df["PREDIAL_ANUAL_2026"],
    )

    # Exceso sobre el límite 80%
    df["EXCESO_80PCT"] = np.where(
        df["_excede_80pct"],
        (df["PREDIAL_ANUAL_2026"] - df["tope_80pct"]).clip(lower=0),
        0.0,
    )

    # Variación % (solo para predios con historial 2025 en EXCEDEN)
    df["VAR_PREDIAL_PCT"] = np.where(
        df["total_anual_2025"].fillna(0) > 0,
        (df["PREDIAL_ANUAL_2026"] - df["total_anual_2025"])
        / df["total_anual_2025"] * 100,
        np.nan,
    )

    # Rangos de base gravable
    def rango_base(v):
        if pd.isna(v) or v == 0: return "Sin dato"
        if v < 5e6:    return "0–5 M"
        if v < 15e6:   return "5–15 M"
        if v < 50e6:   return "15–50 M"
        if v < 150e6:  return "50–150 M"
        if v < 500e6:  return "150–500 M"
        return ">500 M"

    df["RANGO_BASE_2026"] = df["BASE_GRAVABLE"].apply(rango_base)
    orden_rng = ["0–5 M","5–15 M","15–50 M","50–150 M","150–500 M",">500 M","Sin dato"]
    df["_rango_ord"] = pd.Categorical(df["RANGO_BASE_2026"], categories=orden_rng, ordered=True)
    df["DEST_NOM"] = df["DEST_NOM"].fillna("Sin destino").str.title()

    return df


for _f, _lbl in [(RUTA_PREF, "prefacturado.xlsx"), (RUTA_EXC, "PREDIOS EXCEDEN LIMITE 80%.xlsx")]:
    if not os.path.exists(_f):
        st.error(f"No se encontró **{_lbl}**. Verifique que el archivo esté en la misma carpeta.")
        st.stop()

df_all = cargar()


# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🏛️ Predial Ciudad Bolívar 2026")
    st.markdown("**Análisis Prefacturado · Límite 80%**")
    st.divider()
    st.markdown("#### Filtros globales")

    sel_zona = st.selectbox("Zona:", ["Todos", "Solo Urbano", "Solo Rural"])
    sel_lim  = st.selectbox(
        "Estado límite 80%:",
        ["Todos", "Excede límite 80%", "Dentro del límite / sin historial"],
    )

    destinos_disp = sorted(df_all["DEST_NOM"].dropna().unique())
    sel_dest = st.multiselect("Destino IGAC:", destinos_disp, placeholder="Todos")

    rangos_disp = ["0–5 M","5–15 M","15–50 M","50–150 M","150–500 M",">500 M"]
    sel_rng = st.multiselect("Rango base gravable:", rangos_disp, placeholder="Todos")

    st.divider()
    umbral_var = st.slider("Umbral 'incremento alto' predial (%):", 20, 500, 80, 10)

    st.divider()
    st.markdown("""
    <div style='font-size:0.73rem; opacity:0.8;'>
    📂 Fuente: <i>prefacturado.xlsx</i><br>
    + <i>PREDIOS EXCEDEN LIMITE 80%.xlsx</i><br><br>
    📋 Límite aplicado:<br>
    · Predial 2026 ≤ Predial 2025 × 1,80<br>
    · Incremento máximo: 80%
    </div>""", unsafe_allow_html=True)


# ── FILTRAR ───────────────────────────────────────────────────────────────────
df = df_all.copy()
if sel_zona == "Solo Urbano":
    df = df[df["ZONA"] == "URBANO"]
elif sel_zona == "Solo Rural":
    df = df[df["ZONA"] == "RURAL"]

if sel_lim == "Excede límite 80%":
    df = df[df["_excede_80pct"]]
elif sel_lim == "Dentro del límite / sin historial":
    df = df[~df["_excede_80pct"]]

if sel_dest:
    df = df[df["DEST_NOM"].isin(sel_dest)]
if sel_rng:
    df = df[df["RANGO_BASE_2026"].isin(sel_rng)]


# ── HEADER ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-main">
  <h1>🏛️ Análisis Predial — Municipio de Ciudad Bolívar</h1>
  <p>Prefacturado 2026 · Comparativo vs 2025 · Límite máximo de incremento: 80% · Fuente: Sistema de Facturación</p>
</div>""", unsafe_allow_html=True)

n_f  = df["matricula"].nunique()
n_t  = df_all["matricula"].nunique()
n_u  = df[df["ZONA"] == "URBANO"]["matricula"].nunique()
n_r  = df[df["ZONA"] == "RURAL"]["matricula"].nunique()
if n_f < n_t:
    st.info(f"Mostrando **{n_f:,}** de **{n_t:,}** predios únicos según filtros — **{n_u:,} urbanos · {n_r:,} rurales**")
else:
    st.info(f"**{n_t:,}** predios únicos totales — **{n_u:,} urbanos · {n_r:,} rurales**")


# ── HELPERS ───────────────────────────────────────────────────────────────────
def fmt_cop(v):
    if pd.isna(v):
        return "$0"
    if abs(v) >= 1e12:
        return f"${v/1e12:,.2f} Bill."   # billones colombianos (10¹²)
    if abs(v) >= 1e6:
        return f"${v/1e6:,.0f} M"        # millones
    return f"${v:,.0f}"

def kpi(col, val, lab, sub="", color=""):
    with col:
        st.markdown(f"""
        <div class="kpi-card {color}">
          <div class="kpi-valor">{val}</div>
          <div class="kpi-label">{lab}</div>
          <div class="kpi-sub">{sub}</div>
        </div>""", unsafe_allow_html=True)


# ── KPIs PRINCIPALES ──────────────────────────────────────────────────────────
st.markdown('<div class="sec-tit">📊 Indicadores Clave — Prefacturado 2026</div>', unsafe_allow_html=True)

pred_26    = df["PREDIAL_ANUAL_2026"].fillna(0).sum()
corr_26    = df["IMPTO_CORRECTO_2026"].fillna(0).sum()
exceso_tot = df["EXCESO_80PCT"].fillna(0).sum()
n_excede   = df["_excede_80pct"].sum()

cols_kpi = st.columns(5)
kpi(cols_kpi[0], f"{n_f:,}",          "Total Predios",          sel_zona or "Todos",        "oscuro")
kpi(cols_kpi[1], fmt_cop(pred_26),    "Predial Anual 2026",     "Prefacturado (trim × 4)",  "ambar")
kpi(cols_kpi[2], fmt_cop(corr_26),    "Predial Correcto 2026",  "Con límite 80%",           "verde")
kpi(cols_kpi[3], f"{n_excede:,}",     "Exceden Límite 80%",     "Predios a corregir",       "rojo")
kpi(cols_kpi[4], fmt_cop(exceso_tot), "Exceso s/ Límite 80%",   "Cobrado de más",           "naranja")
st.markdown("<br>", unsafe_allow_html=True)

# KPIs comparativo 2025 (solo predios con historial en EXCEDEN)
pred_25    = df["total_anual_2025"].fillna(0).sum()
n_hist     = df["_tiene_hist"].sum()
tope_sum   = df["tope_80pct"].fillna(0).sum()
var_25_26  = (pred_26 - pred_25) / pred_25 * 100 if pred_25 > 0 else 0

if n_hist > 0:
    st.markdown('<div class="sec-tit">📅 Comparativo 2025 vs 2026 (predios con historial en archivo EXCEDEN)</div>',
                unsafe_allow_html=True)
    cols_kpi2 = st.columns(5)
    kpi(cols_kpi2[0], f"{n_hist:,}",         "Predios con historial 2025", "En archivo de análisis",       "oscuro")
    kpi(cols_kpi2[1], fmt_cop(pred_25),       "Predial Base 2025",          "Suma predios con historial",   "")
    kpi(cols_kpi2[2], fmt_cop(tope_sum),      "Tope 80% Acumulado",         "2025 × 1,80",                  "ambar")
    kpi(cols_kpi2[3], f"{var_25_26:+.1f}%",  "Var. Predial 2025→2026",     "Prefacturado vs 2025",         "rojo" if var_25_26 > 80 else "verde")
    kpi(cols_kpi2[4], fmt_cop(exceso_tot),    "Exceso Total s/ 80%",        "A devolver / corregir",        "naranja")
    st.markdown("<br>", unsafe_allow_html=True)


# ── SECCIÓN 1: COMPARATIVO FISCAL ────────────────────────────────────────────
st.markdown('<div class="sec-tit">📈 Comparativo Fiscal Prefacturado vs Correcto</div>', unsafe_allow_html=True)
c1, c2, c3 = st.columns([2.2, 1.8, 1.8])

with c1:
    vals_bar = [pred_25, pred_26, corr_26]
    etiq_bar = ["Predial 2025\n(historial parcial)", "Prefacturado 2026\n(anual)", "Correcto 2026\n(con límite 80%)"]
    colores_bar = [AZUL_CLAR, ROJO if pred_26 > corr_26 else AMBAR, VERDE]
    fig_bar = go.Figure(go.Bar(
        x=etiq_bar,
        y=vals_bar,
        marker_color=colores_bar,
        text=[fmt_cop(v) for v in vals_bar],
        textposition="outside", textfont=dict(size=11, color="black"),
    ))
    fig_bar.update_layout(
        title=dict(text="Comparativo Recaudo Predial ($)", font=dict(size=13, color=AZUL_OSC)),
        xaxis=dict(tickfont=dict(color="black")),
        yaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee", tickfont=dict(color="black")),
        plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(t=50, b=20, l=10, r=20), height=320, showlegend=False,
    )
    st.plotly_chart(fig_bar, use_container_width=True)

with c2:
    n_excede_v = int(df["_excede_80pct"].sum())
    n_ok       = n_f - n_excede_v
    fig_pie = go.Figure(go.Pie(
        labels=["Excede límite 80%", "Dentro del límite / sin historial"],
        values=[n_excede_v, n_ok],
        hole=0.52,
        marker_colors=[ROJO, VERDE],
        textinfo="percent+label", textfont=dict(size=9, color="black"),
        hovertemplate="%{label}: %{value:,}<extra></extra>",
    ))
    fig_pie.update_layout(
        title=dict(text="Estado Límite 80%", font=dict(size=13, color=AZUL_OSC)),
        showlegend=False, margin=dict(t=50, b=0, l=0, r=0), height=340,
        paper_bgcolor="white",
        annotations=[dict(text=f"<b>{n_f:,}</b><br>predios",
                          x=0.5, y=0.5, font_size=12, showarrow=False, font_color=AZUL_OSC)],
    )
    st.plotly_chart(fig_pie, use_container_width=True)

with c3:
    liq_u = float(df[df["ZONA"] == "URBANO"]["IMPTO_CORRECTO_2026"].fillna(0).sum())
    liq_r = float(df[df["ZONA"] == "RURAL"]["IMPTO_CORRECTO_2026"].fillna(0).sum())
    fig_zona = go.Figure(go.Pie(
        labels=["Urbano", "Rural"],
        values=[liq_u, liq_r],
        hole=0.55,
        marker_colors=[AZUL_MED, VERDE],
        textinfo="label+percent", textfont=dict(size=10),
        hovertemplate="%{label}: %{value:$,.0f}<extra></extra>",
    ))
    fig_zona.update_layout(
        title=dict(text="Predial Correcto por Zona", font=dict(size=13, color=AZUL_OSC)),
        showlegend=False, margin=dict(t=50, b=0, l=0, r=0), height=320,
        paper_bgcolor="white",
    )
    st.plotly_chart(fig_zona, use_container_width=True)


# ── SECCIÓN 2: DISTRIBUCIÓN POR DESTINO ──────────────────────────────────────
st.markdown('<div class="sec-tit">📉 Distribución por Destino Económico y Rango de Base</div>',
            unsafe_allow_html=True)
c4, c5 = st.columns(2)

with c4:
    df_dest_agg = (
        df.groupby("DEST_NOM")
        .agg(
            predios           = ("matricula",          "count"),
            predial_correcto  = ("IMPTO_CORRECTO_2026","sum"),
            exceso_total      = ("EXCESO_80PCT",        "sum"),
            n_excede          = ("_excede_80pct",       "sum"),
        )
        .reset_index()
        .sort_values("predial_correcto", ascending=True)
        .tail(14)
    )
    fig_dest = go.Figure(go.Bar(
        x=df_dest_agg["predial_correcto"],
        y=df_dest_agg["DEST_NOM"],
        orientation="h", marker_color=AZUL_MED,
        text=[fmt_cop(v) for v in df_dest_agg["predial_correcto"]],
        textposition="outside",
    ))
    fig_dest.update_layout(
        title=dict(text="Predial Correcto 2026 por Destino", font=dict(size=13, color=AZUL_OSC)),
        plot_bgcolor="white", paper_bgcolor="white",
        xaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee", tickfont=dict(color="black")),
        yaxis=dict(tickfont=dict(color="black")),
        margin=dict(t=50, b=20, l=10, r=90), height=360,
    )
    st.plotly_chart(fig_dest, use_container_width=True)

with c5:
    df_rng_agg = (
        df.groupby("RANGO_BASE_2026")
        .agg(
            total      = ("matricula",           "count"),
            excede     = ("_excede_80pct",        "sum"),
            pred_26_s  = ("PREDIAL_ANUAL_2026",  "sum"),
            corr_26_s  = ("IMPTO_CORRECTO_2026", "sum"),
            exceso_s   = ("EXCESO_80PCT",         "sum"),
        )
        .reset_index()
    )
    orden_r = ["0–5 M","5–15 M","15–50 M","50–150 M","150–500 M",">500 M"]
    df_rng_agg["_ord"] = df_rng_agg["RANGO_BASE_2026"].apply(
        lambda x: orden_r.index(x) if x in orden_r else 99)
    df_rng_agg = df_rng_agg.sort_values("_ord")

    fig_rng = go.Figure()
    fig_rng.add_trace(go.Bar(
        name="Prefacturado 2026", x=df_rng_agg["RANGO_BASE_2026"], y=df_rng_agg["pred_26_s"],
        marker_color=ROJO,
    ))
    fig_rng.add_trace(go.Bar(
        name="Correcto 2026 (80%)", x=df_rng_agg["RANGO_BASE_2026"], y=df_rng_agg["corr_26_s"],
        marker_color=VERDE,
    ))
    fig_rng.update_layout(
        title=dict(text="Predial por Rango de Base Gravable", font=dict(size=13, color=AZUL_OSC)),
        barmode="group", plot_bgcolor="white", paper_bgcolor="white",
        yaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee", tickfont=dict(color="black")),
        xaxis=dict(tickfont=dict(color="black")),
        legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="right", x=1),
        margin=dict(t=60, b=20, l=10, r=10), height=360,
    )
    st.plotly_chart(fig_rng, use_container_width=True)


# ── SECCIÓN 3: ANÁLISIS LÍMITE 80% ───────────────────────────────────────────
st.markdown('<div class="sec-tit">⚖️ Análisis de Predios que Exceden el Límite 80%</div>',
            unsafe_allow_html=True)
c6, c7 = st.columns(2)

with c6:
    df_exc_dest = (
        df[df["_excede_80pct"]]
        .groupby("DEST_NOM")
        .agg(n_excede=("matricula", "count"), exceso=("EXCESO_80PCT", "sum"))
        .reset_index()
        .sort_values("exceso", ascending=True)
    )
    if len(df_exc_dest) > 0:
        fig_vd = go.Figure()
        fig_vd.add_trace(go.Bar(
            x=df_exc_dest["exceso"], y=df_exc_dest["DEST_NOM"],
            orientation="h", marker_color=ROJO,
            text=[fmt_cop(v) for v in df_exc_dest["exceso"]], textposition="outside",
        ))
        fig_vd.update_layout(
            title=dict(text="Exceso sobre Límite 80% por Destino", font=dict(size=13, color=AZUL_OSC)),
            plot_bgcolor="white", paper_bgcolor="white",
            xaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee", tickfont=dict(color="black")),
            yaxis=dict(tickfont=dict(color="black")),
            margin=dict(t=50, b=20, l=10, r=90), height=360,
        )
        st.plotly_chart(fig_vd, use_container_width=True)
    else:
        st.success("No hay predios que excedan el límite 80% en la selección actual.")

with c7:
    # Histograma variación % (solo predios con historial)
    df_var = df[df["VAR_PREDIAL_PCT"].notna() & (df["VAR_PREDIAL_PCT"].abs() < 2000)].copy()
    if len(df_var) > 0:
        fig_hist = px.histogram(
            df_var, x="VAR_PREDIAL_PCT", nbins=50,
            color_discrete_sequence=[AZUL_MED],
            labels={"VAR_PREDIAL_PCT": "Variación Predial (%) 2025→2026"},
            title=f"Distribución Variación Predial — Predios con historial 2025",
        )
        fig_hist.add_vline(x=0,   line_dash="dash", line_color=AZUL_OSC, annotation_text="0%")
        fig_hist.add_vline(x=80,  line_dash="dot",  line_color=ROJO,
                           annotation_text="Límite 80%")
        fig_hist.update_layout(
            plot_bgcolor="white", paper_bgcolor="white",
            margin=dict(t=50, b=20, l=10, r=10), height=360,
            bargap=0.05,
            xaxis=dict(tickfont=dict(color="black")),
            yaxis=dict(title="N° predios", tickfont=dict(color="black")),
        )
        st.plotly_chart(fig_hist, use_container_width=True)
    else:
        st.info("Sin datos de variación 2025→2026 para la selección actual.")


# ── SECCIÓN 4: SCATTER BASE vs PREDIAL ───────────────────────────────────────
st.markdown('<div class="sec-tit">🔍 Relación Base Gravable 2026 vs Predial Anual</div>',
            unsafe_allow_html=True)

df_sc = df[df["BASE_GRAVABLE"].notna() & df["PREDIAL_ANUAL_2026"].notna()].copy()
if len(df_sc) > 0:
    df_sc["_estado"] = df_sc["_excede_80pct"].map(
        {True: "Excede límite 80%", False: "Dentro del límite"})
    muestra = df_sc.sample(min(len(df_sc), 2000), random_state=42)
    fig_sc = px.scatter(
        muestra,
        x="BASE_GRAVABLE", y="PREDIAL_ANUAL_2026",
        color="_estado",
        color_discrete_map={"Excede límite 80%": ROJO, "Dentro del límite": VERDE},
        hover_data={"matricula": True, "DEST_NOM": True, "ZONA": True,
                    "VAR_PREDIAL_PCT": ":.1f", "_estado": False},
        labels={
            "BASE_GRAVABLE":      "Base Gravable 2026 ($)",
            "PREDIAL_ANUAL_2026": "Predial Anual 2026 ($)",
            "_estado":            "Estado",
        },
        title="Base Gravable vs Predial Anual — muestra hasta 2.000 predios",
        opacity=0.6,
    )
    fig_sc.update_layout(
        plot_bgcolor="white", paper_bgcolor="white",
        height=380, margin=dict(t=50, b=20, l=10, r=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.01),
        xaxis=dict(tickformat="$,.0f", tickfont=dict(color="black")),
        yaxis=dict(tickformat="$,.0f", tickfont=dict(color="black")),
    )
    fig_sc.update_traces(marker=dict(size=5))
    st.plotly_chart(fig_sc, use_container_width=True)


# ── SECCIÓN 5: MATRIZ DINÁMICA ────────────────────────────────────────────────
st.markdown('<div class="sec-tit">📊 Matriz Dinámica: Destino × Rango de Base Gravable</div>',
            unsafe_allow_html=True)

tab1, tab2, tab3 = st.tabs(["N° Predios", "Predial Anual 2026 ($)", "Exceso s/ Límite 80% ($)"])

orden_rng_list = ["0–5 M","5–15 M","15–50 M","50–150 M","150–500 M",">500 M"]
df_mat = df[df["RANGO_BASE_2026"].isin(orden_rng_list)].copy()

def pivote(val_col, aggfunc="sum", fmt_fn=None):
    if df_mat.empty or val_col not in df_mat.columns:
        return pd.DataFrame()
    piv = pd.pivot_table(
        df_mat, values=val_col, index="DEST_NOM",
        columns="RANGO_BASE_2026", aggfunc=aggfunc, fill_value=0,
        observed=True,
    )
    cols_ord = [c for c in orden_rng_list if c in piv.columns]
    piv = piv[cols_ord]
    piv["TOTAL"] = piv.sum(axis=1)
    piv = piv.sort_values("TOTAL", ascending=False)
    fila_total = piv.sum(); fila_total.name = "▶ TOTAL"
    piv = pd.concat([piv, fila_total.to_frame().T])
    if fmt_fn:
        return piv.map(fmt_fn)
    return piv

with tab1:
    piv_n = pivote("matricula", aggfunc="count")
    if not piv_n.empty:
        st.dataframe(piv_n.astype(int).style.background_gradient(
            cmap="Blues", axis=None, subset=orden_rng_list,
        ), use_container_width=True, height=380)

with tab2:
    piv_p = pivote("PREDIAL_ANUAL_2026")
    if not piv_p.empty:
        st.dataframe(
            piv_p.style.format("${:,.0f}").background_gradient(cmap="Oranges", axis=None, subset=orden_rng_list),
            use_container_width=True, height=380,
        )

with tab3:
    piv_e = pivote("EXCESO_80PCT")
    if not piv_e.empty:
        st.dataframe(
            piv_e.style.format("${:,.0f}").background_gradient(cmap="Reds", axis=None, subset=orden_rng_list),
            use_container_width=True, height=380,
        )


# ── SECCIÓN 6: PREDIOS QUE EXCEDEN — DETALLE ─────────────────────────────────
st.markdown(f'<div class="sec-tit">🚨 Predios que Exceden el Límite 80% (variación ≥ {umbral_var}%)</div>',
            unsafe_allow_html=True)

df_ext = df[
    df["_excede_80pct"] | (df["VAR_PREDIAL_PCT"].fillna(0).abs() >= umbral_var)
].copy().sort_values("EXCESO_80PCT", ascending=False)

n_ext       = len(df_ext)
n_ext_exc   = int(df_ext["_excede_80pct"].sum())
st.markdown(
    f"<p style='color:#222; font-size:0.95rem;'>"
    f"<b>{n_ext:,}</b> predios con exceso o variación alta · "
    f"de estos <b>{n_ext_exc:,}</b> exceden el límite 80%.</p>",
    unsafe_allow_html=True,
)

cols_ext = [c for c in [
    "matricula", "ZONA", "DEST_NOM", "DIRECCION", "N_PROPIETARIOS",
    "BASE_GRAVABLE", "TARIFA_MIL",
    "total_anual_2025", "PREDIAL_ANUAL_2026", "tope_80pct",
    "IMPTO_CORRECTO_2026", "EXCESO_80PCT", "VAR_PREDIAL_PCT",
] if c in df_ext.columns]

col_cfg_ext = {}
for c in ["BASE_GRAVABLE","total_anual_2025","PREDIAL_ANUAL_2026","tope_80pct",
          "IMPTO_CORRECTO_2026","EXCESO_80PCT"]:
    if c in cols_ext:
        col_cfg_ext[c] = st.column_config.NumberColumn(c, format="$ %,.0f")
if "TARIFA_MIL" in cols_ext:
    col_cfg_ext["TARIFA_MIL"] = st.column_config.NumberColumn("TARIFA_MIL", format="%.2f ‰")
if "VAR_PREDIAL_PCT" in cols_ext:
    col_cfg_ext["VAR_PREDIAL_PCT"] = st.column_config.NumberColumn("VAR %", format="%.1f %%")

st.dataframe(df_ext[cols_ext].reset_index(drop=True),
             use_container_width=True, height=380, column_config=col_cfg_ext)

_buf_ext = io.BytesIO()
with pd.ExcelWriter(_buf_ext, engine="openpyxl") as _w:
    df_ext[cols_ext].to_excel(_w, index=False, sheet_name="Exceden_Limite_80")
st.download_button(
    "⬇️ Descargar predios que exceden límite (.xlsx)",
    data=_buf_ext.getvalue(),
    file_name="ciudad_bolivar_exceden_limite_80pct.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)


# ── SECCIÓN 7: RESUMEN POR DESTINO ────────────────────────────────────────────
st.markdown('<div class="sec-tit">📋 Resumen por Destino Económico</div>', unsafe_allow_html=True)

df_res = (
    df.groupby(["ZONA", "DEST_NOM"])
    .agg(
        N_Predios         = ("matricula",           "count"),
        Exceden_80pct     = ("_excede_80pct",        "sum"),
        Con_Historial_2025= ("_tiene_hist",          "sum"),
        Base_Gravable_2026= ("BASE_GRAVABLE",        "sum"),
        Predial_2025      = ("total_anual_2025",     "sum"),
        Predial_Anual_2026= ("PREDIAL_ANUAL_2026",  "sum"),
        Predial_Correcto  = ("IMPTO_CORRECTO_2026", "sum"),
        Exceso_80pct_Total= ("EXCESO_80PCT",         "sum"),
        Tope_80pct        = ("tope_80pct",           "sum"),
    )
    .reset_index()
    .sort_values(["ZONA", "Exceden_80pct"], ascending=[True, False])
)

df_res["Pct_Exceden"] = (
    df_res["Exceden_80pct"] / df_res["N_Predios"].replace(0, np.nan) * 100
).round(1)

col_cfg_res = {}
for c in ["Base_Gravable_2026","Predial_2025","Predial_Anual_2026",
          "Predial_Correcto","Exceso_80pct_Total","Tope_80pct"]:
    col_cfg_res[c] = st.column_config.NumberColumn(c, format="$ %,.0f")
col_cfg_res["Pct_Exceden"] = st.column_config.NumberColumn("% Exceden", format="%.1f %%")

st.dataframe(df_res.reset_index(drop=True), use_container_width=True,
             height=380, column_config=col_cfg_res)

csv_res = df_res.to_csv(index=False).encode("utf-8-sig")
st.download_button("⬇️ Descargar resumen por destino (.csv)", data=csv_res,
                   file_name="ciudad_bolivar_resumen_destino.csv", mime="text/csv")


# ── SECCIÓN 8: DETALLE COMPLETO ───────────────────────────────────────────────
st.markdown('<div class="sec-tit">📋 Detalle de Predios</div>', unsafe_allow_html=True)

cols_vis = [c for c in [
    "matricula", "N_PROPIETARIOS", "ZONA", "DEST_NOM", "DIRECCION",
    "BASE_GRAVABLE", "TARIFA_MIL",
    "total_anual_2025", "tope_80pct",
    "PREDIAL_ANUAL_2026", "IMPTO_CORRECTO_2026",
    "EXCESO_80PCT", "VAR_PREDIAL_PCT", "RANGO_BASE_2026",
] if c in df.columns]

busq = st.text_input("🔎 Buscar por matrícula, destino o dirección:", placeholder="Escriba para filtrar…")
df_vis = df[cols_vis].copy()
if busq:
    mask = df_vis.apply(
        lambda col: col.astype(str).str.contains(busq, case=False, na=False)
    ).any(axis=1)
    df_vis = df_vis[mask]

st.markdown(f"**{len(df_vis):,}** registros")

col_cfg_vis = {}
for c in ["BASE_GRAVABLE","total_anual_2025","tope_80pct",
          "PREDIAL_ANUAL_2026","IMPTO_CORRECTO_2026","EXCESO_80PCT"]:
    if c in cols_vis:
        col_cfg_vis[c] = st.column_config.NumberColumn(c, format="$ %,.0f")
if "TARIFA_MIL" in cols_vis:
    col_cfg_vis["TARIFA_MIL"] = st.column_config.NumberColumn("TARIFA_MIL", format="%.2f ‰")
if "VAR_PREDIAL_PCT" in cols_vis:
    col_cfg_vis["VAR_PREDIAL_PCT"] = st.column_config.NumberColumn("VAR %", format="%.1f %%")

st.dataframe(df_vis.reset_index(drop=True), use_container_width=True,
             height=450, column_config=col_cfg_vis)

csv_all = df_vis.to_csv(index=False).encode("utf-8-sig")
st.download_button("⬇️ Descargar tabla filtrada (.csv)", data=csv_all,
                   file_name="ciudad_bolivar_predios_2026.csv", mime="text/csv")


# ── NOTA LEGAL ────────────────────────────────────────────────────────────────
st.markdown("""
<div class="nota-legal">
⚖️ <strong>Nota:</strong> El límite analizado corresponde al tope de incremento del 80% sobre el predial anual 2025
(Predial 2026 ≤ Predial 2025 × 1,80). Los predios sin historial 2025 no tienen límite calculable.
El «Predial Anual 2026» se estima como el predial trimestral del prefacturado × 4.
Los datos del archivo <em>PREDIOS EXCEDEN LIMITE 80%.xlsx</em> cubren únicamente los predios que superan dicho tope.
Verificar con el Acuerdo Municipal vigente de Ciudad Bolívar y el sistema de facturación.
</div>""", unsafe_allow_html=True)

st.markdown(
    "<br><center style='color:#aaa; font-size:0.76rem;'>"
    "Municipio de Ciudad Bolívar · Predial 2026 · Límite: incremento máximo 80% vs 2025"
    "</center>", unsafe_allow_html=True,
)

"""
Tablero Ciudad Bolívar — Streamlit
Predial 2026 + ICA Industria y Comercio 2023-2024-2025
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
    page_title="Ciudad Bolívar · Predial & ICA",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

RUTA_AN  = os.path.join(os.path.dirname(__file__), "ANALISIS.xlsx")
RUTA_ICA = os.path.join(os.path.dirname(__file__), "ica")

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
  .kpi-label { font-size: 0.73rem; color: #333; margin-top:5px; font-weight:600;
               text-transform:uppercase; letter-spacing:0.5px; }
  .kpi-sub   { font-size: 0.80rem; color: #1F4E79; font-weight:600; margin-top:3px; }

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


# ── CARGA PREDIAL ─────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Cargando datos Predial 2026...")
def cargar():
    df_raw = pd.read_excel(RUTA_AN, sheet_name="maestro de entes (marzo)2026")

    # Columnas por posición (encoding seguro)
    COL_MAT   = df_raw.columns[7]   # Matrícula
    COL_CONTR = df_raw.columns[4]   # Contribuyente
    COL_DEST  = df_raw.columns[24]  # Destinación 2026
    COL_BASE  = df_raw.columns[17]  # Base gravable 2026
    COL_TAR26 = df_raw.columns[22]  # Tarifa 2026
    COL_DIGAC = df_raw.columns[29]  # Dest. IGAC

    for col in [COL_BASE, COL_TAR26]:
        df_raw[col] = pd.to_numeric(df_raw[col], errors="coerce")
    for col in ["TOTAL ANUAL 2025", "TOTAL ANUAL 2026", "TOPE DEL 80%", "DEBIDO LIQUIDAR 2026 (80%)"]:
        df_raw[col] = pd.to_numeric(df_raw[col], errors="coerce").fillna(0)

    df_raw["_excede"] = df_raw["REGLA 1"].astype(str).str.strip().str.upper() == "SE PASA"

    # Zona desde código de destinación: 01-19 = urbano, 20+ = rural
    def _zona(dest):
        try:
            return "RURAL" if int(str(dest).strip().split("-")[0].strip()) >= 20 else "URBANO"
        except Exception:
            return "URBANO"
    df_raw["_ZONA"] = df_raw[COL_DEST].apply(_zona)

    df = df_raw.groupby(COL_MAT).agg(
        ZONA                  = ("_ZONA",                        "first"),
        DEST_NOM              = (COL_DEST,                       "first"),
        DEST_COD              = (COL_DIGAC,                      "first"),
        N_PROPIETARIOS        = (COL_CONTR,                      "count"),
        BASE_GRAVABLE         = (COL_BASE,                       "sum"),
        TARIFA                = (COL_TAR26,                      "first"),
        total_anual_2025      = ("TOTAL ANUAL 2025",             "sum"),
        PREDIAL_ANUAL_2026    = ("TOTAL ANUAL 2026",             "sum"),
        tope_80pct            = ("TOPE DEL 80%",                 "sum"),
        debido_liquidar_80pct = ("DEBIDO LIQUIDAR 2026 (80%)",   "sum"),
        _excede_80pct         = ("_excede",                      "any"),
    ).reset_index().rename(columns={COL_MAT: "matricula"})

    df["TARIFA_MIL"] = df["TARIFA"] * 1000
    df["_tiene_hist"] = df["total_anual_2025"] > 0

    df["IMPTO_CORRECTO_2026"] = np.where(
        df["_excede_80pct"],
        df["debido_liquidar_80pct"].fillna(df["PREDIAL_ANUAL_2026"]),
        df["PREDIAL_ANUAL_2026"],
    )
    df["EXCESO_80PCT"] = np.where(
        df["_excede_80pct"],
        (df["PREDIAL_ANUAL_2026"] - df["tope_80pct"]).clip(lower=0),
        0.0,
    )
    df["VAR_PREDIAL_PCT"] = np.where(
        df["total_anual_2025"].fillna(0) > 0,
        (df["PREDIAL_ANUAL_2026"] - df["total_anual_2025"]) / df["total_anual_2025"] * 100,
        np.nan,
    )

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


# ── CARGA ICA ─────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Cargando datos ICA...")
def cargar_ica():
    if not os.path.isdir(RUTA_ICA):
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    ica_files = os.listdir(RUTA_ICA)
    partes = []

    for yr in [2023, 2024, 2025]:
        sufijo = " ACTIVOS (1).xls" if yr == 2024 else " ACTIVOS.xls"
        fname_m = os.path.join(RUTA_ICA, f"12_maest_ente AL 31 DIC {yr}{sufijo}")
        fname_d = os.path.join(RUTA_ICA, f"DECLARACIONES {yr}.xls")
        if not (os.path.exists(fname_m) and os.path.exists(fname_d)):
            continue

        # Maestro: header en fila 11; cols por índice (encoding roto)
        dm = pd.read_excel(fname_m, header=11)
        dm = dm[dm.iloc[:, 0].notna()].copy()
        dm["LICENCIA"]       = dm.iloc[:, 0].apply(lambda x: str(int(round(float(x))))[-8:])
        dm["TARIFA_MAESTRO"] = pd.to_numeric(dm.iloc[:, 8], errors="coerce")
        dm["DESTINACION"]    = dm.iloc[:, 12].astype(str).str.strip()
        dm_sel = dm[["LICENCIA", "DESTINACION", "TARIFA_MAESTRO"]].copy()

        # Declaraciones: header en fila 12
        dd = pd.read_excel(fname_d, header=12)
        dd_c = dd[dd.iloc[:, 2].notna()].copy()
        dd_c["LICENCIA"]      = dd_c.iloc[:, 2].astype(str).str.strip().str.zfill(8)
        dd_c["NOMBRE_EST"]    = dd_c.iloc[:, 4].astype(str).str.strip()
        dd_c["DECLARACION"]   = pd.to_numeric(dd_c.iloc[:, 9], errors="coerce").fillna(0)
        dd_c["VALOR_MENSUAL"] = pd.to_numeric(dd_c.iloc[:, 10], errors="coerce").fillna(0)
        dd_c["TARIFA_ACT"]    = pd.to_numeric(
            dd_c.iloc[:, 13].astype(str).str.replace(",", ".", regex=False),
            errors="coerce",
        )
        dd_sel = dd_c[["LICENCIA", "NOMBRE_EST", "DECLARACION", "VALOR_MENSUAL", "TARIFA_ACT"]].copy()

        merged = dm_sel.merge(dd_sel, on="LICENCIA", how="inner")
        merged["AÑO"] = yr
        partes.append(merged)

    if not partes:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    df_ica = pd.concat(partes, ignore_index=True)

    def shorten_dest(s):
        if pd.isna(s) or str(s).strip() in ("", "nan"):
            return "Sin clasificar"
        s = str(s).strip().replace("División", "Div.").replace("Divisón", "Div.")
        if " - " in s:
            prefix, rest = s.split(" - ", 1)
            return f"{prefix.strip()} — {rest[:50].strip()}"
        return s[:65]

    df_ica["DEST_CORTA"] = df_ica["DESTINACION"].apply(shorten_dest)

    # ── Anexo 2 ──────────────────────────────────────────────────────────────
    anexo_f = next((f for f in ica_files if "Anexo" in f), None)
    df_tarifas    = pd.DataFrame()
    df_act_princ  = pd.DataFrame()

    if anexo_f:
        fname_a = os.path.join(RUTA_ICA, anexo_f)
        try:
            df_t = pd.read_excel(fname_a, sheet_name="Tarifas PropuestasETM2026 (2)",
                                  header=None, skiprows=5)
            mask = pd.to_numeric(df_t.iloc[:, 0], errors="coerce").notna()
            df_t = df_t[mask].copy().reset_index(drop=True)
            rename_map = {0: "CIIU", 1: "DESCRIPCION", 2: "TARIFA_ETM2022",
                          3: "TARIFA_ACU2024", 4: "TARIFA_2025", 5: "N_CONTRIB"}
            df_t = df_t.rename(columns=rename_map)
            df_t["CIIU"] = df_t["CIIU"].apply(lambda x: str(int(round(float(x)))))
            for c in ["TARIFA_ETM2022", "TARIFA_ACU2024", "TARIFA_2025", "N_CONTRIB"]:
                df_t[c] = pd.to_numeric(df_t[c], errors="coerce")
            df_t["N_CONTRIB"] = df_t["N_CONTRIB"].fillna(0).astype(int)
            df_tarifas = df_t[["CIIU", "DESCRIPCION", "TARIFA_ETM2022",
                                "TARIFA_ACU2024", "TARIFA_2025", "N_CONTRIB"]].copy()
        except Exception:
            pass

        try:
            df_act = pd.read_excel(fname_a, sheet_name="ActPrincipales", header=2)
            new_cols = ["CANTIDAD", "CIIU", "DESCRIPCION", "TARIFA", "N_DECL", "TOTAL_ICA"]
            df_act.columns = new_cols + list(df_act.columns[6:])
            df_act = df_act[pd.to_numeric(df_act["CANTIDAD"], errors="coerce").notna()].copy()
            df_act["TOTAL_ICA"] = pd.to_numeric(df_act["TOTAL_ICA"], errors="coerce").fillna(0)
            df_act["N_DECL"]    = pd.to_numeric(df_act["N_DECL"],    errors="coerce").fillna(0).astype(int)
            df_act["TARIFA"]    = pd.to_numeric(df_act["TARIFA"],    errors="coerce")
            df_act_princ = df_act[["CANTIDAD", "CIIU", "DESCRIPCION", "TARIFA", "N_DECL", "TOTAL_ICA"]].copy()
        except Exception:
            pass

    return df_ica, df_tarifas, df_act_princ


# ── VERIFICAR ARCHIVOS PREDIAL ─────────────────────────────────────────────────
if not os.path.exists(RUTA_AN):
    st.error("No se encontró **ANALISIS.xlsx**. Verifique que el archivo esté en la misma carpeta.")
    st.stop()

df_all = cargar()


# ── HELPERS ───────────────────────────────────────────────────────────────────
def fmt_cop(v):
    if pd.isna(v):
        return "$0"
    if abs(v) >= 1e12:
        return f"${v/1e12:,.2f} Bill."
    if abs(v) >= 1e6:
        return f"${v/1e6:,.0f} M"
    return f"${v:,.0f}"

def kpi(col, val, lab, sub="", color=""):
    with col:
        st.markdown(f"""
        <div class="kpi-card {color}">
          <div class="kpi-valor">{val}</div>
          <div class="kpi-label">{lab}</div>
          <div class="kpi-sub">{sub}</div>
        </div>""", unsafe_allow_html=True)


# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🏛️ Ciudad Bolívar — Tablero Fiscal")
    st.markdown("**Predial 2026 · ICA 2023-2025**")
    st.divider()
    st.markdown("#### Filtros Predial")

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
    📂 Predial:<br>
    &nbsp;&nbsp;<i>ANALISIS.xlsx</i><br>
    &nbsp;&nbsp;<i>(hoja: maestro de entes (marzo)2026)</i><br><br>
    📂 ICA:<br>
    &nbsp;&nbsp;<i>ica/DECLARACIONES *.xls</i><br>
    &nbsp;&nbsp;<i>ica/12_maest_ente *.xls</i><br>
    &nbsp;&nbsp;<i>ica/Anexo 2 Tarifas.xlsx</i>
    </div>""", unsafe_allow_html=True)


# ── FILTRAR PREDIAL ───────────────────────────────────────────────────────────
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


# ── TABS PRINCIPALES ──────────────────────────────────────────────────────────
tab_predial, tab_ica = st.tabs(["🏠 Predial 2026", "🏭 ICA — Industria y Comercio"])


# ════════════════════════════════════════════════════════════════════════════════
# TAB 1 — PREDIAL 2026
# ════════════════════════════════════════════════════════════════════════════════
with tab_predial:

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

    # ── KPIs ─────────────────────────────────────────────────────────────────
    pred_25    = df["total_anual_2025"].fillna(0).sum()
    pred_26    = df["PREDIAL_ANUAL_2026"].fillna(0).sum()
    corr_26    = df["IMPTO_CORRECTO_2026"].fillna(0).sum()
    exceso_tot = df["EXCESO_80PCT"].fillna(0).sum()
    n_excede   = int(df["_excede_80pct"].sum())
    n_hist     = int(df["_tiene_hist"].sum())
    tope_sum   = df["tope_80pct"].fillna(0).sum()
    aumento    = pred_26 - pred_25
    var_25_26  = aumento / pred_25 * 100 if pred_25 > 0 else 0
    reduccion  = pred_26 - corr_26

    # Fila 1 — Comparativo anual
    st.markdown('<div class="sec-tit">📅 Comparativo Predial 2025 → 2026</div>', unsafe_allow_html=True)
    cols_kpi = st.columns(5)
    kpi(cols_kpi[0], f"{n_f:,}",             "Total Predios",          sel_zona or "Todos",           "oscuro")
    kpi(cols_kpi[1], fmt_cop(pred_25),        "Predial 2025",           "Facturado año anterior",      "")
    kpi(cols_kpi[2], fmt_cop(pred_26),        "Prefacturado 2026",      "Antes de aplicar tope 80%",   "ambar")
    kpi(cols_kpi[3], fmt_cop(aumento),        "Aumento Nominal",        f"vs 2025",                    "rojo" if var_25_26 > 80 else "naranja")
    kpi(cols_kpi[4], f"{var_25_26:+.1f}%",   "Variación 2025→2026",    "Incremento porcentual",       "rojo" if var_25_26 > 80 else "verde")
    st.markdown("<br>", unsafe_allow_html=True)

    # Fila 2 — Análisis límite 80%
    st.markdown('<div class="sec-tit">⚖️ Análisis del Límite Máximo de Incremento (80%)</div>', unsafe_allow_html=True)
    cols_kpi2 = st.columns(5)
    kpi(cols_kpi2[0], f"{n_excede:,}",        "Predios Exceden 80%",    f"De {n_hist:,} con historial",  "rojo")
    kpi(cols_kpi2[1], fmt_cop(tope_sum),       "Tope Máximo Legal",      "Suma 2025 × 1,80",              "ambar")
    kpi(cols_kpi2[2], fmt_cop(exceso_tot),     "Exceso sobre el Tope",   "Cobrado de más (a corregir)",   "naranja")
    kpi(cols_kpi2[3], fmt_cop(corr_26),        "Predial Correcto 2026",  "Con límite 80% aplicado",       "verde")
    kpi(cols_kpi2[4], fmt_cop(reduccion),      "Reducción a Aplicar",    "Prefacturado − Correcto",       "oscuro")
    st.markdown("<br>", unsafe_allow_html=True)

    # ── COMPARATIVO FISCAL ────────────────────────────────────────────────────
    st.markdown('<div class="sec-tit">📈 Comparativo Fiscal Prefacturado vs Correcto</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns([2.2, 1.8, 1.8])

    with c1:
        vals_bar = [pred_25, pred_26, corr_26]
        etiq_bar = ["Predial 2025\n(historial parcial)", "Prefacturado 2026\n(anual)", "Correcto 2026\n(con límite 80%)"]
        colores_bar = [AZUL_CLAR, ROJO if pred_26 > corr_26 else AMBAR, VERDE]
        fig_bar = go.Figure(go.Bar(
            x=etiq_bar, y=vals_bar, marker_color=colores_bar,
            text=[fmt_cop(v) for v in vals_bar], textposition="outside",
            textfont=dict(size=11, color="black"),
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
            values=[n_excede_v, n_ok], hole=0.52,
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
            labels=["Urbano", "Rural"], values=[liq_u, liq_r], hole=0.55,
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

    # ── DISTRIBUCIÓN POR DESTINO ──────────────────────────────────────────────
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
            x=df_dest_agg["predial_correcto"], y=df_dest_agg["DEST_NOM"],
            orientation="h", marker_color=AZUL_MED,
            text=[fmt_cop(v) for v in df_dest_agg["predial_correcto"]], textposition="outside",
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

    # ── ANÁLISIS LÍMITE 80% ───────────────────────────────────────────────────
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
        df_var = df[df["VAR_PREDIAL_PCT"].notna() & (df["VAR_PREDIAL_PCT"].abs() < 2000)].copy()
        if len(df_var) > 0:
            fig_hist = px.histogram(
                df_var, x="VAR_PREDIAL_PCT", nbins=50,
                color_discrete_sequence=[AZUL_MED],
                labels={"VAR_PREDIAL_PCT": "Variación Predial (%) 2025→2026"},
                title="Distribución Variación Predial — Predios con historial 2025",
            )
            fig_hist.add_vline(x=0,  line_dash="dash", line_color=AZUL_OSC, annotation_text="0%")
            fig_hist.add_vline(x=80, line_dash="dot",  line_color=ROJO,     annotation_text="Límite 80%")
            fig_hist.update_layout(
                plot_bgcolor="white", paper_bgcolor="white",
                margin=dict(t=50, b=20, l=10, r=10), height=360, bargap=0.05,
                xaxis=dict(tickfont=dict(color="black")),
                yaxis=dict(title="N° predios", tickfont=dict(color="black")),
            )
            st.plotly_chart(fig_hist, use_container_width=True)
        else:
            st.info("Sin datos de variación 2025→2026 para la selección actual.")

    # ── SCATTER BASE vs PREDIAL ───────────────────────────────────────────────
    st.markdown('<div class="sec-tit">🔍 Relación Base Gravable 2026 vs Predial Anual</div>',
                unsafe_allow_html=True)
    df_sc = df[df["BASE_GRAVABLE"].notna() & df["PREDIAL_ANUAL_2026"].notna()].copy()
    if len(df_sc) > 0:
        df_sc["_estado"] = df_sc["_excede_80pct"].map(
            {True: "Excede límite 80%", False: "Dentro del límite"})
        muestra = df_sc.sample(min(len(df_sc), 2000), random_state=42)
        fig_sc = px.scatter(
            muestra, x="BASE_GRAVABLE", y="PREDIAL_ANUAL_2026", color="_estado",
            color_discrete_map={"Excede límite 80%": ROJO, "Dentro del límite": VERDE},
            hover_data={"matricula": True, "DEST_NOM": True, "ZONA": True,
                        "VAR_PREDIAL_PCT": ":.1f", "_estado": False},
            labels={"BASE_GRAVABLE": "Base Gravable 2026 ($)",
                    "PREDIAL_ANUAL_2026": "Predial Anual 2026 ($)", "_estado": "Estado"},
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

    # ── MATRIZ DINÁMICA ───────────────────────────────────────────────────────
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
            columns="RANGO_BASE_2026", aggfunc=aggfunc, fill_value=0, observed=True,
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

    # ── PREDIOS QUE EXCEDEN — DETALLE ─────────────────────────────────────────
    st.markdown(f'<div class="sec-tit">🚨 Predios que Exceden el Límite 80% (variación ≥ {umbral_var}%)</div>',
                unsafe_allow_html=True)

    df_ext = df[
        df["_excede_80pct"] | (df["VAR_PREDIAL_PCT"].fillna(0).abs() >= umbral_var)
    ].copy().sort_values("EXCESO_80PCT", ascending=False)

    n_ext     = len(df_ext)
    n_ext_exc = int(df_ext["_excede_80pct"].sum())
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
        "⬇️ Descargar predios que exceden límite (.xlsx)", data=_buf_ext.getvalue(),
        file_name="ciudad_bolivar_exceden_limite_80pct.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # ── RESUMEN POR DESTINO ───────────────────────────────────────────────────
    st.markdown('<div class="sec-tit">📋 Resumen por Destino Económico</div>', unsafe_allow_html=True)

    df_res = (
        df.groupby(["ZONA", "DEST_NOM"])
        .agg(
            N_Predios          = ("matricula",           "count"),
            Exceden_80pct      = ("_excede_80pct",        "sum"),
            Con_Historial_2025 = ("_tiene_hist",          "sum"),
            Base_Gravable_2026 = ("BASE_GRAVABLE",        "sum"),
            Predial_2025       = ("total_anual_2025",     "sum"),
            Predial_Anual_2026 = ("PREDIAL_ANUAL_2026",  "sum"),
            Predial_Correcto   = ("IMPTO_CORRECTO_2026", "sum"),
            Exceso_80pct_Total = ("EXCESO_80PCT",         "sum"),
            Tope_80pct         = ("tope_80pct",           "sum"),
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

    # ── DETALLE COMPLETO ──────────────────────────────────────────────────────
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

    st.markdown("""
    <div class="nota-legal">
    ⚖️ <strong>Nota:</strong> El límite analizado corresponde al tope de incremento del 80% sobre el predial anual 2025
    (Predial 2026 ≤ Predial 2025 × 1,80). Los predios sin historial 2025 no tienen límite calculable.
    El «Predial Anual 2026» se estima como el predial trimestral del prefacturado × 4.
    Los datos del archivo <em>PREDIOS EXCEDEN LIMITE 80%.xlsx</em> cubren únicamente los predios que superan dicho tope.
    </div>""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════════
# TAB 2 — ICA
# ════════════════════════════════════════════════════════════════════════════════
with tab_ica:
    df_ica, df_tarifas, df_act_princ = cargar_ica()

    if df_ica.empty:
        st.warning("No se encontraron archivos ICA en la carpeta `ica/`. Verifique que los archivos estén presentes.")
    else:
        # ── HEADER ICA ────────────────────────────────────────────────────────
        st.markdown("""
        <div class="header-main">
          <h1>🏭 Industria y Comercio (ICA) · Ciudad Bolívar</h1>
          <p>Comparativo 2023 · 2024 · 2025 — Declaraciones por Destinación Económica · Maestro de Entes + Declaraciones</p>
        </div>""", unsafe_allow_html=True)

        # ── KPIs POR AÑO ──────────────────────────────────────────────────────
        st.markdown('<div class="sec-tit">📊 Resumen de Declaraciones ICA por Año</div>', unsafe_allow_html=True)

        col_23, col_24, col_25 = st.columns(3)
        colors_yr_kpi = {2023: "", 2024: "ambar", 2025: "verde"}
        for col_k, yr in zip([col_23, col_24, col_25], [2023, 2024, 2025]):
            df_yr = df_ica[df_ica["AÑO"] == yr]
            n_decl    = len(df_yr)
            tot_mens  = df_yr["VALOR_MENSUAL"].sum()
            tot_decl  = df_yr["DECLARACION"].sum()
            ica_anual = tot_mens * 12
            kpi(col_k,
                f"{yr}",
                f"Año Gravable · {n_decl:,} declarantes",
                f"Base declarada {fmt_cop(tot_decl)} · ICA anual est. {fmt_cop(ica_anual)}",
                colors_yr_kpi[yr])
        st.markdown("<br>", unsafe_allow_html=True)

        # ── EVOLUCIÓN 3 AÑOS ──────────────────────────────────────────────────
        st.markdown('<div class="sec-tit">📈 Evolución 2023 — 2024 — 2025</div>', unsafe_allow_html=True)

        resumen_anual = (
            df_ica.groupby("AÑO")
            .agg(N_DECL=("LICENCIA","count"),
                 TOTAL_DECL=("DECLARACION","sum"),
                 TOTAL_MENS=("VALOR_MENSUAL","sum"))
            .reset_index()
        )
        resumen_anual["ICA_ANUAL"] = resumen_anual["TOTAL_MENS"] * 12
        anios_str = resumen_anual["AÑO"].astype(str).tolist()

        ev1, ev2, ev3 = st.columns(3)
        with ev1:
            fig_ev1 = go.Figure(go.Bar(
                x=anios_str, y=resumen_anual["N_DECL"],
                marker_color=[AZUL_CLAR, AZUL_MED, AZUL_OSC],
                text=resumen_anual["N_DECL"].apply(lambda x: f"{x:,}"),
                textposition="outside",
            ))
            fig_ev1.update_layout(
                title=dict(text="N° Establecimientos Declarantes", font=dict(size=13, color=AZUL_OSC)),
                plot_bgcolor="white", paper_bgcolor="white",
                margin=dict(t=50, b=10, l=10, r=10), height=280,
                yaxis=dict(showgrid=True, gridcolor="#eee", tickfont=dict(color="black")),
                xaxis=dict(tickfont=dict(color="black")), showlegend=False,
            )
            st.plotly_chart(fig_ev1, use_container_width=True)

        with ev2:
            fig_ev2 = go.Figure(go.Bar(
                x=anios_str, y=resumen_anual["TOTAL_DECL"],
                marker_color=[AZUL_CLAR, AZUL_MED, AZUL_OSC],
                text=[fmt_cop(v) for v in resumen_anual["TOTAL_DECL"]],
                textposition="outside",
            ))
            fig_ev2.update_layout(
                title=dict(text="∑ Base Declarada ($)", font=dict(size=13, color=AZUL_OSC)),
                plot_bgcolor="white", paper_bgcolor="white",
                margin=dict(t=50, b=10, l=10, r=10), height=280,
                yaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee", tickfont=dict(color="black")),
                xaxis=dict(tickfont=dict(color="black")), showlegend=False,
            )
            st.plotly_chart(fig_ev2, use_container_width=True)

        with ev3:
            fig_ev3 = go.Figure(go.Bar(
                x=anios_str, y=resumen_anual["ICA_ANUAL"],
                marker_color=[VERDE, VERDE, AZUL_OSC],
                text=[fmt_cop(v) for v in resumen_anual["ICA_ANUAL"]],
                textposition="outside",
            ))
            fig_ev3.update_layout(
                title=dict(text="ICA Anual Estimado (V.Mensual × 12)", font=dict(size=13, color=AZUL_OSC)),
                plot_bgcolor="white", paper_bgcolor="white",
                margin=dict(t=50, b=10, l=10, r=10), height=280,
                yaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee", tickfont=dict(color="black")),
                xaxis=dict(tickfont=dict(color="black")), showlegend=False,
            )
            st.plotly_chart(fig_ev3, use_container_width=True)

        # ── COMPARATIVO POR DESTINACIÓN ───────────────────────────────────────
        st.markdown('<div class="sec-tit">🗂️ Comparativo por Destinación Económica — 2023 · 2024 · 2025</div>',
                    unsafe_allow_html=True)

        agg_dest = (
            df_ica.groupby(["DEST_CORTA", "AÑO"])
            .agg(
                N            = ("LICENCIA",       "count"),
                DECLARACION  = ("DECLARACION",    "sum"),
                VALOR_MENSUAL= ("VALOR_MENSUAL",  "sum"),
                TARIFA       = ("TARIFA_MAESTRO",
                                lambda x: x.mode().iloc[0] if len(x) > 0 else np.nan),
            )
            .reset_index()
        )
        agg_dest["ICA_ANUAL"] = agg_dest["VALOR_MENSUAL"] * 12

        # Top 15 por total declaración acumulada
        top_dest = (
            agg_dest.groupby("DEST_CORTA")["DECLARACION"].sum()
            .sort_values(ascending=False).head(15).index.tolist()
        )
        agg_top = agg_dest[agg_dest["DEST_CORTA"].isin(top_dest)].copy()

        _colors_yr = {2023: AZUL_CLAR, 2024: AZUL_MED, 2025: AZUL_OSC}

        tab_d1, tab_d2, tab_d3 = st.tabs(
            ["💰 Base Declarada", "💵 ICA Anual Estimado", "🏪 N° Establecimientos"]
        )

        for tab_sub, metrica, es_moneda in [
            (tab_d1, "DECLARACION",   True),
            (tab_d2, "ICA_ANUAL",     True),
            (tab_d3, "N",             False),
        ]:
            with tab_sub:
                orden = (
                    agg_top[agg_top["AÑO"] == 2025]
                    .set_index("DEST_CORTA")[metrica]
                    .reindex(top_dest, fill_value=0)
                    .sort_values()
                    .index.tolist()
                )
                fig_d = go.Figure()
                for yr in [2023, 2024, 2025]:
                    sub = agg_top[agg_top["AÑO"] == yr].set_index("DEST_CORTA")[metrica]
                    vals = [sub.get(d, 0) for d in orden]
                    fig_d.add_trace(go.Bar(
                        name=str(yr), x=vals, y=orden,
                        orientation="h", marker_color=_colors_yr[yr],
                        text=[fmt_cop(v) if es_moneda else f"{int(v):,}" for v in vals],
                        textposition="auto",
                    ))
                fig_d.update_layout(
                    barmode="group",
                    height=max(420, len(orden) * 42),
                    plot_bgcolor="white", paper_bgcolor="white",
                    margin=dict(t=20, b=20, l=10, r=10),
                    legend=dict(orientation="h", yanchor="bottom", y=1.01),
                    xaxis=dict(
                        tickformat="$,.0f" if es_moneda else ",d",
                        showgrid=True, gridcolor="#eee", tickfont=dict(color="black"),
                    ),
                    yaxis=dict(tickfont=dict(size=9, color="black")),
                )
                st.plotly_chart(fig_d, use_container_width=True)

        # ── TABLA COMPARATIVA PIVOT ────────────────────────────────────────────
        st.markdown('<div class="sec-tit">📋 Tabla Comparativa Destinación × Año</div>', unsafe_allow_html=True)

        piv_n    = agg_dest.pivot(index="DEST_CORTA", columns="AÑO", values="N").fillna(0).astype(int)
        piv_decl = agg_dest.pivot(index="DEST_CORTA", columns="AÑO", values="DECLARACION").fillna(0)
        piv_mens = agg_dest.pivot(index="DEST_CORTA", columns="AÑO", values="VALOR_MENSUAL").fillna(0)
        piv_tar  = agg_dest.pivot(index="DEST_CORTA", columns="AÑO", values="TARIFA")

        piv_n.columns    = [f"N_{c}"     for c in piv_n.columns]
        piv_decl.columns = [f"Decl_{c}"  for c in piv_decl.columns]
        piv_mens.columns = [f"Mens_{c}"  for c in piv_mens.columns]
        piv_tar.columns  = [f"Tarifa_{c}" for c in piv_tar.columns]

        df_pivot = pd.concat([piv_n, piv_decl, piv_mens, piv_tar], axis=1).reset_index()
        df_pivot.columns.name = None
        sort_col = "Decl_2025" if "Decl_2025" in df_pivot.columns else df_pivot.columns[1]
        df_pivot = df_pivot.sort_values(sort_col, ascending=False, na_position="last")

        col_cfg_piv = {}
        for c in df_pivot.columns:
            if c.startswith("Decl_") or c.startswith("Mens_"):
                col_cfg_piv[c] = st.column_config.NumberColumn(c, format="$ %,.0f")
            elif c.startswith("N_"):
                col_cfg_piv[c] = st.column_config.NumberColumn(c, format="%d")
            elif c.startswith("Tarifa_"):
                col_cfg_piv[c] = st.column_config.NumberColumn(c, format="%.1f ‰")

        st.dataframe(df_pivot.reset_index(drop=True), use_container_width=True,
                     height=400, column_config=col_cfg_piv)

        buf_piv = io.BytesIO()
        with pd.ExcelWriter(buf_piv, engine="openpyxl") as _w:
            df_pivot.to_excel(_w, index=False, sheet_name="ICA_Destinacion")
        st.download_button(
            "⬇️ Descargar tabla comparativa (.xlsx)", buf_piv.getvalue(),
            file_name="ica_comparativo_destino.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # ── TARIFAS VIGENTES (ANEXO 2) ─────────────────────────────────────────
        if not df_tarifas.empty:
            st.markdown('<div class="sec-tit">📑 Tarifas por Código CIIU (Anexo 2 — ETM)</div>',
                        unsafe_allow_html=True)

            solo_activos = st.checkbox("Mostrar solo actividades con contribuyentes registrados", value=True,
                                       key="solo_activos_ica")
            df_tar_show = (df_tarifas[df_tarifas["N_CONTRIB"] > 0].copy()
                           if solo_activos else df_tarifas.copy())
            df_tar_show = df_tar_show.sort_values("N_CONTRIB", ascending=False)

            col_cfg_tar = {
                "TARIFA_ETM2022":  st.column_config.NumberColumn("ETM 2022 (‰)",   format="%.1f"),
                "TARIFA_ACU2024":  st.column_config.NumberColumn("Acuerdo 2024 (‰)", format="%.1f"),
                "TARIFA_2025":     st.column_config.NumberColumn("2025 (‰)",        format="%.1f"),
                "N_CONTRIB":       st.column_config.NumberColumn("N° Contribuyentes", format="%d"),
            }
            st.dataframe(
                df_tar_show[["CIIU", "DESCRIPCION", "TARIFA_ETM2022",
                              "TARIFA_ACU2024", "TARIFA_2025", "N_CONTRIB"]].reset_index(drop=True),
                use_container_width=True, height=380, column_config=col_cfg_tar,
            )

            # Resaltar variaciones de tarifa entre períodos
            df_cambios = df_tar_show[
                df_tar_show["N_CONTRIB"] > 0
            ].copy()
            df_cambios = df_cambios[
                (df_cambios["TARIFA_ETM2022"].notna()) &
                (df_cambios["TARIFA_2025"].notna()) &
                (df_cambios["TARIFA_ETM2022"] != df_cambios["TARIFA_2025"])
            ]
            if not df_cambios.empty:
                st.markdown(
                    f'<div class="alerta-rojo">⚠️ <b>{len(df_cambios)}</b> actividades tienen tarifa diferente entre '
                    f'ETM-2022 y 2025. Verifique el acuerdo vigente.</div>',
                    unsafe_allow_html=True,
                )

        # ── TOP ACTIVIDADES (ANEXO 2) ──────────────────────────────────────────
        if not df_act_princ.empty:
            st.markdown('<div class="sec-tit">🏆 Principales Actividades por Recaudo ICA 2025</div>',
                        unsafe_allow_html=True)

            df_top = df_act_princ.sort_values("TOTAL_ICA", ascending=True).tail(15)
            etiq_top = (df_top["CIIU"].astype(str) + " — " +
                        df_top["DESCRIPCION"].astype(str).str[:55])

            fig_top = go.Figure(go.Bar(
                x=df_top["TOTAL_ICA"],
                y=etiq_top,
                orientation="h",
                marker_color=[AZUL_OSC if i >= len(df_top) - 3 else AZUL_MED
                              for i in range(len(df_top))],
                text=[fmt_cop(v) for v in df_top["TOTAL_ICA"]],
                textposition="outside",
                customdata=df_top[["TARIFA", "N_DECL"]].values,
                hovertemplate=(
                    "<b>%{y}</b><br>ICA Total: %{x:$,.0f}<br>"
                    "Tarifa: %{customdata[0]:.0f}‰ · N° Decl: %{customdata[1]}<extra></extra>"
                ),
            ))
            fig_top.update_layout(
                title=dict(text="Top 15 Actividades ICA 2025", font=dict(size=13, color=AZUL_OSC)),
                plot_bgcolor="white", paper_bgcolor="white",
                height=max(350, len(df_top) * 36),
                margin=dict(t=50, b=20, l=10, r=110),
                xaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee",
                           tickfont=dict(color="black")),
                yaxis=dict(tickfont=dict(size=9, color="black")),
                showlegend=False,
            )
            st.plotly_chart(fig_top, use_container_width=True)

            col_cfg_act = {
                "TOTAL_ICA": st.column_config.NumberColumn("Total ICA ($)", format="$ %,.0f"),
                "TARIFA":    st.column_config.NumberColumn("Tarifa (‰)",    format="%.0f"),
                "N_DECL":    st.column_config.NumberColumn("N° Decl.",      format="%d"),
            }
            st.dataframe(
                df_act_princ.sort_values("TOTAL_ICA", ascending=False).reset_index(drop=True),
                use_container_width=True, height=320, column_config=col_cfg_act,
            )

        # ── DETALLE POR ESTABLECIMIENTO ────────────────────────────────────────
        st.markdown('<div class="sec-tit">🔍 Detalle por Establecimiento</div>', unsafe_allow_html=True)

        cf1, cf2 = st.columns([2, 1])
        with cf1:
            busq_ica = st.text_input(
                "🔎 Buscar (licencia, nombre, destinación):",
                placeholder="Escriba para filtrar…", key="busq_ica",
            )
        with cf2:
            sel_anio = st.multiselect(
                "Año:", [2023, 2024, 2025], default=[2023, 2024, 2025], key="sel_anio_ica",
            )

        anios_sel = sel_anio if sel_anio else [2023, 2024, 2025]
        df_det = df_ica[df_ica["AÑO"].isin(anios_sel)].copy()

        if busq_ica:
            mask_ica = df_det.apply(
                lambda col: col.astype(str).str.contains(busq_ica, case=False, na=False)
            ).any(axis=1)
            df_det = df_det[mask_ica]

        df_det["DIF_TARIFA"] = (
            df_det["TARIFA_MAESTRO"].round(1) != df_det["TARIFA_ACT"].round(1)
        )

        n_dif = int(df_det["DIF_TARIFA"].sum())
        msg_dif = (f"  ·  ⚠️ **{n_dif:,}** con tarifa declarada ≠ tarifa en maestro" if n_dif > 0 else "")
        st.markdown(f"**{len(df_det):,}** registros{msg_dif}")

        cols_det = ["AÑO", "LICENCIA", "NOMBRE_EST", "DEST_CORTA",
                    "TARIFA_MAESTRO", "TARIFA_ACT", "DECLARACION", "VALOR_MENSUAL", "DIF_TARIFA"]
        col_cfg_det = {
            "DECLARACION":    st.column_config.NumberColumn("Declaración ($)",      format="$ %,.0f"),
            "VALOR_MENSUAL":  st.column_config.NumberColumn("V. Mensual ($)",       format="$ %,.0f"),
            "TARIFA_MAESTRO": st.column_config.NumberColumn("Tarifa Maestro (‰)",   format="%.1f"),
            "TARIFA_ACT":     st.column_config.NumberColumn("Tarifa Declarada (‰)", format="%.1f"),
            "DIF_TARIFA":     st.column_config.CheckboxColumn("⚠️ Dif. Tarifa"),
        }
        st.dataframe(df_det[cols_det].reset_index(drop=True),
                     use_container_width=True, height=440, column_config=col_cfg_det)

        buf_det = io.BytesIO()
        with pd.ExcelWriter(buf_det, engine="openpyxl") as _w:
            df_det[cols_det].to_excel(_w, index=False, sheet_name="ICA_Detalle")
        st.download_button(
            "⬇️ Descargar detalle ICA (.xlsx)", buf_det.getvalue(),
            file_name="ica_detalle_establecimientos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.markdown("""
        <div class="nota-legal">
        ⚖️ <strong>Nota ICA:</strong> La <em>Base Declarada</em> es el ingreso bruto mensual reportado por el establecimiento.
        El <em>Valor Mensual</em> = Base × Tarifa‰ (o mínimo estatutario si aplica).
        La <em>ICA Anual Estimada</em> = Valor Mensual × 12. Fuente: Maestros de Entes activos al 31 Dic de cada año
        + Declaraciones de establecimientos. Cruce por Licencia (últimos 8 dígitos del Código maestro).
        </div>""", unsafe_allow_html=True)


st.markdown(
    "<br><center style='color:#aaa; font-size:0.76rem;'>"
    "Municipio de Ciudad Bolívar · Análisis Fiscal 2026"
    "</center>", unsafe_allow_html=True,
)

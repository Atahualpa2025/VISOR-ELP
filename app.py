# -*- coding: utf-8 -*-
# VISOR ELECTROPERÚ – CMG + Caudales La Mejorada
# Versión FINAL: leyenda centrada, Santa Rosa por defecto,
# línea límite de último medido, exportación a Excel, CMG en medias horas.

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime, timedelta
import streamlit as st
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from io import BytesIO

st.set_page_config(page_title="Visor CMG + Caudales", layout="wide")
FILE = Path.cwd() / "Fuente.xlsx"

# ============================================================
# UTILIDADES
# ============================================================

def safe_rerun():
    if hasattr(st, "rerun"):
        st.rerun()
    elif hasattr(st, "experimental_rerun"):
        st.experimental_rerun()

def window_filter(df, start, end):
    if df.empty:
        return df
    return df[(df["datetime"] >= start) & (df["datetime"] <= end)]

def max_dt(*dfs):
    vals = [df["datetime"].max() for df in dfs if not df.empty]
    return max(vals) if vals else None

# ============================================================
# CMG – Opción A: medias horas reconstruidas por índice
# ============================================================

def build_datetime_cmg(df, col_fecha="Fecha"):
    df = df.copy()
    df[col_fecha] = pd.to_datetime(df[col_fecha], errors="coerce")
    df = df.dropna(subset=[col_fecha])

    final_dt = []
    for fecha, block in df.groupby(col_fecha):
        n = len(block)
        horas = pd.to_timedelta(np.arange(n) * 30, unit="m")
        final_dt.extend(pd.to_datetime(fecha) + horas)

    df["datetime"] = final_dt
    return df

@st.cache_data(show_spinner=False, ttl=30)
def load_cmg():
    pdo = pd.read_excel(FILE, sheet_name="CMG-PDO", engine="openpyxl")
    cos = pd.read_excel(FILE, sheet_name="CMG-COS", engine="openpyxl")

    pdo = build_datetime_cmg(pdo)
    cos = build_datetime_cmg(cos)

    ignore = ["fecha", "hora", "datetime"]
    barras = [c for c in pdo.columns if c.lower() not in ignore]

    def melt(df):
        out = df[["datetime"] + barras].melt(
            id_vars="datetime",
            value_vars=barras,
            var_name="barra",
            value_name="valor"
        )
        out["valor"] = pd.to_numeric(out["valor"], errors="coerce")
        return out.dropna()

    return melt(pdo), melt(cos), barras

# ============================================================
# HIDRO‑ELP – Lectura EXACTA Fecha + Hora
# ============================================================

def build_datetime_hidro(df):
    df = df.copy()
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce", dayfirst=True)

    # Construcción EXACTA: Fecha + " " + Hora (HH:MM)
    df["datetime"] = pd.to_datetime(
        df["Fecha"].dt.strftime("%Y-%m-%d") + " " + df["Hora"].astype(str),
        errors="coerce"
    )
    return df.dropna(subset=["datetime"])

@st.cache_data(show_spinner=False, ttl=30)
def load_hidro():
    df = pd.read_excel(FILE, sheet_name="Hidro-ELP", engine="openpyxl")
    df = build_datetime_hidro(df)

    col_med = [c for c in df.columns if "Medido" in c][0]
    col_proy = [c for c in df.columns if "Proye" in c][0]

    med = df[["datetime", col_med]].rename(columns={col_med:"valor"})
    proy = df[["datetime", col_proy]].rename(columns={col_proy:"valor"})

    med["valor"] = pd.to_numeric(med["valor"], errors="coerce")
    proy["valor"] = pd.to_numeric(proy["valor"], errors="coerce")

    return med.dropna(), proy.dropna()

# ============================================================
# UI
# ============================================================

st.title("Visor CMG + Caudales – ELECTROPERÚ")

c1, c2, c3 = st.columns([1,1,1])
with c1:
    past_hours = st.slider("Horas hacia atrás", 1, 24, 12)
with c2:
    if st.button("🔄 Refrescar"):
        st.cache_data.clear()
        safe_rerun()
with c3:
    export_btn = st.button("📤 Exportar Excel")

# ============================================================
# CARGAR DATA
# ============================================================

pdo_long, cos_long, barras = load_cmg()
med, proy = load_hidro()

# Santa Rosa por defecto
default_bar = ["Santa Rosa"] if "Santa Rosa" in barras else [barras[0]]

sel_bars = st.multiselect("Selecciona barras CMG:", barras, default=default_bar)

now = datetime.now()
start = now - timedelta(hours=past_hours)

end_CMG = max_dt(cos_long)
end_Caudal = max_dt(proy)
end = max(v for v in [end_CMG, end_Caudal] if v is not None)

pdo_w = window_filter(pdo_long[pdo_long["barra"].isin(sel_bars)], start, end)
cos_w = window_filter(cos_long[cos_long["barra"].isin(sel_bars)], start, end)
med_w = window_filter(med, start, end)
proy_w = window_filter(proy, start, end)

ultima_medida = med_w["datetime"].max() if not med_w.empty else None

# ============================================================
# GRAFICO CMG
# ============================================================

st.subheader("Costos Marginales – PDO vs COS")

fig_cmg = make_subplots(specs=[[{"secondary_y": False}]])
pal = ["#1f77b4","#ff7f0e","#2ca02c","#9467bd","#8c564b","#e377c2"]

for i, b in enumerate(sel_bars):
    color = pal[i % len(pal)]
    seg_pdo = pdo_w[pdo_w["barra"] == b]
    seg_cos = cos_w[cos_w["barra"] == b]

    fig_cmg.add_trace(go.Scatter(
        x=seg_pdo["datetime"], y=seg_pdo["valor"],
        mode="lines", line=dict(color=color), name=f"{b} – PDO"
    ))
    fig_cmg.add_trace(go.Scatter(
        x=seg_cos["datetime"], y=seg_cos["valor"],
        mode="lines", line=dict(color=color, dash="dash"), name=f"{b} – COS"
    ))

fig_cmg.update_layout(
    height=430,
    legend=dict(orientation="h", x=0.5, xanchor="center", y=-0.30),
    margin=dict(l=40, r=40, t=10, b=80)
)
fig_cmg.update_xaxes(title="Fecha‑Hora")
fig_cmg.update_yaxes(title="CMg (PEN/MWh)")

st.plotly_chart(fig_cmg, use_container_width=True)

# ============================================================
# GRAFICO CAUDALES
# ============================================================

st.subheader("Caudal – La Mejorada (Medido vs Proyección)")

fig_cau = make_subplots(specs=[[{"secondary_y": False}]])

fig_cau.add_trace(go.Scatter(
    x=med_w["datetime"], y=med_w["valor"],
    mode="lines", line=dict(color="#1f77b4"),
    name="Medido"
))
fig_cau.add_trace(go.Scatter(
    x=proy_w["datetime"], y=proy_w["valor"],
    mode="lines", line=dict(color="#d62728", dash="dash"),
    name="Proyección"
))

# ---- PARCHE: LÍNEA VERTICAL (shape) ----
if ultima_medida is not None:
    fig_cau.add_shape(
        type="line",
        x0=ultima_medida,
        x1=ultima_medida,
        y0=0,
        y1=1,
        xref="x",
        yref="paper",
        line=dict(color="gray", width=2, dash="dot")
    )

    fig_cau.add_annotation(
        x=ultima_medida,
        y=1.05,
        xref="x",
        yref="paper",
        text="Última medida",
        showarrow=False,
        font=dict(size=11, color="gray"),
        align="center"
    )

fig_cau.update_layout(
    height=430,
    legend=dict(orientation="h", x=0.5, xanchor="center", y=-0.30),
    margin=dict(l=40, r=40, t=10, b=80)
)
fig_cau.update_xaxes(title="Fecha‑Hora")
fig_cau.update_yaxes(title="Caudal (m³/s)")

st.plotly_chart(fig_cau, use_container_width=True)

# ============================================================
# EXPORTAR EXCEL
# ============================================================

if export_btn:

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pdo_w.to_excel(writer, sheet_name="CMG_PDO", index=False)
        cos_w.to_excel(writer, sheet_name="CMG_COS", index=False)
        med_w.to_excel(writer, sheet_name="Caudal_Medido", index=False)
        proy_w.to_excel(writer, sheet_name="Caudal_Proyeccion", index=False)

    st.download_button(
        label="📥 Descargar visor_export.xlsx",
        data=output.getvalue(),
        file_name="visor_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
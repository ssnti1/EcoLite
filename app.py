# -*- coding: utf-8 -*-
import os
import io
import uuid
from datetime import datetime
from typing import Optional, Tuple


import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from flask import (
    Flask, render_template, request, redirect, url_for,
    send_from_directory, flash, session
)





# ---- Flask setup ----
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "devkey-please-change")

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
REPORTS_DIR = os.path.join(STATIC_DIR, "reportes")
os.makedirs(REPORTS_DIR, exist_ok=True)

ALLOWED_EXT = {"xls", "xlsx"}

# Optional logo (put a file called logoecolite.jpg next to app.py)
LOGO_PATH = os.path.join(BASE_DIR, "logoecolite.jpg")
HAS_LOGO = os.path.exists(LOGO_PATH)

AVAILABLE = {
    "reporte_vendedores": "Ventas por Vendedor",
    "reporte_margen_vendedores": "Margen por Vendedor",
    "reporte_margen_productos": "Margen por Producto",
    "reporte_producto_volumen_margen": "Producto por Volumen y Margen",
    "reporte_top_ciudades": "Top Ciudades",
    "reporte_top_departamentos": "Top Departamentos",
    "reporte_comparativo_vendedor": "Comparativo Vendedor",
    "reporte_comparativo_ciudad": "Comparativo Ciudad",
    "reporte_comparativo_departamento": "Comparativo Departamento",
    "reporte_comparativo_linea": "Comparativo Línea",
    "reporte_ventas_semana": "Ventas por Semana",
    "reporte_presupuesto_anio": "Presupuesto anio",
    "reporte_rotacion_inventario": "Rotación de Inventario"
}

REPORT_FIELDS = {
    "reporte_vendedores": ["anio"],
    "reporte_margen_vendedores": ["anio"],
    "reporte_margen_productos": ["anio", "mes_inicio", "dia_inicio", "mes_fin", "dia_fin", "top_n"],
    "reporte_producto_volumen_margen": ["anio", "mes_inicio", "dia_inicio", "mes_fin", "dia_fin", "top_n"],
    "reporte_top_ciudades": ["anio", "top_n"],
    "reporte_top_departamentos": ["anio", "top_n"],
    "reporte_comparativo_vendedor": [],
    "reporte_comparativo_ciudad": [],
    "reporte_comparativo_departamento": [],
    "reporte_comparativo_linea": [],
    "reporte_ventas_semana": [],
    "reporte_presupuesto_anio": [],
    "reporte_rotacion_inventario": []
}



def ensure_excel_loaded() -> Optional[str]:
    """Return absolute path to uploaded Excel or None if not set."""
    excel_path = session.get("excel_path")
    if excel_path and os.path.exists(excel_path):
        return excel_path
    return None


def save_dataframe_xlsx(df: pd.DataFrame, name: str) -> str:
    folder = "static/reportes"
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, f"{name}.xlsx")
    df.to_excel(path, index=False)
    return f"reportes/{name}.xlsx"  # ✅ solo lo relativo

def save_figure(fig, name: str) -> str:
    folder = "static/reportes"
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, f"{name}.png")
    fig.savefig(path, bbox_inches="tight")
    plt.close(fig)
    return f"reportes/{name}.png"


def read_facturacion(excel_path: str) -> pd.DataFrame:
    df = pd.read_excel(excel_path, sheet_name="Facturacion")
    df.columns = df.columns.str.strip().str.upper().str.replace('\xa0', '', regex=True)
    if "FECHA" in df:
        df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
    for c in ["NETO", "COST", "QTYSHIP"]:
        if c in df:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def maybe_add_logo(ax):
    if not HAS_LOGO:
        return
    try:
        import matplotlib.offsetbox as ofb
        img = plt.imread(LOGO_PATH)
        oi = ofb.OffsetImage(img, zoom=0.5)
        ab = ofb.AnnotationBbox(oi, (1.02, 1.02), frameon=False, xycoords='axes fraction', box_alignment=(1, 1))
        ax.add_artist(ab)
    except Exception:
        pass



# ---------------- Report logic (ported from desktop) ----------------

def reporte_vendedores(excel_path: str, anio: int) -> Tuple[str, Optional[str]]:
    excluir = ["ARANGO JULIO CESAR", "LOPEZ GAITAN JORGE HERNAN", "Sin vendedor"]
    df = read_facturacion(excel_path)
    df = df[(df["FECHA"].dt.year == anio) & (~df["VENDEDOR"].isin(excluir))]
    if df.empty:
        raise ValueError(f"No hay datos para el anio {anio}.")

    reporte = df.groupby('VENDEDOR', as_index=False)['NETO'].sum().sort_values('NETO', ascending=False)
    total = reporte['NETO'].sum()

    plt.style.use("seaborn-v0_8-whitegrid")
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.barh(reporte['VENDEDOR'], reporte['NETO'])
    ax.invert_yaxis()
    ax.set_title(f"Ventas por Vendedor - {anio} | Total: ${total:,.0f}", fontsize=14, fontweight="bold")
    ax.set_xlabel("Valor vendido ($)")
    for i, v in enumerate(reporte["NETO"]):
        ax.text(v + (reporte["NETO"].max() * 0.01), i, f"${v:,.0f}", va="center", fontsize=9)
    maybe_add_logo(ax)
    img_rel = save_figure(fig, f"vendedores_{anio}")
    xlsx_rel = save_dataframe_xlsx(reporte, f"vendedores_{anio}")
    return img_rel, xlsx_rel


def reporte_margen_vendedores(excel_path: str, anio: int) -> Tuple[str, Optional[str]]:
    excluir = ["ARANGO JULIO CESAR", "LOPEZ GAITAN JORGE HERNAN", "Sin vendedor"]
    df = read_facturacion(excel_path)
    df = df[df["FECHA"].dt.year == anio]
    df = df[~df["VENDEDOR"].isin(excluir)].copy()
    if df.empty:
        raise ValueError(f"No hay datos para el anio {anio}.")

    df["COSTO_TOTAL"] = df["COST"] * df["QTYSHIP"]
    reporte = (
        df.groupby('VENDEDOR', as_index=False)
          .agg(NETO=('NETO','sum'), COSTO_TOTAL=('COSTO_TOTAL','sum'))
    )
    reporte['MARGEN_%'] = ((reporte['NETO'] - reporte['COSTO_TOTAL']) / reporte['NETO']) * 100
    reporte = reporte.sort_values('MARGEN_%', ascending=False)

    plt.style.use("seaborn-v0_8-whitegrid")
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.barh(reporte['VENDEDOR'], reporte['MARGEN_%'])
    ax.invert_yaxis()
    ax.set_title(f"Margen por Vendedor (%) - {anio}", fontsize=14, fontweight="bold")
    ax.set_xlabel("Margen (%)")
    for i, v in enumerate(reporte["MARGEN_%"]):
        ax.text(v + 0.5, i, f"{v:.1f}%", va="center", fontsize=9)
    maybe_add_logo(ax)
    img_rel = save_figure(fig, f"margen_vendedores_{anio}")
    xlsx_rel = save_dataframe_xlsx(reporte, f"margen_vendedores_{anio}")
    return img_rel, xlsx_rel


def _filtrar_fecha(df: pd.DataFrame, anio:int, mes_i:int, dia_i:int, mes_f:int, dia_f:int) -> pd.DataFrame:
    fi = datetime(anio, mes_i, dia_i)
    ff = datetime(anio, mes_f, dia_f)
    return df[(df["FECHA"] >= fi) & (df["FECHA"] <= ff)].copy()


def reporte_margen_productos(excel_path: str, anio:int, mes_i:int, dia_i:int, mes_f:int, dia_f:int, top_n:int):
    df = read_facturacion(excel_path)
    df = _filtrar_fecha(df, anio, mes_i, dia_i, mes_f, dia_f)
    df["ITEM"] = df["ITEM"].astype(str).str.strip().str.upper()
    patron_excluir = "FLETE|MANTENIMIENTO|DF-DISEÑO|DISEÑO"
    df = df[~df["ITEM"].str.contains(patron_excluir, case=False, na=False, regex=True)]
    if df.empty:
        raise ValueError("No hay datos en el rango seleccionado.")

    df["COSTO_TOTAL"] = df["COST"] * df["QTYSHIP"]
    df["MARGEN_%"] = ((df["NETO"] - df["COSTO_TOTAL"]) / df["NETO"]) * 100
    top_df = (df.groupby("ITEM")["MARGEN_%"].mean().reset_index()
                .sort_values("MARGEN_%", ascending=False).head(top_n))

    plt.style.use("seaborn-v0_8-whitegrid")
    fig, ax = plt.subplots(figsize=(10, 6))
    bars = ax.barh(top_df["ITEM"], top_df["MARGEN_%"])
    ax.invert_yaxis()
    ax.set_title(f"Top {top_n} productos por rentabilidad (%)", fontsize=14, fontweight="bold")
    ax.set_xlabel("Margen (%)")
    for b in bars:
        w = b.get_width()
        ax.text(w + 0.5, b.get_y()+b.get_height()/2, f"{w:.2f}%", va="center", fontsize=11)
    maybe_add_logo(ax)
    img_rel = save_figure(fig, "margen_productos")
    xlsx_rel = save_dataframe_xlsx(top_df, "margen_productos")
    return img_rel, xlsx_rel


def reporte_producto_volumen_margen(excel_path:str, anio:int, mes_i:int, dia_i:int, mes_f:int, dia_f:int, top_n:int):
    df = read_facturacion(excel_path)
    df = _filtrar_fecha(df, anio, mes_i, dia_i, mes_f, dia_f)
    df["ITEM"] = df["ITEM"].astype(str).str.strip().str.upper()
    patron_excluir = "FLETE|MANTENIMIENTO|DF-DISEÑO|DISEÑO"
    df = df[~df["ITEM"].str.contains(patron_excluir, case=False, na=False, regex=True)]
    df["COSTO_TOTAL"] = df["COST"] * df["QTYSHIP"]
    df["MARGEN_%"] = ((df["NETO"] - df["COSTO_TOTAL"]) / df["NETO"]) * 100

    top_df = (
        df
        .groupby("ITEM", as_index=False)
        .agg({
            "NETO": "sum",
            "MARGEN_%": "mean"
        })
        .sort_values("NETO", ascending=False)
        .head(top_n)
    )


    margen_escalado = (top_df["MARGEN_%"] / 100.0) * top_df["NETO"]
    margen_escalado = np.minimum(margen_escalado, top_df["NETO"] * 0.9)

    plt.style.use("seaborn-v0_8-whitegrid")
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.barh(top_df["ITEM"], top_df["NETO"], label="Valor vendido ($)")
    ax.barh(top_df["ITEM"], margen_escalado, height=0.4, label="Margen (%)")
    ax.invert_yaxis()
    ax.set_title(f"Top {top_n} productos por Valor y Margen", fontsize=14, fontweight="bold")
    ax.set_xlabel("Valor vendido ($)")
    ax.legend()
    for i, (valor, margen, margen_px) in enumerate(zip(top_df["NETO"], top_df["MARGEN_%"], margen_escalado)):
        ax.text(valor + (top_df["NETO"].max()*0.01), i, f"${valor:,.0f}", va="center", fontsize=9)
        ax.text(margen_px + (top_df["NETO"].max()*0.01), i, f"{margen:.1f}%", va="center", fontsize=9)
    maybe_add_logo(ax)
    img_rel = save_figure(fig, "producto_valor_margen")
    xlsx_rel = save_dataframe_xlsx(top_df, "producto_valor_margen")
    return img_rel, xlsx_rel


def _top_group(excel_path:str, anio:int, group_col:str, top_n:int, color:str, titulo_prefix:str):
    df = read_facturacion(excel_path)
    df = df[df["FECHA"].dt.year == anio].copy()
    if df.empty:
        raise ValueError(f"No hay datos para el anio {anio}.")
    rep = (df.groupby(group_col)["NETO"].sum().sort_values(ascending=False).head(top_n))
    total = rep.sum()

    plt.style.use("seaborn-v0_8-whitegrid")
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.barh(rep.index, rep.values)
    ax.invert_yaxis()
    ax.set_title(f"{titulo_prefix} {anio} (${total:,.0f})", fontsize=14, fontweight="bold")
    ax.set_xlabel("Valor facturado ($)")
    for i, v in enumerate(rep.values):
        ax.text(v + (rep.values.max()*0.01), i, f"${v:,.0f}", va="center", fontsize=9)
    maybe_add_logo(ax)
    img_rel = save_figure(fig, f"top_{group_col.lower()}_{anio}")
    xlsx_rel = save_dataframe_xlsx(rep.reset_index(name="TOTAL_FACTURACION"), f"top_{group_col.lower()}_{anio}")
    return img_rel, xlsx_rel


def reporte_top_ciudades(excel_path, anio:int, top_n:int):
    return _top_group(excel_path, anio, "CITY", top_n, "skyblue", "Top Ciudades - Facturación")


def reporte_top_departamentos(excel_path, anio:int, top_n:int):
    return _top_group(excel_path, anio, "DEPARTAMENTO", top_n, "lightgreen", "Top Departamentos - Facturación")


def _comparativo_generico(excel_path: str, eje: str) -> Tuple[Optional[str], Optional[str]]:
    df = read_facturacion(excel_path)
    hoy = datetime.today()
    fi_a = datetime(hoy.year, 1, 1)
    ff_a = hoy
    fi_b = fi_a.replace(year=fi_a.year - 1)
    ff_b = ff_a.replace(year=ff_a.year - 1)

    df_a = df[(df["FECHA"] >= fi_a) & (df["FECHA"] <= ff_a)]
    df_b = df[(df["FECHA"] >= fi_b) & (df["FECHA"] <= ff_b)]

    va = df_a.groupby(eje)["NETO"].sum().reset_index()
    vb = df_b.groupby(eje)["NETO"].sum().reset_index()

    comp = pd.merge(vb, va, on=eje, how="outer", suffixes=(f"_{fi_b.year}", f"_{fi_a.year}")).fillna(0)
    comp["DIFERENCIA"] = comp[f"NETO_{fi_a.year}"] - comp[f"NETO_{fi_b.year}"]
    comp = comp.sort_values(by=f"NETO_{fi_a.year}", ascending=True)

    plt.style.use("seaborn-v0_8-whitegrid")
    fig, ax = plt.subplots(figsize=(12, 7))
    y = range(len(comp))
    h = 0.4
    ax.barh([yy + h/2 for yy in y], comp[f"NETO_{fi_b.year}"], height=h, label=str(fi_b.year))
    ax.barh([yy - h/2 for yy in y], comp[f"NETO_{fi_a.year}"], height=h, label=str(fi_a.year))
    for i, (vbv, vav) in enumerate(zip(comp[f"NETO_{fi_b.year}"], comp[f"NETO_{fi_a.year}"])):
        ax.text(vbv + (vbv*0.01 if vbv else 5), i + h/2, f"${vbv:,.0f}", va="center", fontsize=8)
        ax.text(vav + (vav*0.01 if vav else 5), i - h/2, f"${vav:,.0f}", va="center", fontsize=8)
    ax.set_yticks(list(y))
    ax.set_yticklabels(comp[eje])
    ax.set_xlabel("Ventas ($)")
    ax.set_title(f"Comparativo YTD de Ventas por {eje}\n{fi_b.date()}–{ff_b.date()} vs {fi_a.date()}–{ff_a.date()}",
                 fontsize=14, fontweight="bold")
    ax.legend()
    maybe_add_logo(ax)

    img_rel = save_figure(fig, f"comparativo_{eje.lower()}")
    xlsx_rel = save_dataframe_xlsx(comp, f"comparativo_{eje.lower()}")
    return img_rel, xlsx_rel

def reporte_comparativo_vendedor(excel_path: str, params: dict) -> Tuple[Optional[str], Optional[str]]:
    eje = "VENDEDOR"
    df = read_facturacion(excel_path)
    hoy = datetime.today()
    fi_a = datetime(hoy.year, 1, 1)
    ff_a = hoy
    fi_b = fi_a.replace(year=fi_a.year - 1)
    ff_b = ff_a.replace(year=ff_a.year - 1)

    df_a = df[(df["FECHA"] >= fi_a) & (df["FECHA"] <= ff_a)]
    df_b = df[(df["FECHA"] >= fi_b) & (df["FECHA"] <= ff_b)]

    va = df_a.groupby(eje)["NETO"].sum().reset_index()
    vb = df_b.groupby(eje)["NETO"].sum().reset_index()

    comp = pd.merge(vb, va, on=eje, how="outer", suffixes=(f"_{fi_b.year}", f"_{fi_a.year}")).fillna(0)
    comp["DIFERENCIA"] = comp[f"NETO_{fi_a.year}"] - comp[f"NETO_{fi_b.year}"]
    comp = comp.sort_values(by=f"NETO_{fi_a.year}", ascending=True)

    plt.style.use("seaborn-v0_8-whitegrid")
    fig, ax = plt.subplots(figsize=(12, 7))
    y = range(len(comp))
    h = 0.4
    ax.barh([yy + h/2 for yy in y], comp[f"NETO_{fi_b.year}"], height=h, label=str(fi_b.year))
    ax.barh([yy - h/2 for yy in y], comp[f"NETO_{fi_a.year}"], height=h, label=str(fi_a.year))
    for i, (vbv, vav) in enumerate(zip(comp[f"NETO_{fi_b.year}"], comp[f"NETO_{fi_a.year}"])):
        ax.text(vbv + (vbv*0.01 if vbv else 5), i + h/2, f"${vbv:,.0f}", va="center", fontsize=8)
        ax.text(vav + (vav*0.01 if vav else 5), i - h/2, f"${vav:,.0f}", va="center", fontsize=8)
    ax.set_yticks(list(y))
    ax.set_yticklabels(comp[eje])
    ax.set_xlabel("Ventas ($)")
    ax.set_title(f"Comparativo YTD de Ventas por {eje}\n{fi_b.date()}–{ff_b.date()} vs {fi_a.date()}–{ff_a.date()}",
                 fontsize=14, fontweight="bold")
    ax.legend()
    maybe_add_logo(ax)

    img_rel = save_figure(fig, "comparativo_vendedor")
    xlsx_rel = save_dataframe_xlsx(comp, "comparativo_vendedor")
    return img_rel, xlsx_rel

def reporte_comparativo_ciudad(excel_path: str, params: dict) -> Tuple[Optional[str], Optional[str]]:
    eje = "CITY"
    df = read_facturacion(excel_path)
    hoy = datetime.today()
    fi_a = datetime(hoy.year, 1, 1)
    ff_a = hoy
    fi_b = fi_a.replace(year=fi_a.year - 1)
    ff_b = ff_a.replace(year=ff_a.year - 1)

    df_a = df[(df["FECHA"] >= fi_a) & (df["FECHA"] <= ff_a)]
    df_b = df[(df["FECHA"] >= fi_b) & (df["FECHA"] <= ff_b)]

    va = df_a.groupby(eje)["NETO"].sum().reset_index()
    vb = df_b.groupby(eje)["NETO"].sum().reset_index()

    comp = pd.merge(vb, va, on=eje, how="outer", suffixes=(f"_{fi_b.year}", f"_{fi_a.year}")).fillna(0)
    comp["DIFERENCIA"] = comp[f"NETO_{fi_a.year}"] - comp[f"NETO_{fi_b.year}"]
    comp = comp.sort_values(by=f"NETO_{fi_a.year}", ascending=True)

    plt.style.use("seaborn-v0_8-whitegrid")
    fig, ax = plt.subplots(figsize=(12, 7))
    y = range(len(comp))
    h = 0.4
    ax.barh([yy + h/2 for yy in y], comp[f"NETO_{fi_b.year}"], height=h, label=str(fi_b.year))
    ax.barh([yy - h/2 for yy in y], comp[f"NETO_{fi_a.year}"], height=h, label=str(fi_a.year))
    for i, (vbv, vav) in enumerate(zip(comp[f"NETO_{fi_b.year}"], comp[f"NETO_{fi_a.year}"])):
        ax.text(vbv + (vbv*0.01 if vbv else 5), i + h/2, f"${vbv:,.0f}", va="center", fontsize=8)
        ax.text(vav + (vav*0.01 if vav else 5), i - h/2, f"${vav:,.0f}", va="center", fontsize=8)
    ax.set_yticks(list(y))
    ax.set_yticklabels(comp[eje])
    ax.set_xlabel("Ventas ($)")
    ax.set_title(f"Comparativo YTD de Ventas por {eje}\n{fi_b.date()}–{ff_b.date()} vs {fi_a.date()}–{ff_a.date()}",
                 fontsize=14, fontweight="bold")
    ax.legend()
    maybe_add_logo(ax)

    img_rel = save_figure(fig, "comparativo_ciudad")
    xlsx_rel = save_dataframe_xlsx(comp, "comparativo_ciudad")
    return img_rel, xlsx_rel

def reporte_comparativo_departamento(excel_path: str, params: dict) -> Tuple[Optional[str], Optional[str]]:
    eje = "DEPARTAMENTO"
    df = read_facturacion(excel_path)
    hoy = datetime.today()
    fi_a = datetime(hoy.year, 1, 1)
    ff_a = hoy
    fi_b = fi_a.replace(year=fi_a.year - 1)
    ff_b = ff_a.replace(year=ff_a.year - 1)

    df_a = df[(df["FECHA"] >= fi_a) & (df["FECHA"] <= ff_a)]
    df_b = df[(df["FECHA"] >= fi_b) & (df["FECHA"] <= ff_b)]

    va = df_a.groupby(eje)["NETO"].sum().reset_index()
    vb = df_b.groupby(eje)["NETO"].sum().reset_index()

    comp = pd.merge(vb, va, on=eje, how="outer", suffixes=(f"_{fi_b.year}", f"_{fi_a.year}")).fillna(0)
    comp["DIFERENCIA"] = comp[f"NETO_{fi_a.year}"] - comp[f"NETO_{fi_b.year}"]
    comp = comp.sort_values(by=f"NETO_{fi_a.year}", ascending=True)

    plt.style.use("seaborn-v0_8-whitegrid")
    fig, ax = plt.subplots(figsize=(12, 7))
    y = range(len(comp))
    h = 0.4
    ax.barh([yy + h/2 for yy in y], comp[f"NETO_{fi_b.year}"], height=h, label=str(fi_b.year))
    ax.barh([yy - h/2 for yy in y], comp[f"NETO_{fi_a.year}"], height=h, label=str(fi_a.year))
    for i, (vbv, vav) in enumerate(zip(comp[f"NETO_{fi_b.year}"], comp[f"NETO_{fi_a.year}"])):
        ax.text(vbv + (vbv*0.01 if vbv else 5), i + h/2, f"${vbv:,.0f}", va="center", fontsize=8)
        ax.text(vav + (vav*0.01 if vav else 5), i - h/2, f"${vav:,.0f}", va="center", fontsize=8)
    ax.set_yticks(list(y))
    ax.set_yticklabels(comp[eje])
    ax.set_xlabel("Ventas ($)")
    ax.set_title(f"Comparativo YTD de Ventas por {eje}\n{fi_b.date()}–{ff_b.date()} vs {fi_a.date()}–{ff_a.date()}",
                 fontsize=14, fontweight="bold")
    ax.legend()
    maybe_add_logo(ax)

    img_rel = save_figure(fig, "comparativo_departamento")
    xlsx_rel = save_dataframe_xlsx(comp, "comparativo_departamento")
    return img_rel, xlsx_rel

def reporte_comparativo_linea(excel_path: str, params: dict) -> Tuple[Optional[str], Optional[str]]:
    eje = "DESCLINEA"
    df = read_facturacion(excel_path)
    hoy = datetime.today()
    fi_a = datetime(hoy.year, 1, 1)
    ff_a = hoy
    fi_b = fi_a.replace(year=fi_a.year - 1)
    ff_b = ff_a.replace(year=ff_a.year - 1)

    df_a = df[(df["FECHA"] >= fi_a) & (df["FECHA"] <= ff_a)]
    df_b = df[(df["FECHA"] >= fi_b) & (df["FECHA"] <= ff_b)]

    va = df_a.groupby(eje)["NETO"].sum().reset_index()
    vb = df_b.groupby(eje)["NETO"].sum().reset_index()

    comp = pd.merge(vb, va, on=eje, how="outer", suffixes=(f"_{fi_b.year}", f"_{fi_a.year}")).fillna(0)
    comp["DIFERENCIA"] = comp[f"NETO_{fi_a.year}"] - comp[f"NETO_{fi_b.year}"]
    comp = comp.sort_values(by=f"NETO_{fi_a.year}", ascending=True)

    plt.style.use("seaborn-v0_8-whitegrid")
    fig, ax = plt.subplots(figsize=(12, 7))
    y = range(len(comp))
    h = 0.4
    ax.barh([yy + h/2 for yy in y], comp[f"NETO_{fi_b.year}"], height=h, label=str(fi_b.year))
    ax.barh([yy - h/2 for yy in y], comp[f"NETO_{fi_a.year}"], height=h, label=str(fi_a.year))
    for i, (vbv, vav) in enumerate(zip(comp[f"NETO_{fi_b.year}"], comp[f"NETO_{fi_a.year}"])):
        ax.text(vbv + (vbv*0.01 if vbv else 5), i + h/2, f"${vbv:,.0f}", va="center", fontsize=8)
        ax.text(vav + (vav*0.01 if vav else 5), i - h/2, f"${vav:,.0f}", va="center", fontsize=8)
    ax.set_yticks(list(y))
    ax.set_yticklabels(comp[eje])
    ax.set_xlabel("Ventas ($)")
    ax.set_title(f"Comparativo YTD de Ventas por {eje}\n{fi_b.date()}–{ff_b.date()} vs {fi_a.date()}–{ff_a.date()}",
                 fontsize=14, fontweight="bold")
    ax.legend()
    maybe_add_logo(ax)

    img_rel = save_figure(fig, "comparativo_linea")
    xlsx_rel = save_dataframe_xlsx(comp, "comparativo_linea")
    return img_rel, xlsx_rel

def reporte_ventas_semana(excel_path: str, params: dict) -> Tuple[Optional[str], Optional[str]]:
    df = read_facturacion(excel_path)
    hoy = datetime.today()
    anyo_full = hoy.year - 1
    anyo_actual = hoy.year

    df = df[df["FECHA"].dt.year.isin([anyo_full, anyo_actual])]
    meses_map = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
                 7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}
    df["MES"] = df["FECHA"].dt.month.map(meses_map)

    def semana_del_mes(fecha):
        dia_semana_inicio = fecha.replace(day=1).weekday()
        desplaz = fecha.day + dia_semana_inicio - 1
        return (desplaz // 7) + 1

    df["SEMANA_MES"] = df["FECHA"].apply(semana_del_mes)
    df["anio"] = df["FECHA"].dt.year

    meses_ord = list(meses_map.values())
    max_semana = int(df["SEMANA_MES"].max())
    semanas = list(range(1, max_semana + 1))

    bloques = []
    for anyo in [anyo_actual, anyo_full]:
        df_a = df[df["anio"] == anyo]
        if df_a.empty:
            continue
        pivot = (df_a.pivot_table(index="SEMANA_MES", columns="MES", values="NETO", aggfunc="sum", fill_value=0)
                      .reindex(index=semanas, columns=meses_ord, fill_value=0))
        pivot["Total"] = pivot.sum(axis=1)
        pivot.insert(0, "anio", anyo)
        pivot.insert(1, "Semana", pivot.index)
        bloques.append(pivot.reset_index(drop=True))
    if not bloques:
        raise ValueError("No hay datos para construir el reporte.")
    salida = pd.concat(bloques, ignore_index=True)

    xlsx_rel = save_dataframe_xlsx(salida, "ventas_semana")
    return None, xlsx_rel

def reporte_presupuesto_anio(excel_path: str, params: dict) -> Tuple[Optional[str], Optional[str]]:
    df = pd.read_excel(excel_path, sheet_name="Ppto año", header=None)
    df = df.iloc[6:18, [3,4,5,6,7,8,9,10,11,12]]
    df.columns = [
        "Mes","2024","Acum 2024",
        "Ppto Mes 2025","Ppto Acum 2025",
        "Real 2025","Acum 2025",
        "% Cumpl Mes","Acum 2025 vs Acum 2024",
        "Acum 2025 vs Ppto Acum 2025"
    ]
    for col in ["2024","Acum 2024","Ppto Mes 2025","Ppto Acum 2025","Real 2025","Acum 2025"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
    for col in ["% Cumpl Mes","Acum 2025 vs Acum 2024","Acum 2025 vs Ppto Acum 2025"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    xlsx_rel = save_dataframe_xlsx(df, "presupuesto_anio")
    return None, xlsx_rel

def reporte_rotacion_inventario(excel_path: str, params: dict) -> Tuple[Optional[str], Optional[str]]:
    df_fac = pd.read_excel(excel_path, sheet_name="Facturacion")
    df_fac.columns = df_fac.columns.str.strip().str.upper().str.replace('\xa0','', regex=True)
    columnas_fact = {"ITEM","DESCRIPCION","FECHA","QTYSHIP"}
    if columnas_fact - set(df_fac.columns):
        raise ValueError(f"Faltan columnas en Facturacion: {columnas_fact - set(df_fac.columns)}")
    df_fac = df_fac.dropna(subset=["FECHA","QTYSHIP","ITEM"]).copy()
    df_fac["FECHA"] = pd.to_datetime(df_fac["FECHA"], errors="coerce")
    df_fac["QTYSHIP"] = pd.to_numeric(df_fac["QTYSHIP"], errors="coerce").fillna(0)
    df_fac["ITEM"] = df_fac["ITEM"].astype(str).str.strip()
    df_fac["anio_MES"] = df_fac["FECHA"].dt.to_period("M")

    meses_activos = (df_fac[df_fac["QTYSHIP"]>0]
                        .groupby(["ITEM","DESCRIPCION"])["anio_MES"].nunique()
                        .reset_index().rename(columns={"anio_MES":"MESES"}))
    ventas_totales = (df_fac.groupby(["ITEM","DESCRIPCION"], as_index=False)["QTYSHIP"].sum()
                        .rename(columns={"QTYSHIP":"VENTA_TOTAL"}))
    resumen = pd.merge(ventas_totales, meses_activos, on=["ITEM","DESCRIPCION"], how="left").fillna(0.0)
    resumen["PROMEDIO_MES"] = resumen.apply(
        lambda r: r["VENTA_TOTAL"]/r["MESES"] if r["MESES"]>0 else 0.0, axis=1
    )

    df_inv = pd.read_excel(excel_path, sheet_name="Inventario SAI", header=None)
    df_inv = df_inv.iloc[6:1200, 2:5]
    df_inv.columns = ["ITEM","OTRA_COL","INVENTARIO"]
    df_inv["ITEM"] = df_inv["ITEM"].astype(str).str.strip()
    df_inv["INVENTARIO"] = pd

def run_report(reporte, excel_path, params):
    if not excel_path or not os.path.exists(excel_path):
        return []  # no devolvemos un Response

    p = {k: int(v) if v and v.isdigit() else v for k, v in params.items()}

    if reporte == "reporte_vendedores":
        return reporte_vendedores(excel_path, p["anio"])
    elif reporte == "reporte_margen_vendedores":
        return reporte_margen_vendedores(excel_path, p["anio"])
    elif reporte == "reporte_margen_productos":
        return reporte_margen_productos(excel_path, p["anio"], p["mes_inicio"], p["dia_inicio"], p["mes_fin"], p["dia_fin"], p["top_n"])
    elif reporte == "reporte_producto_volumen_margen":
        return reporte_producto_volumen_margen(excel_path, p["anio"], p["mes_inicio"], p["dia_inicio"], p["mes_fin"], p["dia_fin"], p["top_n"])
    elif reporte == "reporte_top_ciudades":
        return reporte_top_ciudades(excel_path, p["anio"], p["top_n"])
    elif reporte == "reporte_top_departamentos":
        return reporte_top_departamentos(excel_path, p["anio"], p["top_n"])
    elif reporte == "reporte_comparativo_vendedor":
        return reporte_comparativo_vendedor(excel_path, params)
    elif reporte == "reporte_comparativo_ciudad":
        return reporte_comparativo_ciudad(excel_path, params)
    elif reporte == "reporte_comparativo_departamento":
        return reporte_comparativo_departamento(excel_path, params)
    elif reporte == "reporte_comparativo_linea":
        return reporte_comparativo_linea(excel_path, params)
    elif reporte == "reporte_ventas_semana":
        return reporte_ventas_semana(excel_path, params)
    elif reporte == "reporte_presupuesto_anio":
        return reporte_presupuesto_anio(excel_path, params)
    elif reporte == "reporte_rotacion_inventario":
        return reporte_rotacion_inventario(excel_path, params)

    raise ValueError(f"Reporte desconocido: {reporte}")




LATEST_EXCEL = os.path.join(REPORTS_DIR, "archivo.xlsx")

@app.route("/", methods=["GET", "POST"])
def index():
    notice = None
    outputs = {}
    selected = None

    # 1) Handle file upload
    if request.method == "POST" and request.form.get("accion") == "upload":
        file = request.files.get("archivo")
        if file and file.filename:
            os.makedirs(REPORTS_DIR, exist_ok=True)
            file.save(LATEST_EXCEL)
            notice = f"Archivo cargado: {file.filename}"

    has_excel = os.path.exists(LATEST_EXCEL)

    # 2) Handle report execution
    if request.method == "POST" and request.form.get("accion") == "run":
        selected = request.form.get("reporte")
        if not has_excel:
            notice = "Primero sube un archivo Excel."
        else:
            params = dict(
                anio = request.form.get("anio") or "",
                mes_inicio = request.form.get("mes_inicio") or "",
                dia_inicio = request.form.get("dia_inicio") or "",
                mes_fin = request.form.get("mes_fin") or "",
                dia_fin = request.form.get("dia_fin") or "",
                top_vendedores = request.form.get("top_vendedores") or "",
                top_productos = request.form.get("top_productos") or "",
                top_departamentos = request.form.get("top_departamentos") or "",
                top_ciudades = request.form.get("top_ciudades") or "",
                top_n = request.form.get("top_n") or "",
            )
    try: 
        result = run_report(selected, LATEST_EXCEL, params)
        if isinstance(result, tuple):
            files = [f for f in result if f]
        elif isinstance(result, list):
            files = result
        else:
            files = [result] if result else []

        outputs[selected] = files

        if not files:
            notice = f"Se ejecutó '{selected}' pero no se detectaron archivos nuevos."
    except Exception as e:
        notice = f"Error al ejecutar {selected}: {e}"

    return render_template(
        "index.html",
        available=AVAILABLE.items(),
        has_excel=has_excel,
        notice=notice,
        outputs=outputs,
        report_fields=REPORT_FIELDS
    )
# -*- coding: utf-8 -*-
import os
import uuid
from datetime import datetime
from typing import Optional, Tuple

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from flask import Flask, render_template, request, redirect, url_for, session

# ---- Flask setup & constantes ----
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "devkey-please-change")

BASE_DIR    = os.path.abspath(os.path.dirname(__file__))
STATIC_DIR  = os.path.join(BASE_DIR, "static")
REPORTS_DIR = os.path.join(STATIC_DIR, "reportes")
os.makedirs(REPORTS_DIR, exist_ok=True)

# Aquí DEFINES la ruta del Excel antes de cualquier ruta
LATEST_EXCEL = os.path.join(REPORTS_DIR, "archivo.xlsx")

# ---- Decorador para desactivar cache ----
@app.after_request
def no_cache(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response

# ---- Ruta principal con PRG ----
@app.route("/", methods=["GET", "POST"])
def index():
    notice  = session.pop("notice", None)
    outputs = session.pop("outputs", {})

    if request.method == "POST":
        accion = request.form.get("accion")

        if accion == "upload":
            f = request.files.get("archivo")
            if f and f.filename:
                f.save(LATEST_EXCEL)
                session["notice"] = f"Archivo cargado: {f.filename}"
            return redirect(url_for("index"))

        if accion == "run":
            reporte = request.form.get("reporte")
            if not os.path.exists(LATEST_EXCEL):
                session["notice"] = "Primero sube un archivo Excel."
            else:
                # Recoger y convertir params
                params = {k: request.form.get(k) or "" for k in (
                    "anio","mes_inicio","dia_inicio",
                    "mes_fin","dia_fin",
                    "top_vendedores","top_productos",
                    "top_departamentos","top_ciudades","top_n"
                )}
                try:
                    archivos = run_report(reporte, LATEST_EXCEL, params)
                    if archivos:
                        outputs[reporte] = archivos
                        session["outputs"] = outputs
                    else:
                        session["notice"] = f"No se generaron archivos para '{reporte}'."
                except Exception as e:
                    session["notice"] = f"Error al ejecutar {reporte}: {e}"
            return redirect(url_for("index"))

    has_excel = os.path.exists(LATEST_EXCEL)
    return render_template(
        "index.html",
        available=AVAILABLE.items(),
        has_excel=has_excel,
        notice=notice,
        outputs=outputs,
        report_fields=REPORT_FIELDS
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=True)

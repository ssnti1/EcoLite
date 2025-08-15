# -*- coding: utf-8 -*-
import io
import os
from datetime import datetime
from typing import Optional, Tuple
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from flask import (
    Flask, render_template, request, redirect, url_for,
    send_file, session
)

# ---- Flask setup ----
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "devkey-please-change")



# --- Variables en memoria ---
uploaded_excel_bytes: Optional[bytes] = None
generated_reports = {}  # {reporte: [(filename, BytesIO), ...]}

ALLOWED_EXT = {"xls", "xlsx"}

# --- Configuración de reportes ---
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
    "reporte_presupuesto_anio": "Presupuesto Año",
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

# --- Utilidades ---
def read_facturacion(excel_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="Facturacion")
    df.columns = df.columns.str.strip().str.upper().str.replace('\xa0', '', regex=True)
    if "FECHA" in df:
        df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
    for c in ["NETO", "COST", "QTYSHIP"]:
        if c in df:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

def save_dataframe_xlsx(df: pd.DataFrame, name: str):
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return (f"{name}.xlsx", buf)

def save_figure(fig, name: str):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return (f"{name}.png", buf)

def maybe_add_logo(ax):
    # puedes eliminarlo o adaptarlo a memoria si quieres
    pass

# --- Reportes ---
def reporte_vendedores(excel_bytes: bytes, anio: int):
    excluir = ["ARANGO JULIO CESAR", "LOPEZ GAITAN JORGE HERNAN", "Sin vendedor"]
    df = read_facturacion(excel_bytes)
    df = df[(df["FECHA"].dt.year == anio) & (~df["VENDEDOR"].isin(excluir))]
    if df.empty:
        raise ValueError(f"No hay datos para el año {anio}.")
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
    img_file = save_figure(fig, f"vendedores_{anio}")
    xlsx_file = save_dataframe_xlsx(reporte, f"vendedores_{anio}")
    return [img_file, xlsx_file]

def reporte_margen_vendedores(excel_bytes: bytes, anio: int):
    excluir = ["ARANGO JULIO CESAR", "LOPEZ GAITAN JORGE HERNAN", "Sin vendedor"]
    df = read_facturacion(excel_bytes)
    df = df[df["FECHA"].dt.year == anio]
    df = df[~df["VENDEDOR"].isin(excluir)].copy()
    if df.empty:
        raise ValueError(f"No hay datos para el año {anio}.")
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
    img_file = save_figure(fig, f"margen_vendedores_{anio}")
    xlsx_file = save_dataframe_xlsx(reporte, f"margen_vendedores_{anio}")
    return [img_file, xlsx_file]

# aquí seguirán todos los demás reportes adaptados...
# ---- FIN PARTE 1 ----
def _filtrar_fecha(df: pd.DataFrame, anio:int, mes_i:int, dia_i:int, mes_f:int, dia_f:int) -> pd.DataFrame:
    fi = datetime(anio, mes_i, dia_i)
    ff = datetime(anio, mes_f, dia_f)
    return df[(df["FECHA"] >= fi) & (df["FECHA"] <= ff)].copy()

def reporte_margen_productos(excel_bytes: bytes, anio:int, mes_i:int, dia_i:int, mes_f:int, dia_f:int, top_n:int):
    df = read_facturacion(excel_bytes)
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

    fig, ax = plt.subplots(figsize=(10, 6))
    bars = ax.barh(top_df["ITEM"], top_df["MARGEN_%"])
    ax.invert_yaxis()
    ax.set_title(f"Top {top_n} productos por rentabilidad (%)", fontsize=14, fontweight="bold")
    ax.set_xlabel("Margen (%)")
    for b in bars:
        w = b.get_width()
        ax.text(w + 0.5, b.get_y()+b.get_height()/2, f"{w:.2f}%", va="center", fontsize=11)
    img_file = save_figure(fig, "margen_productos")
    xlsx_file = save_dataframe_xlsx(top_df, "margen_productos")
    return [img_file, xlsx_file]

def reporte_producto_volumen_margen(excel_bytes: bytes, anio:int, mes_i:int, dia_i:int, mes_f:int, dia_f:int, top_n:int):
    df = read_facturacion(excel_bytes)
    df = _filtrar_fecha(df, anio, mes_i, dia_i, mes_f, dia_f)
    df["ITEM"] = df["ITEM"].astype(str).str.strip().str.upper()
    patron_excluir = "FLETE|MANTENIMIENTO|DF-DISEÑO|DISEÑO"
    df = df[~df["ITEM"].str.contains(patron_excluir, case=False, na=False, regex=True)]
    df["COSTO_TOTAL"] = df["COST"] * df["QTYSHIP"]
    df["MARGEN_%"] = ((df["NETO"] - df["COSTO_TOTAL"]) / df["NETO"]) * 100
    top_df = (
        df
        .groupby("ITEM", as_index=False)
        .agg({"NETO": "sum", "MARGEN_%": "mean"})
        .sort_values("NETO", ascending=False)
        .head(top_n)
    )

    margen_escalado = (top_df["MARGEN_%"] / 100.0) * top_df["NETO"]
    margen_escalado = np.minimum(margen_escalado, top_df["NETO"] * 0.9)

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
    img_file = save_figure(fig, "producto_valor_margen")
    xlsx_file = save_dataframe_xlsx(top_df, "producto_valor_margen")
    return [img_file, xlsx_file]

def _top_group(excel_bytes: bytes, anio:int, group_col:str, top_n:int, titulo_prefix:str):
    df = read_facturacion(excel_bytes)
    df = df[df["FECHA"].dt.year == anio].copy()
    if df.empty:
        raise ValueError(f"No hay datos para el año {anio}.")
    rep = (df.groupby(group_col)["NETO"].sum().sort_values(ascending=False).head(top_n))
    total = rep.sum()

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.barh(rep.index, rep.values)
    ax.invert_yaxis()
    ax.set_title(f"{titulo_prefix} {anio} (${total:,.0f})", fontsize=14, fontweight="bold")
    ax.set_xlabel("Valor facturado ($)")
    for i, v in enumerate(rep.values):
        ax.text(v + (rep.values.max()*0.01), i, f"${v:,.0f}", va="center", fontsize=9)
    img_file = save_figure(fig, f"top_{group_col.lower()}_{anio}")
    xlsx_file = save_dataframe_xlsx(rep.reset_index(name="TOTAL_FACTURACION"), f"top_{group_col.lower()}_{anio}")
    return [img_file, xlsx_file]

def reporte_top_ciudades(excel_bytes, anio:int, top_n:int):
    return _top_group(excel_bytes, anio, "CITY", top_n, "Top Ciudades - Facturación")

def reporte_top_departamentos(excel_bytes, anio:int, top_n:int):
    return _top_group(excel_bytes, anio, "DEPARTAMENTO", top_n, "Top Departamentos - Facturación")

# ---- FIN PARTE 2 ----
def _comparativo_generico(excel_bytes: bytes, eje: str):
    df = read_facturacion(excel_bytes)
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
    ax.set_title(f"Comparativo YTD de Ventas por {eje}", fontsize=14, fontweight="bold")
    ax.legend()

    img_file = save_figure(fig, f"comparativo_{eje.lower()}")
    xlsx_file = save_dataframe_xlsx(comp, f"comparativo_{eje.lower()}")
    return [img_file, xlsx_file]

def reporte_comparativo_vendedor(excel_bytes, params):
    return _comparativo_generico(excel_bytes, "VENDEDOR")

def reporte_comparativo_ciudad(excel_bytes, params):
    return _comparativo_generico(excel_bytes, "CITY")

def reporte_comparativo_departamento(excel_bytes, params):
    return _comparativo_generico(excel_bytes, "DEPARTAMENTO")

def reporte_comparativo_linea(excel_bytes, params):
    return _comparativo_generico(excel_bytes, "DESCLINEA")

def reporte_ventas_semana(excel_bytes, params):
    df = read_facturacion(excel_bytes)
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
    xlsx_file = save_dataframe_xlsx(salida, "ventas_semana")
    return [xlsx_file]

def reporte_presupuesto_anio(excel_bytes, params):
    df = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="Ppto anio", header=None)
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
    xlsx_file = save_dataframe_xlsx(df, "presupuesto_anio")
    return [xlsx_file]

# ---- Ejecutar reportes ----
def run_report(reporte, excel_bytes, params):
    if reporte == "reporte_vendedores":
        return reporte_vendedores(excel_bytes, int(params["anio"]))
    elif reporte == "reporte_margen_vendedores":
        return reporte_margen_vendedores(excel_bytes, int(params["anio"]))
    elif reporte == "reporte_margen_productos":
        return reporte_margen_productos(excel_bytes, int(params["anio"]), int(params["mes_inicio"]),
                                        int(params["dia_inicio"]), int(params["mes_fin"]), int(params["dia_fin"]), int(params["top_n"]))
    elif reporte == "reporte_producto_volumen_margen":
        return reporte_producto_volumen_margen(excel_bytes, int(params["anio"]), int(params["mes_inicio"]),
                                               int(params["dia_inicio"]), int(params["mes_fin"]), int(params["dia_fin"]), int(params["top_n"]))
    elif reporte == "reporte_top_ciudades":
        return reporte_top_ciudades(excel_bytes, int(params["anio"]), int(params["top_n"]))
    elif reporte == "reporte_top_departamentos":
        return reporte_top_departamentos(excel_bytes, int(params["anio"]), int(params["top_n"]))
    elif reporte == "reporte_comparativo_vendedor":
        return reporte_comparativo_vendedor(excel_bytes, params)
    elif reporte == "reporte_comparativo_ciudad":
        return reporte_comparativo_ciudad(excel_bytes, params)
    elif reporte == "reporte_comparativo_departamento":
        return reporte_comparativo_departamento(excel_bytes, params)
    elif reporte == "reporte_comparativo_linea":
        return reporte_comparativo_linea(excel_bytes, params)
    elif reporte == "reporte_ventas_semana":
        return reporte_ventas_semana(excel_bytes, params)
    elif reporte == "reporte_presupuesto_anio":
        return reporte_presupuesto_anio(excel_bytes, params)
    else:
        raise ValueError(f"Reporte desconocido: {reporte}")


@app.route("/", methods=["GET", "POST"])
def index():
    global uploaded_excel_bytes, generated_reports
    notice = session.pop('notice', None) 
    
    if request.method == "POST":
        accion = request.form.get("accion")

        if accion == "upload":
            f = request.files.get("archivo")
            if f and f.filename and f.filename.split(".")[-1].lower() in ALLOWED_EXT:
                uploaded_excel_bytes = f.read()
                generated_reports.clear()
                session['notice'] = f"Archivo cargado: {f.filename}"
            else:
                session['notice'] = "Formato no permitido."
            return redirect(url_for("index"))

        elif accion == "run":
            if uploaded_excel_bytes is None:
                session['notice'] = "Primero sube un archivo Excel."
            else:
                reporte = request.form.get("reporte")
                params = {k: request.form.get(k) or "" for k in REPORT_FIELDS.get(reporte, [])}
                try:
                    archivos = run_report(reporte, uploaded_excel_bytes, params)
                    generated_reports[reporte] = archivos
                    session['notice'] = f"Reporte '{AVAILABLE[reporte]}' generado."
                except Exception as e:
                    session['notice'] = f"Error: {e}"
            return redirect(url_for("index"))

    # Solo llega aquí si es GET
    return render_template(
        "index.html",
        available=AVAILABLE.items(),
        has_excel=uploaded_excel_bytes is not None,
        notice=notice,
        outputs=generated_reports,
        report_fields=REPORT_FIELDS
    )



@app.route("/download/<reporte>/<filename>")
def download_file(reporte, filename):
    for fname, buf in generated_reports.get(reporte, []):
        if fname == filename:
            return send_file(buf, as_attachment=True, download_name=fname)
    return "Archivo no encontrado", 404

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=True)
# ---- FIN PARTE 3 ----

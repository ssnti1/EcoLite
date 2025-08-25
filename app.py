import io
import os
from datetime import datetime
import re
import tempfile
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

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "devkey-please-change")

uploaded_excel_bytes: Optional[bytes] = None
generated_reports = {}

ALLOWED_EXT = {"xls", "xlsx"}

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
    "reporte_presupuesto_año": "Presupuesto Año",
    "reporte_rotacion_inventario": "Rotación de Inventario",
    "reporte_analisis_lista_inv_sin_venta": "Lista + Inventario + Venta",
    "reporte_pendientes_lista_precios": "Pendientes Lista de Precios"
}
REPORT_FIELDS = {
    "reporte_vendedores": ["año"],
    "reporte_margen_vendedores": ["año"],
    "reporte_margen_productos": ["año", "mes_inicio", "dia_inicio", "mes_fin", "dia_fin", "top_n"],
    "reporte_producto_volumen_margen": ["año", "mes_inicio", "dia_inicio", "mes_fin", "dia_fin", "top_n"],
    "reporte_top_ciudades": ["año", "top_n"],
    "reporte_top_departamentos": ["año", "top_n"],
    "reporte_comparativo_vendedor": [],
    "reporte_comparativo_ciudad": [],
    "reporte_comparativo_departamento": [],
    "reporte_comparativo_linea": [],
    "reporte_ventas_semana": [],
    "reporte_presupuesto_año": [],
    "reporte_rotacion_inventario": [],
    "reporte_analisis_lista_inv_sin_venta": ["año", "lista"],
    "reporte_pendientes_lista_precios": ["lista"]
}

def _normalizar_cols(dff: pd.DataFrame) -> pd.DataFrame:
    dff = dff.copy()
    dff.columns = dff.columns.astype(str).str.strip().str.upper().str.replace('\xa0', '', regex=True)
    return dff


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
    pass

def reporte_vendedores(excel_bytes: bytes, año: int):
    excluir = ["ARANGO JULIO CESAR", "LOPEZ GAITAN JORGE HERNAN", "Sin vendedor"]
    df = read_facturacion(excel_bytes)
    df = df[(df["FECHA"].dt.year == año) & (~df["VENDEDOR"].isin(excluir))]
    if df.empty:
        raise ValueError(f"No hay datos para el año {año}.")
    reporte = df.groupby('VENDEDOR', as_index=False)['NETO'].sum().sort_values('NETO', ascending=False)
    total = reporte['NETO'].sum()

    plt.style.use("seaborn-v0_8-whitegrid")
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.barh(reporte['VENDEDOR'], reporte['NETO'])
    ax.invert_yaxis()
    ax.set_title(f"Ventas por Vendedor - {año} | Total: ${total:,.0f}", fontsize=14, fontweight="bold")
    ax.set_xlabel("Valor vendido ($)")
    for i, v in enumerate(reporte["NETO"]):
        ax.text(v + (reporte["NETO"].max() * 0.01), i, f"${v:,.0f}", va="center", fontsize=9)
    img_file = save_figure(fig, f"vendedores_{año}")
    xlsx_file = save_dataframe_xlsx(reporte, f"vendedores_{año}")
    return [img_file, xlsx_file]

def reporte_margen_vendedores(excel_bytes: bytes, año: int):
    excluir = ["ARANGO JULIO CESAR", "LOPEZ GAITAN JORGE HERNAN", "Sin vendedor"]
    df = read_facturacion(excel_bytes)
    df = df[df["FECHA"].dt.year == año]
    df = df[~df["VENDEDOR"].isin(excluir)].copy()
    if df.empty:
        raise ValueError(f"No hay datos para el año {año}.")
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
    ax.set_title(f"Margen por Vendedor (%) - {año}", fontsize=14, fontweight="bold")
    ax.set_xlabel("Margen (%)")
    for i, v in enumerate(reporte["MARGEN_%"]):
        ax.text(v + 0.5, i, f"{v:.1f}%", va="center", fontsize=9)
    img_file = save_figure(fig, f"margen_vendedores_{año}")
    xlsx_file = save_dataframe_xlsx(reporte, f"margen_vendedores_{año}")
    return [img_file, xlsx_file]

def _filtrar_fecha(df: pd.DataFrame, año:int, mes_i:int, dia_i:int, mes_f:int, dia_f:int) -> pd.DataFrame:
    fi = datetime(año, mes_i, dia_i)
    ff = datetime(año, mes_f, dia_f)
    return df[(df["FECHA"] >= fi) & (df["FECHA"] <= ff)].copy()

def reporte_margen_productos(excel_bytes: bytes, año:int, mes_i:int, dia_i:int, mes_f:int, dia_f:int, top_n:int):
    df = read_facturacion(excel_bytes)
    df = _filtrar_fecha(df, año, mes_i, dia_i, mes_f, dia_f)
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

def reporte_producto_volumen_margen(excel_bytes: bytes, año:int, mes_i:int, dia_i:int, mes_f:int, dia_f:int, top_n:int):
    df = read_facturacion(excel_bytes)
    df = _filtrar_fecha(df, año, mes_i, dia_i, mes_f, dia_f)
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

def _top_group(excel_bytes: bytes, año:int, group_col:str, top_n:int, titulo_prefix:str):
    df = read_facturacion(excel_bytes)
    df = df[df["FECHA"].dt.year == año].copy()
    if df.empty:
        raise ValueError(f"No hay datos para el año {año}.")
    rep = (df.groupby(group_col)["NETO"].sum().sort_values(ascending=False).head(top_n))
    total = rep.sum()

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.barh(rep.index, rep.values)
    ax.invert_yaxis()
    ax.set_title(f"{titulo_prefix} {año} (${total:,.0f})", fontsize=14, fontweight="bold")
    ax.set_xlabel("Valor facturado ($)")

    for i, v in enumerate(rep.values):
        ax.text(v + (rep.values.max()*0.01), i, f"${v:,.0f}", va="center", fontsize=9)
    img_file = save_figure(fig, f"top_{group_col.lower()}_{año}")
    xlsx_file = save_dataframe_xlsx(rep.reset_index(name="TOTAL_FACTURACION"), f"top_{group_col.lower()}_{año}")
    return [img_file, xlsx_file]

def reporte_top_ciudades(excel_bytes, año:int, top_n:int):
    return _top_group(excel_bytes, año, "CITY", top_n, "Top Ciudades - Facturación")

def reporte_top_departamentos(excel_bytes, año:int, top_n:int):
    return _top_group(excel_bytes, año, "DEPARTAMENTO", top_n, "Top Departamentos - Facturación")

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
    df["año"] = df["FECHA"].dt.year
    meses_ord = list(meses_map.values())
    max_semana = int(df["SEMANA_MES"].max())
    semanas = list(range(1, max_semana + 1))
    bloques = []
    for anyo in [anyo_actual, anyo_full]:
        df_a = df[df["año"] == anyo]
        if df_a.empty:
            continue
        pivot = (df_a.pivot_table(index="SEMANA_MES", columns="MES", values="NETO", aggfunc="sum", fill_value=0)
                      .reindex(index=semanas, columns=meses_ord, fill_value=0))
        pivot["Total"] = pivot.sum(axis=1)
        pivot.insert(0, "año", anyo)
        pivot.insert(1, "Semana", pivot.index)
        bloques.append(pivot.reset_index(drop=True))
    if not bloques:
        raise ValueError("No hay datos para construir el reporte.")
    salida = pd.concat(bloques, ignore_index=True)
    xlsx_file = save_dataframe_xlsx(salida, "ventas_semana")
    return [xlsx_file]

def reporte_presupuesto_año(excel_bytes, params):
    df = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="Ppto año", header=None)
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
    xlsx_file = save_dataframe_xlsx(df, "presupuesto_año")
    return [xlsx_file]

def extraer_lista_precios_pdf_bytes(pdf_bytes: bytes) -> pd.DataFrame:
    try:
        from PyPDF2 import PdfReader
    except Exception:
        raise RuntimeError("Instala PyPDF2:  pip install PyPDF2")

    rows = []
    patron = re.compile(r"^([A-Z0-9\-\/\.]+)\s+(.*?)\s*\$\s*([\d\.\,]+)\s*$")
    reader = PdfReader(io.BytesIO(pdf_bytes))
    for page in reader.pages:
        text = page.extract_text() or ""
        for line in text.split("\n"):
            m = patron.search(line.strip())
            if not m:
                continue
            item = m.group(1).strip().upper()
            desc = m.group(2).strip()
            precio = m.group(3).strip().replace(".", "").replace(",", ".")
            try:
                precio = float(precio)
            except:
                precio = None
            rows.append([item, desc, precio])

    df = pd.DataFrame(rows, columns=["ITEM", "DESCRIPCION_LISTA", "PRECIO_LISTA"])
    if not df.empty:
        df["ITEM"] = df["ITEM"].astype(str).str.strip().str.upper()
    return df

# --- reporte: Lista + Inventario > 0 SIN ventas (estilo Flask) ---
def reporte_analisis_lista_inv_sin_venta(excel_bytes: bytes, params: dict):
    """
    Ítems que:
      (1) Están en la lista de precios (archivo 'lista' en el form),
      (2) Tienen inventario > 0 (hoja 'Inventario SAI' C7:E1200),
      (3) NO fueron facturados en el año elegido (hoja 'Facturacion').

    Devuelve: [("analisis_lista_inv_sin_venta_YYYY.xlsx", buffer)]
    Requiere en el request: <input type="file" name="lista"> (xlsx/xls/csv/pdf)
    Parámetros: params['año'] o params['anio']
    """
    # --- año ---
    anio_str = (params.get("año") or params.get("anio") or "").strip()
    try:
        anio = int(anio_str)
    except:
        raise ValueError("Parámetro 'año'/'anio' inválido.")

    # --- cargar Facturacion desde bytes (como tus otros reportes) ---
    df_fac = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="Facturacion")
    df_fac.columns = df_fac.columns.str.strip().str.upper().str.replace('\xa0', '', regex=True)
    req_fac = {"ITEM", "FECHA", "QTYSHIP"}
    faltan = req_fac - set(df_fac.columns)
    if faltan:
        raise ValueError(f"Faltan columnas en 'Facturacion': {faltan}")

    df_fac = df_fac.dropna(subset=["ITEM", "FECHA"]).copy()
    df_fac["ITEM"] = df_fac["ITEM"].astype(str).str.strip().str.upper()
    df_fac["FECHA"] = pd.to_datetime(df_fac["FECHA"], errors="coerce")
    df_fac["QTYSHIP"] = pd.to_numeric(df_fac["QTYSHIP"], errors="coerce").fillna(0)
    df_fac_anio = df_fac[df_fac["FECHA"].dt.year == anio]
    vendidos = set(df_fac_anio.loc[df_fac_anio["QTYSHIP"] > 0, "ITEM"].unique())

    # --- cargar Inventario (C7:E1200) ---
    df_inv = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="Inventario SAI", header=None)
    df_inv = df_inv.iloc[6:1200, 2:5].copy()
    df_inv.columns = ["ITEM", "OTRA_COL", "INVENTARIO"]
    df_inv["ITEM"] = df_inv["ITEM"].astype(str).str.strip().str.upper()
    df_inv["INVENTARIO"] = pd.to_numeric(df_inv["INVENTARIO"], errors="coerce").fillna(0)
    df_inv_pos = df_inv[df_inv["INVENTARIO"] > 0][["ITEM", "INVENTARIO"]].drop_duplicates("ITEM")

    # --- leer lista de precios desde request.files['lista'] ---
    fs = request.files.get("lista")
    if not fs or not fs.filename:
        raise ValueError("Adjunta la lista de precios como archivo 'lista' (xlsx/xls/csv/pdf).")

    filename = fs.filename.lower()
    ext = os.path.splitext(filename)[1]
    def _norm(dff: pd.DataFrame) -> pd.DataFrame:
        dff = dff.copy()
        dff.columns = dff.columns.astype(str).str.strip().str.upper()
        return dff

    if ext in [".xlsx", ".xls"]:
        xls_bytes = fs.read()
        xls = pd.ExcelFile(io.BytesIO(xls_bytes))
        hojas = []
        for sn in xls.sheet_names:
            try:
                hojas.append(_norm(pd.read_excel(xls, sn)))
            except Exception:
                pass
        df_lista_raw = pd.concat(hojas, ignore_index=True) if hojas else pd.DataFrame()

    elif ext == ".csv":
        csv_bytes = fs.read()
        df_lista_raw = pd.read_csv(io.BytesIO(csv_bytes), encoding="utf-8", sep=None, engine="python")

    elif ext == ".pdf":
        fs.stream.seek(0)
        pdf_bytes = fs.read()            # ← bytes del PDF
        df_lista_raw = extraer_lista_precios_pdf_bytes(pdf_bytes)


    else:
        raise ValueError("Formato de lista no soportado. Usa xlsx/xls/csv/pdf.")

    if df_lista_raw is None or df_lista_raw.empty:
        # devolver archivo vacío con headers esperados
        cols_out = ["ITEM", "DESCRIPCION_LISTA", "PRECIO_LISTA", "INVENTARIO"]
        return [save_dataframe_xlsx(pd.DataFrame(columns=cols_out), f"analisis_lista_inv_sin_venta_{anio}")]

    # --- mapear columnas ITEM / PRECIO / DESCRIPCION ---
    candidatos_item = {"ITEM", "CODIGO", "CÓDIGO", "SKU", "REFERENCIA"}
    candidatos_precio = {"PRECIO", "PRECIO_LISTA", "PRICE", "PVP", "VALOR", "LISTA", "PRECIO UNITARIO"}
    candidatos_desc = {"DESCRIPCION", "DESCRIPCIÓN", "NOMBRE", "PRODUCTO", "DETALLE"}

    def pick_col(cands, cols):
        for c in cols:
            if c in cands: return c
        for c in cols:
            for k in cands:
                if k in c: return c
        return None

    cols = list(df_lista_raw.columns)
    col_item  = pick_col(candidatos_item, cols)
    col_prec  = pick_col(candidatos_precio, cols)
    col_desc  = pick_col(candidatos_desc, cols)
    if not col_item:
        raise ValueError("No encontré columna de ITEM en la lista de precios.")

    keep = [col_item] + ([col_prec] if col_prec else []) + ([col_desc] if col_desc else [])
    df_lista = df_lista_raw[keep].copy()
    new_cols = ["ITEM"] + (["PRECIO_LISTA"] if col_prec else []) + (["DESCRIPCION_LISTA"] if col_desc else [])
    df_lista.columns = new_cols

    df_lista["ITEM"] = df_lista["ITEM"].astype(str).str.strip().str.upper()
    if "PRECIO_LISTA" in df_lista.columns:
        df_lista["PRECIO_LISTA"] = (
            df_lista["PRECIO_LISTA"].astype(str)
            .str.replace(r"[^\d,.\-]", "", regex=True)
            .str.replace(",", ".", regex=False)
        )
        df_lista["PRECIO_LISTA"] = pd.to_numeric(df_lista["PRECIO_LISTA"], errors="coerce")
    if "DESCRIPCION_LISTA" not in df_lista.columns:
        df_lista["DESCRIPCION_LISTA"] = ""
    df_lista = df_lista.dropna(subset=["ITEM"]).drop_duplicates("ITEM", keep="first")

    # --- lógica principal ---
    base = pd.merge(df_lista, df_inv_pos, on="ITEM", how="inner")
    sin_venta = base[~base["ITEM"].isin(vendidos)].copy()

    # --- salida XLSX ---
    for c in ["DESCRIPCION_LISTA", "PRECIO_LISTA", "INVENTARIO"]:
        if c not in sin_venta.columns:
            sin_venta[c] = None
    cols_out = ["ITEM", "DESCRIPCION_LISTA", "PRECIO_LISTA", "INVENTARIO"]
    sin_venta = sin_venta[cols_out]

    return [save_dataframe_xlsx(sin_venta, f"analisis_lista_inv_sin_venta_{anio}")]

def reporte_rotacion_inventario(excel_bytes: bytes, params: dict):
    """
    Calcula rotación de inventario por ITEM:
      - MESES: meses con ventas (>0) por ítem
      - VENTA_TOTAL: sum(QTYSHIP)
      - PROMEDIO_MES: VENTA_TOTAL / MESES
      - INVENTARIO: tomado de 'Inventario SAI' (C7:E1200)
      - MESES_DIVISION: INVENTARIO / PROMEDIO_MES  (0 si PROMEDIO_MES=0)

    Devuelve: [("rotacion_inventario.xlsx", BytesIO)]
    """
    # === 1) FACTURACIÓN ===
    import io
    import pandas as pd

    df_fac = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="Facturacion")
    df_fac.columns = (
        df_fac.columns.astype(str).str.strip().str.upper().str.replace("\xa0", "", regex=True)
    )

    columnas_fact = {"ITEM", "DESCRIPCION", "FECHA", "QTYSHIP"}
    faltan = columnas_fact - set(df_fac.columns)
    if faltan:
        raise ValueError(f"Faltan columnas en 'Facturacion': {faltan}")

    df_fac = df_fac.dropna(subset=["FECHA", "QTYSHIP", "ITEM"]).copy()
    df_fac["FECHA"] = pd.to_datetime(df_fac["FECHA"], errors="coerce")
    df_fac["QTYSHIP"] = pd.to_numeric(df_fac["QTYSHIP"], errors="coerce").fillna(0)
    df_fac["ITEM"] = df_fac["ITEM"].astype(str).str.strip().str.upper()
    df_fac["DESCRIPCION"] = df_fac["DESCRIPCION"].astype(str).str.strip()
    df_fac["AÑO_MES"] = df_fac["FECHA"].dt.to_period("M")

    meses_activos = (
        df_fac[df_fac["QTYSHIP"] > 0]
        .groupby(["ITEM", "DESCRIPCION"])["AÑO_MES"]
        .nunique()
        .reset_index()
        .rename(columns={"AÑO_MES": "MESES"})
    )

    ventas_totales = (
        df_fac.groupby(["ITEM", "DESCRIPCION"], as_index=False)["QTYSHIP"].sum()
              .rename(columns={"QTYSHIP": "VENTA_TOTAL"})
    )

    resumen = pd.merge(ventas_totales, meses_activos, on=["ITEM", "DESCRIPCION"], how="left").fillna(0)
    resumen["PROMEDIO_MES"] = resumen.apply(
        lambda row: row["VENTA_TOTAL"] / row["MESES"] if row["MESES"] > 0 else 0, axis=1
    )

    # === 2) INVENTARIO (C7:E1200 de 'Inventario SAI') ===
    df_inv = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="Inventario SAI", header=None)
    df_inv = df_inv.iloc[6:1200, 2:5].copy()  # C7:E1200
    df_inv.columns = ["ITEM", "OTRA_COL", "INVENTARIO"]
    df_inv["ITEM"] = df_inv["ITEM"].astype(str).str.strip().str.upper()
    df_inv["INVENTARIO"] = pd.to_numeric(df_inv["INVENTARIO"], errors="coerce").fillna(0)

    resumen = pd.merge(resumen, df_inv[["ITEM", "INVENTARIO"]], on="ITEM", how="left").fillna(0)

    # === 3) Métrica de rotación (meses de división) ===
    resumen["MESES_DIVISION"] = resumen.apply(
        lambda row: (row["INVENTARIO"] / row["PROMEDIO_MES"]) if row["PROMEDIO_MES"] != 0 else 0, axis=1
    )

    # Cast a int como el original
    for col in resumen.select_dtypes(include=["float", "int"]).columns:
        try:
            resumen[col] = pd.to_numeric(resumen[col], errors="coerce").fillna(0).astype(int)
        except Exception:
            pass

    # Orden de columnas
    resumen = resumen[[
        "ITEM", "DESCRIPCION", "MESES", "VENTA_TOTAL",
        "PROMEDIO_MES", "INVENTARIO", "MESES_DIVISION"
    ]]

    # === 4) Exportar ===
    return [save_dataframe_xlsx(resumen, "rotacion_inventario")]



import io, os, re
import pandas as pd
from flask import request


def _normalizar_cols(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = df.columns.astype(str).str.strip().str.upper()
    return df

def _normalizar_item(col: pd.Series) -> pd.Series:
    # Deja SOLO letras y números (limpia NBSP, espacios, guiones raros, etc.)
    return (col.astype(str)
               .str.upper()
               .str.replace(r"[^A-Z0-9]", "", regex=True))

def _codigo_base(item: str) -> str:
    # Elimina sufijos comunes para agrupar variantes (APSUN6-C -> APSUN6)
    if not isinstance(item, str):
        return ""
    item = item.upper().strip()
    return re.sub(r"(-C|-N|-D|-LD|-LN|-SP)$", "", item)

def _leer_lista_precios_desde_request(fact_items: set[str] | None = None) -> tuple[pd.DataFrame, tuple[str, io.BytesIO] | None]:
    """
    Lee la lista (xlsx/xls/csv/pdf). Si es PDF, primero lo convierte a DataFrame
    y además retorna un Excel convertido para descargar.
    Devuelve: (df_estandarizado, xlsx_convertido | None)
              df_estandarizado con columnas: ITEM, PRECIO_LISTA?, DESCRIPCION_LISTA?
    """
    fs = request.files.get("lista")
    if not fs or not fs.filename:
        raise ValueError("Adjunta la lista de precios como archivo 'lista'.")

    filename = fs.filename.lower()
    ext = os.path.splitext(filename)[1]
    converted_xlsx = None

    # --- lectura / conversión ---
    if ext in [".xlsx", ".xls"]:
        xls_bytes = fs.read()
        xls = pd.ExcelFile(io.BytesIO(xls_bytes))
        hojas = []
        for sn in xls.sheet_names:
            try:
                hojas.append(_normalizar_cols(pd.read_excel(xls, sn)))
            except Exception:
                pass
        df_lista_raw = pd.concat(hojas, ignore_index=True) if hojas else pd.DataFrame()

    elif ext == ".csv":
        csv_bytes = fs.read()
        df_lista_raw = pd.read_csv(io.BytesIO(csv_bytes), encoding="utf-8", sep=None, engine="python")
        df_lista_raw = _normalizar_cols(df_lista_raw)

    elif ext == ".pdf":
        pdf_bytes = fs.read()
        df_lista_raw = _pdf_lista_to_dataframe(pdf_bytes)
        # si se pudo leer algo, arma un Excel convertido para que el usuario lo descargue
        if not df_lista_raw.empty:
            buf = io.BytesIO()
            df_lista_raw.to_excel(buf, index=False)
            buf.seek(0)
            converted_xlsx = ("lista_convertida_desde_pdf.xlsx", buf)

    else:
        raise ValueError("Formato no soportado: usa xlsx/xls/csv/pdf.")

    if df_lista_raw is None or df_lista_raw.empty:
        raise ValueError("No pude leer una tabla válida desde la lista de precios.")

    # --- heurística columnas ---
    candidatos_item   = {"ITEM","CODIGO","CÓDIGO","SKU","REFERENCIA","REF","CÓDIGO ITEM"}
    candidatos_precio = {"PRECIO","PRECIO_LISTA","PRICE","PVP","VALOR","LISTA","PRECIO UNITARIO","PRECIO VENTA"}
    candidatos_desc   = {"DESCRIPCION","DESCRIPCIÓN","NOMBRE","PRODUCTO","DETALLE"}

    cols = list(df_lista_raw.columns)

    def pick_col(cands, cols_list):
        for c in cols_list:
            if c in cands: return c
        for c in cols_list:
            up = str(c).upper()
            for k in cands:
                if k in up:
                    return c
        return None

    posibles = [c for c in cols if any(k in str(c).upper() for k in candidatos_item)]
    if not posibles:
        raise ValueError("No se encontró columna de ITEM en la lista.")

    mejor_col = posibles[0]
    if fact_items:
        fact_items_norm = set(_normalizar_item(pd.Series(list(fact_items))))
        max_inter = -1
        for c in posibles:
            norm_vals = _normalizar_item(pd.Series(df_lista_raw[c]).dropna())
            inter = len(set(norm_vals) & fact_items_norm)
            if inter > max_inter:
                mejor_col, max_inter = c, inter

    col_item = mejor_col
    col_prec = pick_col(candidatos_precio, cols)
    col_desc = pick_col(candidatos_desc, cols)

    keep = [col_item] + ([col_prec] if col_prec else []) + ([col_desc] if col_desc else [])
    df_lista = df_lista_raw[keep].copy()
    new_cols = ["ITEM"] + (["PRECIO_LISTA"] if col_prec else []) + (["DESCRIPCION_LISTA"] if col_desc else [])
    df_lista.columns = new_cols

    # normalizar
    df_lista["ITEM"] = _normalizar_item(df_lista["ITEM"])
    if "PRECIO_LISTA" in df_lista.columns:
        df_lista["PRECIO_LISTA"] = (
            df_lista["PRECIO_LISTA"].astype(str)
            .str.replace(r"[^\d,.\-]", "", regex=True)
            .str.replace(",", ".", regex=False)
        )
        df_lista["PRECIO_LISTA"] = pd.to_numeric(df_lista["PRECIO_LISTA"], errors="coerce")

    if "DESCRIPCION_LISTA" not in df_lista.columns:
        df_lista["DESCRIPCION_LISTA"] = ""

    df_lista = df_lista.dropna(subset=["ITEM"]).drop_duplicates(subset=["ITEM"], keep="first")
    return df_lista, converted_xlsx

def reporte_pendientes_lista_precios(excel_bytes: bytes, params: dict):
    """
    Ítems facturados que NO existen en la lista de precios,
    tolerando diferencias de sufijo y caracteres invisibles.
    Ahora incluye INVENTARIO desde la hoja 'Inventario SAI'.
    """
    # --- 1) Facturación ---
    df_fac = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="Facturacion")
    df_fac = _normalizar_cols(df_fac)

    req = {"ITEM", "DESCRIPCION", "FECHA", "NETO", "QTYSHIP"}
    faltan = req - set(df_fac.columns)
    if faltan:
        raise ValueError(f"Faltan columnas en 'Facturacion': {faltan}")

    df_fac = df_fac.dropna(subset=["ITEM", "FECHA"]).copy()
    df_fac["ITEM"] = _normalizar_item(df_fac["ITEM"])
    df_fac["DESCRIPCION"] = df_fac["DESCRIPCION"].astype(str).str.strip()
    df_fac["FECHA"] = pd.to_datetime(df_fac["FECHA"], errors="coerce")
    df_fac["NETO"] = pd.to_numeric(df_fac["NETO"], errors="coerce")
    df_fac["QTYSHIP"] = pd.to_numeric(df_fac["QTYSHIP"], errors="coerce")

    df_fac = df_fac[(df_fac["QTYSHIP"] > 0) & (df_fac["NETO"].notna())].copy()
    df_fac["PRECIO_UNIT"] = df_fac["NETO"] / df_fac["QTYSHIP"]

    ultimos = (df_fac.sort_values(["ITEM", "FECHA"])
                     .drop_duplicates("ITEM", keep="last")
                     .loc[:, ["ITEM", "DESCRIPCION", "FECHA", "PRECIO_UNIT"]]
                     .rename(columns={
                         "DESCRIPCION": "DESCRIPCION_VENTA",
                         "FECHA": "FECHA_ULT_VENTA",
                         "PRECIO_UNIT": "ULTIMO_PRECIO_FACTURADO"
                     }))

    # --- 2) Inventario (hoja Inventario SAI) ---
    try:
        df_inv = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="Inventario SAI", header=None)
        df_inv = df_inv.iloc[6:1200, 2:5].copy()  # rango C7:E1200
        df_inv.columns = ["ITEM", "OTRA_COL", "INVENTARIO"]
        df_inv["ITEM"] = _normalizar_item(df_inv["ITEM"])
        df_inv["INVENTARIO"] = pd.to_numeric(df_inv["INVENTARIO"], errors="coerce").fillna(0)
        df_inv = df_inv[["ITEM", "INVENTARIO"]].drop_duplicates("ITEM")
    except Exception:
        df_inv = pd.DataFrame(columns=["ITEM", "INVENTARIO"])

    # --- 3) Lista (soporta PDF→Excel) ---
    df_lista, lista_convertida = _leer_lista_precios_desde_request(set(ultimos["ITEM"]))

    # --- 4) Cruce robusto por código base ---
    fact_norm  = ultimos.assign(COD_BASE=ultimos["ITEM"].apply(_codigo_base))
    lista_norm = df_lista.assign(COD_BASE=df_lista["ITEM"].apply(_codigo_base))

    comparacion = fact_norm.merge(
        lista_norm[["COD_BASE"]], on="COD_BASE", how="left", indicator=True
    )
    pendientes = comparacion.query('_merge == "left_only"').drop(columns=["_merge"])

    # --- 5) Agregar inventario ---
    pendientes = pendientes.merge(df_inv, on="ITEM", how="left")
    pendientes["INVENTARIO"] = pendientes["INVENTARIO"].fillna(0).astype(int)

    # --- 6) Exportar ---
    cols_out = [
        "ITEM", "DESCRIPCION_VENTA", "FECHA_ULT_VENTA",
        "ULTIMO_PRECIO_FACTURADO", "INVENTARIO"
    ]
    outputs = []

    buf = io.BytesIO()
    if pendientes.empty:
        pd.DataFrame(columns=cols_out).to_excel(buf, index=False)
    else:
        pendientes[cols_out].to_excel(buf, index=False)
    buf.seek(0)
    outputs.append(("pendientes_lista_precios.xlsx", buf))

    # Si la lista fue PDF, agrega el Excel convertido
    if lista_convertida is not None:
        outputs.append(lista_convertida)

    return outputs

def _pdf_lista_to_dataframe(pdf_bytes: bytes) -> pd.DataFrame:
    """
    Extrae la lista de precios desde un PDF:
      1) tabula-py (si hay Java)
      2) camelot (si hay Ghostscript)
      3) pdfplumber + regex (fallback)
    Devuelve un DataFrame tabular crudo. No guarda archivo.
    """
    # --- 1) TABULA ---
    try:
        import tabula  # requiere Java
        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=True) as tf:
            tf.write(pdf_bytes); tf.flush()
            dfs = tabula.read_pdf(tf.name, pages="all", lattice=True, multiple_tables=True)
            if not dfs:
                dfs = tabula.read_pdf(tf.name, pages="all", stream=True, multiple_tables=True)
        if dfs:
            df = pd.concat(dfs, ignore_index=True)
            return _normalizar_cols(df)
    except Exception:
        pass

    # --- 2) CAMELOT ---
    try:
        import camelot  # requiere ghostscript
        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=True) as tf:
            tf.write(pdf_bytes); tf.flush()
            tables = camelot.read_pdf(tf.name, pages="all", flavor="lattice")
            if not tables or tables.n == 0:
                tables = camelot.read_pdf(tf.name, pages="all", flavor="stream")
        if tables and tables.n > 0:
            frames = [t.df for t in tables]
            df = pd.concat(frames, ignore_index=True)
            if not df.empty:
                df.columns = [str(c).strip() for c in df.iloc[0].tolist()]
                df = df.iloc[1:].reset_index(drop=True)
            return _normalizar_cols(df)
    except Exception:
        pass

    # --- 3) pdfplumber (fallback texto + regex) ---
    try:
        import pdfplumber
        rows = []
        patron = re.compile(r"^([A-Z0-9\-/\.]+)\s+(.*?)\s*\$?\s*([\d\.\,]+)\s*$")
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                for line in text.split("\n"):
                    s = line.strip()
                    if not s:
                        continue
                    m = patron.match(s)
                    if not m:
                        continue
                    item = m.group(1).strip()
                    desc = m.group(2).strip()
                    precio = m.group(3).strip().replace(".", "").replace(",", ".")
                    try:
                        precio = float(precio)
                    except:
                        precio = None
                    rows.append([item, desc, precio])
        if rows:
            df = pd.DataFrame(rows, columns=["ITEM", "DESCRIPCION_LISTA", "PRECIO_LISTA"])
            return _normalizar_cols(df)
    except Exception:
        pass

    return pd.DataFrame()


def run_report(reporte, excel_bytes, params):
    if reporte == "reporte_vendedores":
        return reporte_vendedores(excel_bytes, int(params["año"]))
    elif reporte == "reporte_margen_vendedores":
        return reporte_margen_vendedores(excel_bytes, int(params["año"]))
    elif reporte == "reporte_margen_productos":
        return reporte_margen_productos(excel_bytes, int(params["año"]), int(params["mes_inicio"]),
                                        int(params["dia_inicio"]), int(params["mes_fin"]), int(params["dia_fin"]), int(params["top_n"]))
    elif reporte == "reporte_producto_volumen_margen":
        return reporte_producto_volumen_margen(excel_bytes, int(params["año"]), int(params["mes_inicio"]),
                                               int(params["dia_inicio"]), int(params["mes_fin"]), int(params["dia_fin"]), int(params["top_n"]))
    elif reporte == "reporte_top_ciudades":
        return reporte_top_ciudades(excel_bytes, int(params["año"]), int(params["top_n"]))
    elif reporte == "reporte_top_departamentos":
        return reporte_top_departamentos(excel_bytes, int(params["año"]), int(params["top_n"]))
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
    elif reporte == "reporte_presupuesto_año":
        return reporte_presupuesto_año(excel_bytes, params)
    elif reporte == "reporte_analisis_lista_inv_sin_venta":
        return reporte_analisis_lista_inv_sin_venta(excel_bytes, params)
    elif reporte == "reporte_pendientes_lista_precios":
        return reporte_pendientes_lista_precios(excel_bytes, params)
    elif reporte == "reporte_rotacion_inventario":
        return reporte_rotacion_inventario(excel_bytes, params)
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


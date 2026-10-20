"""Lógica de negocio del STATUS por SO.

Lee los Excel crudos (ventas y stock), cruza un SO contra producción,
detecta producto retenido y arma el resultado que muestra la web y el Excel
descargable. No guarda nada en disco.
"""
import io
import json
import re
from pathlib import Path

import pandas as pd

_RAIZ = Path(__file__).parent
CONFIG_DIR = _RAIZ if (_RAIZ / "columnas.json").exists() else _RAIZ / "config"


class StatusError(Exception):
    """Error que se muestra tal cual a quien usa la app."""


# ---------------------------------------------------------------- configuración
def cargar_config():
    def _leer(nombre):
        with open(CONFIG_DIR / nombre, encoding="utf-8") as f:
            return json.load(f)

    return _leer("columnas.json"), _leer("reglas.json"), _leer("equivalencias.json")


# ---------------------------------------------------------------- normalización
def norm_so(valor):
    if valor is None:
        return ""
    if isinstance(valor, float) and valor != valor:  # NaN
        return ""
    return re.sub(r"\.0+$", "", str(valor).strip())


def norm_codigo(valor):
    return str(valor).strip().upper()


def limpiar_descripcion(desc, prefijos):
    texto = re.sub(r"\s+", " ", str(desc)).strip()
    for p in prefijos:
        p = re.sub(r"\s+", " ", p).strip()
        if texto.upper().startswith(p.upper()):
            resto = texto[len(p):].strip()
            if resto:
                return resto
    return texto


# ---------------------------------------------------------------- formato es-AR
def fmt_num(x):
    if abs(x - round(x)) < 0.005:
        s = f"{round(x):,}"
    else:
        s = f"{x:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(p):
    if p is None:
        return "s/SC"
    v = p * 100
    if abs(v) < 0.005:
        return "0%"
    s = f"{v:+.2f}" if abs(v) < 0.05 else f"{v:+.1f}"
    s = s.replace(".", ",")
    if s.endswith(",0"):
        s = s[:-2]
    return s + "%"


def texto_estado(estado, pct):
    p = fmt_pct(pct)
    if estado == "ok":
        return f"🟢 {p}"
    if estado == "pend":
        return f"🔴🚨 {p}"
    return f"🟠🚨 Excede about {p}"


# ---------------------------------------------------------------- lectura de Excel
def _abrir_hojas(data, etiqueta):
    """Devuelve una función por hoja que entrega sus filas. Usa python-calamine
    (muy rápido) y, si no está instalado, openpyxl."""
    try:
        from python_calamine import CalamineWorkbook
    except ImportError:
        CalamineWorkbook = None

    try:
        if CalamineWorkbook is not None:
            wb = CalamineWorkbook.from_filelike(io.BytesIO(data))
            return [(lambda n=n: wb.get_sheet_by_name(n).iter_rows()) for n in wb.sheet_names]
        from openpyxl import load_workbook

        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        return [(lambda ws=ws: ws.iter_rows(values_only=True)) for ws in wb.worksheets]
    except Exception as exc:
        raise StatusError(
            f"No se pudo abrir el Excel de {etiqueta}. Verifique que sea un archivo .xlsx válido."
        ) from exc


def _leer_tabla(data, mapa, requeridas, opcionales, etiqueta, so):
    """Recorre el Excel una sola vez, detecta la fila de encabezados (puede haber
    títulos arriba) y conserva solo las filas del SO pedido y las columnas configuradas."""
    esperadas = [mapa[k] for k in requeridas]
    anclas = {n.strip().lower() for n in esperadas}
    deseadas = {mapa[k].strip().lower(): k for k in requeridas + opcionales if k in mapa}

    filas_so = []
    posiciones = None
    for abrir in _abrir_hojas(data, etiqueta):
        it = iter(abrir())
        for n_fila, fila in enumerate(it):
            if n_fila >= 40:
                break
            celdas = {str(v).strip().lower(): i for i, v in enumerate(fila) if v not in (None, "")}
            if anclas <= set(celdas):
                posiciones = {deseadas[n]: i for n, i in celdas.items() if n in deseadas}
                break
        if posiciones is None:
            continue
        idx_so = posiciones["so"]
        for fila in it:
            if idx_so < len(fila) and norm_so(fila[idx_so]) == so:
                filas_so.append({k: (fila[i] if i < len(fila) else None) for k, i in posiciones.items()})
        break

    if posiciones is None:
        raise StatusError(
            f"No se encontraron las columnas esperadas en el Excel de {etiqueta}: "
            f"{', '.join(esperadas)}. Si cambiaron de nombre, se ajustan en config/columnas.json."
        )
    return pd.DataFrame(filas_so, columns=list(posiciones))


# ---------------------------------------------------------------- cálculo
def clasificar(sc, producido, tolerancia):
    diff = producido - sc
    if sc <= 0:
        return diff, None, ("exc" if producido > 0 else "ok")
    pct = diff / sc
    if abs(pct) <= tolerancia + 1e-9:
        return diff, pct, "ok"
    return diff, pct, ("pend" if diff < 0 else "exc")


def generar_status(ventas_bytes, stock_bytes, so):
    cols, reglas, equivalencias = cargar_config()
    tolerancia = float(reglas.get("tolerancia", 0.10))
    aptos = {e.strip().upper() for e in reglas.get("estados_aptos", ["APTO"])}
    prefijos = reglas.get("prefijos_descripcion", [])
    so = norm_so(so)
    if not so:
        raise StatusError("Ingrese el número de SO.")

    ventas = _leer_tabla(
        ventas_bytes, cols["ventas"], ["so", "codigo", "descripcion", "kg"], ["consignee"], "ventas", so
    )
    stock = _leer_tabla(
        stock_bytes, cols["stock"], ["so", "codigo", "kg"], ["descripcion", "retenido"], "stock", so
    )

    # --- ventas del SO
    v = ventas[ventas["so"].map(norm_so) == so].copy()
    if v.empty:
        raise StatusError(f"No se encontró el SO {so} en el Excel de ventas.")
    v["codigo"] = v["codigo"].map(norm_codigo)
    v["kg"] = pd.to_numeric(v["kg"], errors="coerce").fillna(0.0)
    lineas_venta = v.groupby("codigo", sort=False).agg(desc=("descripcion", "first"), sc=("kg", "sum")).reset_index()
    consignee = ""
    if "consignee" in v.columns:
        no_nulos = v["consignee"].dropna()
        consignee = str(no_nulos.iloc[0]).strip() if not no_nulos.empty else ""

    # --- stock del SO
    s = stock[stock["so"].map(norm_so) == so].copy()
    s["codigo"] = s["codigo"].map(norm_codigo)
    s["kg"] = pd.to_numeric(s["kg"], errors="coerce").fillna(0.0)
    if "retenido" in s.columns:
        estado = s["retenido"].fillna("").astype(str).str.strip().str.upper()
        s["ret"] = (estado != "") & (~estado.isin(aptos))
    else:
        s["ret"] = False

    prod = s.groupby("codigo")["kg"].sum().to_dict()
    ret = s[s["ret"]].groupby("codigo")["kg"].sum().to_dict()
    desc_stock = {}
    if "descripcion" in s.columns:
        desc_stock = s.groupby("codigo")["descripcion"].first().to_dict()

    equiv = {norm_codigo(k): [norm_codigo(c) for c in lst] for k, lst in equivalencias.get("*", {}).items()}
    equiv.update(
        {norm_codigo(k): [norm_codigo(c) for c in lst] for k, lst in equivalencias.get(so, {}).items()}
    )

    # --- líneas
    usados = set()
    lineas = []
    for _, fila in lineas_venta.iterrows():
        codigo = fila["codigo"]
        codigos_prod = equiv.get(codigo, [codigo])
        usados.update(codigos_prod)
        producido = float(sum(prod.get(c, 0.0) for c in codigos_prod))
        retenido = float(sum(ret.get(c, 0.0) for c in codigos_prod))
        sc = float(fila["sc"])
        diff, pct, est = clasificar(sc, producido, tolerancia)
        lineas.append(
            {
                "codigo": codigo,
                "producto": limpiar_descripcion(fila["desc"], prefijos),
                "sc": sc,
                "producido": producido,
                "retenido": retenido,
                "diff": float(diff),
                "pct": None if pct is None else float(pct),
                "estado": est,
                "sc_txt": fmt_num(sc),
                "prod_txt": fmt_num(producido),
                "ret_txt": fmt_num(retenido),
                "diff_txt": fmt_num(diff),
                "status": texto_estado(est, pct),
            }
        )

    # --- producción del SO sin línea de venta
    sin_venta = []
    for codigo, kg in prod.items():
        if codigo not in usados:
            sin_venta.append(
                {
                    "codigo": codigo,
                    "descripcion": str(desc_stock.get(codigo, "")),
                    "kg_txt": fmt_num(float(kg)),
                    "ret_txt": fmt_num(float(ret.get(codigo, 0.0))),
                }
            )

    # --- totales y alertas
    tot_sc = sum(l["sc"] for l in lineas)
    tot_prod = sum(l["producido"] for l in lineas)
    tot_ret = sum(l["retenido"] for l in lineas)
    ret_total_so = float(s.loc[s["ret"], "kg"].sum())
    ret_codigos = int(s.loc[s["ret"], "codigo"].nunique())
    hay_retenido = ret_total_so > 0.005

    pendientes = [l for l in lineas if l["estado"] == "pend"]
    excesos = [l for l in lineas if l["estado"] == "exc"]
    ok = (not hay_retenido) and not pendientes and not excesos

    resumen_pend = ""
    if pendientes:
        partes = [f"{l['producto']} ({fmt_num(-l['diff'])} kg)" for l in pendientes]
        resumen_pend = "Pendiente real: " + ", ".join(partes) + "."

    avisos = []
    if s.empty:
        avisos.append(f"No hay stock cargado con PED. ORIG = {so}. Si la producción está en otros códigos, declare la equivalencia en config/equivalencias.json.")
    if sin_venta:
        avisos.append("Producción de este SO en códigos sin línea de venta: " + ", ".join(x["codigo"] for x in sin_venta) + ".")

    return {
        "so": so,
        "consignee": consignee,
        "lineas": lineas,
        "con_retenido": tot_ret > 0.005,
        "tot": {
            "sc_txt": fmt_num(tot_sc),
            "prod_txt": fmt_num(tot_prod),
            "ret_txt": fmt_num(tot_ret),
            "diff_txt": fmt_num(tot_prod - tot_sc),
            "sc": tot_sc,
            "producido": tot_prod,
            "retenido": tot_ret,
        },
        "hay_retenido": hay_retenido,
        "retenido_txt": (
            f"{fmt_num(ret_total_so)} kg retenidos en {ret_codigos} código" + ("s" if ret_codigos != 1 else "")
        )
        if hay_retenido
        else "",
        "ok": ok,
        "resumen_pend": resumen_pend,
        "sin_venta": sin_venta,
        "avisos": avisos,
    }

"""
aplicacion.py - Aplicación web principal en Streamlit para el 
Sistema de Conciliación y Tesorería Multidivisa con:
- Aislamiento previo de Asientos Anulados.
- Clasificación de Naturaleza Contable (distinguiendo Transferencias entre Cuentas Propias,
  Ingresos por Ventas/Clientes, Ingresos Bancarios/Intereses, Devoluciones y Egresos Operativos).
- Reporte Excel Consolidado con peso porcentual (%) por partida sobre Total Ingresos y Total Egresos,
  fila de Totales y pestañas de Top / Distribución Porcentual.
- Asistente IA conectado a datos reales del Mayor y exportación directa a Word (.docx) y Excel (.xlsx).
"""
from __future__ import annotations

import io
import json
import os
import re
import zipfile
from datetime import datetime
from xml.sax.saxutils import escape as xml_escape

import pandas as pd
import streamlit as st

# Importación de motores lógicos locales
from conciliar import (
    Config,
    ErrorDeParseo,
    calcular_cuadre,
    cargar_movimientos,
    conciliar,
)
from tasas import TablaTasas, TasaNoDisponible

# Cliente OpenAI moderno (requiere openai>=1.0.0)
try:
    from openai import OpenAI
except ImportError:
    OpenAI = None

# =============================================================================
# 1. CONFIGURACIÓN INICIAL DE LA PÁGINA Y PERSISTENCIA
# =============================================================================
st.set_page_config(
    page_title="Conciliación - AI Reconciliation",
    page_icon="📊",
    layout="wide",
)

HISTORIAL_FILE = "historial_conciliaciones.json"


def cargar_historial_disco() -> list[dict]:
    if os.path.exists(HISTORIAL_FILE):
        try:
            with open(HISTORIAL_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []


def guardar_en_historial(resumen_dict: dict) -> None:
    historial = cargar_historial_disco()
    historial.insert(0, resumen_dict)
    try:
        with open(HISTORIAL_FILE, "w", encoding="utf-8") as f:
            json.dump(historial, f, ensure_ascii=False, indent=4)
    except Exception:
        pass


# =============================================================================
# 2. CLASIFICACIÓN CONTABLE DE INGRESOS Y EGRESOS (CUENTAS PROPIAS VS TERCEROS)
# =============================================================================
_PATRON_BANCOS_PROPIOS = re.compile(
    r"\b(banco|banesco|mercantil|provincial|bnc|nacional de credito|venezuela|"
    r"bancamiga|tesoro|bancaribe|exterior|plaza|caroni|activo|sofitasa|100%\s*banco)\b",
    re.IGNORECASE,
)
_PATRON_INGRESO_BANCARIO = re.compile(
    r"\b(interes\w*|rendimiento\w*|abono intereses|intereses bancarios)\b",
    re.IGNORECASE,
)
_PATRON_DEVOLUCION = re.compile(
    r"\b(devoluc\w*|reintegro\w*|reembolso\w*|reverso|error|nota de credito)\b",
    re.IGNORECASE,
)
_PATRON_ANTICIPO_VENTAS = re.compile(
    r"\b(anticipo\w*|viatico\w*|representantes de ventas|gastos de viaje)\b",
    re.IGNORECASE,
)
_PATRON_NOMINA = re.compile(
    r"\b(nomina|nómina|sueldo\w*|salario\w*|vacacion\w*|liquidacion\w*|utilidades|cestaticket)\b",
    re.IGNORECASE,
)
_PATRON_CXC_INTERNAS = re.compile(
    r"\b(cuentas por cobrar accionistas|cuentas por cobrar personal|prestamo\w*)\b",
    re.IGNORECASE,
)
_PATRON_PARAFISCAL_IMPUESTO = re.compile(
    r"\b(i\.?v\.?s\.?s|f\.?a\.?o\.?v|inces|seniat|insai|igtf|impuesto\w*|retencion\w*|alcaldia)\b",
    re.IGNORECASE,
)


def clasificar_naturaleza_contable(desc_mayor: str, desc_banco: str, monto_mayor: float, monto_banco: float, marca: str) -> str:
    """
    Distingue con precisión la naturaleza del movimiento para separar:
    - Transferencias entre Cuentas Propias (Fondeos entre bancos de la empresa)
    - Ingresos por Ventas / Cobranzas de Clientes
    - Ingresos Bancarios (Intereses)
    - Ingresos por Devoluciones / Reintegros de Proveedores
    - Categorías clave de Egresos
    """
    if marca == "Asiento Anulado":
        return "Asiento Anulado (Neutralizado)"

    texto = (str(desc_mayor or "").strip() or str(desc_banco or "").strip())
    monto = float(monto_mayor) if float(monto_mayor) != 0.0 else float(monto_banco)

    if monto > 0:
        if _PATRON_BANCOS_PROPIOS.search(texto):
            return "Ingreso por Transferencia entre Cuentas Propias"
        if _PATRON_INGRESO_BANCARIO.search(texto):
            return "Ingreso Bancario / Intereses"
        if _PATRON_DEVOLUCION.search(texto):
            return "Ingreso por Devolución / Reintegro"
        return "Ingreso por Ventas / Cobranza a Clientes"
    elif monto < 0:
        if _PATRON_BANCOS_PROPIOS.search(texto):
            return "Egreso por Transferencia entre Cuentas Propias"
        if _PATRON_ANTICIPO_VENTAS.search(texto):
            return "Egreso - Anticipos y Viáticos de Ventas"
        if _PATRON_NOMINA.search(texto):
            return "Egreso - Nómina y Personal"
        if _PATRON_CXC_INTERNAS.search(texto):
            return "Egreso - Cuentas por Cobrar Accionistas / Personal"
        if _PATRON_PARAFISCAL_IMPUESTO.search(texto):
            return "Egreso - Impuestos y Parafiscales"
        return "Egreso - Pagos a Proveedores y Operativos"
    return "Sin monto"


def construir_dataframe_consolidado(res, tabla_t) -> pd.DataFrame:
    """
    Construye el DataFrame consolidado completo e incorpora:
    - Naturaleza Contable (Cuentas Propias vs Ventas Clientes vs Bancarios vs Proveedores)
    - % sobre Total Ingresos Mayor y % sobre Total Egresos Mayor para cada partida.
    """
    filas_excel = []

    # 1. Pares cerrados (Conciliados)
    for par in res.pares:
        b_movs = par.banco
        l_movs = par.libros
        max_len = max(len(b_movs), len(l_movs))
        for i in range(max_len):
            mb = b_movs[i] if i < len(b_movs) else None
            ml = l_movs[i] if i < len(l_movs) else None
            f_op = mb.fecha if mb else (ml.fecha if ml else None)

            tasa_usd_val = 0.0
            monto_usd = 0.0
            if f_op:
                try:
                    t_obj = tabla_t.obtener(f_op, "USD", tipo="BCV")
                    tasa_usd_val = float(t_obj.valor)
                    monto_base = float(ml.monto) if ml else (float(mb.monto) if mb else 0.0)
                    monto_usd = round(monto_base / tasa_usd_val, 2) if tasa_usd_val > 0 else 0.0
                except Exception:
                    pass

            m_bco = float(mb.monto) if mb else 0.0
            m_may = float(ml.monto) if ml else 0.0
            d_bco = mb.descripcion if mb else ""
            d_may = ml.descripcion if ml else ""
            marca = f"Conciliado ({par.regla})"

            filas_excel.append({
                "Banco - Fecha": mb.fecha if mb else "",
                "Banco - Descripción": d_bco,
                "Banco - Referencia": mb.referencia if mb else "",
                "Banco - Monto": m_bco,
                "Mayor - Fecha": ml.fecha if ml else "",
                "Mayor - Descripción": d_may,
                "Mayor - Referencia": ml.referencia if ml else "",
                "Mayor - Monto": m_may,
                "Marca Conciliación": marca,
                "Naturaleza Contable": clasificar_naturaleza_contable(d_may, d_bco, m_may, m_bco, marca),
                "Tasa BCV USD": tasa_usd_val,
                "Monto Convertido USD": monto_usd,
            })

    # 2. Partidas pendientes del Banco
    for mb in res.banco_pendiente:
        tasa_usd_val = 0.0
        monto_usd = 0.0
        try:
            t_obj = tabla_t.obtener(mb.fecha, "USD", tipo="BCV")
            tasa_usd_val = float(t_obj.valor)
            monto_usd = round(float(mb.monto) / tasa_usd_val, 2) if tasa_usd_val > 0 else 0.0
        except Exception:
            pass

        m_bco = float(mb.monto)
        marca = "Por conciliar (Banco)"
        filas_excel.append({
            "Banco - Fecha": mb.fecha,
            "Banco - Descripción": mb.descripcion,
            "Banco - Referencia": mb.referencia,
            "Banco - Monto": m_bco,
            "Mayor - Fecha": "",
            "Mayor - Descripción": "",
            "Mayor - Referencia": "",
            "Mayor - Monto": 0.0,
            "Marca Conciliación": marca,
            "Naturaleza Contable": clasificar_naturaleza_contable("", mb.descripcion, 0.0, m_bco, marca),
            "Tasa BCV USD": tasa_usd_val,
            "Monto Convertido USD": monto_usd,
        })

    # 3. Partidas pendientes de Libros
    for ml in res.libros_pendiente:
        tasa_usd_val = 0.0
        monto_usd = 0.0
        try:
            t_obj = tabla_t.obtener(ml.fecha, "USD", tipo="BCV")
            tasa_usd_val = float(t_obj.valor)
            monto_usd = round(float(ml.monto) / tasa_usd_val, 2) if tasa_usd_val > 0 else 0.0
        except Exception:
            pass

        m_may = float(ml.monto)
        marca = "Por conciliar (Libros)"
        filas_excel.append({
            "Banco - Fecha": "",
            "Banco - Descripción": "",
            "Banco - Referencia": "",
            "Banco - Monto": 0.0,
            "Mayor - Fecha": ml.fecha,
            "Mayor - Descripción": ml.descripcion,
            "Mayor - Referencia": ml.referencia,
            "Mayor - Monto": m_may,
            "Marca Conciliación": marca,
            "Naturaleza Contable": clasificar_naturaleza_contable(ml.descripcion, "", m_may, 0.0, marca),
            "Tasa BCV USD": tasa_usd_val,
            "Monto Convertido USD": monto_usd,
        })

    # 4. Asientos Anulados detectados en Libros
    for ml, marca_anulado in getattr(res, "libros_anulados", []):
        tasa_usd_val = 0.0
        monto_usd = 0.0
        try:
            t_obj = tabla_t.obtener(ml.fecha, "USD", tipo="BCV")
            tasa_usd_val = float(t_obj.valor)
            monto_usd = round(float(ml.monto) / tasa_usd_val, 2) if tasa_usd_val > 0 else 0.0
        except Exception:
            pass

        m_may = float(ml.monto)
        filas_excel.append({
            "Banco - Fecha": "",
            "Banco - Descripción": "",
            "Banco - Referencia": "",
            "Banco - Monto": 0.0,
            "Mayor - Fecha": ml.fecha,
            "Mayor - Descripción": ml.descripcion,
            "Mayor - Referencia": ml.referencia,
            "Mayor - Monto": m_may,
            "Marca Conciliación": marca_anulado,
            "Naturaleza Contable": "Asiento Anulado (Neutralizado)",
            "Tasa BCV USD": tasa_usd_val,
            "Monto Convertido USD": monto_usd,
        })

    df = pd.DataFrame(filas_excel)
    if df.empty:
        return df

    # Calcular totales activos de Mayor (excluyendo asientos anulados) para los pesos porcentuales
    mask_activo = df["Marca Conciliación"] != "Asiento Anulado"
    total_ingresos_mayor = df.loc[mask_activo & (df["Mayor - Monto"] > 0), "Mayor - Monto"].sum()
    total_egresos_mayor = df.loc[mask_activo & (df["Mayor - Monto"] < 0), "Mayor - Monto"].abs().sum()

    pct_ing_list = []
    pct_egr_list = []
    for _, row in df.iterrows():
        if row["Marca Conciliación"] == "Asiento Anulado":
            pct_ing_list.append(0.0)
            pct_egr_list.append(0.0)
            continue
        val_m = float(row["Mayor - Monto"])
        if val_m > 0 and total_ingresos_mayor > 0:
            pct_ing_list.append(round((val_m / total_ingresos_mayor) * 100.0, 2))
            pct_egr_list.append(0.0)
        elif val_m < 0 and total_egresos_mayor > 0:
            pct_ing_list.append(0.0)
            pct_egr_list.append(round((abs(val_m) / total_egresos_mayor) * 100.0, 2))
        else:
            pct_ing_list.append(0.0)
            pct_egr_list.append(0.0)

    df["% s/ Total Ingresos Mayor"] = pct_ing_list
    df["% s/ Total Egresos Mayor"] = pct_egr_list
    return df


def generar_tablas_distribucion(df_rep: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Genera 3 DataFrames analíticos para exportar en pestañas de Excel:
    1) Resumen Ejecutivo de Totales (Ingresos por Ventas vs Transferencias Cuentas Propias vs Egresos)
    2) Top y Distribución Porcentual de Egresos por 'Mayor - Descripción'
    3) Top y Distribución Porcentual de Ingresos por 'Mayor - Descripción' (con peso sobre Total y sobre Ventas Clientes)
    """
    if df_rep.empty:
        vacio = pd.DataFrame()
        return vacio, vacio, vacio

    df_act = df_rep[
        (df_rep["Mayor - Descripción"].astype(str).str.strip() != "")
        & (df_rep["Marca Conciliación"] != "Asiento Anulado")
    ].copy()

    # 1. Tabla de Egresos Agrupada
    egresos = df_act[df_act["Mayor - Monto"] < 0].copy()
    if not egresos.empty:
        egresos["Monto_VES"] = egresos["Mayor - Monto"].abs()
        egresos["Monto_USD"] = egresos["Monto Convertido USD"].abs()
        tot_eg_ves = egresos["Monto_VES"].sum()
        tot_eg_usd = egresos["Monto_USD"].sum()

        df_dist_egresos = (
            egresos.groupby(["Mayor - Descripción", "Naturaleza Contable"], as_index=False)
            .agg(
                Movimientos=("Monto_VES", "count"),
                Monto_Total_VES=("Monto_VES", "sum"),
                Monto_Total_USD=("Monto_USD", "sum"),
            )
            .sort_values("Monto_Total_VES", ascending=False)
        )
        df_dist_egresos["% sobre Total Egresos"] = (
            (df_dist_egresos["Monto_Total_VES"] / tot_eg_ves * 100.0).round(2) if tot_eg_ves > 0 else 0.0
        )
        # Fila de Total
        fila_tot_eg = pd.DataFrame([{
            "Mayor - Descripción": "TOTAL EGRESOS MAYOR (SIN ANULADOS)",
            "Naturaleza Contable": "TOTAL",
            "Movimientos": int(df_dist_egresos["Movimientos"].sum()),
            "Monto_Total_VES": round(tot_eg_ves, 2),
            "Monto_Total_USD": round(tot_eg_usd, 2),
            "% sobre Total Egresos": 100.0,
        }])
        df_dist_egresos = pd.concat([df_dist_egresos, fila_tot_eg], ignore_index=True)
    else:
        df_dist_egresos = pd.DataFrame()

    # 2. Tabla de Ingresos Agrupada (separando Transferencias Propias de Ventas Clientes)
    ingresos = df_act[df_act["Mayor - Monto"] > 0].copy()
    if not ingresos.empty:
        ingresos["Monto_VES"] = ingresos["Mayor - Monto"]
        ingresos["Monto_USD"] = ingresos["Monto Convertido USD"]
        tot_in_ves = ingresos["Monto_VES"].sum()
        tot_in_usd = ingresos["Monto_USD"].sum()

        tot_ventas_clientes = ingresos.loc[
            ingresos["Naturaleza Contable"] == "Ingreso por Ventas / Cobranza a Clientes", "Monto_VES"
        ].sum()

        df_dist_ingresos = (
            ingresos.groupby(["Mayor - Descripción", "Naturaleza Contable"], as_index=False)
            .agg(
                Movimientos=("Monto_VES", "count"),
                Monto_Total_VES=("Monto_VES", "sum"),
                Monto_Total_USD=("Monto_USD", "sum"),
            )
            .sort_values("Monto_Total_VES", ascending=False)
        )
        df_dist_ingresos["% s/ Total Ingresos General"] = (
            (df_dist_ingresos["Monto_Total_VES"] / tot_in_ves * 100.0).round(2) if tot_in_ves > 0 else 0.0
        )
        df_dist_ingresos["% s/ Ingresos por Ventas (Clientes)"] = df_dist_ingresos.apply(
            lambda r: round((r["Monto_Total_VES"] / tot_ventas_clientes) * 100.0, 2)
            if r["Naturaleza Contable"] == "Ingreso por Ventas / Cobranza a Clientes" and tot_ventas_clientes > 0
            else 0.0,
            axis=1,
        )
        fila_tot_in = pd.DataFrame([{
            "Mayor - Descripción": "TOTAL INGRESOS MAYOR (SIN ANULADOS)",
            "Naturaleza Contable": "TOTAL",
            "Movimientos": int(df_dist_ingresos["Movimientos"].sum()),
            "Monto_Total_VES": round(tot_in_ves, 2),
            "Monto_Total_USD": round(tot_in_usd, 2),
            "% s/ Total Ingresos General": 100.0,
            "% s/ Ingresos por Ventas (Clientes)": 100.0 if tot_ventas_clientes > 0 else 0.0,
        }])
        df_dist_ingresos = pd.concat([df_dist_ingresos, fila_tot_in], ignore_index=True)
    else:
        df_dist_ingresos = pd.DataFrame()

    # 3. Resumen Ejecutivo por Naturaleza Contable
    if not df_act.empty:
        df_res_nat = (
            df_act.groupby("Naturaleza Contable", as_index=False)
            .agg(
                Cantidad_Partidas=("Mayor - Monto", "count"),
                Monto_Neto_VES=("Mayor - Monto", "sum"),
                Monto_Neto_USD=("Monto Convertido USD", "sum"),
            )
            .sort_values("Monto_Neto_VES", ascending=False)
        )
    else:
        df_res_nat = pd.DataFrame()

    return df_res_nat, df_dist_egresos, df_dist_ingresos


# =============================================================================
# 3. EXPORTADOR NATIVO A WORD (.DOCX) CON SOPORTE DE TABLAS Y FORMATO EJECUTIVO
# =============================================================================
def generar_word_analisis_ia(pregunta: str, respuesta_md: str, df_res_nat: pd.DataFrame, df_eg: pd.DataFrame, df_in: pd.DataFrame) -> bytes:
    """
    Genera un documento Word (.docx) válido y profesional con el análisis de la IA
    y las tablas numéricas de distribución porcentual de Ingresos y Egresos.
    Funciona nativamente en cualquier servidor sin requerir librerías externas adicionales.
    """
    parrafos_xml = []

    def _p_titulo(texto: str, nivel: int = 1) -> str:
        tam = "32" if nivel == 1 else ("26" if nivel == 2 else "22")
        color = "1B4332" if nivel == 1 else "2D6A4F"
        t_esc = xml_escape(texto)
        return (
            f'<w:p><w:pPr><w:spacing w:before="200" w:after="100"/></w:pPr>'
            f'<w:r><w:rPr><w:b/><w:color w:val="{color}"/><w:sz w:val="{tam}"/></w:rPr>'
            f'<w:t>{t_esc}</w:t></w:r></w:p>'
        )

    def _p_texto(texto: str, negrita: bool = False) -> str:
        t_esc = xml_escape(texto)
        b_tag = "<w:b/>" if negrita else ""
        return (
            f'<w:p><w:pPr><w:spacing w:after="80"/></w:pPr>'
            f'<w:r><w:rPr>{b_tag}<w:sz w:val="21"/></w:rPr>'
            f'<w:t xml:space="preserve">{t_esc}</w:t></w:r></w:p>'
        )

    def _tabla_xml(encabezados: list[str], filas: list[list[str]]) -> str:
        bord = (
            '<w:tblBorders>'
            '<w:top w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
            '<w:left w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
            '<w:right w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
            '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
            '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
            '</w:tblBorders>'
        )
        xml_t = [f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>{bord}</w:tblPr>']
        # Encabezado
        xml_t.append("<w:tr>")
        for h in encabezados:
            h_esc = xml_escape(str(h))
            xml_t.append(
                f'<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="1B4332"/></w:tcPr>'
                f'<w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="19"/></w:rPr>'
                f'<w:t>{h_esc}</w:t></w:r></w:p></w:tc>'
            )
        xml_t.append("</w:tr>")
        # Filas
        for idx_f, fila in enumerate(filas):
            fill_c = "F8FAFC" if idx_f % 2 == 1 else "FFFFFF"
            xml_t.append("<w:tr>")
            for celda in fila:
                c_esc = xml_escape(str(celda))
                xml_t.append(
                    f'<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="{fill_c}"/></w:tcPr>'
                    f'<w:p><w:r><w:rPr><w:sz w:val="19"/></w:rPr>'
                    f'<w:t xml:space="preserve">{c_esc}</w:t></w:r></w:p></w:tc>'
                )
            xml_t.append("</w:tr>")
        xml_t.append("</w:tbl>")
        return "".join(xml_t)

    # Encabezado del documento
    parrafos_xml.append(_p_titulo("Informe Ejecutivo de Conciliación y Tesorería (IA)", 1))
    parrafos_xml.append(_p_texto(f"Fecha de emisión: {datetime.now().strftime('%Y-%m-%d %H:%M')}", negrita=True))
    if pregunta:
        parrafos_xml.append(_p_texto(f"Consulta analizada: {pregunta}", negrita=True))

    # Parsear el Markdown del asistente (incluyendo tablas Markdown si las hubiera)
    lineas_md = (respuesta_md or "").splitlines()
    buffer_tabla = []

    def _vaciar_buffer_tabla():
        nonlocal buffer_tabla
        if not buffer_tabla:
            return
        filas_limpias = []
        for l_tab in buffer_tabla:
            partes = [c.strip() for c in l_tab.strip().strip("|").split("|")]
            # Ignorar fila separadora de markdown (---|---|---)
            if all(set(p) <= set("-: ") for p in partes if p):
                continue
            filas_limpias.append(partes)
        if len(filas_limpias) >= 2:
            parrafos_xml.append(_tabla_xml(filas_limpias[0], filas_limpias[1:]))
        buffer_tabla = []

    for lin in lineas_md:
        s = lin.strip()
        if s.startswith("|") and s.endswith("|") and "|" in s[1:-1]:
            buffer_tabla.append(s)
            continue
        else:
            _vaciar_buffer_tabla()

        if not s:
            continue
        if s.startswith("### "):
            parrafos_xml.append(_p_titulo(s.replace("### ", "").replace("**", ""), 3))
        elif s.startswith("## "):
            parrafos_xml.append(_p_titulo(s.replace("## ", "").replace("**", ""), 2))
        elif s.startswith("# "):
            parrafos_xml.append(_p_titulo(s.replace("# ", "").replace("**", ""), 1))
        else:
            limpio = s.replace("**", "")
            parrafos_xml.append(_p_texto(limpio))

    _vaciar_buffer_tabla()

    # Añadir tablas numéricas estructuradas al final del Word
    if not df_res_nat.empty:
        parrafos_xml.append(_p_titulo("Resumen por Naturaleza Contable (Cuentas Propias vs Operativo)", 2))
        filas_nat = [
            [
                r["Naturaleza Contable"],
                str(int(r["Cantidad_Partidas"])),
                f"{float(r['Monto_Neto_VES']):,.2f}",
                f"{float(r['Monto_Neto_USD']):,.2f}",
            ]
            for _, r in df_res_nat.iterrows()
        ]
        parrafos_xml.append(_tabla_xml(["Naturaleza Contable", "Partidas", "Monto Neto (VES)", "Monto Neto (USD)"], filas_nat))

    if not df_in.empty:
        parrafos_xml.append(_p_titulo("Top y Distribución Porcentual de Ingresos (Mayor)", 2))
        filas_in = [
            [
                r["Mayor - Descripción"],
                r["Naturaleza Contable"],
                str(int(r["Movimientos"])),
                f"{float(r['Monto_Total_VES']):,.2f}",
                f"{float(r['% s/ Total Ingresos General']):.2f}%",
                f"{float(r['% s/ Ingresos por Ventas (Clientes)']):.2f}%",
            ]
            for _, r in df_in.head(20).iterrows()
        ]
        parrafos_xml.append(
            _tabla_xml(
                ["Mayor - Descripción", "Clasificación", "Mov.", "Total VES", "% Total Ing.", "% Ventas Clientes"],
                filas_in,
            )
        )

    if not df_eg.empty:
        parrafos_xml.append(_p_titulo("Top y Distribución Porcentual de Egresos (Mayor)", 2))
        filas_eg = [
            [
                r["Mayor - Descripción"],
                r["Naturaleza Contable"],
                str(int(r["Movimientos"])),
                f"{float(r['Monto_Total_VES']):,.2f}",
                f"{float(r['% sobre Total Egresos']):.2f}%",
            ]
            for _, r in df_eg.head(20).iterrows()
        ]
        parrafos_xml.append(
            _tabla_xml(
                ["Mayor - Descripción", "Clasificación", "Mov.", "Total VES", "% Total Egresos"],
                filas_eg,
            )
        )

    body_xml = "".join(parrafos_xml)
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f"<w:body>{body_xml}<w:sectPr/></w:body></w:document>"
    )
    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        "</Relationships>"
    )

    mem_zip = io.BytesIO()
    with zipfile.ZipFile(mem_zip, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", rels_xml)
        zf.writestr("word/document.xml", document_xml)
    return mem_zip.getvalue()


def construir_contexto_ia_real(df_rep: pd.DataFrame, res) -> str:
    """
    Construye un contexto detallado basado EXCLUSIVAMENTE en los datos reales
    cargados por el usuario, separando explícitamente en los Ingresos:
    1) Transferencias entre Cuentas Propias (Fondeos bancarios internos)
    2) Ingresos por Ventas / Cobranzas a Clientes
    3) Ingresos Bancarios (Intereses)
    4) Devoluciones / Reintegros
    """
    if df_rep.empty:
        return "No hay movimientos en el reporte."

    df_act = df_rep[
        (df_rep["Mayor - Descripción"].astype(str).str.strip() != "")
        & (df_rep["Marca Conciliación"] != "Asiento Anulado")
    ].copy()

    tot_ing_general = df_act.loc[df_act["Mayor - Monto"] > 0, "Mayor - Monto"].sum()
    tot_egr_general = df_act.loc[df_act["Mayor - Monto"] < 0, "Mayor - Monto"].abs().sum()

    ing_propias = df_act.loc[
        df_act["Naturaleza Contable"] == "Ingreso por Transferencia entre Cuentas Propias", "Mayor - Monto"
    ].sum()
    ing_clientes = df_act.loc[
        df_act["Naturaleza Contable"] == "Ingreso por Ventas / Cobranza a Clientes", "Mayor - Monto"
    ].sum()
    ing_bancarios = df_act.loc[
        df_act["Naturaleza Contable"] == "Ingreso Bancario / Intereses", "Mayor - Monto"
    ].sum()
    ing_devoluciones = df_act.loc[
        df_act["Naturaleza Contable"] == "Ingreso por Devolución / Reintegro", "Mayor - Monto"
    ].sum()

    lineas = []
    lineas.append("=== RESUMEN EJECUTIVO Y DESGLOSE DE INGRESOS (CUENTAS PROPIAS VS VENTAS REALES) ===")
    lineas.append(f"- Pares conciliados: {len(res.pares)}")
    lineas.append(f"- Pendientes en Banco: {len(res.banco_pendiente)}")
    lineas.append(f"- Pendientes en Libros: {len(res.libros_pendiente)}")
    lineas.append(f"- Asientos Anulados aislados (neutralizados): {len(getattr(res, 'libros_anulados', []))}")
    lineas.append(f"- TOTAL INGRESOS BRUTOS EN MAYOR (sin anulados): {tot_ing_general:,.2f} VES")
    lineas.append(
        f"   * De los cuales: TRANSFERENCIAS ENTRE CUENTAS PROPIAS / FONDEOS INTERNOS (ej. Banco Banesco, BNC): "
        f"{ing_propias:,.2f} VES ({(ing_propias/tot_ing_general*100 if tot_ing_general else 0):.2f}% del total entradas)"
    )
    lineas.append(
        f"   * De los cuales: INGRESOS REALES POR VENTAS / COBRANZAS A CLIENTES: "
        f"{ing_clientes:,.2f} VES ({(ing_clientes/tot_ing_general*100 if tot_ing_general else 0):.2f}% del total entradas)"
    )
    lineas.append(f"   * De los cuales: INGRESOS BANCARIOS / INTERESES: {ing_bancarios:,.2f} VES")
    lineas.append(f"   * De los cuales: INGRESOS POR DEVOLUCIONES / REINTEGROS: {ing_devoluciones:,.2f} VES")
    lineas.append(f"- TOTAL EGRESOS EN MAYOR (sin anulados): {tot_egr_general:,.2f} VES")
    lineas.append(f"- FLUJO NETO TOTAL EN MAYOR: {(tot_ing_general - tot_egr_general):,.2f} VES")

    df_res_nat, df_eg, df_in = generar_tablas_distribucion(df_rep)

    if not df_in.empty:
        lineas.append("\n=== DISTRIBUCIÓN PORCENTUAL DE INGRESOS POR 'Mayor - Descripción' ===")
        for _, r in df_in.iterrows():
            if r["Naturaleza Contable"] == "TOTAL":
                continue
            lineas.append(
                f"  * '{r['Mayor - Descripción']}' | Tipo: {r['Naturaleza Contable']} | "
                f"Movimientos: {int(r['Movimientos'])} | Monto VES: {r['Monto_Total_VES']:,.2f} | "
                f"USD: {r['Monto_Total_USD']:,.2f} | % s/ Total Entradas: {r['% s/ Total Ingresos General']:.2f}% | "
                f"% s/ Ventas Clientes: {r['% s/ Ingresos por Ventas (Clientes)']:.2f}%"
            )

    if not df_eg.empty:
        lineas.append("\n=== DISTRIBUCIÓN PORCENTUAL DE EGRESOS POR 'Mayor - Descripción' ===")
        for _, r in df_eg.iterrows():
            if r["Naturaleza Contable"] == "TOTAL":
                continue
            lineas.append(
                f"  * '{r['Mayor - Descripción']}' | Tipo: {r['Naturaleza Contable']} | "
                f"Movimientos: {int(r['Movimientos'])} | Monto VES: {r['Monto_Total_VES']:,.2f} | "
                f"USD: {r['Monto_Total_USD']:,.2f} | % s/ Total Egresos: {r['% sobre Total Egresos']:.2f}%"
            )

    return "\n".join(lineas)


# Estilos CSS profesionales
st.markdown(
    """
    <style>
        .stApp { background-color: #F4F6F9; }
        .dashboard-card {
            background-color: #FFFFFF;
            padding: 20px;
            border-radius: 10px;
            border: 1px solid #E2E8F0;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            margin-bottom: 15px;
        }
        .stButton button { border-radius: 8px; font-weight: 500; }
    </style>
""",
    unsafe_allow_html=True,
)

# =============================================================================
# 4. GESTIÓN DE ESTADO Y SEGURIDAD (AUTENTICACIÓN)
# =============================================================================
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.title("🔒 Acceso al Sistema de Conciliación")
    st.markdown("Por favor, ingrese su contraseña para continuar.")

    password_correcta = st.secrets.get("PASSWORD_ACCESO", "admin123")
    password = st.text_input("Contraseña:", type="password")

    if st.button("Ingresar"):
        if password and password == password_correcta:
            st.session_state.autenticado = True
            st.rerun()
        else:
            st.error("Contraseña incorrecta. Intente nuevamente.")
    st.stop()

if "tabla_tasas" not in st.session_state:
    st.session_state.tabla_tasas = TablaTasas(max_antiguedad_dias=7)

if "cuentas_bancarias" not in st.session_state:
    st.session_state.cuentas_bancarias = {
        "Banco Nacional": ["0102-1234-56-1234567890"],
        "Banco Regional": ["0105-9876-54-0987654321"],
    }

if "resultado_conciliacion" not in st.session_state:
    st.session_state.resultado_conciliacion = None

if "movs_banco" not in st.session_state:
    st.session_state.movs_banco = []

if "movs_libros" not in st.session_state:
    st.session_state.movs_libros = []

if "ultimo_analisis_ia" not in st.session_state:
    st.session_state.ultimo_analisis_ia = ""

if "ultima_pregunta_ia" not in st.session_state:
    st.session_state.ultima_pregunta_ia = ""


# =============================================================================
# 5. BARRA LATERAL DE NAVEGACIÓN
# =============================================================================
with st.sidebar:
    st.markdown("### 📊 Conciliacion")
    st.caption("AI Reconciliation")
    st.markdown("---")

    st.markdown(
        "<p style='font-size:11px; color:#64748B; font-weight:bold;margin-bottom:5px;'>PROCESO</p>",
        unsafe_allow_html=True,
    )
    menu_proceso = st.radio(
        "Proceso",
        [
            "Dashboard",
            "Importar",
            "Conciliación",
            "Reglas",
            "Partidas pendientes",
            "Historial",
        ],
        label_visibility="collapsed",
    )

    st.markdown("---")
    st.markdown(
        "<p style='font-size:11px; color:#64748B; font-weight:bold;margin-bottom:5px;'>ANÁLISIS & TASAS</p>",
        unsafe_allow_html=True,
    )
    modulo_analisis = st.selectbox(
        "Análisis", ["Resumen histórico", "Gestión de Tasas BCV / Multidivisa"]
    )

    st.markdown("---")
    st.markdown(
        "<p style='font-size:11px; color:#64748B; font-weight:bold;margin-bottom:5px;'>ARCHIVOS & MAESTROS</p>",
        unsafe_allow_html=True,
    )
    modulo_maestros = st.selectbox(
        "Maestros", ["Ninguno", "Bancos y Cuentas", "Estados de cuenta"]
    )

    st.markdown("---")
    st.caption("Soporte técnico • Configuración")


# =============================================================================
# 6. MÓDULOS Y VISTAS DE LA APLICACIÓN
# =============================================================================

if modulo_analisis == "Gestión de Tasas BCV / Multidivisa":
    st.title("💱 Tasas de Cambio BCV y Conversión Multidivisa")
    st.markdown("Cargue masivamente las tasas oficiales desde Excel y realice conversiones seguras.")

    archivo_tasas = st.file_uploader("Seleccione el archivo Excel de Tasas", type=["xlsx", "xls"])
    if archivo_tasas is not None:
        try:
            df_t = pd.read_excel(archivo_tasas)
            st.write("🔍 **Vista previa de tasas:**")
            st.dataframe(df_t.head())

            cols_t = list(df_t.columns)
            col_f_t = st.selectbox("Columna de Fecha", cols_t)
            col_u_t = st.selectbox("Columna Tasa USD", ["Ninguna"] + cols_t)
            col_e_t = st.selectbox("Columna Tasa EUR", ["Ninguna"] + cols_t)

            map_monedas = {}
            if col_u_t != "Ninguna":
                map_monedas["USD"] = col_u_t
            if col_e_t != "Ninguna":
                map_monedas["EUR"] = col_e_t

            if st.button("Procesar y Guardar Tasas"):
                filas_t = df_t.to_dict(orient="records")
                errores_t = st.session_state.tabla_tasas.cargar_filas(
                    filas_t, col_f_t, map_monedas, tipo="BCV", fuente=archivo_tasas.name
                )
                if errores_t:
                    st.warning(f"Se omitieron {len(errores_t)} filas con errores de formato.")
                else:
                    st.success("🎉 ¡Todas las tasas se cargaron correctamente al sistema!")
        except Exception as ex:
            st.error(f"Error al leer el archivo de tasas: {ex}")

elif modulo_maestros == "Bancos y Cuentas":
    st.title("🏦 Maestros: Bancos y Cuentas")
    st.markdown("Gestión de cuentas bancarias corporativas agrupadas por institución.")

    with st.form("form_cuenta"):
        banco_input = st.selectbox("Banco", ["Banco Nacional", "Banco Regional", "Banco Mercantil"])
        nro_cta = st.text_input("Número de Cuenta")
        moneda_cta = st.selectbox("Moneda", ["VES", "USD", "EUR"])
        obs_cta = st.text_area("Observaciones (Sucursal, tipo, titular)")

        if st.form_submit_button("Registrar Cuenta"):
            if nro_cta.strip():
                nro_limpio = "".join(filter(str.isalnum, nro_cta)).upper()
                existentes = [
                    "".join(filter(str.isalnum, c)).upper()
                    for c in st.session_state.cuentas_bancarias.get(banco_input, [])
                ]
                if nro_limpio in existentes:
                    st.error("⚠️ Ya existe una cuenta registrada con este número en el sistema.")
                else:
                    st.session_state.cuentas_bancarias.setdefault(banco_input, []).append(nro_cta)
                    st.success(f"¡Cuenta {nro_cta} ({moneda_cta}) registrada con éxito para {banco_input}!")
            else:
                st.warning("Debe ingresar un número de cuenta válido.")

    st.markdown("---")
    st.subheader("Cuentas Activas")
    for bnk, ctas in st.session_state.cuentas_bancarias.items():
        st.write(f"**{bnk}**")
        for c in ctas:
            st.info(f"Número: {c}")

elif menu_proceso == "Dashboard":
    st.title("Dashboard")
    st.markdown(
        "<p style='color: #64748B; margin-top: -15px;'>Resumen general de conciliaciones bancarias</p>",
        unsafe_allow_html=True,
    )
    st.success(
        "🚀 **Sistema de Conciliación Inteligente Activo** — Bienvenido al "
        "entorno de trabajo multimoneda. Utilice el menú lateral para iniciar "
        "la importación o consultar las tasas.",
        icon="💼",
    )
    res = st.session_state.resultado_conciliacion
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
        st.metric(
            label="Partidas abiertas",
            value=len(res.banco_pendiente) if res else len(st.session_state.movs_banco),
        )
        st.markdown("</div>", unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
        st.metric(label="Conciliaciones cerradas", value="1" if res else "0")
        st.markdown("</div>", unsafe_allow_html=True)
    with c3:
        st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
        st.metric(
            label="Movimientos conciliados",
            value=len(res.pares) if res else 0,
        )
        st.markdown("</div>", unsafe_allow_html=True)
    with c4:
        st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
        st.metric(label="Diferencias registradas", value="0")
        st.markdown("</div>", unsafe_allow_html=True)

    col_izq, col_der = st.columns([2, 1])
    with col_izq:
        st.markdown("### Conciliaciones recientes")
        with st.container():
            st.markdown(
                '<div class="dashboard-card" style="text-align: center; padding: 30px;">',
                unsafe_allow_html=True,
            )
            if not res:
                st.markdown(
                    "📄<br>**No hay conciliaciones cerradas**<br><span style='color: #64748B; font-size: 14px;'>Importá un extracto y realizá tu primera conciliación.</span>",
                    unsafe_allow_html=True,
                )
            else:
                st.success("✅ Conciliación procesada correctamente. Revise el detalle en el menú **Conciliación**.")
            st.markdown("</div>", unsafe_allow_html=True)

    with col_der:
        st.markdown("### Acciones rápidas")
        with st.container():
            st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
            if st.button("📂 Importar extracto o mayor", use_container_width=True):
                pass
            if st.button("⚡ Iniciar conciliación", use_container_width=True):
                pass
            st.markdown("</div>", unsafe_allow_html=True)

elif menu_proceso == "Importar":
    st.title("📂 Importar Extractos y Registros Contables")
    st.markdown("Cargue archivos en formato Excel o CSV para procesar los movimientos automáticamente.")

    tipo_origen = st.selectbox(
        "Seleccione el tipo de archivo",
        ["Extracto Bancario", "Registro Contable (Mayor)"],
    )
    archivo_subido = st.file_uploader("Seleccione el archivo", type=["xlsx", "xls", "csv"])

    if archivo_subido is not None:
        try:
            df_imp = pd.read_excel(archivo_subido)
            st.write("🔍 **Vista previa de los datos:**")
            st.dataframe(df_imp.head())

            st.subheader("⚙️ Mapeo de Columnas")
            cols_disponibles = list(df_imp.columns)
            c_fec = st.selectbox("Columna de Fecha", cols_disponibles)
            c_desc = st.selectbox("Columna de Descripción", ["Ninguna"] + cols_disponibles, index=1)
            c_ref = st.selectbox("Columna de Referencia", ["Ninguna"] + cols_disponibles, index=2)
            esquema_monto = st.radio(
                "Estructura de Montos",
                [
                    "Columna única (Monto firmado)",
                    "Columnas separadas (Débito / Crédito)",
                ],
            )

            mapeo = {"fecha": c_fec}
            if c_desc != "Ninguna":
                mapeo["descripcion"] = c_desc
            if c_ref != "Ninguna":
                mapeo["referencia"] = c_ref

            if esquema_monto == "Columna única (Monto firmado)":
                mapeo["monto"] = st.selectbox("Columna de Monto", cols_disponibles)
            else:
                mapeo["entrada"] = st.selectbox("Columna Entrada (Crédito / Debe)", cols_disponibles)
                mapeo["salida"] = st.selectbox("Columna Salida (Débito / Haber)", cols_disponibles)

            if st.button("Procesar e Importar al Sistema"):
                origen_val = "banco" if "Extracto" in tipo_origen else "libros"
                resultado_carga = cargar_movimientos(
                    df_imp.to_dict(orient="records"), origen_val, mapeo
                )

                st.success(f"¡Se importaron {len(resultado_carga.movimientos)} movimientos con éxito!")

                if resultado_carga.errores:
                    st.warning(f"Se detectaron {len(resultado_carga.errores)} filas con advertencias omitidas.")
                    st.dataframe(pd.DataFrame([vars(e) for e in resultado_carga.errores]))

                if origen_val == "banco":
                    st.session_state.movs_banco = resultado_carga.movimientos
                else:
                    st.session_state.movs_libros = resultado_carga.movimientos

        except Exception as err:
            st.error(f"Error al procesar el archivo cargado: {err}")

elif menu_proceso == "Reglas":
    st.title("📋 Configuración y Criterios del Motor de Conciliación")
    st.markdown("Detalle automatizado de las reglas contables y de cruce aplicadas por el sistema en Venezuela.")

    st.markdown("""
    ### 1. Aislamiento Previo de Asientos Anulados en Libros (Anti-Falsos Positivos)
    * **Condición:** Antes de cruzar nada con el banco, el motor escanea el mayor contable buscando parejas de movimientos con montos exactos de signo opuesto ($+X$ y $-X$) que compartan la misma referencia (`Nº documento`) y descripción/contrapartida.
    * **Propósito:** Neutralizar las anulaciones internas desde la raíz para que **nunca consuman movimientos bancarios reales** ni inflen los pendientes. Quedan etiquetadas como **`Asiento Anulado`**.

    ### 2. Clasificación Automática de Naturaleza Contable (Cuentas Propias vs Operativo Real)
    * **Transferencias entre Cuentas Propias (Fondeos):** Identifica traslados entre bancos de la empresa (`BANCO BANESCO`, `BNC`, `MERCANTIL`, etc.) para no confundirlos con ingresos por ventas.
    * **Ingresos por Ventas / Cobranzas de Clientes:** Entradas provenientes de clientes y distribuidoras comerciales.
    * **Ingresos Bancarios / Intereses y Devoluciones:** Distingue abonos por intereses bancarios y reintegros o devoluciones de proveedores por errores de pago.

    ### 3. Cruce Exacto (Regla C1) y Tolerancia de Fechas (Regla C2)
    * **C1:** Coincidencia exacta de referencia, monto idéntico y fecha.
    * **C2:** Mismo monto y signo dentro de la ventana de días hábiles configurada.
    """)

elif menu_proceso == "Conciliación":
    st.title("⚡ Motor de Cruce, Distribución Porcentual y Asistente IA")
    st.markdown("Cruce inteligente con aislamiento de anulados, peso porcentual por partida y distinción de cuentas propias.")

    st.info(
        f"📊 Movimientos cargados en memoria — Banco: {len(st.session_state.movs_banco)} | Libros: {len(st.session_state.movs_libros)}"
    )

    tolerancia_c2 = st.slider("Tolerancia de días para regla C2", min_value=0, max_value=7, value=3)

    if st.button("Ejecutar Conciliación Completa", type="primary"):
        if not st.session_state.movs_banco or not st.session_state.movs_libros:
            st.warning("⚠️ Debe importar tanto el extracto bancario como el registro contable antes de ejecutar la conciliación.")
        else:
            try:
                configuracion = Config(tolerancia_dias_c2=tolerancia_c2)
                res_cruces = conciliar(
                    st.session_state.movs_banco,
                    st.session_state.movs_libros,
                    configuracion,
                )
                st.session_state.resultado_conciliacion = res_cruces
                st.success("🎉 ¡Conciliación ejecutada con éxito!")
            except Exception as e:
                st.error(f"Error durante la ejecución del motor: {e}")

    if st.session_state.resultado_conciliacion:
        res = st.session_state.resultado_conciliacion
        datos_resumen = res.resumen()
        df_reporte_actual = construir_dataframe_consolidado(res, st.session_state.tabla_tasas)
        df_res_nat, df_dist_egresos, df_dist_ingresos = generar_tablas_distribucion(df_reporte_actual)

        st.markdown("---")
        st.subheader("📈 Resultados del Cruce")

        kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
        kpi1.metric("Pares Cerrados (C1/C2)", datos_resumen.get("pares_cerrados", 0))
        kpi2.metric("Sugerencias (C3/C4)", datos_resumen.get("sugerencias", 0))
        kpi3.metric("Banco Pendiente", datos_resumen.get("banco_pendiente", 0))
        kpi4.metric("Libros Pendiente", datos_resumen.get("libros_pendiente", 0))
        kpi5.metric("Asientos Anulados", datos_resumen.get("libros_anulados", 0))

        if st.button("💾 Guardar Conciliación Actual en Historial"):
            registro_historial = {
                "fecha_cierre": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "pares_cerrados": datos_resumen.get("pares_cerrados", 0),
                "banco_pendiente": datos_resumen.get("banco_pendiente", 0),
                "libros_pendiente": datos_resumen.get("libros_pendiente", 0),
                "libros_anulados": datos_resumen.get("libros_anulados", 0),
            }
            guardar_en_historial(registro_historial)
            st.success("¡Conciliación guardada permanentemente en el Historial del sistema!")

        # =====================================================================
        # EXPORTAR REPORTE CONSOLIDADO EN EXCEL (CON % POR PARTIDA Y PESTAÑAS TOP)
        # =====================================================================
        st.markdown("---")
        st.subheader("📥 Exportar Reporte Consolidado + Distribución Porcentual en Excel")
        st.caption(
            "Incluye en el mismo Excel: 1) Conciliación Consolidada con el % de cada partida sobre el Total de Ingresos y Egresos, "
            "2) Resumen por Naturaleza Contable (Cuentas Propias vs Ventas), 3) Top Distribución Egresos y 4) Top Distribución Ingresos."
        )

        # Agregar fila final de Totales a una copia de la hoja principal para el Excel
        df_hoja_principal = df_reporte_actual.copy()
        if not df_hoja_principal.empty:
            mask_no_anul = df_hoja_principal["Marca Conciliación"] != "Asiento Anulado"
            tot_ing_m = df_hoja_principal.loc[mask_no_anul & (df_hoja_principal["Mayor - Monto"] > 0), "Mayor - Monto"].sum()
            tot_egr_m = df_hoja_principal.loc[mask_no_anul & (df_hoja_principal["Mayor - Monto"] < 0), "Mayor - Monto"].sum()
            tot_net_m = tot_ing_m + tot_egr_m
            tot_bco_m = df_hoja_principal["Banco - Monto"].sum()
            tot_usd_m = df_hoja_principal.loc[mask_no_anul, "Monto Convertido USD"].sum()

            filas_totales = pd.DataFrame([
                {
                    "Banco - Fecha": "",
                    "Banco - Descripción": "TOTAL INGRESOS MAYOR (SIN ANULADOS)",
                    "Banco - Referencia": "",
                    "Banco - Monto": 0.0,
                    "Mayor - Fecha": "",
                    "Mayor - Descripción": "TOTAL INGRESOS MAYOR",
                    "Mayor - Referencia": "",
                    "Mayor - Monto": round(tot_ing_m, 2),
                    "Marca Conciliación": "TOTAL INGRESOS",
                    "Naturaleza Contable": "Total Entradas Mayor",
                    "Tasa BCV USD": 0.0,
                    "Monto Convertido USD": 0.0,
                    "% s/ Total Ingresos Mayor": 100.0 if tot_ing_m > 0 else 0.0,
                    "% s/ Total Egresos Mayor": 0.0,
                },
                {
                    "Banco - Fecha": "",
                    "Banco - Descripción": "TOTAL EGRESOS MAYOR (SIN ANULADOS)",
                    "Banco - Referencia": "",
                    "Banco - Monto": 0.0,
                    "Mayor - Fecha": "",
                    "Mayor - Descripción": "TOTAL EGRESOS MAYOR",
                    "Mayor - Referencia": "",
                    "Mayor - Monto": round(tot_egr_m, 2),
                    "Marca Conciliación": "TOTAL EGRESOS",
                    "Naturaleza Contable": "Total Salidas Mayor",
                    "Tasa BCV USD": 0.0,
                    "Monto Convertido USD": 0.0,
                    "% s/ Total Ingresos Mayor": 0.0,
                    "% s/ Total Egresos Mayor": 100.0 if tot_egr_m < 0 else 0.0,
                },
                {
                    "Banco - Fecha": "",
                    "Banco - Descripción": "FLUJO NETO / SALDO MAYOR (INGRESOS - EGRESOS)",
                    "Banco - Referencia": "",
                    "Banco - Monto": round(tot_bco_m, 2),
                    "Mayor - Fecha": "",
                    "Mayor - Descripción": "FLUJO NETO MAYOR",
                    "Mayor - Referencia": "",
                    "Mayor - Monto": round(tot_net_m, 2),
                    "Marca Conciliación": "FLUJO NETO",
                    "Naturaleza Contable": "Resultado Neto Periodo",
                    "Tasa BCV USD": 0.0,
                    "Monto Convertido USD": round(tot_usd_m, 2),
                    "% s/ Total Ingresos Mayor": 0.0,
                    "% s/ Total Egresos Mayor": 0.0,
                },
            ])
            df_hoja_principal = pd.concat([df_hoja_principal, filas_totales], ignore_index=True)

        output_excel = io.BytesIO()
        with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
            df_hoja_principal.to_excel(writer, index=False, sheet_name="Conciliacion_Consolidada")
            if not df_res_nat.empty:
                df_res_nat.to_excel(writer, index=False, sheet_name="Resumen_Naturaleza")
            if not df_dist_ingresos.empty:
                df_dist_ingresos.to_excel(writer, index=False, sheet_name="Top_Distribucion_Ingresos")
            if not df_dist_egresos.empty:
                df_dist_egresos.to_excel(writer, index=False, sheet_name="Top_Distribucion_Egresos")
            if st.session_state.ultimo_analisis_ia:
                df_ia_hoja = pd.DataFrame([
                    {
                        "Consulta Realizada": st.session_state.ultima_pregunta_ia,
                        "Análisis del Asistente IA": st.session_state.ultimo_analisis_ia,
                    }
                ])
                df_ia_hoja.to_excel(writer, index=False, sheet_name="Analisis_IA")
        processed_data = output_excel.getvalue()

        st.download_button(
            label="📥 Descargar Reporte Consolidado + Distribución % en Excel (.xlsx)",
            data=processed_data,
            file_name="Reporte_Conciliacion_Consolidado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Vista previa interactiva de las tablas de distribución porcentual
        with st.expander("📊 Ver Tablas de Distribución Porcentual (Ingresos vs Cuentas Propias y Egresos)", expanded=False):
            st.markdown("#### 1. Resumen por Naturaleza Contable")
            st.dataframe(df_res_nat, use_container_width=True)
            col_ti, col_te = st.columns(2)
            with col_ti:
                st.markdown("#### 2. Distribución % de Ingresos (Mayor)")
                st.dataframe(df_dist_ingresos, use_container_width=True)
            with col_te:
                st.markdown("#### 3. Distribución % de Egresos (Mayor)")
                st.dataframe(df_dist_egresos, use_container_width=True)

        # =====================================================================
        # ASISTENTE DE IA INTEGRADO + DESCARGA EN WORD (.DOCX) Y EXCEL (.XLSX)
        # =====================================================================
        st.markdown("---")
        with st.expander("🤖 Asistente IA para Análisis de Conciliación (Descargable en Word y Excel)", expanded=True):
            st.markdown(
                "Realice preguntas sobre el **Top de movimientos mayores**, **distribución porcentual** "
                "o **análisis de ingresos reales vs transferencias entre cuentas propias**."
            )
            pregunta_ai = st.text_input(
                "Consulta para el Asistente IA:",
                placeholder="Ejemplo: Dame un análisis general con el Top 10 de egresos e ingresos, separando transferencias entre cuentas propias de las cobranzas a clientes",
            )

            if st.button("Consultar Asistente IA", type="primary"):
                if not pregunta_ai.strip():
                    st.warning("Por favor ingrese una pregunta para el asistente.")
                else:
                    if OpenAI is None:
                        st.error("La librería `openai` no está instalada o configurada.")
                    else:
                        try:
                            client = OpenAI(
                                api_key=st.secrets.get("OPENAI_API_KEY", ""),
                                base_url=st.secrets.get("OPENAI_BASE_URL", "https://openrouter.ai/api/v1"),
                            )
                            contexto_real = construir_contexto_ia_real(df_reporte_actual, res)

                            prompt_sistema = (
                                "Eres un auditor financiero y analista de tesorería experto en Venezuela. "
                                "REGLAS ESTRICTAS E INQUEBRANTABLES: "
                                "1) Usa ÚNICAMENTE los textos literales exactos de 'Mayor - Descripción' y los montos reales provistos. Prohibido inventar categorías o cifras. "
                                "2) DISTINCIÓN CONTABLE CRÍTICA EN INGRESOS: Cuando en 'Mayor - Descripción' aparezca el nombre de un banco "
                                "(ej. 'BANCO BANESCO', 'BANCO NACIONAL DE CREDITO-395', etc.), NO es una venta ni un cliente: es una "
                                "'Transferencia entre Cuentas Propias / Fondeo Interno'. Por lo tanto, al analizar los ingresos, presenta siempre por separado: "
                                "a) Las Transferencias entre Cuentas Propias (fondeos bancarios), "
                                "b) Los Ingresos Reales por Ventas / Cobranzas a Clientes (ej. Inversiones Vetglobal, Agrodistribuidora, etc., indicando tanto su % sobre el total de entradas como su % sobre las ventas a clientes), "
                                "c) Los Ingresos Bancarios / Intereses (ej. 'INTERESES BANCARIOS') y posibles Devoluciones/Reintegros de proveedores. "
                                "Nunca recomiendes 'diversificar ingresos porque dependen de Banco Banesco', sino explica que 'BANCO BANESCO' corresponde a fondeos o traslados internos entre cuentas propias. "
                                "3) En las tablas Markdown que generes, incluye siempre Monto en VES, Monto en USD (si aplica) y el Porcentaje (%)."
                            )

                            response = client.chat.completions.create(
                                model="openai/gpt-4o-mini",
                                temperature=0.0,
                                messages=[
                                    {"role": "system", "content": prompt_sistema},
                                    {
                                        "role": "user",
                                        "content": f"DATOS REALES DE LA CONCILIACIÓN ACTUAL:\n{contexto_real}\n\nPREGUNTA DEL ANALISTA:\n{pregunta_ai}",
                                    },
                                ],
                            )
                            respuesta_texto = response.choices[0].message.content
                            st.session_state.ultimo_analisis_ia = respuesta_texto
                            st.session_state.ultima_pregunta_ia = pregunta_ai
                        except Exception as ai_err:
                            st.error(f"Error al conectar con la API (Verifique su API Key y base_url en st.secrets): {ai_err}")

            if st.session_state.ultimo_analisis_ia:
                st.markdown("#### 📄 Resultado del Análisis IA")
                st.markdown(st.session_state.ultimo_analisis_ia)

                # Preparar archivo Word (.docx) nativo con formato ejecutivo y tablas
                bytes_docx = generar_word_analisis_ia(
                    st.session_state.ultima_pregunta_ia,
                    st.session_state.ultimo_analisis_ia,
                    df_res_nat,
                    df_dist_egresos,
                    df_dist_ingresos,
                )

                # Preparar archivo Excel (.xlsx) con el análisis y todas las tablas numéricas de distribución
                out_ia_xlsx = io.BytesIO()
                with pd.ExcelWriter(out_ia_xlsx, engine="openpyxl") as writer_ia:
                    if not df_res_nat.empty:
                        df_res_nat.to_excel(writer_ia, index=False, sheet_name="Resumen_Naturaleza")
                    if not df_dist_ingresos.empty:
                        df_dist_ingresos.to_excel(writer_ia, index=False, sheet_name="Top_Ingresos_Porcentaje")
                    if not df_dist_egresos.empty:
                        df_dist_egresos.to_excel(writer_ia, index=False, sheet_name="Top_Egresos_Porcentaje")
                    pd.DataFrame([
                        {
                            "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
                            "Consulta": st.session_state.ultima_pregunta_ia,
                            "Respuesta IA": st.session_state.ultimo_analisis_ia,
                        }
                    ]).to_excel(writer_ia, index=False, sheet_name="Texto_Analisis_IA")

                col_w, col_x, col_m = st.columns(3)
                with col_w:
                    st.download_button(
                        label="📝 Descargar Resumen IA en Word (.docx)",
                        data=bytes_docx,
                        file_name="Resumen_Ejecutivo_IA_Conciliacion.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )
                with col_x:
                    st.download_button(
                        label="📊 Descargar Tablas Top y % en Excel (.xlsx)",
                        data=out_ia_xlsx.getvalue(),
                        file_name="Distribucion_Porcentual_y_Top_IA.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                with col_m:
                    st.download_button(
                        label="📄 Descargar Texto (.md)",
                        data=st.session_state.ultimo_analisis_ia.encode("utf-8"),
                        file_name="Analisis_IA_Conciliacion.md",
                        mime="text/markdown",
                        use_container_width=True,
                    )

        if res.sugerencias:
            st.markdown("### 💡 Sugerencias pendientes de aprobación (C3 / C4)")
            for idx, sug in enumerate(res.sugerencias):
                with st.expander(f"Sugerencia {idx+1} — Regla: {sug.regla} | Motivo: {sug.motivo}"):
                    col_b, col_l = st.columns(2)
                    with col_b:
                        st.write("**Movimiento(s) Banco:**")
                        for mb in sug.banco:
                            st.text(f"Fecha: {mb.fecha} | Monto: {mb.monto} | {mb.descripcion}")
                    with col_l:
                        st.write("**Movimiento(s) Libros:**")
                        for ml in sug.libros:
                            st.text(f"Fecha: {ml.fecha} | Monto: {ml.monto} | {ml.descripcion}")

                    if st.button(f"Aprobar sugerencia {idx+1}", key=f"aprobar_{idx}"):
                        res.aprobar(sug)
                        st.success("¡Sugerencia aprobada con éxito y movida a pares cerrados!")
                        st.rerun()

elif menu_proceso == "Partidas pendientes":
    st.title("📌 Partidas Abiertas, Pendientes y Asientos Anulados")
    st.markdown("Gestión detallada de partidas sin resolver y asientos contables aislados por anulación.")
    if st.session_state.resultado_conciliacion:
        res = st.session_state.resultado_conciliacion
        col_pb, col_pl, col_pa = st.columns(3)
        with col_pb:
            st.subheader("Pendientes de Banco")
            if res.banco_pendiente:
                df_pb = pd.DataFrame([vars(m) for m in res.banco_pendiente])
                st.dataframe(df_pb[["id", "fecha", "monto", "descripcion", "referencia"]])
            else:
                st.info("Sin partidas pendientes en banco.")
        with col_pl:
            st.subheader("Pendientes de Libros")
            if res.libros_pendiente:
                df_pl = pd.DataFrame([vars(m) for m in res.libros_pendiente])
                st.dataframe(df_pl[["id", "fecha", "monto", "descripcion", "referencia"]])
            else:
                st.info("Sin partidas pendientes en libros.")
        with col_pa:
            st.subheader("Asientos Anulados")
            if getattr(res, "libros_anulados", None):
                lista_anulados = [
                    {
                        "id": m.id,
                        "fecha": m.fecha,
                        "monto": m.monto,
                        "descripcion": m.descripcion,
                        "referencia": m.referencia,
                        "marca": marca,
                    }
                    for m, marca in res.libros_anulados
                ]
                df_pa = pd.DataFrame(lista_anulados)
                st.dataframe(df_pa[["id", "fecha", "monto", "descripcion", "referencia", "marca"]])
            else:
                st.info("Sin asientos anulados detectados.")
    else:
        st.info("Ejecute una conciliación primero para visualizar las partidas.")

elif menu_proceso == "Historial":
    st.title("📜 Historial Permanente de Conciliaciones Cerradas")
    st.markdown("Registro histórico persistente almacenado para auditoría, seguimiento y control.")

    historial_disco = cargar_historial_disco()
    if historial_disco:
        st.success(f"Se encontraron {len(historial_disco)} conciliación(es) guardada(s) en el sistema.")
        for idx, reg in enumerate(historial_disco):
            with st.expander(f"Cierre del {reg.get('fecha_cierre', 'Sin fecha')} — Pares cerrados: {reg.get('pares_cerrados', 0)}"):
                st.write(f"* **Pares cerrados (C1/C2):** {reg.get('pares_cerrados', 0)}")
                st.write(f"* **Banco pendiente:** {reg.get('banco_pendiente', 0)}")
                st.write(f"* **Libros pendiente:** {reg.get('libros_pendiente', 0)}")
                st.write(f"* **Asientos anulados aislados:** {reg.get('libros_anulados', 0)}")
    else:
        st.info("No hay conciliaciones guardadas en el historial todavía. Ejecute una conciliación y haga clic en 'Guardar Conciliación Actual en Historial'.")

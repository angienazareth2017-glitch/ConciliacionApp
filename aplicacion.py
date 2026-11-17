"""
aplicacion.py - Aplicación web principal en Streamlit para el 
Sistema de Conciliación y Tesorería Multidivisa con:
- Navegación interactiva real en Dashboard (Acciones Rápidas y Conciliaciones Recientes accionables).
- Historial Persistente Completo: guarda el consolidado y los movimientos base, permitiendo
  ver el consolidado, editarlo en vivo (st.data_editor), descargarlo en Excel o retomar/reabrir
  cualquier conciliación anterior en el motor.
- Aislamiento previo de Asientos Anulados y clasificación de Naturaleza Contable (Cuentas Propias vs Ventas).
- Exportación del Resumen IA en Word (.docx) y Excel (.xlsx) con distribución porcentual (%).
"""
from __future__ import annotations

import io
import json
import os
import re
import zipfile
from datetime import datetime
from decimal import Decimal
from xml.sax.saxutils import escape as xml_escape

import pandas as pd
import streamlit as st

# Importación de motores lógicos locales
from conciliar import (
    Config,
    ErrorDeParseo,
    Movimiento,
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
# 1. CONFIGURACIÓN INICIAL DE LA PÁGINA Y PERSISTENCIA COMPLETA
# =============================================================================
st.set_page_config(
    page_title="Tesora Pro - Sistema de Conciliación y Tesorería",
    page_icon="📊",
    layout="wide",
)

HISTORIAL_FILE = "historial_conciliaciones.json"


def cargar_historial_disco() -> list[dict]:
    if os.path.exists(HISTORIAL_FILE):
        try:
            with open(HISTORIAL_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data if isinstance(data, list) else []
        except Exception:
            return []
    return []


def guardar_lista_historial_disco(historial: list[dict]) -> None:
    try:
        with open(HISTORIAL_FILE, "w", encoding="utf-8") as f:
            json.dump(historial, f, ensure_ascii=False, indent=2, default=str)
    except Exception:
        pass


def guardar_en_historial(resumen_dict: dict) -> None:
    historial = cargar_historial_disco()
    historial.insert(0, resumen_dict)
    guardar_lista_historial_disco(historial)


def serializar_movimiento(m: Movimiento) -> dict:
    return {
        "id": str(m.id),
        "origen": str(m.origen),
        "fecha": m.fecha.isoformat() if hasattr(m.fecha, "isoformat") else str(m.fecha),
        "monto": str(m.monto),
        "descripcion": str(m.descripcion),
        "referencia": str(m.referencia),
        "moneda": str(getattr(m, "moneda", "VES")),
    }


def deserializar_movimiento(d: dict) -> Movimiento:
    return Movimiento(
        id=str(d.get("id", "")),
        origen=str(d.get("origen", "banco")),
        fecha=d.get("fecha", "2026-01-01"),
        monto=Decimal(str(d.get("monto", "0"))),
        descripcion=str(d.get("descripcion", "")),
        referencia=str(d.get("referencia", "")),
        moneda=str(d.get("moneda", "VES")),
    )


def ir_a_menu(destino: str) -> None:
    """Cambia de pestaña de forma segura (tanto en callbacks on_click como antes de st.rerun)."""
    st.session_state._nav_pending = {
        "menu_proceso": destino,
        "modulo_analisis": "Resumen histórico",
        "modulo_maestros": "Ninguno",
    }


def retomar_conciliacion_historica(reg: dict) -> bool:
    """
    Restaura en memoria una conciliación guardada en el historial para poder
    editarla, aprobar sugerencias o volver a cruzarla en el motor.
    """
    try:
        movs_b_raw = reg.get("movs_banco_raw", [])
        movs_l_raw = reg.get("movs_libros_raw", [])
        if movs_b_raw and movs_l_raw:
            st.session_state.movs_banco = [deserializar_movimiento(x) for x in movs_b_raw]
            st.session_state.movs_libros = [deserializar_movimiento(x) for x in movs_l_raw]
            tol = int(reg.get("tolerancia_c2", 3))
            cfg = Config(tolerancia_dias_c2=tol)
            st.session_state.resultado_conciliacion = conciliar(
                st.session_state.movs_banco,
                st.session_state.movs_libros,
                cfg,
            )
            st.session_state.ultimo_analisis_ia = reg.get("analisis_ia", "")
            st.session_state.ultima_pregunta_ia = reg.get("pregunta_ia", "")
            return True
    except Exception:
        pass
    return False


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
                "Banco - Fecha": str(mb.fecha) if mb else "",
                "Banco - Descripción": d_bco,
                "Banco - Referencia": mb.referencia if mb else "",
                "Banco - Monto": m_bco,
                "Mayor - Fecha": str(ml.fecha) if ml else "",
                "Mayor - Descripción": d_may,
                "Mayor - Referencia": ml.referencia if ml else "",
                "Mayor - Monto": m_may,
                "Marca Conciliación": marca,
                "Naturaleza Contable": clasificar_naturaleza_contable(d_may, d_bco, m_may, m_bco, marca),
                "Tasa BCV USD": tasa_usd_val,
                "Monto Convertido USD": monto_usd,
                "Observaciones / Ajustes": "",
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
            "Banco - Fecha": str(mb.fecha),
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
            "Observaciones / Ajustes": "",
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
            "Mayor - Fecha": str(ml.fecha),
            "Mayor - Descripción": ml.descripcion,
            "Mayor - Referencia": ml.referencia,
            "Mayor - Monto": m_may,
            "Marca Conciliación": marca,
            "Naturaleza Contable": clasificar_naturaleza_contable(ml.descripcion, "", m_may, 0.0, marca),
            "Tasa BCV USD": tasa_usd_val,
            "Monto Convertido USD": monto_usd,
            "Observaciones / Ajustes": "",
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
            "Mayor - Fecha": str(ml.fecha),
            "Mayor - Descripción": ml.descripcion,
            "Mayor - Referencia": ml.referencia,
            "Mayor - Monto": m_may,
            "Marca Conciliación": marca_anulado,
            "Naturaleza Contable": "Asiento Anulado (Neutralizado)",
            "Tasa BCV USD": tasa_usd_val,
            "Monto Convertido USD": monto_usd,
            "Observaciones / Ajustes": "",
        })

    df = pd.DataFrame(filas_excel)
    if df.empty:
        return df

    return recalcular_porcentajes_consolidado(df)


def recalcular_porcentajes_consolidado(df: pd.DataFrame) -> pd.DataFrame:
    """Recalcula los porcentajes sobre Total Ingresos y Total Egresos en el DataFrame consolidado."""
    if df.empty:
        return df
    df = df.copy()
    mask_activo = ~df["Marca Conciliación"].isin(["Asiento Anulado", "TOTAL INGRESOS", "TOTAL EGRESOS", "FLUJO NETO"])
    total_ingresos_mayor = pd.to_numeric(df.loc[mask_activo & (pd.to_numeric(df["Mayor - Monto"], errors="coerce") > 0), "Mayor - Monto"], errors="coerce").sum()
    total_egresos_mayor = pd.to_numeric(df.loc[mask_activo & (pd.to_numeric(df["Mayor - Monto"], errors="coerce") < 0), "Mayor - Monto"], errors="coerce").abs().sum()

    pct_ing_list = []
    pct_egr_list = []
    for _, row in df.iterrows():
        if row["Marca Conciliación"] in ("Asiento Anulado", "TOTAL INGRESOS", "TOTAL EGRESOS", "FLUJO NETO"):
            pct_ing_list.append(0.0)
            pct_egr_list.append(0.0)
            continue
        try:
            val_m = float(row["Mayor - Monto"])
        except Exception:
            val_m = 0.0
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
    if df_rep.empty:
        vacio = pd.DataFrame()
        return vacio, vacio, vacio

    df_act = df_rep[
        (df_rep["Mayor - Descripción"].astype(str).str.strip() != "")
        & (~df_rep["Marca Conciliación"].isin(["Asiento Anulado", "TOTAL INGRESOS", "TOTAL EGRESOS", "FLUJO NETO"]))
    ].copy()
    df_act["Mayor - Monto"] = pd.to_numeric(df_act["Mayor - Monto"], errors="coerce").fillna(0.0)
    df_act["Monto Convertido USD"] = pd.to_numeric(df_act["Monto Convertido USD"], errors="coerce").fillna(0.0)

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

    # 2. Tabla de Ingresos Agrupada
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


def empaquetar_excel_consolidado(df_base: pd.DataFrame, pregunta_ia: str = "", respuesta_ia: str = "") -> bytes:
    """Genera el archivo Excel multi-hoja completo a partir de un DataFrame consolidado."""
    df_rep = recalcular_porcentajes_consolidado(df_base)
    df_res_nat, df_dist_egresos, df_dist_ingresos = generar_tablas_distribucion(df_rep)

    df_hoja_principal = df_rep.copy()
    if not df_hoja_principal.empty:
        mask_no_anul = ~df_hoja_principal["Marca Conciliación"].isin(
            ["Asiento Anulado", "TOTAL INGRESOS", "TOTAL EGRESOS", "FLUJO NETO"]
        )
        df_sin_tot = df_hoja_principal[
            ~df_hoja_principal["Marca Conciliación"].isin(["TOTAL INGRESOS", "TOTAL EGRESOS", "FLUJO NETO"])
        ].copy()
        montos_m = pd.to_numeric(df_sin_tot["Mayor - Monto"], errors="coerce").fillna(0.0)
        tot_ing_m = montos_m[mask_no_anul & (montos_m > 0)].sum()
        tot_egr_m = montos_m[mask_no_anul & (montos_m < 0)].sum()
        tot_net_m = tot_ing_m + tot_egr_m
        tot_bco_m = pd.to_numeric(df_sin_tot["Banco - Monto"], errors="coerce").fillna(0.0).sum()
        tot_usd_m = pd.to_numeric(df_sin_tot.loc[mask_no_anul, "Monto Convertido USD"], errors="coerce").fillna(0.0).sum()

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
                "Observaciones / Ajustes": "",
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
                "Observaciones / Ajustes": "",
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
                "Observaciones / Ajustes": "",
                "% s/ Total Ingresos Mayor": 0.0,
                "% s/ Total Egresos Mayor": 0.0,
            },
        ])
        df_hoja_principal = pd.concat([df_sin_tot, filas_totales], ignore_index=True)

    output_excel = io.BytesIO()
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        df_hoja_principal.to_excel(writer, index=False, sheet_name="Conciliacion_Consolidada")
        if not df_res_nat.empty:
            df_res_nat.to_excel(writer, index=False, sheet_name="Resumen_Naturaleza")
        if not df_dist_ingresos.empty:
            df_dist_ingresos.to_excel(writer, index=False, sheet_name="Top_Distribucion_Ingresos")
        if not df_dist_egresos.empty:
            df_dist_egresos.to_excel(writer, index=False, sheet_name="Top_Distribucion_Egresos")
        if respuesta_ia:
            df_ia_hoja = pd.DataFrame([
                {
                    "Consulta Realizada": pregunta_ia,
                    "Análisis del Asistente IA": respuesta_ia,
                }
            ])
            df_ia_hoja.to_excel(writer, index=False, sheet_name="Analisis_IA")
    return output_excel.getvalue()


# =============================================================================
# 3. EXPORTADOR NATIVO A WORD (.DOCX) CON SOPORTE DE TABLAS Y FORMATO EJECUTIVO
# =============================================================================
def generar_word_analisis_ia(pregunta: str, respuesta_md: str, df_res_nat: pd.DataFrame, df_eg: pd.DataFrame, df_in: pd.DataFrame) -> bytes:
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
        xml_t.append("<w:tr>")
        for h in encabezados:
            h_esc = xml_escape(str(h))
            xml_t.append(
                f'<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="1B4332"/></w:tcPr>'
                f'<w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="19"/></w:rPr>'
                f'<w:t>{h_esc}</w:t></w:r></w:p></w:tc>'
            )
        xml_t.append("</w:tr>")
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

    parrafos_xml.append(_p_titulo("Tesora Pro — Informe Ejecutivo de Conciliación y Tesorería (IA)", 1))
    parrafos_xml.append(_p_texto(f"Fecha de emisión: {datetime.now().strftime('%Y-%m-%d %H:%M')}", negrita=True))
    if pregunta:
        parrafos_xml.append(_p_texto(f"Consulta analizada: {pregunta}", negrita=True))

    lineas_md = (respuesta_md or "").splitlines()
    buffer_tabla = []

    def _vaciar_buffer_tabla():
        nonlocal buffer_tabla
        if not buffer_tabla:
            return
        filas_limpias = []
        for l_tab in buffer_tabla:
            partes = [c.strip() for c in l_tab.strip().strip("|").split("|")]
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


# Estilos CSS de Alto Contraste: Azul Oscuro (#0A1E46), Blanco (#FFFFFF) y Negro (#0B0F19) + Tipografía Ampliada
st.markdown(
    """
    <style>
        /* Fondo general claro y texto negro nítido */
        .stApp {
            background-color: #F8FAFC;
            color: #0B0F19;
        }

        /* Títulos de sección en Azul Oscuro y Negro con tamaño apropiado */
        h1 {
            color: #0A1E46 !important;
            font-size: 32px !important;
            font-weight: 800 !important;
        }
        h2 {
            color: #0A1E46 !important;
            font-size: 25px !important;
            font-weight: 700 !important;
        }
        h3 {
            color: #0B0F19 !important;
            font-size: 21px !important;
            font-weight: 700 !important;
        }
        h4 {
            color: #0A1E46 !important;
            font-size: 18px !important;
            font-weight: 700 !important;
        }

        /* Tarjetas blancas con borde definido y sombra elegante */
        div[data-testid="stVerticalBlockBorderWrapper"] {
            background-color: #FFFFFF;
            border-radius: 12px;
            border: 1.5px solid #CBD5E1;
            box-shadow: 0 4px 12px rgba(10, 30, 70, 0.06);
            padding: 18px;
            margin-bottom: 14px;
        }

        /* Barra lateral (Sidebar) con contraste limpio */
        section[data-testid="stSidebar"] {
            background-color: #FFFFFF;
            border-right: 2px solid #E2E8F0;
        }
        section[data-testid="stSidebar"] .stRadio label p {
            font-size: 16px !important;
            font-weight: 600 !important;
            color: #0B0F19 !important;
        }

        /* Botones principales y secundarios en Azul Oscuro */
        .stButton button[kind="primary"] {
            background-color: #0A1E46 !important;
            color: #FFFFFF !important;
            border: 1.5px solid #0A1E46 !important;
            border-radius: 10px !important;
            font-size: 16px !important;
            font-weight: 700 !important;
            padding: 0.55rem 1.2rem !important;
            box-shadow: 0 2px 6px rgba(10, 30, 70, 0.25);
        }
        .stButton button[kind="primary"] p,
        .stButton button[kind="primary"] span {
            color: #FFFFFF !important;
            font-size: 16px !important;
            font-weight: 700 !important;
        }
        .stButton button[kind="primary"]:hover {
            background-color: #15357A !important;
            border-color: #15357A !important;
        }
        .stButton button[kind="secondary"] {
            background-color: #FFFFFF !important;
            color: #0A1E46 !important;
            border: 1.5px solid #0A1E46 !important;
            border-radius: 10px !important;
            font-size: 15px !important;
            font-weight: 700 !important;
            padding: 0.5rem 1rem !important;
        }
        .stButton button[kind="secondary"] p,
        .stButton button[kind="secondary"] span {
            color: #0A1E46 !important;
            font-size: 15px !important;
            font-weight: 700 !important;
        }
        .stButton button[kind="secondary"]:hover {
            background-color: #EFF6FF !important;
        }

        /* Botones de descarga en Azul Oscuro */
        .stDownloadButton button {
            background-color: #0A1E46 !important;
            color: #FFFFFF !important;
            border-radius: 10px !important;
            font-size: 15px !important;
            font-weight: 700 !important;
            border: none !important;
        }
        .stDownloadButton button p,
        .stDownloadButton button span {
            color: #FFFFFF !important;
            font-size: 15px !important;
            font-weight: 700 !important;
        }

        /* Banner Superior Azul Oscuro Corporativo */
        .tesora-banner {
            background: linear-gradient(135deg, #071636 0%, #0F2C67 60%, #153E85 100%);
            padding: 28px 32px;
            border-radius: 14px;
            margin-bottom: 24px;
            box-shadow: 0 8px 22px rgba(7, 22, 54, 0.22);
            border-left: 6px solid #38BDF8;
        }
        .tesora-banner .tesora-banner-title,
        .tesora-banner .tesora-banner-title * {
            color: #FFFFFF !important;
            font-size: 32px !important;
            font-weight: 800 !important;
            line-height: 1.25 !important;
            margin: 0 0 10px 0 !important;
            display: block;
            text-shadow: 0 2px 4px rgba(0, 0, 0, 0.25);
        }
        .tesora-banner .tesora-banner-sub,
        .tesora-banner .tesora-banner-sub * {
            color: #E0F2FE !important;
            font-size: 18px !important;
            font-weight: 500 !important;
            line-height: 1.45 !important;
            margin: 0 !important;
            display: block;
        }

        /* Tarjetas de Indicadores (Recuadros KPI) con fuente grande y clara */
        .tesora-kpi-card {
            background-color: #FFFFFF;
            border: 1.5px solid #CBD5E1;
            border-top: 5px solid #0A1E46;
            border-radius: 12px;
            padding: 20px 22px;
            box-shadow: 0 4px 12px rgba(10, 30, 70, 0.06);
            margin-bottom: 16px;
        }
        .tesora-kpi-card .tesora-kpi-label,
        .tesora-kpi-card .tesora-kpi-label * {
            color: #1E293B !important;
            font-size: 17px !important;
            font-weight: 700 !important;
            margin-bottom: 10px !important;
            line-height: 1.35 !important;
            display: block;
        }
        .tesora-kpi-card .tesora-kpi-value,
        .tesora-kpi-card .tesora-kpi-value * {
            color: #0A1E46 !important;
            font-size: 38px !important;
            font-weight: 800 !important;
            line-height: 1.1 !important;
            display: block;
        }

        /* Respaldo para st.metric nativos */
        div[data-testid="stMetricValue"],
        div[data-testid="stMetricValue"] > div,
        div[data-testid="stMetricValue"] p {
            color: #0A1E46 !important;
            font-size: 36px !important;
            font-weight: 800 !important;
        }
        div[data-testid="stMetricLabel"] p,
        div[data-testid="stMetricLabel"] label {
            color: #0B0F19 !important;
            font-size: 17px !important;
            font-weight: 700 !important;
        }

        @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,600;0,700;1,400;1,600&display=swap');

        /* Insignia Elegante y Femenina Tesora Pro en la barra lateral */
        .sidebar-brand {
            background: linear-gradient(135deg, #09193A 0%, #132B5E 55%, #1E3F7F 100%);
            padding: 20px 16px;
            border-radius: 18px;
            margin-bottom: 18px;
            box-shadow: 0 8px 22px rgba(9, 25, 58, 0.22);
            border: 1px solid rgba(243, 198, 182, 0.45);
            border-left: 5px solid #E8B4B8;
        }
        .sidebar-section-label {
            font-size: 13px !important;
            color: #0A1E46 !important;
            font-weight: 800 !important;
            letter-spacing: 0.8px;
            text-transform: uppercase;
            margin: 12px 0 6px 0 !important;
        }
    </style>
""",
    unsafe_allow_html=True,
)


def render_banner(titulo: str, subtitulo: str) -> None:
    """Renderiza un banner institucional en Azul Oscuro con título blanco de 32px y subtítulo de 18px."""
    html_banner = (
        f'<div class="tesora-banner">'
        f'<div class="tesora-banner-title" style="color:#FFFFFF; font-size:32px; font-weight:800; margin-bottom:8px;">{titulo}</div>'
        f'<div class="tesora-banner-sub" style="color:#E0F2FE; font-size:18px; font-weight:500;">{subtitulo}</div>'
        f'</div>'
    )
    st.markdown(html_banner, unsafe_allow_html=True)


def render_kpi_card(etiqueta: str, valor: object) -> None:
    """Renderiza un recuadro KPI de alto contraste con etiqueta de 17px y número grande de 38px."""
    html_kpi = (
        f'<div class="tesora-kpi-card">'
        f'<div class="tesora-kpi-label" style="color:#1E293B; font-size:17px; font-weight:700; margin-bottom:8px;">{etiqueta}</div>'
        f'<div class="tesora-kpi-value" style="color:#0A1E46; font-size:38px; font-weight:800;">{valor}</div>'
        f'</div>'
    )
    st.markdown(html_kpi, unsafe_allow_html=True)


# =============================================================================
# 4. GESTIÓN DE ESTADO Y SEGURIDAD (AUTENTICACIÓN)
# =============================================================================
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    render_banner("✨ Tesora Pro", "Sistema Corporativo de Conciliación Bancaria, Mayor Analítico y Tesorería")
    st.markdown("Por favor, ingrese su contraseña para continuar.")

    password_correcta = st.secrets.get("PASSWORD_ACCESO", "admin123")
    password = st.text_input("Contraseña:", type="password")

    if st.button("🔑 Ingresar a Tesora Pro", type="primary"):
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

if "menu_proceso" not in st.session_state:
    st.session_state.menu_proceso = "Dashboard"

if "modulo_analisis" not in st.session_state:
    st.session_state.modulo_analisis = "Resumen histórico"

if "modulo_maestros" not in st.session_state:
    st.session_state.modulo_maestros = "Ninguno"

# Procesar cualquier cambio de pestaña solicitado antes de instanciar los widgets del sidebar
if st.session_state.get("_nav_pending"):
    _nav = st.session_state.pop("_nav_pending")
    if "menu_proceso" in _nav:
        st.session_state.menu_proceso = _nav["menu_proceso"]
    if "modulo_analisis" in _nav:
        st.session_state.modulo_analisis = _nav["modulo_analisis"]
    if "modulo_maestros" in _nav:
        st.session_state.modulo_maestros = _nav["modulo_maestros"]


# =============================================================================
# 5. BARRA LATERAL DE NAVEGACIÓN CON LOGO VECTORIAL IMPONENTE TESORA PRO
# =============================================================================
with st.sidebar:
    # Logo SVG vectorial delicado y femenino (Oro Rosa / Champagne + Serif Elegante + Subtítulo en Cursiva)
    st.markdown(
        '<div class="sidebar-brand">'
        '<svg width="100%" height="72" viewBox="0 0 265 72" xmlns="http://www.w3.org/2000/svg">'
        '<defs>'
        '<linearGradient id="roseGoldGrad" x1="0%" y1="0%" x2="100%" y2="100%">'
        '<stop offset="0%" stop-color="#FCE7F3" />'
        '<stop offset="50%" stop-color="#F5D0C5" />'
        '<stop offset="100%" stop-color="#E8B4B8" />'
        '</linearGradient>'
        '<linearGradient id="iconSoftGrad" x1="0%" y1="0%" x2="100%" y2="100%">'
        '<stop offset="0%" stop-color="#FFFFFF" />'
        '<stop offset="100%" stop-color="#FFF1F2" />'
        '</linearGradient>'
        '</defs>'
        '<rect x="2" y="8" width="54" height="54" rx="27" fill="url(#iconSoftGrad)" stroke="#E8B4B8" stroke-width="1.5" />'
        '<path d="M29 17 C32 26, 36 30, 45 33 C36 36, 32 40, 29 49 C26 40, 22 36, 13 33 C22 30, 26 26, 29 17 Z" fill="#0F295E" />'
        '<path d="M29 22 C31 28, 33 30, 39 33 C33 36, 31 38, 29 44 C27 38, 25 36, 19 33 C25 30, 27 28, 29 22 Z" fill="#E8B4B8" />'
        '<circle cx="42" cy="21" r="2.2" fill="#E8B4B8" />'
        '<text x="68" y="36" fill="#FFFFFF" font-family="Playfair Display, Georgia, Times New Roman, serif" font-size="28" font-weight="700" letter-spacing="0.4">Tesora</text>'
        '<rect x="166" y="16" width="48" height="24" rx="12" fill="url(#roseGoldGrad)" />'
        '<text x="190" y="33" fill="#09193A" font-family="Playfair Display, Georgia, serif" font-size="14.5" font-weight="700" font-style="italic" text-anchor="middle" letter-spacing="0.5">Pro</text>'
        '<text x="68" y="57" fill="#F5D0C5" font-family="Playfair Display, Georgia, Times New Roman, serif" font-size="15.5" font-style="italic" font-weight="400" letter-spacing="0.3">Tesorería &amp; Conciliación</text>'
        '</svg>'
        '</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        "<p class='sidebar-section-label'>PROCESO</p>",
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
        key="menu_proceso",
        label_visibility="collapsed",
    )

    st.markdown("---")
    st.markdown(
        "<p class='sidebar-section-label'>ANÁLISIS & TASAS</p>",
        unsafe_allow_html=True,
    )
    modulo_analisis = st.selectbox(
        "Análisis",
        ["Resumen histórico", "Gestión de Tasas BCV / Multidivisa"],
        key="modulo_analisis",
    )

    st.markdown("---")
    st.markdown(
        "<p class='sidebar-section-label'>ARCHIVOS & MAESTROS</p>",
        unsafe_allow_html=True,
    )
    modulo_maestros = st.selectbox(
        "Maestros",
        ["Ninguno", "Bancos y Cuentas", "Estados de cuenta"],
        key="modulo_maestros",
    )

    st.markdown("---")
    st.caption("Soporte técnico • Configuración")


# =============================================================================
# 6. MÓDULOS Y VISTAS DE LA APLICACIÓN
# =============================================================================

if modulo_analisis == "Gestión de Tasas BCV / Multidivisa":
    render_banner(
        "💱 Tasas de Cambio BCV y Conversión Multidivisa",
        "Cargue masivamente las tasas oficiales desde Excel y realice conversiones multimoneda seguras en Tesora Pro.",
    )

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
    render_banner(
        "🏦 Bancos y Cuentas",
        "Maestro de cuentas bancarias de la empresa. Registre sus cuentas para vincularlas al importar extractos.",
    )

    col_f_cta, col_l_cta = st.columns([1.1, 1.5])
    with col_f_cta:
        with st.container(border=True):
            st.markdown("#### ➕ Nueva Cuenta Bancaria")
            with st.form("form_cuenta"):
                banco_input = st.selectbox(
                    "Institución Bancaria",
                    ["Banesco", "BNC (Banco Nacional de Crédito)", "Banco Mercantil", "BBVA Provincial", "Banco de Venezuela", "Bancamiga", "Banco Nacional", "Banco Regional"],
                )
                nro_cta = st.text_input("Número de Cuenta (20 dígitos o código IBAN)")
                moneda_cta = st.selectbox("Moneda de la Cuenta", ["VES", "USD", "EUR"])
                obs_cta = st.text_input("Alias / Observaciones (Ej. Cuenta Principal Recaudadora)")

                if st.form_submit_button("➕ Registrar Cuenta", use_container_width=True):
                    if nro_cta.strip():
                        nro_limpio = "".join(filter(str.isalnum, nro_cta)).upper()
                        existentes = [
                            "".join(filter(str.isalnum, c)).upper()
                            for c in st.session_state.cuentas_bancarias.get(banco_input, [])
                        ]
                        if nro_limpio in existentes:
                            st.error("⚠️ Ya existe una cuenta registrada con este número en el sistema.")
                        else:
                            etiqueta_cta = f"{nro_cta.strip()} ({moneda_cta})" + (f" — {obs_cta.strip()}" if obs_cta.strip() else "")
                            st.session_state.cuentas_bancarias.setdefault(banco_input, []).append(etiqueta_cta)
                            st.success(f"¡Cuenta registrada con éxito en {banco_input}!")
                            st.rerun()
                    else:
                        st.warning("Debe ingresar un número de cuenta válido.")

    with col_l_cta:
        with st.container(border=True):
            st.markdown("#### 🔍 Cuentas Registradas")
            filtro_cta = st.text_input("🔎 Buscar por banco, número de cuenta o moneda...", placeholder="Ej: Banesco, 0134, VES...")
            total_mostradas = 0
            for bnk, ctas in st.session_state.cuentas_bancarias.items():
                ctas_filtradas = [
                    c for c in ctas
                    if not filtro_cta.strip()
                    or filtro_cta.lower() in bnk.lower()
                    or filtro_cta.lower() in c.lower()
                ]
                if ctas_filtradas:
                    st.markdown(f"**🏛️ {bnk}**")
                    for c in ctas_filtradas:
                        total_mostradas += 1
                        st.markdown(
                            f"<div style='background:#F1F5F9; border-left:4px solid #0A1E46; padding:10px 14px; border-radius:8px; margin-bottom:8px; font-weight:600;'>"
                            f"💳 {c}</div>",
                            unsafe_allow_html=True,
                        )
            if total_mostradas == 0:
                st.info("No se encontraron cuentas que coincidan con la búsqueda.")

# --- VISTA: DASHBOARD PRINCIPAL CON ACCIONES RÁPIDAS Y LINKS ACCIONABLES ---
elif menu_proceso == "Dashboard":
    render_banner(
        "📊 Tesora Pro — Panel de Control Ejecutivo",
        "Resumen general de conciliaciones bancarias, partidas abiertas, asientos anulados e historial corporativo.",
    )
    res = st.session_state.resultado_conciliacion
    historial_actual = cargar_historial_disco()

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        val_abiertas = (
            (len(res.banco_pendiente) + len(res.libros_pendiente))
            if res
            else (len(st.session_state.movs_banco) + len(st.session_state.movs_libros))
        )
        render_kpi_card("Partidas abiertas (Banco + Libros)", val_abiertas)
    with c2:
        render_kpi_card("Conciliaciones guardadas", len(historial_actual))
    with c3:
        render_kpi_card("Movimientos conciliados (Sesión)", len(res.pares) if res else 0)
    with c4:
        render_kpi_card("Asientos anulados aislados", len(getattr(res, "libros_anulados", [])) if res else 0)

    col_izq, col_der = st.columns([2, 1])
    with col_izq:
        st.markdown("### Conciliaciones recientes")
        with st.container(border=True):
            if res:
                st.success(
                    f"✅ **Conciliación activa en memoria:** {len(res.pares)} pares conciliados | "
                    f"{len(res.banco_pendiente)} pend. banco | {len(res.libros_pendiente)} pend. libros | "
                    f"{len(getattr(res, 'libros_anulados', []))} anulados."
                )
                b_ir1, b_ir2 = st.columns(2)
                with b_ir1:
                    st.button(
                        "🔍 Ir al Detalle de la Conciliación Actual",
                        on_click=ir_a_menu,
                        args=("Conciliación",),
                        type="primary",
                        use_container_width=True,
                    )
                with b_ir2:
                    st.button(
                        "📌 Ver Partidas Pendientes y Anulados",
                        on_click=ir_a_menu,
                        args=("Partidas pendientes",),
                        use_container_width=True,
                    )

            if historial_actual:
                st.markdown("#### 🗂️ Últimos cierres guardados (acceso directo)")
                for idx_h, reg_h in enumerate(historial_actual[:5]):
                    etiqueta_h = reg_h.get("etiqueta", f"Conciliación #{len(historial_actual) - idx_h}")
                    fecha_h = reg_h.get("fecha_cierre", "Sin fecha")
                    c_info, c_btn1, c_btn2 = st.columns([3, 1.3, 1.3])
                    with c_info:
                        st.markdown(
                            f"**{etiqueta_h}** ({fecha_h})  \n"
                            f"<span style='color:#64748B; font-size:13px;'>"
                            f"Conciliados: {reg_h.get('pares_cerrados', 0)} • "
                            f"Pend. Banco: {reg_h.get('banco_pendiente', 0)} • "
                            f"Pend. Libros: {reg_h.get('libros_pendiente', 0)} • "
                            f"Anulados: {reg_h.get('libros_anulados', 0)}</span>",
                            unsafe_allow_html=True,
                        )
                    with c_btn1:
                        if st.button(
                            "🔄 Retomar / Editar",
                            key=f"dash_retomar_{idx_h}",
                            use_container_width=True,
                        ):
                            if retomar_conciliacion_historica(reg_h):
                                ir_a_menu("Conciliación")
                                st.rerun()
                            else:
                                st.session_state.historial_seleccionado_idx = idx_h
                                ir_a_menu("Historial")
                                st.rerun()
                    with c_btn2:
                        if st.button(
                            "📜 Ver Consolidado",
                            key=f"dash_ver_hist_{idx_h}",
                            use_container_width=True,
                        ):
                            st.session_state.historial_seleccionado_idx = idx_h
                            ir_a_menu("Historial")
                            st.rerun()
                    st.markdown("<hr style='margin:6px 0; border-color:#F1F5F9;'>", unsafe_allow_html=True)
            elif not res:
                st.info(
                    "📄 **No hay conciliaciones recientes.** Utilice los botones de **Acciones rápidas** a la derecha para importar sus archivos e iniciar su primera conciliación."
                )

    with col_der:
        st.markdown("### Acciones rápidas")
        with st.container(border=True):
            st.button(
                "📂 Importar extracto o mayor",
                on_click=ir_a_menu,
                args=("Importar",),
                use_container_width=True,
                type="primary",
            )
            st.button(
                "⚡ Ir al Motor de Conciliación",
                on_click=ir_a_menu,
                args=("Conciliación",),
                use_container_width=True,
            )
            st.button(
                "📌 Revisar Partidas Pendientes",
                on_click=ir_a_menu,
                args=("Partidas pendientes",),
                use_container_width=True,
            )
            st.button(
                "📜 Abrir Historial y Consolidados",
                on_click=ir_a_menu,
                args=("Historial",),
                use_container_width=True,
            )
            st.button(
                "💱 Cargar Tasas BCV / Multidivisa",
                on_click=lambda: st.session_state.update({"modulo_analisis": "Gestión de Tasas BCV / Multidivisa"}),
                use_container_width=True,
            )

elif menu_proceso == "Importar":
    render_banner(
        "📂 Importar Estado de Cuenta o Mayor Contable",
        "Suba el extracto bancario o el registro contable (Mayor), asocie la cuenta bancaria y parametrice sus columnas.",
    )

    with st.container(border=True):
        st.markdown("#### 1. Tipo de archivo y Cuenta Asociada")
        col_t1, col_t2, col_t3 = st.columns(3)
        with col_t1:
            tipo_origen = st.selectbox(
                "Tipo de documento a importar",
                ["Extracto Bancario (Movimientos del banco)", "Registro Contable (Mayor / Libro contable)"],
            )
        with col_t2:
            bancos_disp = list(st.session_state.cuentas_bancarias.keys())
            banco_sel_imp = st.selectbox("Banco asociado", bancos_disp if bancos_disp else ["Banesco"])
        with col_t3:
            ctas_disp = st.session_state.cuentas_bancarias.get(banco_sel_imp, ["Cuenta Principal"])
            cta_sel_imp = st.selectbox("Cuenta bancaria", ctas_disp if ctas_disp else ["Cuenta Principal"])

    with st.container(border=True):
        st.markdown("#### 2. Período y Saldos de Control (Opcional)")
        cp1, cp2, cp3, cp4 = st.columns(4)
        with cp1:
            st.date_input("Período desde", value=None, key="imp_fecha_desde")
        with cp2:
            st.date_input("Período hasta", value=None, key="imp_fecha_hasta")
        with cp3:
            st.number_input("Saldo inicial (VES)", value=0.0, step=100.0, format="%.2f", key="imp_saldo_ini")
        with cp4:
            st.number_input("Saldo final (VES)", value=0.0, step=100.0, format="%.2f", key="imp_saldo_fin")

    with st.container(border=True):
        st.markdown("#### 3. Seleccione el archivo (Excel o CSV)")
        archivo_subido = st.file_uploader("Arrastre el archivo aquí o haga clic para seleccionar", type=["xlsx", "xls", "csv"])

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

            if st.button("Procesar e Importar al Sistema", type="primary"):
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
    render_banner(
        "🏷️ Reglas de Conciliación y Etiquetas (Tags)",
        "Defina reglas paramétricas automáticas, consulte los algoritmos activos del motor y cree etiquetas personalizadas para auditoría.",
    )

    if "reglas_personalizadas" not in st.session_state:
        st.session_state.reglas_personalizadas = [
            {"nombre": "Aislamiento de Asientos Anulados (Debe = Haber)", "patron_banco": "N/A (Interno Libros)", "patron_contable": "Mismo Nº Doc + Misma Contrapartida + Signo Opuesto", "activa": True},
            {"nombre": "Transferencias entre Cuentas Propias (Fondeos)", "patron_banco": "TRANSFERENCIA / FONDEO", "patron_contable": "BANCO BANESCO / BNC / MERCANTIL", "activa": True},
        ]
    if "tags_personalizados" not in st.session_state:
        st.session_state.tags_personalizados = [
            {"nombre": "Pendiente Auditoría", "color": "#0A1E46"},
            {"nombre": "Comisión Bancaria", "color": "#1D4ED8"},
            {"nombre": "Fondeo Cuentas Propias", "color": "#047857"},
        ]

    with st.container(border=True):
        st.markdown("#### ⚙️ Reglas del Motor y Reglas Paramétricas Activas")
        st.markdown(
            "* **Regla Previa Anti-Falsos Positivos (Asientos Anulados):** Aísla antes del cruce bancario los pares en libros con igual `Nº documento`, igual `Contrapartida` y montos opuestos ($+X$ y $-X$).\n"
            "* **Regla C1 (Cruce Exacto):** Referencia + Monto idéntico + Fecha exacta.\n"
            "* **Regla C2 (Ventana de Días):** Mismo monto y signo dentro de los días de tolerancia.\n"
            "* **Reglas C3 / C4 (Sugerencias IA):** Aproximación de montos y agrupación de múltiples partidas."
        )
        df_reglas = pd.DataFrame(st.session_state.reglas_personalizadas)
        st.dataframe(df_reglas, use_container_width=True)

    col_nr, col_nt = st.columns([1.3, 1])
    with col_nr:
        with st.container(border=True):
            st.markdown("#### ➕ Nueva Regla Paramétrica")
            nom_r = st.text_input("Nombre de la regla", placeholder="Ej: Liquidación Punto de Venta / TDD")
            c_pb, c_pc = st.columns(2)
            with c_pb:
                pat_b = st.text_input("Patrón en Banco", placeholder="Ej: LIQUIDACION POS")
            with c_pc:
                pat_c = st.text_input("Patrón en Mayor Contable", placeholder="Ej: VENTAS POS")
            if st.button("➕ Crear Regla", type="primary"):
                if nom_r.strip():
                    st.session_state.reglas_personalizadas.append({
                        "nombre": nom_r.strip(),
                        "patron_banco": pat_b.strip() or "Cualquiera",
                        "patron_contable": pat_c.strip() or "Cualquiera",
                        "activa": True,
                    })
                    st.success("¡Regla paramétrica agregada con éxito!")
                    st.rerun()
                else:
                    st.warning("Ingrese un nombre para la regla.")

    with col_nt:
        with st.container(border=True):
            st.markdown("#### 🏷️ Mis Tags de Conciliación")
            tags_html = " ".join(
                f"<span style='display:inline-block; background:{t['color']}; color:#FFFFFF; padding:6px 14px; border-radius:20px; font-weight:700; font-size:14px; margin:4px;'>{t['nombre']}</span>"
                for t in st.session_state.tags_personalizados
            )
            st.markdown(tags_html, unsafe_allow_html=True)
            st.markdown("---")
            nom_t = st.text_input("Nombre del nuevo Tag", placeholder="Ej: Revisar con Proveedor")
            col_t_sel = st.selectbox("Color del Tag", ["Azul Oscuro (#0A1E46)", "Azul Rey (#1D4ED8)", "Negro (#0B0F19)", "Verde (#047857)", "Ámbar (#B45309)", "Rojo (#B91C1C)"])
            if st.button("➕ Crear Tag"):
                if nom_t.strip():
                    hex_c = col_t_sel.split("(")[-1].replace(")", "")
                    st.session_state.tags_personalizados.append({"nombre": nom_t.strip(), "color": hex_c})
                    st.rerun()

elif menu_proceso == "Conciliación":
    render_banner(
        "⚡ Conciliación Bancaria y Asistente IA",
        "Cruce inteligente con aislamiento previo de asientos anulados, distribución porcentual (%) y distinción de cuentas propias.",
    )

    with st.container(border=True):
        st.markdown("#### 📂 Archivos a Conciliar en Memoria")
        ca1, ca2 = st.columns(2)
        with ca1:
            n_bco = len(st.session_state.movs_banco)
            if n_bco > 0:
                st.success(f"🏛️ **EXTRACTO BANCARIO:** {n_bco} movimientos cargados y listos.")
            else:
                st.warning("🏛️ **EXTRACTO BANCARIO:** No hay extractos importados.")
                st.button("📤 Importar Extracto", on_click=ir_a_menu, args=("Importar",), key="btn_imp_bco_c")
        with ca2:
            n_lib = len(st.session_state.movs_libros)
            if n_lib > 0:
                st.success(f"📖 **REGISTROS CONTABLES (MAYOR):** {n_lib} movimientos cargados y listos.")
            else:
                st.warning("📖 **REGISTROS CONTABLES (MAYOR):** No hay registros importados.")
                st.button("📤 Importar Mayor", on_click=ir_a_menu, args=("Importar",), key="btn_imp_lib_c")

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
        with kpi1:
            render_kpi_card("Pares Cerrados (C1/C2)", datos_resumen.get("pares_cerrados", 0))
        with kpi2:
            render_kpi_card("Sugerencias (C3/C4)", datos_resumen.get("sugerencias", 0))
        with kpi3:
            render_kpi_card("Banco Pendiente", datos_resumen.get("banco_pendiente", 0))
        with kpi4:
            render_kpi_card("Libros Pendiente", datos_resumen.get("libros_pendiente", 0))
        with kpi5:
            render_kpi_card("Asientos Anulados", datos_resumen.get("libros_anulados", 0))

        # Guardado completo en Historial (con etiqueta identificadora, movimientos crudos y consolidado completo)
        with st.container(border=True):
            st.markdown("#### 💾 Guardar Conciliación Completa en el Historial (Para Retomar o Editar después)")
            col_et, col_bt = st.columns([3, 1])
            with col_et:
                etiqueta_cierre = st.text_input(
                    "Nombre o referencia de este cierre (Ej. Banesco Octubre 2026):",
                    value=f"Conciliación {datetime.now().strftime('%Y-%m-%d %H:%M')}",
                )
            with col_bt:
                st.write("")
                st.write("")
                if st.button("💾 Guardar en Historial", type="primary", use_container_width=True):
                    registro_historial = {
                        "id_cierre": datetime.now().strftime("%Y%m%d%H%M%S"),
                        "etiqueta": etiqueta_cierre.strip() or f"Conciliación {datetime.now().strftime('%Y-%m-%d')}",
                        "fecha_cierre": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "tolerancia_c2": tolerancia_c2,
                        "pares_cerrados": datos_resumen.get("pares_cerrados", 0),
                        "banco_pendiente": datos_resumen.get("banco_pendiente", 0),
                        "libros_pendiente": datos_resumen.get("libros_pendiente", 0),
                        "libros_anulados": datos_resumen.get("libros_anulados", 0),
                        "consolidado_rows": df_reporte_actual.to_dict(orient="records"),
                        "movs_banco_raw": [serializar_movimiento(m) for m in st.session_state.movs_banco],
                        "movs_libros_raw": [serializar_movimiento(m) for m in st.session_state.movs_libros],
                        "pregunta_ia": st.session_state.ultima_pregunta_ia,
                        "analisis_ia": st.session_state.ultimo_analisis_ia,
                    }
                    guardar_en_historial(registro_historial)
                    st.success("🎉 ¡Conciliación completa guardada! Puede ver su consolidado o editarla cuando quiera en el menú 'Historial' o desde el 'Dashboard'.")

        # =====================================================================
        # EXPORTAR REPORTE CONSOLIDADO EN EXCEL (CON % POR PARTIDA Y PESTAÑAS TOP)
        # =====================================================================
        st.markdown("---")
        st.subheader("📥 Exportar Reporte Consolidado + Distribución Porcentual en Excel")

        processed_data = empaquetar_excel_consolidado(
            df_reporte_actual,
            st.session_state.ultima_pregunta_ia,
            st.session_state.ultimo_analisis_ia,
        )

        st.download_button(
            label="📥 Descargar Reporte Consolidado + Distribución % en Excel (.xlsx)",
            data=processed_data,
            file_name="Reporte_Conciliacion_Consolidado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

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

                bytes_docx = generar_word_analisis_ia(
                    st.session_state.ultima_pregunta_ia,
                    st.session_state.ultimo_analisis_ia,
                    df_res_nat,
                    df_dist_egresos,
                    df_dist_ingresos,
                )

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
    render_banner(
        "📌 Partidas Conciliatorias Abiertas y Asientos Anulados",
        "Partidas sin conciliar que requieren seguimiento contable y asientos anulados aislados automáticamente.",
    )
    if st.session_state.resultado_conciliacion:
        res = st.session_state.resultado_conciliacion
        with st.container(border=True):
            f_col1, f_col2 = st.columns([2, 1])
            with f_col1:
                busq_pend = st.text_input("🔎 Buscar por descripción, referencia o monto...", placeholder="Ej: PS 15469, GASTOS, 500000...")
            with f_col2:
                lado_filtro = st.selectbox(" Filtrar por Lado / Estado", ["Todos", "Solo Pendientes de Banco", "Solo Pendientes de Libros", "Solo Asientos Anulados"])

        def _filtrar_df(df_in: pd.DataFrame) -> pd.DataFrame:
            if df_in.empty or not busq_pend.strip():
                return df_in
            q = busq_pend.strip().lower()
            mask = df_in.astype(str).apply(lambda col: col.str.lower().str.contains(q, na=False)).any(axis=1)
            return df_in[mask]

        if lado_filtro in ("Todos", "Solo Pendientes de Banco"):
            with st.container(border=True):
                st.subheader(f"🏛️ Pendientes de Banco ({len(res.banco_pendiente)})")
                if res.banco_pendiente:
                    df_pb = _filtrar_df(pd.DataFrame([vars(m) for m in res.banco_pendiente]))
                    st.dataframe(df_pb[["id", "fecha", "monto", "descripcion", "referencia"]], use_container_width=True)
                else:
                    st.info("Sin partidas pendientes en banco.")

        if lado_filtro in ("Todos", "Solo Pendientes de Libros"):
            with st.container(border=True):
                st.subheader(f"📖 Pendientes de Libros ({len(res.libros_pendiente)})")
                if res.libros_pendiente:
                    df_pl = _filtrar_df(pd.DataFrame([vars(m) for m in res.libros_pendiente]))
                    st.dataframe(df_pl[["id", "fecha", "monto", "descripcion", "referencia"]], use_container_width=True)
                else:
                    st.info("Sin partidas pendientes en libros.")

        if lado_filtro in ("Todos", "Solo Asientos Anulados"):
            with st.container(border=True):
                st.subheader(f"🚫 Asientos Anulados Neutralizados ({len(getattr(res, 'libros_anulados', []))})")
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
                    df_pa = _filtrar_df(pd.DataFrame(lista_anulados))
                    st.dataframe(df_pa[["id", "fecha", "monto", "descripcion", "referencia", "marca"]], use_container_width=True)
                else:
                    st.info("Sin asientos anulados detectados.")
    else:
        with st.container(border=True):
            st.info("Ejecute una conciliación primero para visualizar y filtrar las partidas abiertas.")

# --- VISTA: HISTORIAL INTERACTIVO CON CONSOLIDADO COMPLETO Y EDICIÓN EN VIVO ---
elif menu_proceso == "Historial":
    render_banner(
        "📜 Historial de Conciliaciones y Editor de Consolidados",
        "Consulte cierres anteriores, edite marcas u observaciones en vivo, descargue en Excel o reabra la conciliación en el motor.",
    )

    historial_disco = cargar_historial_disco()
    if historial_disco:
        opciones_hist = [
            f"{idx + 1}. {reg.get('etiqueta', 'Conciliación')} ({reg.get('fecha_cierre', 'Sin fecha')}) — Conciliados: {reg.get('pares_cerrados', 0)}"
            for idx, reg in enumerate(historial_disco)
        ]
        idx_defecto = st.session_state.get("historial_seleccionado_idx", 0)
        if idx_defecto >= len(opciones_hist):
            idx_defecto = 0

        seleccion = st.selectbox(
            "Seleccione la conciliación histórica que desea inspeccionar o editar:",
            options=range(len(opciones_hist)),
            format_func=lambda i: opciones_hist[i],
            index=idx_defecto,
        )
        reg_sel = historial_disco[seleccion]

        with st.container(border=True):
            c_m1, c_m2, c_m3, c_m4 = st.columns(4)
            with c_m1:
                render_kpi_card("Pares Cerrados", reg_sel.get("pares_cerrados", 0))
            with c_m2:
                render_kpi_card("Banco Pendiente", reg_sel.get("banco_pendiente", 0))
            with c_m3:
                render_kpi_card("Libros Pendiente", reg_sel.get("libros_pendiente", 0))
            with c_m4:
                render_kpi_card("Asientos Anulados", reg_sel.get("libros_anulados", 0))

            col_acc1, col_acc2, col_acc3 = st.columns(3)
            with col_acc1:
                if st.button("🔄 Retomar / Reabrir en el Motor de Conciliación", type="primary", use_container_width=True):
                    if retomar_conciliacion_historica(reg_sel):
                        ir_a_menu("Conciliación")
                        st.rerun()
                    else:
                        st.warning(
                            "Este registro fue guardado con una versión anterior que solo almacenaba el resumen numérico. "
                            "Guarde su conciliación actual con el botón '💾 Guardar en Historial' para almacenar el consolidado completo."
                        )
            with col_acc2:
                nuevo_nombre = st.text_input(
                    "Renombrar etiqueta del cierre:",
                    value=reg_sel.get("etiqueta", ""),
                    key=f"renombrar_{seleccion}",
                    label_visibility="collapsed",
                )
            with col_acc3:
                if st.button("🗑️ Eliminar este registro del Historial", use_container_width=True):
                    historial_disco.pop(seleccion)
                    guardar_lista_historial_disco(historial_disco)
                    st.success("Registro eliminado del historial.")
                    st.rerun()

        # Tabla Consolidada Editable si el registro contiene las filas consolidadas
        filas_guardadas = reg_sel.get("consolidado_rows", [])
        if filas_guardadas:
            st.markdown("### 📋 Tabla Consolidada de la Conciliación (Editable)")
            st.caption(
                "Puede modificar directamente la columna **Marca Conciliación**, **Naturaleza Contable** o añadir notas en **Observaciones / Ajustes**. "
                "Luego presione **💾 Guardar Cambios en el Consolidado**."
            )
            df_hist_consolidado = pd.DataFrame(filas_guardadas)
            if "Observaciones / Ajustes" not in df_hist_consolidado.columns:
                df_hist_consolidado["Observaciones / Ajustes"] = ""

            df_editado = st.data_editor(
                df_hist_consolidado,
                use_container_width=True,
                num_rows="dynamic",
                key=f"editor_hist_{seleccion}",
                column_config={
                    "Marca Conciliación": st.column_config.SelectboxColumn(
                        "Marca Conciliación",
                        options=[
                            "Conciliado (C1)",
                            "Conciliado (C2)",
                            "Conciliado (Manual)",
                            "Por conciliar (Banco)",
                            "Por conciliar (Libros)",
                            "Asiento Anulado",
                        ],
                        required=True,
                    ),
                },
            )

            c_save, c_dl = st.columns(2)
            with c_save:
                if st.button("💾 Guardar Cambios en el Consolidado Histórico", type="primary", use_container_width=True):
                    df_recalc = recalcular_porcentajes_consolidado(df_editado)
                    historial_disco[seleccion]["etiqueta"] = nuevo_nombre.strip() or reg_sel.get("etiqueta", "Conciliación")
                    historial_disco[seleccion]["consolidado_rows"] = df_recalc.to_dict(orient="records")
                    # Actualizar conteos según las marcas editadas
                    marcas_s = df_recalc["Marca Conciliación"].astype(str)
                    historial_disco[seleccion]["pares_cerrados"] = int(marcas_s.str.startswith("Conciliado").sum())
                    historial_disco[seleccion]["banco_pendiente"] = int((marcas_s == "Por conciliar (Banco)").sum())
                    historial_disco[seleccion]["libros_pendiente"] = int((marcas_s == "Por conciliar (Libros)").sum())
                    historial_disco[seleccion]["libros_anulados"] = int((marcas_s == "Asiento Anulado").sum())
                    guardar_lista_historial_disco(historial_disco)
                    st.success("✅ ¡Cambios guardados exitosamente en el consolidado histórico!")
                    st.rerun()

            with c_dl:
                excel_hist_bytes = empaquetar_excel_consolidado(
                    df_editado,
                    reg_sel.get("pregunta_ia", ""),
                    reg_sel.get("analisis_ia", ""),
                )
                st.download_button(
                    label="📥 Descargar este Consolidado Editado en Excel (.xlsx)",
                    data=excel_hist_bytes,
                    file_name=f"Consolidado_{reg_sel.get('id_cierre', 'Historico')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        else:
            st.info(
                "ℹ️ Este registro fue creado antes de activar el guardado de tabla consolidada completa. "
                "Vaya a **Conciliación** y presione **💾 Guardar en Historial** para almacenar el consolidado editable."
            )
    else:
        st.info("No hay conciliaciones guardadas en el historial todavía. Ejecute una conciliación y haga clic en '💾 Guardar en Historial'.")

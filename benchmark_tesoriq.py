import streamlit as st
import pandas as pd
import random
import time
import tracemalloc
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path
import io
import csv

from conciliar import Movimiento, Config, conciliar

st.set_page_config(page_title="TESORIQ | Banco de pruebas", page_icon="🧪", layout="wide")
st.title("TESORIQ — Banco de pruebas de rendimiento")
st.warning(
    "PRUEBA CONTROLADA: utiliza datos sintéticos. No cargues extractos reales ni datos de clientes. "
    "Esta primera versión mide el motor en volúmenes pequeños/medios; no ejecuta todavía pruebas de 50.000 x 50.000."
)

st.subheader("1. Selección del escenario")
col1, col2 = st.columns(2)
with col1:
    n_banco = st.selectbox("Movimientos bancarios", [100, 250, 500, 1000, 2000], index=1)
with col2:
    n_libros = st.selectbox("Movimientos del libro mayor", [100, 250, 500, 1000, 2000], index=1)

st.caption("Los movimientos se generan con referencias únicas y una mezcla de coincidencias exactas y partidas pendientes.")
seed = st.number_input("Semilla reproducible", min_value=1, max_value=999999, value=20261121, step=1)
run = st.button("Ejecutar prueba", type="primary")

def crear_datos(n_banco, n_libros, seed):
    rng = random.Random(seed)
    fecha_base = date(2026, 1, 1)
    banco, libros = [], []
    # Una parte de los movimientos bancarios tendrá pareja exacta en libros.
    parejas = min(int(min(n_banco, n_libros) * 0.70), n_banco, n_libros)
    for i in range(n_banco):
        monto = Decimal(str(100 + i * 1.37)).quantize(Decimal("0.01"))
        f = fecha_base + timedelta(days=i % 28)
        ref = f"REF{i:08d}"
        banco.append(Movimiento(
            id=f"B{i:08d}", origen="banco", fecha=f, monto=monto,
            descripcion="TRANSFERENCIA PRUEBA", referencia=ref, moneda="VES"
        ))
        if i < parejas:
            # Coincidencia exacta de monto, fecha y referencia
            libros.append(Movimiento(
                id=f"L{i:08d}", origen="libros", fecha=f, monto=monto,
                descripcion="TRANSFERENCIA PRUEBA", referencia=ref, moneda="VES"
            ))
    # Completa libros con movimientos no coincidentes
    while len(libros) < n_libros:
        i = len(libros)
        monto = Decimal(str(500000 + i * 2.19)).quantize(Decimal("0.01"))
        f = fecha_base + timedelta(days=(i * 3) % 28)
        libros.append(Movimiento(
            id=f"L{i:08d}", origen="libros", fecha=f, monto=monto,
            descripcion="PAGO SIN CONTRAPARTIDA", referencia=f"OTR{i:08d}", moneda="VES"
        ))
    return banco, libros, parejas

def exportar_csv(movs, origen):
    s = io.StringIO()
    w = csv.writer(s)
    w.writerow(["id", "fecha", "monto", "descripcion", "referencia", "moneda"])
    for m in movs:
        w.writerow([m.id, m.fecha.isoformat(), str(m.monto), m.descripcion, m.referencia, m.moneda])
    return s.getvalue().encode("utf-8")

if run:
    with st.spinner("Generando datos sintéticos y ejecutando el motor..."):
        banco, libros, parejas_esperadas = crear_datos(int(n_banco), int(n_libros), int(seed))
        # Validación de serialización CSV y Excel en memoria; no se usa información real.
        csv_bytes = exportar_csv(banco, "banco")
        df = pd.DataFrame([{
            "id": m.id, "fecha": m.fecha.isoformat(), "monto": float(m.monto),
            "descripcion": m.descripcion, "referencia": m.referencia, "moneda": m.moneda
        } for m in libros])
        xlsx_buffer = io.BytesIO()
        df.to_excel(xlsx_buffer, index=False, engine="openpyxl")
        xlsx_bytes = xlsx_buffer.getvalue()

        tracemalloc.start()
        t0 = time.perf_counter()
        resultado = conciliar(banco, libros, Config())
        elapsed = time.perf_counter() - t0
        current, peak = tracemalloc.get_traced_memory()
        tracemalloc.stop()

    resumen = resultado.resumen()
    c1_c2 = sum(1 for p in resultado.pares if p.regla in ("C1", "C2"))
    st.subheader("2. Resultados")
    a, b, c, d = st.columns(4)
    a.metric("Tiempo motor", f"{elapsed:.2f} s")
    b.metric("Pares cerrados", f"{len(resultado.pares):,}")
    c.metric("Sugerencias", f"{len(resultado.sugerencias):,}")
    d.metric("Pendientes banco/libros", f"{len(resultado.banco_pendiente):,} / {len(resultado.libros_pendiente):,}")

    st.write("**Memoria Python (tracemalloc):**", f"{peak / (1024*1024):.1f} MB pico; {current / (1024*1024):.1f} MB al terminar.")
    st.write("**Prueba de formatos:**", f"CSV generado: {len(csv_bytes):,} bytes · Excel generado: {len(xlsx_bytes):,} bytes.")
    st.write("**Coincidencias exactas esperadas en los datos sintéticos:**", f"{parejas_esperadas:,}")
    st.write("**Pares cerrados C1/C2 obtenidos:**", f"{c1_c2:,}")
    st.json(resumen)
    if c1_c2 != parejas_esperadas:
        st.error("La cantidad de coincidencias C1/C2 no coincide con la esperada. Revisar antes de escalar.")
    else:
        st.success("La cantidad de coincidencias C1/C2 coincide con el conjunto sintético.")
    if resultado.advertencias:
        st.warning("Advertencias del motor")
        for aviso in resultado.advertencias:
            st.write("•", aviso)

    st.download_button(
        "Descargar resultados de la prueba (CSV)",
        data=pd.DataFrame([{
            "banco": n_banco, "libros": n_libros, "tiempo_segundos": round(elapsed, 4),
            "memoria_pico_mb": round(peak / (1024*1024), 2),
            "pares_cerrados": len(resultado.pares), "sugerencias": len(resultado.sugerencias),
            "pendientes_banco": len(resultado.banco_pendiente),
            "pendientes_libros": len(resultado.libros_pendiente),
            "parejas_esperadas": parejas_esperadas, "parejas_obtenidas_c1_c2": c1_c2
        }]).to_csv(index=False).encode("utf-8"),
        file_name="tesoriq_resultado_prueba.csv",
        mime="text/csv"
    )

st.divider()
st.subheader("3. Progresión recomendada")
st.markdown("""
1. Ejecuta primero **250 + 250**.
2. Si termina correctamente, prueba **500 + 500**.
3. Después prueba **1.000 + 1.000**.
4. No ejecutes todavía 25.000 o 50.000 por archivo en Community Cloud.
5. Guarda el CSV de resultados de cada ejecución y los registros de Cloud.
""")
st.caption("Este banco mide el motor de conciliación y la creación de archivos de prueba. No sustituye pruebas de importación real, concurrencia ni una prueba de estrés de producción.")

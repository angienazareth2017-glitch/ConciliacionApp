"""
Tests del motor. Ejecutar con:   pip install pytest   y luego   pytest -q
Cada test de "regresión" corresponde a un error real del script anterior.
"""
import random
import time
from datetime import date
from decimal import Decimal as D

import pytest

import tasas
from conciliar import (
    Config, ErrorDeParseo, Movimiento, calcular_cuadre, cargar_movimientos,
    clasificar_movimiento, conciliar, neutralizar_reversos, normalizar_referencia,
    parsear_fecha, parsear_monto, referencias_coinciden,
)


def B(id_, fecha, monto, desc="", ref=""):
    return Movimiento(id_, "banco", fecha, D(str(monto)), desc, ref)


def L(id_, fecha, monto, desc="", ref=""):
    return Movimiento(id_, "libros", fecha, D(str(monto)), desc, ref)


# ----------------------------------------------------------------------------
# Parseo de montos  (regresión: limpiar_monto devolvía 0.0 o valores 100x mal)
# ----------------------------------------------------------------------------
@pytest.mark.parametrize("entrada, esperado", [
    ("1.234,56", "1234.56"),
    ("-1.234,56", "-1234.56"),
    ("1,234.56", "1234.56"),          # antes daba 1.23456
    ("1234.56", "1234.56"),           # antes daba 123456.0
    ("(1.234,56)", "-1234.56"),       # antes daba 0.0
    ("1.234,56-", "-1234.56"),        # antes daba 0.0
    ("Bs.S 1.234,56", "1234.56"),     # antes daba 0.0
    ("VES 500,00", "500.00"),         # antes daba 0.0
    ("Bs. 50", "50.00"),
    ("$1,234.56", "1234.56"),
    ("1.234.567,89", "1234567.89"),
    ("1.234", "1234.00"),             # ambiguo: con coma decimal, el punto es de miles
    (1234.5, "1234.50"),
    (10, "10.00"),
    (D("2.675"), "2.68"),             # ROUND_HALF_UP (round(2.675, 2) en float da 2.67)
])
def test_parsear_monto(entrada, esperado):
    assert parsear_monto(entrada) == D(esperado)


def test_parsear_monto_ambiguo_con_punto_decimal():
    assert parsear_monto("1,234", separador_decimal=".") == D("1234.00")
    assert parsear_monto("1.234", separador_decimal=".") == D("1.23")


@pytest.mark.parametrize("malo", ["abc", "", "   ", None, float("nan"), "1-2", True, "Bs."])
def test_parsear_monto_error_no_cero(malo):
    with pytest.raises(ErrorDeParseo):
        parsear_monto(malo)


# ----------------------------------------------------------------------------
# Fechas  (regresión: 05/03/2025 se leía como 3 de mayo)
# ----------------------------------------------------------------------------
def test_fecha_dia_primero():
    assert parsear_fecha("05/03/2025") == date(2025, 3, 5)
    assert parsear_fecha("13/03/2025") == date(2025, 3, 13)
    assert parsear_fecha("2025-03-05") == date(2025, 3, 5)
    assert parsear_fecha("20250305") == date(2025, 3, 5)
    assert parsear_fecha(45000) == date(2023, 3, 15)  # serial de Excel
    assert parsear_fecha("05/03/2025", dayfirst=False) == date(2025, 5, 3)


def test_fecha_invalida():
    with pytest.raises(ErrorDeParseo):
        parsear_fecha("no es fecha")
    with pytest.raises(ErrorDeParseo):
        parsear_fecha(None)


# ----------------------------------------------------------------------------
# Clasificación  (regresión: "com" marcaba COMERCIO, COMPRA, COMPAÑIA como comisión)
# ----------------------------------------------------------------------------
@pytest.mark.parametrize("desc, monto, esperado", [
    ("PAGO COMERCIO FERRETERIA", "-100", "Egresos"),
    ("COMPRA POS", "-50", "Egresos"),
    ("TRANSFERENCIA RECIBIDA COMPAÑIA XYZ", "300", "Ingresos"),
    ("DEPOSITO CANAL 3", "200", "Ingresos"),
    ("COMISIÓN MANTENIMIENTO", "-5", "Comisiones"),
    ("COM. SERV. MTO", "-5", "Comisiones"),
    ("COMSERVMTO", "-5", "Comisiones"),
    ("SMS ALERTAS", "-1", "Comisiones"),
    ("IGTF 3%", "-3", "Impuestos"),
    ("PAGO NOMINA", "-900", "Egresos"),
])
def test_clasificacion(desc, monto, esperado):
    assert clasificar_movimiento(desc, D(monto)) == esperado


# ----------------------------------------------------------------------------
# Referencias
# ----------------------------------------------------------------------------
def test_referencias():
    assert normalizar_referencia("000123456") == "123456"
    assert normalizar_referencia(123456.0) == "123456"
    assert referencias_coinciden("987654321", "654321")      # truncada por el banco
    assert not referencias_coinciden("987654321", "4321")    # sufijo muy corto
    assert not referencias_coinciden("", "123456")


# ----------------------------------------------------------------------------
# Motor de cruce
# ----------------------------------------------------------------------------
def test_signo_invertido_no_concilia():
    """Regresión: abs() == abs() conciliaba un egreso con un ingreso."""
    r = conciliar([B("b1", "2025-03-03", -100)], [L("l1", "2025-03-03", 100)])
    assert r.pares == [] and r.sugerencias == []
    assert len(r.banco_pendiente) == 1 and len(r.libros_pendiente) == 1


def test_c1_referencia_truncada():
    r = conciliar([B("b1", "2025-03-03", 250, ref="000987654321")],
                  [L("l1", "2025-03-03", 250, ref="654321")])
    assert [p.regla for p in r.pares] == ["C1"]
    assert r.pares[0].confianza == "alta"


def test_c2_viernes_a_lunes():
    r = conciliar([B("b1", "2025-03-07", -300)], [L("l1", "2025-03-10", -300)])
    assert [p.regla for p in r.pares] == ["C2"]
    assert r.pares[0].confianza == "media"  # sin referencia


def test_misma_cifra_fecha_lejana_no_concilia():
    """Regresión: el 2º pase comparaba solo el monto, sin importar la fecha."""
    r = conciliar([B("b1", "2025-03-03", -300)], [L("l1", "2025-04-02", -300)])
    assert r.pares == [] and r.sugerencias == []


def test_no_reutiliza_partidas():
    r = conciliar(
        [B("b1", "2025-03-03", 100), B("b2", "2025-03-03", 100)],
        [L("l1", "2025-03-03", 100)],
    )
    assert len(r.pares) == 1 and len(r.banco_pendiente) == 1


def test_c2_prefiere_la_referencia_correcta():
    r = conciliar(
        [B("b1", "2025-03-03", 80, ref="BBB222")],
        [L("l1", "2025-03-04", 80, ref="AAA111"), L("l2", "2025-03-04", 80, ref="BBB222")],
    )
    assert r.pares[0].libros[0].id == "l2"
    assert r.pares[0].confianza == "alta"


def test_reverso_se_neutraliza_con_su_original():
    """Regresión: se descartaba el reverso y el original quedaba huérfano."""
    libros = [
        L("l1", "2025-03-03", -500, "PAGO PROVEEDOR"),
        L("l2", "2025-03-04", 500, "ANULACION PAGO PROVEEDOR"),
    ]
    r = conciliar([], libros)
    assert r.libros_pendiente == []
    assert len(r.reversos_libros) == 1
    assert r.reversos_libros[0][0].id == "l1" and r.reversos_libros[0][1].id == "l2"


def test_reverso_sin_original_no_se_pierde():
    libros = [L("l1", "2025-03-04", 500, "ANULACION CHEQUE")]
    restantes, pares, sin = neutralizar_reversos(libros)
    assert [m.id for m in restantes] == ["l1"] and pares == [] and len(sin) == 1
    r = conciliar([], libros)
    assert len(r.libros_pendiente) == 1 and r.advertencias


def test_c3_es_sugerencia_y_requiere_aprobacion():
    banco = [B("b1", "2025-03-03", "-1000.50", "PAGO PROVEEDOR ACME")]
    libros = [L("l1", "2025-03-03", "-1000.00", "PAGO PROVEEDOR ACME C.A.")]
    r = conciliar(banco, libros)
    assert r.pares == []                       # nada se cierra solo
    assert len(r.sugerencias) == 1 and r.sugerencias[0].regla == "C3"
    assert r.sugerencias[0].requiere_revision
    assert len(r.banco_pendiente) == 1         # sigue pendiente hasta aprobar

    r.aprobar(r.sugerencias[0])
    assert len(r.pares) == 1 and r.pares[0].diferencia_monto == D("-0.50")
    assert r.banco_pendiente == [] and r.libros_pendiente == [] and r.sugerencias == []


def test_c3_monto_exacto_fecha_fuera_de_tolerancia():
    r = conciliar([B("b1", "2025-03-03", 700)], [L("l1", "2025-03-09", 700)])
    assert r.pares == []
    assert len(r.sugerencias) == 1 and "fuera de tolerancia" in r.sugerencias[0].motivo


def test_c4_uno_a_varios_y_varios_a_uno():
    r = conciliar([B("b1", "2025-03-03", 300)],
                  [L("l1", "2025-03-03", 100), L("l2", "2025-03-04", 200)])
    assert [(s.regla, len(s.banco), len(s.libros)) for s in r.sugerencias] == [("C4", 1, 2)]

    r = conciliar([B("b1", "2025-03-03", 100), B("b2", "2025-03-03", 200)],
                  [L("l1", "2025-03-03", 300)])
    assert [(s.regla, len(s.banco), len(s.libros)) for s in r.sugerencias] == [("C4", 2, 1)]


def test_c4_respeta_el_presupuesto_y_avisa():
    banco = [B(f"b{i}", "2025-03-03", 1000 + i) for i in range(30)]
    libros = [L(f"l{i}", "2025-03-03", 7 + i) for i in range(30)]
    r = conciliar(banco, libros, Config(presupuesto_combinaciones_c4=50))
    assert any("C4" in a for a in r.advertencias)


def test_ids_duplicados_error():
    with pytest.raises(ValueError):
        conciliar([B("b1", "2025-03-03", 1), B("b1", "2025-03-03", 2)], [])


# ----------------------------------------------------------------------------
# Cuadre
# ----------------------------------------------------------------------------
def _caso_cuadre():
    banco = [B("b1", "2025-03-03", -50, "CHEQUE 123", "123"),
             B("b2", "2025-03-31", -5, "COMISION MANTENIMIENTO")]
    libros = [L("l1", "2025-03-03", -50, "CHEQUE 123", "123"),
              L("l2", "2025-03-31", 100, "DEPOSITO EN TRANSITO")]
    return banco, libros


def test_cuadre_cuadrado():
    banco, libros = _caso_cuadre()
    r = conciliar(banco, libros)
    c = calcular_cuadre(r, saldo_banco="945,00", saldo_libros="1.050,00")
    assert c.saldo_banco_ajustado == D("1045.00") and c.saldo_libros_ajustado == D("1045.00")
    assert c.cuadrado


def test_cuadre_detecta_diferencia():
    banco, libros = _caso_cuadre()
    c = calcular_cuadre(conciliar(banco, libros), saldo_banco="945,00", saldo_libros="1.060,00")
    assert not c.cuadrado and c.diferencia == D("-10.00")


def test_cuadre_con_sugerencia_aprobada_con_diferencia():
    banco = [B("b1", "2025-03-03", "-1000.50", "PAGO ACME")]
    libros = [L("l1", "2025-03-03", "-1000.00", "PAGO ACME")]
    r = conciliar(banco, libros)
    r.aprobar(r.sugerencias[0])
    # banco: 5000 - 1000.50 = 3999.50 ; libros: 5000 - 1000.00 = 4000.00
    c = calcular_cuadre(r, saldo_banco="3999,50", saldo_libros="4000,00")
    assert c.cuadrado


# ----------------------------------------------------------------------------
# Carga desde filas  (regresión: errores silenciosos, columnas débito/crédito)
# ----------------------------------------------------------------------------
def test_carga_entrada_salida_y_errores_visibles():
    filas = [
        {"Fecha": "05/03/2025", "Detalle": "DEPOSITO", "Ref": 123456.0, "Cred": "1.000,00", "Deb": ""},
        {"Fecha": "06/03/2025", "Detalle": "CHEQUE", "Ref": "777", "Cred": None, "Deb": "250,00"},
        {"Fecha": "fecha mala", "Detalle": "X", "Ref": "", "Cred": "5,00", "Deb": None},
        {"Fecha": "07/03/2025", "Detalle": "Y", "Ref": "", "Cred": "abc", "Deb": None},
        {"Fecha": None, "Detalle": None, "Ref": None, "Cred": None, "Deb": None},
        {"Fecha": "08/03/2025", "Detalle": "CERO", "Ref": "", "Cred": "0,00", "Deb": None},
    ]
    c = cargar_movimientos(
        filas, "banco",
        {"fecha": "Fecha", "descripcion": "Detalle", "referencia": "Ref", "entrada": "Cred", "salida": "Deb"},
    )
    assert [(m.fecha, m.monto) for m in c.movimientos] == [
        (date(2025, 3, 5), D("1000.00")), (date(2025, 3, 6), D("-250.00"))
    ]
    assert c.movimientos[0].referencia_norm == "123456"  # no "1234560"
    assert [e.fila for e in c.errores] == [3, 4]          # los ilegibles NO se pierden
    assert c.omitidas == 2


def test_carga_columna_inexistente():
    with pytest.raises(ValueError, match="Columnas no encontradas"):
        cargar_movimientos([{"A": 1}], "banco", {"fecha": "Fecha", "monto": "Monto"})


# ----------------------------------------------------------------------------
# Rendimiento  (regresión: doble iterrows O(n*m); 3000x3000 tardaba ~4 min)
# ----------------------------------------------------------------------------
def test_rendimiento_5000_movimientos():
    rnd = random.Random(7)
    n = 5000
    montos = [D(rnd.randint(1, 10**7)) / 100 for _ in range(n)]
    base = date(2025, 3, 1).toordinal()
    fechas = [date.fromordinal(base + rnd.randint(0, 30)) for _ in range(n)]
    banco = [B(f"b{i}", fechas[i], montos[i]) for i in range(n)]
    libros = [L(f"l{i}", fechas[i], montos[i]) for i in range(n)]
    rnd.shuffle(libros)
    t = time.perf_counter()
    r = conciliar(banco, libros)
    dt = time.perf_counter() - t
    assert len(r.pares) == n
    assert dt < 10, f"demasiado lento: {dt:.1f}s"


# ----------------------------------------------------------------------------
# Tasas
# ----------------------------------------------------------------------------
@pytest.fixture
def tabla():
    t = tasas.TablaTasas(max_antiguedad_dias=5)
    t.agregar("2025-03-07", "USD", "36,5000")   # viernes
    t.agregar("2025-03-07", "EUR", "40,0000")
    t.agregar("2025-03-10", "USD", "37,0000")   # lunes
    t.agregar("2025-03-10", "EUR", "41,0000")
    return t


def test_tasa_exacta_y_arrastrada(tabla):
    ex = tabla.obtener("2025-03-07", "USD")
    assert ex.valor == D("36.5") and not ex.arrastrada
    sab = tabla.obtener("2025-03-08", "USD")  # sábado -> usa la del viernes
    assert sab.fecha_tasa == date(2025, 3, 7) and sab.arrastrada and sab.dias_arrastre == 1


def test_no_hay_tasa_por_defecto(tabla):
    with pytest.raises(tasas.TasaNoDisponible):
        tabla.obtener("2025-03-06", "USD")      # antes de la primera: no se usa una posterior
    with pytest.raises(tasas.TasaNoDisponible):
        tabla.obtener("2025-04-01", "USD")      # muy antigua
    with pytest.raises(tasas.TasaNoDisponible):
        tabla.obtener("2025-03-10", "USD", tipo="MERCADO")


def test_conversiones(tabla):
    assert tabla.convertir("3.650,00", "VES", "USD", "2025-03-07") == D("100.00")
    assert tabla.convertir(100, "USD", "VES", "2025-03-07") == D("3650.00")
    # cruce USD -> EUR vía bolívares: 100 * 36.5 / 40 = 91.25
    assert tabla.convertir(100, "USD", "EUR", "2025-03-07") == D("91.25")
    assert tabla.convertir(100, "USD", "USD", "2025-03-07") == D("100.00")


def test_tres_monedas_y_auditoria(tabla):
    assert tabla.tres_monedas("3.700,00", "VES", "2025-03-10") == {
        "VES": D("3700.00"), "USD": D("100.00"), "EUR": D("90.24"),
    }
    valor, usadas = tabla.convertir_con_detalle(100, "USD", "EUR", "2025-03-08")
    assert {u.moneda for u in usadas} == {"USD", "EUR"} and all(u.arrastrada for u in usadas)


def test_tasa_con_tres_decimales_no_se_confunde_con_miles():
    t = tasas.TablaTasas()
    t.agregar("2025-03-07", "USD", "36.285")
    assert t.obtener("2025-03-07", "USD").valor == D("36.285")  # no 36285


def test_cargar_tasas_reporta_errores():
    t = tasas.TablaTasas()
    filas = [
        {"Fecha": "07/03/2025", "Dolar": "36,50", "Euro": "40,00"},
        {"Fecha": "08/03/2025", "Dolar": "xyz", "Euro": ""},
        {"Fecha": "09/03/2025", "Dolar": "-5", "Euro": None},
        {"Fecha": "mal", "Dolar": "1", "Euro": "1"},
    ]
    errores = t.cargar_filas(filas, "Fecha", {"USD": "Dolar", "EUR": "Euro"})
    assert sorted(e.fila for e in errores) == [2, 3, 4]
    assert t.obtener("2025-03-07", "EUR").valor == D("40")


def test_rechaza_tasa_cero_o_base():
    t = tasas.TablaTasas()
    with pytest.raises(ValueError):
        t.agregar("2025-03-07", "USD", "0")
    with pytest.raises(ValueError):
        t.agregar("2025-03-07", "VES", "1")

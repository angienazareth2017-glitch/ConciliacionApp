"""
tasas.py - Tabla de tasas de cambio para Bs. / USD / EUR.

Diferencias clave con el código anterior:
    * NO existe una tasa por defecto. Si no hay tasa válida para una fecha,
      se lanza TasaNoDisponible. Un valor inventado (el 37.03 fijo) contamina
      todos los montos convertidos sin que nadie lo note.
    * Una fecha sin tasa (fin de semana, feriado) usa la última tasa ANTERIOR,
      hasta `max_antiguedad_dias`, y lo deja registrado (`arrastrada`).
      Nunca usa una tasa posterior a la fecha consultada.
    * Cada conversión puede devolver las tasas exactas que usó (auditoría).
    * Se admiten varios tipos de tasa (p. ej. "BCV", "MERCADO") por moneda.
    * La conversión USD <-> EUR se hace por tasa cruzada vía bolívares.

Convención: `valor` = bolívares por 1 unidad de la moneda extranjera.
"""
from __future__ import annotations

from bisect import bisect_right, insort
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from typing import Any, Iterable, Optional

from conciliar import CENTAVO, ErrorDeParseo, _es_vacio, parsear_fecha, parsear_monto, parsear_numero

MONEDA_BASE = "VES"
MONEDAS_REPORTE = ("VES", "USD", "EUR")


class TasaNoDisponible(LookupError):
    """No hay una tasa válida para la fecha / moneda / tipo pedidos."""


@dataclass(frozen=True)
class TasaAplicada:
    moneda: str
    tipo: str
    valor: Decimal  # Bs. por 1 unidad de `moneda`
    fecha_tasa: date  # fecha de la tasa realmente usada
    fecha_consulta: date  # fecha para la que se pidió
    fuente: str

    @property
    def arrastrada(self) -> bool:
        return self.fecha_tasa != self.fecha_consulta

    @property
    def dias_arrastre(self) -> int:
        return (self.fecha_consulta - self.fecha_tasa).days


@dataclass
class ErrorTasa:
    fila: int
    columna: str
    valor: Any
    motivo: str


class TablaTasas:
    def __init__(self, max_antiguedad_dias: int = 7) -> None:
        self.max_antiguedad_dias = max_antiguedad_dias
        self._fechas: dict[tuple[str, str], list[date]] = {}
        self._valores: dict[tuple[str, str, date], tuple[Decimal, str]] = {}

    # ---- carga ----------------------------------------------------------
    def agregar(self, fecha: Any, moneda: str, valor: Any, tipo: str = "BCV", fuente: str = "manual") -> None:
        f = parsear_fecha(fecha)
        v = valor if isinstance(valor, Decimal) else parsear_numero(valor, un_separador_es_decimal=True)
        if v <= 0:
            raise ValueError(f"La tasa debe ser mayor que cero ({moneda} {f}: {valor!r}).")
        moneda = moneda.upper()
        if moneda == MONEDA_BASE:
            raise ValueError("La moneda base (VES) no lleva tasa.")
        clave = (tipo, moneda)
        fechas = self._fechas.setdefault(clave, [])
        if (tipo, moneda, f) not in self._valores:
            insort(fechas, f)
        self._valores[(tipo, moneda, f)] = (v, fuente)

    def cargar_filas(
        self,
        filas: Iterable[dict],
        columna_fecha: str,
        columnas_moneda: dict[str, str],
        tipo: str = "BCV",
        fuente: str = "archivo",
    ) -> list[ErrorTasa]:
        """
        Carga una tabla tipo  Fecha | USD | EUR  con columnas_moneda={"USD": "Dolar", "EUR": "Euro"}.
        Una celda vacía se ignora; una celda ilegible se reporta (no se pone en 0).
        """
        errores: list[ErrorTasa] = []
        for n, fila in enumerate(filas, start=1):
            try:
                f = parsear_fecha(fila.get(columna_fecha))
            except ErrorDeParseo as e:
                errores.append(ErrorTasa(n, columna_fecha, fila.get(columna_fecha), str(e)))
                continue
            for moneda, col in columnas_moneda.items():
                celda = fila.get(col)
                if _es_vacio(celda):
                    continue
                try:
                    self.agregar(f, moneda, celda, tipo, fuente)
                except (ErrorDeParseo, ValueError) as e:
                    errores.append(ErrorTasa(n, col, celda, str(e)))
        return errores

    # ---- consulta -------------------------------------------------------
    def obtener(self, fecha: Any, moneda: str, tipo: str = "BCV") -> TasaAplicada:
        f = parsear_fecha(fecha)
        moneda = moneda.upper()
        fechas = self._fechas.get((tipo, moneda))
        if not fechas:
            raise TasaNoDisponible(f"No hay tasas cargadas para {moneda} ({tipo}).")
        i = bisect_right(fechas, f) - 1
        if i < 0:
            raise TasaNoDisponible(
                f"No hay tasa de {moneda} ({tipo}) en o antes de {f}; la primera es del {fechas[0]}."
            )
        f_tasa = fechas[i]
        if (f - f_tasa).days > self.max_antiguedad_dias:
            raise TasaNoDisponible(
                f"La última tasa de {moneda} ({tipo}) anterior a {f} es del {f_tasa}, "
                f"con más de {self.max_antiguedad_dias} días de antigüedad."
            )
        valor, fuente = self._valores[(tipo, moneda, f_tasa)]
        return TasaAplicada(moneda, tipo, valor, f_tasa, f, fuente)

    # ---- conversión -----------------------------------------------------
    def convertir_con_detalle(
        self, monto: Any, desde: str, a: str, fecha: Any, tipo: str = "BCV"
    ) -> tuple[Decimal, tuple[TasaAplicada, ...]]:
        """Devuelve (monto convertido a 2 decimales, tasas usadas)."""
        desde, a = desde.upper(), a.upper()
        m = parsear_monto(monto)
        if desde == a:
            return m, ()
        usadas: list[TasaAplicada] = []
        en_bs = m
        if desde != MONEDA_BASE:
            t = self.obtener(fecha, desde, tipo)
            usadas.append(t)
            en_bs = m * t.valor
        resultado = en_bs
        if a != MONEDA_BASE:
            t = self.obtener(fecha, a, tipo)
            usadas.append(t)
            resultado = en_bs / t.valor
        return resultado.quantize(CENTAVO, rounding=ROUND_HALF_UP), tuple(usadas)

    def convertir(self, monto: Any, desde: str, a: str, fecha: Any, tipo: str = "BCV") -> Decimal:
        return self.convertir_con_detalle(monto, desde, a, fecha, tipo)[0]

    def tres_monedas(self, monto: Any, moneda: str, fecha: Any, tipo: str = "BCV") -> dict[str, Decimal]:
        """{'VES': ..., 'USD': ..., 'EUR': ...} para un monto en cualquiera de las tres."""
        return {m: self.convertir(monto, moneda, m, fecha, tipo) for m in MONEDAS_REPORTE}

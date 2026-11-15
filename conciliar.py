"""
conciliar.py - Motor de conciliación (núcleo puro).
"""
from __future__ import annotations

import math
import numbers
import re
import unicodedata
from bisect import bisect_left, bisect_right
from collections import defaultdict
from dataclasses import dataclass, field, replace
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from difflib import SequenceMatcher
from itertools import combinations
from typing import Any, Iterable, Optional, Sequence

CENTAVO = Decimal("0.01")


class ErrorDeParseo(ValueError):
    """Un valor no pudo interpretarse."""


def _es_vacio(v: Any) -> bool:
    if v is None:
        return True
    if isinstance(v, str):
        return not v.strip()
    try:
        return bool(v != v)
    except Exception:
        return False


def parsear_numero(
    valor: Any,
    separador_decimal: str = ",",
    un_separador_es_decimal: bool = False,
) -> Decimal:
    if _es_vacio(valor):
        raise ErrorDeParseo("valor vacío")
    if isinstance(valor, bool):
        raise ErrorDeParseo(f"valor no numérico: {valor!r}")
    if isinstance(valor, Decimal):
        if not valor.is_finite():
            raise ErrorDeParseo(f"número no finito: {valor!r}")
        return valor
    if isinstance(valor, numbers.Integral):
        return Decimal(int(valor))
    if isinstance(valor, numbers.Real):
        f = float(valor)
        if not math.isfinite(f):
            raise ErrorDeParseo(f"número no finito: {valor!r}")
        return Decimal(repr(f))

    texto = str(valor).strip()
    negativo = texto.startswith("(") and texto.endswith(")")
    texto = re.sub(r"[^\W\d_]+\.?|[$€£]", " ", texto)
    limpio = re.sub(r"[^\d.,+\-]", "", texto)

    if "-" in limpio:
        if limpio.count("-") == 1 and (limpio.startswith("-") or limpio.endswith("-")):
            negativo = True
        else:
            raise ErrorDeParseo(f"signo '-' en posición inválida: {valor!r}")
        limpio = limpio.replace("-", "")
    limpio = limpio.replace("+", "").rstrip(".,")

    if not re.search(r"\d", limpio):
        raise ErrorDeParseo(f"no contiene dígitos: {valor!r}")

    punto, coma = "." in limpio, "," in limpio
    if punto and coma:
        decimal: Optional[str] = "." if limpio.rfind(".") > limpio.rfind(",") else ","
    elif punto or coma:
        sep = "." if punto else ","
        partes = limpio.split(sep)
        if len(partes) > 2:
            decimal = None
        elif un_separador_es_decimal or len(partes[1]) != 3:
            decimal = sep
        else:
            decimal = sep if sep == separador_decimal else None
    else:
        decimal = None

    if decimal is None:
        numero = limpio.replace(".", "").replace(",", "")
    else:
        miles = "," if decimal == "." else "."
        numero = limpio.replace(miles, "").replace(decimal, ".")

    try:
        d = Decimal(numero)
    except InvalidOperation:
        raise ErrorDeParseo(f"no se pudo interpretar el número: {valor!r}") from None
    return -d if negativo else d


def parsear_monto(valor: Any, separador_decimal: str = ",") -> Decimal:
    return parsear_numero(valor, separador_decimal).quantize(CENTAVO, rounding=ROUND_HALF_UP)


_FORMATOS_ISO = ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S")
_FORMATOS_DIA_PRIMERO = (
    "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d-%m-%y",
    "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d-%m-%Y %H:%M:%S",
)
_FORMATOS_MES_PRIMERO = ("%m/%d/%Y", "%m-%d-%Y", "%m/%d/%y", "%m/%d/%Y %H:%M:%S")


def parsear_fecha(valor: Any, dayfirst: bool = True) -> date:
    if _es_vacio(valor):
        raise ErrorDeParseo("fecha vacía")
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    if isinstance(valor, numbers.Real) and not isinstance(valor, bool):
        n = int(valor)
        if 30000 <= n <= 80000:
            return date(1899, 12, 30) + timedelta(days=n)
        valor = str(n)

    texto = str(valor).strip()
    if re.fullmatch(r"(19|20)\d{6}", texto):
        try:
            return datetime.strptime(texto, "%Y%m%d").date()
        except ValueError:
            pass
    formatos = _FORMATOS_ISO + (_FORMATOS_DIA_PRIMERO if dayfirst else _FORMATOS_MES_PRIMERO)
    for fmt in formatos:
        try:
            return datetime.strptime(texto, fmt).date()
        except ValueError:
            continue
    raise ErrorDeParseo(f"fecha no reconocida: {valor!r}")


def normalizar_texto(texto: Any) -> str:
    if _es_vacio(texto):
        return ""
    t = unicodedata.normalize("NFKD", str(texto))
    t = "".join(c for c in t if not unicodedata.combining(c)).lower()
    return re.sub(r"[^a-z0-9]+", " ", t).strip()


def normalizar_referencia(ref: Any) -> str:
    if _es_vacio(ref):
        return ""
    if isinstance(ref, float) and ref.is_integer():
        ref = int(ref)
    return re.sub(r"[^A-Z0-9]", "", str(ref).upper()).lstrip("0")


def referencias_coinciden(a: str, b: str, min_sufijo: int = 6) -> bool:
    if not a or not b:
        return False
    if a == b:
        return True
    corta, larga = sorted((a, b), key=len)
    return len(corta) >= min_sufijo and larga.endswith(corta)


_RE_COMISION = re.compile(
    r"\b(comis\w*|com|comserv\w*|sms|tarifa\w*|mantenimiento (de )?cuenta|"
    r"emision (de )?(estado|edo) (de )?(cuenta|cta))\b"
)
_RE_IMPUESTO = re.compile(r"\b(igtf|itf|impuesto\w*)\b")
PATRON_REVERSO_BASE = r"anul\w*|revers\w*|cancelar entrada|nul[oa]"
_RE_REVERSO = re.compile(r"\b(" + PATRON_REVERSO_BASE + r")\b")


def clasificar_movimiento(descripcion: Any, monto: Decimal) -> str:
    t = normalizar_texto(descripcion)
    if _RE_COMISION.search(t):
        return "Comisiones"
    if _RE_IMPUESTO.search(t):
        return "Impuestos"
    if monto > 0:
        return "Ingresos"
    if monto < 0:
        return "Egresos"
    return "Sin movimiento"


@dataclass(frozen=True, eq=False)
class Movimiento:
    id: str
    origen: str
    fecha: date
    monto: Decimal
    descripcion: str = ""
    referencia: str = ""
    moneda: str = "VES"
    datos: dict = field(default_factory=dict)
    referencia_norm: str = field(init=False, default="")
    descripcion_norm: str = field(init=False, default="")

    def __post_init__(self) -> None:
        object.__setattr__(self, "fecha", parsear_fecha(self.fecha))
        object.__setattr__(self, "monto", parsear_monto(self.monto))
        object.__setattr__(self, "referencia_norm", normalizar_referencia(self.referencia))
        object.__setattr__(self, "descripcion_norm", normalizar_texto(self.descripcion))

    @property
    def tipo(self) -> str:
        return clasificar_movimiento(self.descripcion, self.monto)


@dataclass
class Coincidencia:
    regla: str
    banco: tuple[Movimiento, ...]
    libros: tuple[Movimiento, ...]
    confianza: str
    requiere_revision: bool
    motivo: str = ""
    ambigua: bool = False

    @property
    def diferencia_monto(self) -> Decimal:
        return sum((m.monto for m in self.banco), Decimal("0")) - sum(
            (m.monto for m in self.libros), Decimal("0")
        )

    @property
    def dias_diferencia(self) -> int:
        fechas = [m.fecha for m in self.banco + self.libros]
        return (max(fechas) - min(fechas)).days


@dataclass
class Config:
    tolerancia_dias_c2: int = 3
    exigir_referencia_c2: bool = False
    ventana_dias_c3: int = 7
    tolerancia_monto_c3: Decimal = Decimal("1.00")
    similitud_minima_c3: float = 0.5
    max_grupo_c4: int = 4
    max_candidatos_c4: int = 20
    ventana_dias_c4: int = 5
    presupuesto_combinaciones_c4: int = 1_000_000
    neutralizar_reversos: bool = True
    ventana_reversos_dias: int = 10
    patron_reverso_banco: str = PATRON_REVERSO_BASE + r"|devoluc\w*"
    patron_reverso_libros: str = PATRON_REVERSO_BASE


@dataclass
class Resultado:
    pares: list[Coincidencia]
    sugerencias: list[Coincidencia]
    banco_pendiente: list[Movimiento]
    libros_pendiente: list[Movimiento]
    reversos_banco: list[tuple[Movimiento, Movimiento]]
    reversos_libros: list[tuple[Movimiento, Movimiento]]
    advertencias: list[str]

    def resumen(self) -> dict[str, int]:
        por_regla: dict[str, int] = defaultdict(int)
        for p in self.pares:
            por_regla[p.regla] += 1
        return {
            "pares_cerrados": len(self.pares),
            **{f"cerrados_{k}": v for k, v in sorted(por_regla.items())},
            "sugerencias": len(self.sugerencias),
            "banco_pendiente": len(self.banco_pendiente),
            "libros_pendiente": len(self.libros_pendiente),
            "reversos_neutralizados": len(self.reversos_banco) + len(self.reversos_libros),
        }

    def aprobar(self, sugerencia: Coincidencia) -> None:
        if sugerencia not in self.sugerencias:
            raise ValueError("La sugerencia no pertenece a este resultado.")
        self.sugerencias.remove(sugerencia)
        self.pares.append(
            replace(sugerencia, requiere_revision=False, motivo=sugerencia.motivo + " | aprobada")
        )
        ids_b = {m.id for m in sugerencia.banco}
        ids_l = {m.id for m in sugerencia.libros}
        self.banco_pendiente = [m for m in self.banco_pendiente if m.id not in ids_b]
        self.libros_pendiente = [m for m in self.libros_pendiente if m.id not in ids_l]
        self.sugerencias = [
            s for s in self.sugerencias
            if not ({m.id for m in s.banco} & ids_b or {m.id for m in s.libros} & ids_l)
        ]

    def rechazar(self, sugerencia: Coincidencia) -> None:
        if sugerencia not in self.sugerencias:
            raise ValueError("La sugerencia no pertenece a este resultado.")
        self.sugerencias.remove(sugerencia)


def _orden(m: Movimiento) -> tuple:
    return (m.fecha, m.id)


def neutralizar_reversos(
    movs: Sequence[Movimiento], ventana_dias: int = 10, patron: Optional[str] = None
) -> tuple[list[Movimiento], list[tuple[Movimiento, Movimiento]], list[Movimiento]]:
    movs = list(movs)
    re_rev = re.compile(r"\b(" + patron + r")\b") if patron else _RE_REVERSO
    es_rev = {m.id for m in movs if re_rev.search(m.descripcion_norm)}
    por_monto: dict[Decimal, list[Movimiento]] = defaultdict(list)
    for m in movs:
        por_monto[m.monto].append(m)

    usados: set[str] = set()
    pares: list[tuple[Movimiento, Movimiento]] = []
    for r in sorted((m for m in movs if m.id in es_rev), key=_orden):
        if r.id in usados:
            continue
        mejor, mejor_clave = None, None
        for o in por_monto.get(-r.monto, ()):
            if o.id in usados or o.id == r.id:
                continue
            dias = abs((o.fecha - r.fecha).days)
            if dias > ventana_dias:
                continue
            clave = (
                o.id in es_rev,
                not referencias_coinciden(o.referencia_norm, r.referencia_norm),
                dias,
                o.id,
            )
            if mejor_clave is None or clave < mejor_clave:
                mejor, mejor_clave = o, clave
        if mejor is not None:
            usados.update((r.id, mejor.id))
            pares.append((mejor, r))

    restantes = [m for m in movs if m.id not in usados]
    sin_original = [m for m in restantes if m.id in es_rev]
    return restantes, pares, sin_original


def _validar_ids(movs: Sequence[Movimiento], nombre: str) -> None:
    vistos: set[str] = set()
    for m in movs:
        if m.id in vistos:
            raise ValueError(f"id duplicado en {nombre}: {m.id!r}")
        vistos.add(m.id)


def _indice_por_monto(libros: Sequence[Movimiento]) -> dict[Decimal, tuple[list[int], list[Movimiento]]]:
    cubos: dict[Decimal, list[Movimiento]] = defaultdict(list)
    for m in sorted(libros, key=_orden):
        cubos[m.monto].append(m)
    return {k: ([m.fecha.toordinal() for m in v], v) for k, v in cubos.items()}


def _cruzar_exacto(
    banco: Sequence[Movimiento],
    libros: Sequence[Movimiento],
    tolerancia_dias: int,
    regla: str,
    exigir_referencia: bool,
) -> tuple[list[Coincidencia], list[Movimiento], list[Movimiento]]:
    indice = _indice_por_monto(libros)
    usados_l: set[str] = set()
    pares: list[Coincidencia] = []
    sin_par: list[Movimiento] = []

    for b in banco:
        cubo = indice.get(b.monto)
        mejor, mejor_clave, mejor_ref = None, None, False
        claves: list[tuple] = []
        if cubo is not None:
            ords, lista = cubo
            o = b.fecha.toordinal()
            lo, hi = bisect_left(ords, o - tolerancia_dias), bisect_right(ords, o + tolerancia_dias)
            for l in lista[lo:hi]:
                if l.id in usados_l:
                    continue
                ref_ok = referencias_coinciden(b.referencia_norm, l.referencia_norm)
                if exigir_referencia and not ref_ok:
                    continue
                clave = (0 if ref_ok else 1, abs((b.fecha - l.fecha).days), l.id)
                claves.append(clave[:2])
                if mejor_clave is None or clave < mejor_clave:
                    mejor, mejor_clave, mejor_ref = l, clave, ref_ok
        if mejor is None:
            sin_par.append(b)
            continue
        usados_l.add(mejor.id)
        equivalentes = claves.count(mejor_clave[:2])
        ambigua = equivalentes > 1
        if regla == "C1":
            confianza, motivo = "alta", "referencia + monto + signo + fecha"
        else:
            confianza = "alta" if mejor_ref else "media"
            motivo = ("referencia + " if mejor_ref else "") + (
                f"monto + signo + fecha (±{abs((b.fecha - mejor.fecha).days)} d)"
            )
        if ambigua:
            confianza = "media"
            motivo += f" | ambigua: {equivalentes} candidatos equivalentes"
        pares.append(Coincidencia(regla, (b,), (mejor,), confianza, False, motivo, ambigua))

    libros_sin_par = [l for l in libros if l.id not in usados_l]
    return pares, sin_par, libros_sin_par


def _sugerir_aproximadas(
    banco: Sequence[Movimiento], libros: Sequence[Movimiento], cfg: Config
) -> list[Coincidencia]:
    libros_ord = sorted(libros, key=lambda m: m.fecha)
    ords = [m.fecha.toordinal() for m in libros_ord]
    usados_l: set[str] = set()
    sugerencias: list[Coincidencia] = []

    for b in banco:
        o = b.fecha.toordinal()
        lo = bisect_left(ords, o - cfg.ventana_dias_c3)
        hi = bisect_right(ords, o + cfg.ventana_dias_c3)
        mejor, mejor_clave, mejor_info = None, None, None
        for l in libros_ord[lo:hi]:
            if l.id in usados_l or (b.monto > 0) != (l.monto > 0):
                continue
            dif = abs(b.monto - l.monto)
            if dif > cfg.tolerancia_monto_c3:
                continue
            ref_ok = referencias_coinciden(b.referencia_norm, l.referencia_norm)
            sim = 0.0
            if b.descripcion_norm and l.descripcion_norm:
                sim = SequenceMatcher(None, b.descripcion_norm, l.descripcion_norm).ratio()
            if not (ref_ok or sim >= cfg.similitud_minima_c3 or dif == 0):
                continue
            dias = abs((b.fecha - l.fecha).days)
            clave = (ref_ok, -dif, sim, -dias, l.id)
            if mejor_clave is None or clave > mejor_clave:
                mejor, mejor_clave, mejor_info = l, clave, (ref_ok, dif, sim, dias)
        if mejor is not None:
            ref_ok, dif, sim, dias = mejor_info
            partes = []
            if ref_ok:
                partes.append("referencia coincide")
            if sim >= cfg.similitud_minima_c3:
                partes.append(f"descripción similar ({sim:.2f})")
            if dif:
                partes.append(f"diferencia de monto {dif}")
            if dias > cfg.tolerancia_dias_c2:
                partes.append(f"fecha a {dias} días (fuera de tolerancia C2)")
            usados_l.add(mejor.id)
            sugerencias.append(Coincidencia(
                "C3", (b,), (mejor,), "media" if ref_ok else "baja", True, "; ".join(partes),
            ))
    return sugerencias


def _buscar_grupo(
    unico: Movimiento,
    pool_ord: Sequence[Movimiento],
    pool_ords: Sequence[int],
    bloqueados: set[str],
    cfg: Config,
    presupuesto: list[int],
) -> Optional[tuple[Movimiento, ...]]:
    if presupuesto[0] < 0:
        return None
    o = unico.fecha.toordinal()
    lo = bisect_left(pool_ords, o - cfg.ventana_dias_c4)
    hi = bisect_right(pool_ords, o + cfg.ventana_dias_c4)
    cands = [
        m for m in pool_ord[lo:hi]
        if m.id not in bloqueados
        and (m.monto > 0) == (unico.monto > 0)
        and abs(m.monto) < abs(unico.monto)
    ]
    if len(cands) < 2:
        return None
    cands.sort(key=lambda m: (abs((m.fecha - unico.fecha).days), m.id))
    cands = cands[: cfg.max_candidatos_c4]
    for n in range(2, min(cfg.max_grupo_c4, len(cands)) + 1):
        for combo in combinations(cands, n):
            presupuesto[0] -= 1
            if presupuesto[0] < 0:
                return None
            if sum((m.monto for m in combo), Decimal("0")) == unico.monto:
                return combo
    return None


def _sugerir_agrupadas(
    banco: Sequence[Movimiento],
    libros: Sequence[Movimiento],
    cfg: Config,
    bloqueados_b: set[str],
    bloqueados_l: set[str],
) -> tuple[list[Coincidencia], bool]:
    sug: list[Coincidencia] = []
    usados_b, usados_l = set(bloqueados_b), set(bloqueados_l)
    presupuesto = [cfg.presupuesto_combinaciones_c4]

    libros_ord = sorted(libros, key=lambda m: m.fecha)
    libros_ords = [m.fecha.toordinal() for m in libros_ord]
    for b in banco:
        if b.id in usados_b:
            continue
        g = _buscar_grupo(b, libros_ord, libros_ords, usados_l, cfg, presupuesto)
        if g:
            usados_b.add(b.id)
            usados_l.update(m.id for m in g)
            sug.append(Coincidencia("C4", (b,), g, "baja", True, f"1 banco = {len(g)} partidas de libros"))

    banco_ord = sorted(banco, key=lambda m: m.fecha)
    banco_ords = [m.fecha.toordinal() for m in banco_ord]
    for l in libros:
        if l.id in usados_l:
            continue
        g = _buscar_grupo(l, banco_ord, banco_ords, usados_b, cfg, presupuesto)
        if g:
            usados_l.add(l.id)
            usados_b.update(m.id for m in g)
            sug.append(Coincidencia("C4", g, (l,), "baja", True, f"{len(g)} movimientos de banco = 1 partida de libros"))
    return sug, presupuesto[0] < 0


def conciliar(
    banco: Iterable[Movimiento],
    libros: Iterable[Movimiento],
    config: Optional[Config] = None,
) -> Resultado:
    cfg = config or Config()
    banco, libros = list(banco), list(libros)
    _validar_ids(banco, "banco")
    _validar_ids(libros, "libros")
    advertencias: list[str] = []

    rev_b: list[tuple[Movimiento, Movimiento]] = []
    rev_l: list[tuple[Movimiento, Movimiento]] = []
    if cfg.neutralizar_reversos:
        banco, rev_b, sin_b = neutralizar_reversos(
            banco, cfg.ventana_reversos_dias, cfg.patron_reverso_banco)
        libros, rev_l, sin_l = neutralizar_reversos(
            libros, cfg.ventana_reversos_dias, cfg.patron_reverso_libros)
        for nombre, sin in (("banco", sin_b), ("libros", sin_l)):
            if sin:
                advertencias.append(
                    f"{len(sin)} movimiento(s) de {nombre} parecen reversos pero no se encontró "
                    f"su original; se dejan en el cruce normal: {', '.join(m.id for m in sin[:5])}"
                )

    banco = sorted(banco, key=_orden)
    libros = sorted(libros, key=_orden)

    c1, banco, libros = _cruzar_exacto(banco, libros, 0, "C1", exigir_referencia=True)
    c2, banco, libros = _cruzar_exacto(banco, libros, cfg.tolerancia_dias_c2, "C2", True)
    if not cfg.exigir_referencia_c2:
        for dias in range(cfg.tolerancia_dias_c2 + 1):
            par, banco, libros = _cruzar_exacto(banco, libros, dias, "C2", False)
            c2 += par
    c3 = _sugerir_aproximadas(banco, libros, cfg)
    c4, agotado = _sugerir_agrupadas(
        banco, libros, cfg,
        bloqueados_b={m.id for s in c3 for m in s.banco},
        bloqueados_l={m.id for s in c3 for m in s.libros},
    )

    return Resultado(
        pares=c1 + c2,
        sugerencias=c3 + c4,
        banco_pendiente=banco,
        libros_pendiente=libros,
        reversos_banco=rev_b,
        reversos_libros=rev_l,
        advertencias=advertencias,
    )


@dataclass(frozen=True)
class Cuadre:
    saldo_banco: Decimal
    saldo_libros: Decimal
    partidas_libros_no_en_banco: Decimal
    partidas_banco_no_en_libros: Decimal
    diferencias_en_pares: Decimal
    saldo_banco_ajustado: Decimal
    saldo_libros_ajustado: Decimal
    diferencia: Decimal

    @property
    def cuadrado(self) -> bool:
        return self.diferencia == 0


def calcular_cuadre(resultado: Resultado, saldo_banco: Any, saldo_libros: Any) -> Cuadre:
    sb, sl = parsear_monto(saldo_banco), parsear_monto(saldo_libros)
    cero = Decimal("0")
    p_libros = sum((m.monto for m in resultado.libros_pendiente), cero)
    p_banco = sum((m.monto for m in resultado.banco_pendiente), cero)
    d_pares = sum((p.diferencia_monto for p in resultado.pares), cero)
    aj_b = sb + p_libros
    aj_l = sl + p_banco + d_pares
    return Cuadre(sb, sl, p_libros, p_banco, d_pares, aj_b, aj_l, aj_b - aj_l)


@dataclass
class ErrorFila:
    fila: int
    columna: str
    valor: Any
    motivo: str


@dataclass
class SaldoDetectado:
    fila: int
    etiqueta: str
    fecha: Optional[date]
    monto: Decimal


@dataclass
class Carga:
    movimientos: list[Movimiento]
    errores: list[ErrorFila]
    omitidas: int
    saldos: list[SaldoDetectado] = field(default_factory=list)
    omitidas_detalle: list[tuple[int, str]] = field(default_factory=list)
    saldo_final_columna: Optional[Decimal] = None


def _texto(v: Any) -> str:
    if _es_vacio(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def cargar_movimientos(
    filas: Iterable[dict],
    origen: str,
    mapeo: dict[str, str],
    *,
    separador_decimal: str = ",",
    dayfirst: bool = True,
    moneda: str = "VES",
    prefijo_id: Optional[str] = None,
    patron_saldo: Optional[str] = r"\bsaldo (inicial|final|anterior)\b",
) -> Carga:
    if "fecha" not in mapeo:
        raise ValueError("El mapeo debe incluir 'fecha'.")
    if not (mapeo.get("monto") or mapeo.get("entrada") or mapeo.get("salida")):
        raise ValueError("El mapeo debe incluir 'monto' o 'entrada'/'salida'.")

    prefijo = prefijo_id or ("B" if origen == "banco" else "L")
    movs: list[Movimiento] = []
    errores: list[ErrorFila] = []
    omitidas = 0
    columnas_ok = False
    saldos: list[SaldoDetectado] = []
    omitidas_detalle: list[tuple[int, str]] = []
    saldo_final_columna: Optional[Decimal] = None
    col_saldo = mapeo.get("saldo")
    re_saldo = re.compile(patron_saldo) if patron_saldo else None

    for n, fila in enumerate(filas, start=1):
        if not columnas_ok:
            faltan = [c for c in mapeo.values() if c and c not in fila]
            if faltan:
                raise ValueError(
                    f"Columnas no encontradas en el archivo de {origen}: {faltan}. "
                    f"Disponibles: {list(fila.keys())}"
                )
            columnas_ok = True

        col_monto = [c for c in (mapeo.get("monto"), mapeo.get("entrada"), mapeo.get("salida")) if c]
        celdas_monto = [fila.get(c) for c in col_monto]
        sin_monto = all(_es_vacio(c) for c in celdas_monto)
        valor_fecha = fila.get(mapeo["fecha"])
        desc_cruda = _texto(fila.get(mapeo["descripcion"])) if mapeo.get("descripcion") else ""

        if re_saldo is not None and re_saldo.search(normalizar_texto(desc_cruda)):
            celda = fila.get(col_saldo) if col_saldo and not _es_vacio(fila.get(col_saldo)) else (
                fila.get(mapeo["monto"]) if mapeo.get("monto") else None
            )
            try:
                f_saldo = parsear_fecha(valor_fecha, dayfirst) if not _es_vacio(valor_fecha) else None
            except ErrorDeParseo:
                f_saldo = None
            try:
                saldos.append(SaldoDetectado(n, normalizar_texto(desc_cruda), f_saldo,
                                             parsear_monto(celda, separador_decimal)))
            except ErrorDeParseo as e:
                errores.append(ErrorFila(n, "saldo", celda, str(e)))
            continue

        if sin_monto:
            try:
                fecha_ok = not _es_vacio(valor_fecha) and parsear_fecha(valor_fecha, dayfirst) is not None
            except ErrorDeParseo:
                fecha_ok = False
            if fecha_ok:
                errores.append(ErrorFila(n, "/".join(col_monto), celdas_monto,
                                         "fila con fecha pero sin monto"))
                continue
            omitidas += 1
            if _es_vacio(valor_fecha):
                omitidas_detalle.append((n, "fila vacía"))
            else:
                omitidas_detalle.append((n, f"sin monto y sin fecha válida ({_texto(valor_fecha)!r})"))
            continue

        try:
            fecha = parsear_fecha(valor_fecha, dayfirst)
        except ErrorDeParseo as e:
            errores.append(ErrorFila(n, mapeo["fecha"], valor_fecha, str(e)))
            continue

        try:
            if mapeo.get("monto"):
                monto = parsear_monto(fila.get(mapeo["monto"]), separador_decimal)
            else:
                partes = {}
                for clave in ("entrada", "salida"):
                    col = mapeo.get(clave)
                    v = fila.get(col) if col else None
                    partes[clave] = (
                        Decimal("0") if _es_vacio(v) else abs(parsear_monto(v, separador_decimal))
                    )
                if all(_es_vacio(c) for c in celdas_monto):
                    raise ErrorDeParseo("fila sin monto")
                monto = partes["entrada"] - partes["salida"]
        except ErrorDeParseo as e:
            errores.append(ErrorFila(n, "/".join(col_monto), celdas_monto, str(e)))
            continue

        if monto == 0:
            omitidas += 1
            omitidas_detalle.append((n, "monto cero"))
            continue

        if col_saldo and not _es_vacio(fila.get(col_saldo)):
            try:
                saldo_final_columna = parsear_monto(fila.get(col_saldo), separador_decimal)
            except ErrorDeParseo:
                pass

        movs.append(Movimiento(
            id=f"{prefijo}-{n:05d}",
            origen=origen,
            fecha=fecha,
            monto=monto,
            descripcion=desc_cruda,
            referencia=_texto(fila.get(mapeo["referencia"])) if mapeo.get("referencia") else "",
            moneda=moneda,
            datos=dict(fila),
        ))
    return Carga(movs, errores, omitidas, saldos, omitidas_detalle, saldo_final_columna)
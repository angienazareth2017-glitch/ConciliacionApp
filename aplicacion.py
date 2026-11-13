"""
conciliar.py - Motor de conciliación (núcleo puro).

Sin Streamlit ni pandas: solo biblioteca estándar. Así se puede probar
con tests automáticos y reutilizar desde Streamlit, FastAPI o un proceso batch.

Convención de signos (única para banco y libros):
    monto > 0  -> entra dinero a la cuenta (ingreso)
    monto < 0  -> sale dinero de la cuenta (egreso)

Principios:
    1. Nada se convierte en cero en silencio: un valor ilegible genera un error
       reportado, no un 0.0.
    2. El dinero es Decimal con 2 decimales, nunca float.
    3. Solo C1 y C2 cierran partidas solas. C3 y C4 son SUGERENCIAS que un
       humano aprueba o rechaza.
    4. Cada coincidencia indica qué regla la produjo.
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


# =============================================================================
# 1. PARSEO DE VALORES (montos, fechas, textos)
# =============================================================================
class ErrorDeParseo(ValueError):
    """Un valor no pudo interpretarse. Nunca se reemplaza por cero en silencio."""


def _es_vacio(v: Any) -> bool:
    if v is None:
        return True
    if isinstance(v, str):
        return not v.strip()
    try:
        return bool(v != v)  # NaN / NaT
    except Exception:
        return False


def parsear_numero(
    valor: Any,
    separador_decimal: str = ",",
    un_separador_es_decimal: bool = False,
) -> Decimal:
    """
    Interpreta un número escrito en formato venezolano (1.234,56), anglosajón
    (1,234.56), con símbolos de moneda (Bs., Bs.S, VES, $, USD, €), negativo con
    paréntesis (1.234,56) o con signo al final (1.234,56-).

    Regla para el caso ambiguo (un solo separador seguido de exactamente 3
    dígitos, como "1.234"): se usa `separador_decimal` para decidir. Con
    `un_separador_es_decimal=True` (útil para tasas de cambio) un separador
    único siempre es decimal.

    Lanza ErrorDeParseo si el valor no es interpretable. No devuelve 0.
    """
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
        return Decimal(repr(f))  # repr evita el ruido binario de float

    texto = str(valor).strip()
    negativo = texto.startswith("(") and texto.endswith(")")
    # Quita letras (Bs., Bs.S, VES, USD...) y símbolos de moneda.
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
            decimal = None  # 1.234.567 -> separador de miles
        elif un_separador_es_decimal or len(partes[1]) != 3:
            decimal = sep
        else:  # ambiguo: 1.234 / 1,234
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
    """Como parsear_numero, redondeado a centavos (ROUND_HALF_UP)."""
    return parsear_numero(valor, separador_decimal).quantize(CENTAVO, rounding=ROUND_HALF_UP)


_FORMATOS_ISO = ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S")
_FORMATOS_DIA_PRIMERO = (
    "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d-%m-%y",
    "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d-%m-%Y %H:%M:%S",
)
_FORMATOS_MES_PRIMERO = ("%m/%d/%Y", "%m-%d-%Y", "%m/%d/%y", "%m/%d/%Y %H:%M:%S")


def parsear_fecha(valor: Any, dayfirst: bool = True) -> date:
    """
    Convierte a `date`. Por defecto asume día/mes/año (formato venezolano):
    "05/03/2025" es 5 de marzo, no 3 de mayo. Los formatos ISO (2025-03-05)
    no son ambiguos y se aceptan siempre. Acepta también datetime, date,
    serial de Excel y yyyymmdd.
    """
    if _es_vacio(valor):
        raise ErrorDeParseo("fecha vacía")
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    if isinstance(valor, numbers.Real) and not isinstance(valor, bool):
        n = int(valor)
        if 30000 <= n <= 80000:  # serial de Excel (1982-2119)
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
    """Minúsculas, sin tildes ni signos: 'COMISIÓN  Mant.' -> 'comision mant'."""
    if _es_vacio(texto):
        return ""
    t = unicodedata.normalize("NFKD", str(texto))
    t = "".join(c for c in t if not unicodedata.combining(c)).lower()
    return re.sub(r"[^a-z0-9]+", " ", t).strip()


def normalizar_referencia(ref: Any) -> str:
    """Solo letras y dígitos en mayúscula, sin ceros a la izquierda."""
    if _es_vacio(ref):
        return ""
    if isinstance(ref, float) and ref.is_integer():
        ref = int(ref)  # Excel entrega 123456.0
    return re.sub(r"[^A-Z0-9]", "", str(ref).upper()).lstrip("0")


def referencias_coinciden(a: str, b: str, min_sufijo: int = 6) -> bool:
    """
    Iguales, o una es sufijo de la otra (los bancos suelen truncar referencias
    a los últimos dígitos). El sufijo debe tener al menos `min_sufijo` caracteres.
    """
    if not a or not b:
        return False
    if a == b:
        return True
    corta, larga = sorted((a, b), key=len)
    return len(corta) >= min_sufijo and larga.endswith(corta)


# =============================================================================
# 2. CLASIFICACIÓN DE MOVIMIENTOS
# =============================================================================
# Se busca por PALABRA COMPLETA (\b). El código anterior buscaba la subcadena
# "com", que marcaba como comisión "COMERCIO", "COMPRA POS" o "COMPAÑIA".
_RE_COMISION = re.compile(
    r"\b(comis\w*|com|comserv\w*|sms|tarifa\w*|mantenimiento (de )?cuenta|"
    r"emision (de )?(estado|edo) (de )?(cuenta|cta))\b"
)
_RE_IMPUESTO = re.compile(r"\b(igtf|itf|impuesto\w*)\b")
# "cancelación" NO está aquí a propósito: en el uso venezolano suele significar
# "pago" (cancelar una factura). Solo la frase de SAP "cancelar entrada".
PATRON_REVERSO_BASE = r"anul\w*|revers\w*|cancelar entrada|nul[oa]"
_RE_REVERSO = re.compile(r"\b(" + PATRON_REVERSO_BASE + r")\b")


def clasificar_movimiento(descripcion: Any, monto: Decimal) -> str:
    """Devuelve: Comisiones | Impuestos | Ingresos | Egresos | Sin movimiento."""
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


# =============================================================================
# 3. MODELOS
# =============================================================================
@dataclass(frozen=True, eq=False)
class Movimiento:
    id: str
    origen: str  # "banco" | "libros"
    fecha: date
    monto: Decimal  # con signo, 2 decimales
    descripcion: str = ""
    referencia: str = ""
    moneda: str = "VES"
    datos: dict = field(default_factory=dict)  # fila original, para trazabilidad
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
    regla: str  # C1 | C2 | C3 | C4
    banco: tuple[Movimiento, ...]
    libros: tuple[Movimiento, ...]
    confianza: str  # alta | media | baja
    requiere_revision: bool
    motivo: str = ""
    ambigua: bool = False  # había varios candidatos equivalentes (mismo monto, misma cercanía)

    @property
    def diferencia_monto(self) -> Decimal:
        """Banco menos libros. Cero en coincidencias exactas."""
        return sum((m.monto for m in self.banco), Decimal("0")) - sum(
            (m.monto for m in self.libros), Decimal("0")
        )

    @property
    def dias_diferencia(self) -> int:
        fechas = [m.fecha for m in self.banco + self.libros]
        return (max(fechas) - min(fechas)).days


@dataclass
class Config:
    tolerancia_dias_c2: int = 3  # cubre viernes -> lunes
    exigir_referencia_c2: bool = False
    ventana_dias_c3: int = 7
    tolerancia_monto_c3: Decimal = Decimal("1.00")
    similitud_minima_c3: float = 0.5
    max_grupo_c4: int = 4
    max_candidatos_c4: int = 20
    ventana_dias_c4: int = 5
    presupuesto_combinaciones_c4: int = 1_000_000  # tope de trabajo; evita corridas eternas
    neutralizar_reversos: bool = True
    ventana_reversos_dias: int = 10
    # Un banco devuelve transferencias fallidas como "DEVOLUCION ..."; en los libros esa
    # palabra puede ser una devolución legítima de un proveedor, así que no se usa allí.
    patron_reverso_banco: str = PATRON_REVERSO_BASE + r"|devoluc\w*"
    patron_reverso_libros: str = PATRON_REVERSO_BASE


@dataclass
class Resultado:
    pares: list[Coincidencia]  # cerradas (C1, C2 y sugerencias aprobadas)
    sugerencias: list[Coincidencia]  # C3 / C4: esperan decisión humana
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
        """Una persona aprueba una sugerencia: pasa a cerrada."""
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
        # Descarta otras sugerencias que usaban alguna de esas partidas.
        self.sugerencias = [
            s for s in self.sugerencias
            if not ({m.id for m in s.banco} & ids_b or {m.id for m in s.libros} & ids_l)
        ]

    def rechazar(self, sugerencia: Coincidencia) -> None:
        if sugerencia not in self.sugerencias:
            raise ValueError("La sugerencia no pertenece a este resultado.")
        self.sugerencias.remove(sugerencia)

    def filas_coincidencias(self) -> list[dict]:
        """Lista plana (pares + sugerencias), lista para pd.DataFrame(...)."""
        filas = []
        for estado, lista in (("cerrada", self.pares), ("sugerida", self.sugerencias)):
            for c in lista:
                filas.append({
                    "estado": estado,
                    "regla": c.regla,
                    "confianza": c.confianza,
                    "requiere_revision": c.requiere_revision,
                    "ambigua": c.ambigua,
                    "banco_ids": ",".join(m.id for m in c.banco),
                    "libros_ids": ",".join(m.id for m in c.libros),
                    "monto_banco": sum((m.monto for m in c.banco), Decimal("0")),
                    "monto_libros": sum((m.monto for m in c.libros), Decimal("0")),
                    "diferencia": c.diferencia_monto,
                    "dias": c.dias_diferencia,
                    "motivo": c.motivo,
                })
        return filas


# =============================================================================
# 4. REVERSOS
# =============================================================================
def _orden(m: Movimiento) -> tuple:
    return (m.fecha, m.id)


def neutralizar_reversos(
    movs: Sequence[Movimiento], ventana_dias: int = 10, patron: Optional[str] = None
) -> tuple[list[Movimiento], list[tuple[Movimiento, Movimiento]], list[Movimiento]]:
    """
    Empareja cada reverso/anulación con su movimiento original (mismo monto,
    signo contrario, dentro de la ventana) y retira AMBOS.
    El código anterior descartaba solo el reverso, dejando huérfano al original.

    Devuelve (restantes, pares_(original, reverso), reversos_sin_original).
    Un reverso sin original NO se descarta: queda entre los restantes.
    """
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
                o.id in es_rev,  # prefiere un original que no sea otro reverso
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


# =============================================================================
# 5. MOTOR DE CRUCE
# =============================================================================
def _validar_ids(movs: Sequence[Movimiento], nombre: str) -> None:
    vistos: set[str] = set()
    for m in movs:
        if m.id in vistos:
            raise ValueError(f"id duplicado en {nombre}: {m.id!r}")
        vistos.add(m.id)


def _indice_por_monto(libros: Sequence[Movimiento]) -> dict[Decimal, tuple[list[int], list[Movimiento]]]:
    """monto -> (fechas ordinales ordenadas, movimientos), para buscar por ventana."""
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
    """
    Mismo monto y signo (monto con signo idéntico), fecha dentro de la
    tolerancia. Entre varios candidatos prefiere: referencia que coincide,
    luego la fecha más cercana. Cada movimiento se usa una sola vez.
    """
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
    """
    C3: mismo signo, fecha dentro de la ventana y (referencia coincide, o
    descripción parecida, o monto exacto con fecha fuera de la tolerancia C2).
    Monto dentro de la tolerancia. Son SUGERENCIAS, nunca cierran solas.
    """
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
    """
    Busca 2..max_grupo_c4 movimientos del pool cuya suma sea exactamente `unico.monto`.
    `presupuesto` es un contador compartido ([restante]) de combinaciones a probar.
    """
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
    """
    C4: 1 banco = N libros, o N banco = 1 libro (suma exacta). Siempre sugerencia.
    Devuelve (sugerencias, presupuesto_agotado).
    """
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
    """
    Cruza movimientos del banco contra los libros.

    Orden de reglas:
        C1  referencia + monto + signo + fecha exacta            (cierra sola)
        C2  monto + signo, fecha dentro de ±tolerancia_dias_c2   (cierra sola)
        C3  aproximada: monto/fecha tolerados + texto            (sugerencia)
        C4  agrupada 1-N / N-1 con suma exacta                   (sugerencia)

    El signo se compara siempre: un egreso nunca concilia con un ingreso.
    """
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
                    + ("..." if len(sin) > 5 else "")
                )

    banco = sorted(banco, key=_orden)
    libros = sorted(libros, key=_orden)

    c1, banco, libros = _cruzar_exacto(banco, libros, 0, "C1", exigir_referencia=True)
    # C2 en dos etapas, para que un cruce lejano no le quite la partida a uno exacto:
    #   (a) los que comparten referencia, dentro de la tolerancia;
    #   (b) el resto, de la fecha más cercana (0 días) a la más lejana.
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
    if agotado:
        advertencias.append(
            "La búsqueda de coincidencias agrupadas (C4) se detuvo al agotar su presupuesto de "
            "cálculo; pueden existir más agrupaciones. Concilie primero las partidas simples "
            "o aumente Config.presupuesto_combinaciones_c4."
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


# =============================================================================
# 6. CUADRE (el entregable real de una conciliación)
# =============================================================================
@dataclass(frozen=True)
class Cuadre:
    saldo_banco: Decimal
    saldo_libros: Decimal
    partidas_libros_no_en_banco: Decimal  # depósitos / cheques en tránsito (con signo)
    partidas_banco_no_en_libros: Decimal  # comisiones, notas no registradas (con signo)
    diferencias_en_pares: Decimal  # banco - libros en pares aprobados con diferencia
    saldo_banco_ajustado: Decimal
    saldo_libros_ajustado: Decimal
    diferencia: Decimal

    @property
    def cuadrado(self) -> bool:
        return self.diferencia == 0


def calcular_cuadre(resultado: Resultado, saldo_banco: Any, saldo_libros: Any) -> Cuadre:
    """
    saldo banco + partidas de libros no reflejadas en el banco
        = saldo libros + partidas del banco no registradas + diferencias en pares

    Solo cuadra si ambos saldos son de la misma fecha de corte y los
    movimientos cargados cubren todo el periodo desde la última conciliación.
    """
    sb, sl = parsear_monto(saldo_banco), parsear_monto(saldo_libros)
    cero = Decimal("0")
    p_libros = sum((m.monto for m in resultado.libros_pendiente), cero)
    p_banco = sum((m.monto for m in resultado.banco_pendiente), cero)
    d_pares = sum((p.diferencia_monto for p in resultado.pares), cero)
    aj_b = sb + p_libros
    aj_l = sl + p_banco + d_pares
    return Cuadre(sb, sl, p_libros, p_banco, d_pares, aj_b, aj_l, aj_b - aj_l)


# =============================================================================
# 7. CARGA DESDE FILAS (puente con pandas / Streamlit)
# =============================================================================
@dataclass
class ErrorFila:
    fila: int  # 1 = primera fila de datos
    columna: str
    valor: Any
    motivo: str


@dataclass
class SaldoDetectado:
    """Fila que es un saldo (inicial/final), no un movimiento."""
    fila: int
    etiqueta: str
    fecha: Optional[date]
    monto: Decimal


@dataclass
class Carga:
    movimientos: list[Movimiento]
    errores: list[ErrorFila]
    omitidas: int  # filas vacías, de encabezado o de monto cero
    saldos: list[SaldoDetectado] = field(default_factory=list)
    omitidas_detalle: list[tuple[int, str]] = field(default_factory=list)
    saldo_final_columna: Optional[Decimal] = None  # saldo de la última fila con movimiento

    def saldo(self, patron: str) -> Optional[Decimal]:
        """Busca un saldo detectado por texto en su etiqueta, p. ej. carga.saldo("final")."""
        for sd in self.saldos:
            if patron in sd.etiqueta:
                return sd.monto
        return None


def _texto(v: Any) -> str:
    if _es_vacio(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))  # Excel entrega la referencia 123456 como 123456.0
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
    """
    Convierte filas (p. ej. df.to_dict("records")) en Movimientos.

    mapeo:
        "fecha"       (obligatorio)
        "descripcion", "referencia"   (opcionales)
        y UNA de estas dos formas de monto:
        "monto"                  columna con monto firmado
        "entrada" + "salida"     columnas separadas; monto = entrada - salida
            banco : entrada = Crédito/Abono      salida = Débito/Cargo
            libros: entrada = Debe/Cargo (a la cuenta de banco)  salida = Haber/Abono
        Ojo: Débito/Crédito significan lo CONTRARIO en el extracto del banco y
        en el mayor de la empresa; por eso aquí se usan entrada/salida.

    Opcional "saldo": columna de saldo acumulado (típica del mayor). Sirve para detectar
    el saldo inicial (fila de encabezado de cuenta, sin movimiento) y el saldo final.

    Filas que NO son movimientos:
        * descripción que coincide con `patron_saldo` ("SALDO INICIAL", "SALDO FINAL")
          -> no se cargan como movimiento; quedan en `Carga.saldos` para el cuadre.
        * sin ningún monto y con fecha ilegible (encabezado de cuenta, subtotal)
          -> se omiten y quedan listadas en `omitidas_detalle`.
    Las filas ilegibles con dinero NO se descartan en silencio: van a `errores`.
    Una fila con fecha válida pero sin monto también es un error visible.
    """
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

        # (a) "SALDO INICIAL" / "SALDO FINAL": no es un movimiento.
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

        # (b) fila sin dinero: encabezado de cuenta, subtotal o relleno.
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
                omitidas_detalle.append((n, f"sin monto y sin fecha válida ({_texto(valor_fecha)!r}): "
                                            "posible encabezado de cuenta o subtotal"))
                if col_saldo and not _es_vacio(fila.get(col_saldo)):
                    try:
                        saldos.append(SaldoDetectado(
                            n, "saldo inicial (fila de encabezado sin movimiento)", None,
                            parsear_monto(fila.get(col_saldo), separador_decimal)))
                    except ErrorDeParseo:
                        pass
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

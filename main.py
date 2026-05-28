import io
import math
import time
import hashlib
import unicodedata
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Set, Tuple

import pandas as pd
from fastapi import Body, FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field

try:
    from ortools.sat.python import cp_model
except Exception:  # Permite correr modo heurístico si OR-Tools no está instalado.
    cp_model = None

app = FastAPI(title="CineOps DSS Ultra Fast Optimizer")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["GET", "POST", "PUT", "DELETE"],
    allow_headers=["*"],
)

# =====================================================================
# ESTADO EN MEMORIA: EL EXCEL ORIGINAL NO SE MODIFICA
# =====================================================================
DB_EMPLEADOS: List[dict] = []
DF_DEMANDA = None
ULTIMO_EXCEL_BYTES: Optional[bytes] = None

PLAN_OPTIMIZADO: Dict[str, Dict[str, Dict[str, str]]] = {}
PLAN_EXCEL_BASE: Dict[str, Dict[str, Dict[str, str]]] = {}
ROTACIONES_PLAN: Dict[str, Dict[str, Dict[str, str]]] = {}

ASISTENCIAS_REGISTRADAS: Dict[str, str] = {}
DESCANSOS_MANUALES: Dict[str, Set[str]] = {}

# Cache rápido. Si no cambió Excel/faltas/descansos/presupuesto, no recalcula.
PLAN_CACHE_KEY: Optional[Tuple] = None
PLAN_HORAS_TOTAL = 0.0
PLAN_HORAS_EMPLEADO: Dict[str, float] = {}
PLAN_HORAS_DIA: Dict[str, float] = {}
PLAN_HORAS_DIA_EMPLEADO: Dict[str, Dict[str, float]] = {}
LAST_OPTIMIZATION_META: Dict[str, object] = {}

DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
HORAS_BLOQUE = [f"{h:02d}:00" for h in range(8, 24)]
AREAS_CINE = ["Taquilla", "Dulcería", "Lobby", "Baños", "Entrada"]
AREA_TO_M = {"Taquilla": "M1", "Dulcería": "M2", "Lobby": "M3", "Baños": "M4", "Entrada": "M5"}

AREA_COLS = {
    "M1 Taquilla": "Taquilla",
    "M2 Dulcería": "Dulcería",
    "M3 Lobby": "Lobby",
    "M4 Baños": "Baños",
    "M5 Entrada": "Entrada",
}

DISTRIBUCION_DIARIA_PROPORCION = {
    "Miércoles": 0.0829,
    "Jueves": 0.1060,
    "Viernes": 0.1391,
    "Sábado": 0.2304,
    "Domingo": 0.2350,
    "Lunes": 0.1155,
    "Martes": 0.0909,
}

DISTRIBUCION_HORARIA = {
    "08:00": 0.0000,
    "09:00": 0.0009,
    "10:00": 0.0077,
    "11:00": 0.0141,
    "12:00": 0.0362,
    "13:00": 0.0399,
    "14:00": 0.0555,
    "15:00": 0.0925,
    "16:00": 0.0840,
    "17:00": 0.1180,
    "18:00": 0.1508,
    "19:00": 0.1362,
    "20:00": 0.1210,
    "21:00": 0.1044,
    "22:00": 0.0390,
    "23:00": 0.0000,
}

SUELDO_HORA_BASE = 46.50
AXH_META = 7.0
PRESUPUESTO_HORAS_GLOBAL = 3032.0
TOTAL_ASISTENTES_SEMANA = 28152
MAX_HORAS_DEFAULT = 48
MAX_DIAS_DESCANSO_DSS = 2
MIN_PERSONAL_TOTAL_POR_HORA = 1
MIN_AREAS_CUBIERTAS_SI_HAY_PERSONAL = True
PORC_TRANSACCIONES_DULCERIA = 0.8052

# Distribución operativa basada en que Dulcería concentró 80.52% de transacciones 2025.
# No se usa 80.52% literal porque también deben cubrirse taquilla, lobby, baños y entrada.
AREA_WEIGHTS = {
    "Dulcería": 0.46,
    "Taquilla": 0.22,
    "Lobby": 0.14,
    "Entrada": 0.10,
    "Baños": 0.08,
}


# =====================================================================
# MODELOS
# =====================================================================
class OptimizationRequest(BaseModel):
    tipo_semana: str = "Normal"
    hora_seleccionada: str = "16:00"
    descansos_automaticos: bool = True
    descanso_dias_objetivo: int = Field(default=2, ge=0, le=2)


class AttendanceRequest(BaseModel):
    empleado_id: str
    dia: str
    estado: str


class RestDaysRequest(BaseModel):
    empleado_id: str
    dias: List[str] = Field(default_factory=list)
    modo: str = "manual"


class ManualAssignmentRequest(BaseModel):
    empleado_id: str
    dia: str
    hora: str
    area: str


class KPIRequest(BaseModel):
    dia: str = "Sábado"
    hora: str = "16:00"
    asistencias: Dict[str, str] = Field(default_factory=dict)


# =====================================================================
# NORMALIZACIÓN
# =====================================================================
def reparar_mojibake(txt: object) -> str:
    if txt is None:
        return ""
    s = str(txt).strip()
    rep = {
        "Ã¡": "á", "Ã©": "é", "Ã­": "í", "Ã³": "ó", "Ãº": "ú", "Ã±": "ñ",
        "Ã": "Á", "Ã‰": "É", "Ã": "Í", "Ã“": "Ó", "Ãš": "Ú", "Ã‘": "Ñ",
        "DÃ­a": "Día", "invÃ¡lido": "inválido", "MiÃ©rcoles": "Miércoles",
        "SÃ¡bado": "Sábado", "DulcerÃ­a": "Dulcería", "BaÃ±os": "Baños",
        "OptimizaciÃ³n": "Optimización", "Ã¡rea": "área", "mÃ¡ximo": "máximo",
    }
    for a, b in rep.items():
        s = s.replace(a, b)
    return s


def sin_acentos(txt: object) -> str:
    s = reparar_mojibake(txt)
    s = unicodedata.normalize("NFKD", s)
    return "".join(c for c in s if not unicodedata.combining(c)).lower().strip()


def limpiar_columna(col: object) -> str:
    return " ".join(reparar_mojibake(col).replace("\n", " ").replace("\r", " ").split())


def normalizar_dia(dia: object) -> Optional[str]:
    clave = sin_acentos(dia)
    return {
        "lunes": "Lunes", "martes": "Martes", "miercoles": "Miércoles",
        "jueves": "Jueves", "viernes": "Viernes", "sabado": "Sábado", "domingo": "Domingo",
    }.get(clave)


def normalizar_area(area: object) -> str:
    clave = sin_acentos(area)
    return {
        "taquilla": "Taquilla", "taquillas": "Taquilla",
        "dulceria": "Dulcería", "lobby": "Lobby", "banos": "Baños", "bano": "Baños",
        "entrada": "Entrada", "descanso": "Descanso", "descanso general": "Descanso General",
        "ninguna": "Ninguna",
    }.get(clave, reparar_mojibake(area))


def normalizar_estado_asistencia(estado: object) -> str:
    clave = sin_acentos(estado).upper()
    return "FALTA" if clave in {"FALTA", "AUSENTE", "A"} else "PRESENTE"


def minutes(hora: str) -> int:
    h, m = str(hora).split(":")[:2]
    return int(h) * 60 + int(m)


def bloque_a_hora(hora: object) -> str:
    try:
        return f"{int(str(hora).split(':')[0]):02d}:00"
    except Exception:
        return "16:00"


def calcular_horas_reales(rango_horario: str) -> float:
    try:
        if not rango_horario or "-" not in str(rango_horario):
            return 0.0
        entrada, salida = str(rango_horario).split("-")[:2]
        ini = minutes(entrada.strip())
        fin = minutes(salida.strip())
        if fin < ini:
            fin += 24 * 60
        return max(0.0, (fin - ini) / 60)
    except Exception:
        return 0.0


def esta_en_horario(hora_check: str, rango_es: str) -> bool:
    try:
        if not rango_es or "-" not in str(rango_es):
            return False
        entrada, salida = str(rango_es).split("-")[:2]
        h_check = minutes(bloque_a_hora(hora_check))
        ini = minutes(entrada.strip())
        fin = minutes(salida.strip())
        if fin < ini:
            fin += 24 * 60
            if h_check < ini:
                h_check += 24 * 60
        return ini <= h_check < fin
    except Exception:
        return False


def clave_estado(dia: str, empleado_id: str) -> str:
    return f"{normalizar_dia(dia) or reparar_mojibake(dia)}:{str(empleado_id)}"


def esta_en_falta(empleado_id: str, dia: str, asistencias_extra: Optional[Dict[str, str]] = None) -> bool:
    emp_id = str(empleado_id)
    dia_ok = normalizar_dia(dia) or dia
    estado_global = ASISTENCIAS_REGISTRADAS.get(clave_estado(dia_ok, emp_id), "PRESENTE")
    estado_extra = (asistencias_extra or {}).get(emp_id, estado_global)
    return normalizar_estado_asistencia(estado_extra) == "FALTA"


def esta_en_descanso_manual(empleado_id: str, dia: str) -> bool:
    dia_ok = normalizar_dia(dia) or dia
    return dia_ok in DESCANSOS_MANUALES.get(str(empleado_id), set())


# =====================================================================
# DEMANDA Y REQUERIMIENTOS
# =====================================================================
def calcular_demanda(dia: str, hora: Optional[str] = None) -> int:
    dia_ok = normalizar_dia(dia) or "Sábado"
    total = int(TOTAL_ASISTENTES_SEMANA)
    asistentes_dia = total * DISTRIBUCION_DIARIA_PROPORCION.get(dia_ok, 0.11)
    if hora is None:
        return int(round(asistentes_dia))
    hora_b = bloque_a_hora(hora)
    total_prop = sum(DISTRIBUCION_HORARIA.values()) or 1
    return int(math.ceil(asistentes_dia * (DISTRIBUCION_HORARIA.get(hora_b, 0.05) / total_prop)))


def requerimiento_total_hora(dia: str, hora: str, disponibles: int) -> int:
    """Mínimo duro: nunca menos de 1 si hay gente; ideal se maneja visualmente, no bloquea el modelo."""
    if disponibles <= 0:
        return 0
    demanda = calcular_demanda(dia, hora)
    ideal = max(MIN_PERSONAL_TOTAL_POR_HORA, math.ceil(demanda / AXH_META)) if demanda > 0 else MIN_PERSONAL_TOTAL_POR_HORA
    # Se mantiene factible y rápido: duro = mínimo de operación + áreas; ideal se reporta como déficit.
    minimo_area = min(5, disponibles) if MIN_AREAS_CUBIERTAS_SI_HAY_PERSONAL and disponibles >= 5 else 1
    return min(disponibles, max(MIN_PERSONAL_TOTAL_POR_HORA, minimo_area, min(ideal, 12)))


def requerimiento_por_area(demanda_hora: int, activos_posibles: int = 0) -> Dict[str, int]:
    if demanda_hora <= 0:
        base = {a: 0 for a in AREAS_CINE}
        if activos_posibles >= 5:
            for a in AREAS_CINE:
                base[a] = 1
        return base
    total = max(1, math.ceil(demanda_hora / AXH_META))
    if activos_posibles > 0:
        total = min(total, activos_posibles)
    req = {a: max(0, int(round(total * AREA_WEIGHTS[a]))) for a in AREAS_CINE}
    if activos_posibles >= 5:
        for a in AREAS_CINE:
            req[a] = max(1, req[a])
    # Ajuste para que la suma no exceda total.
    while sum(req.values()) > max(1, total):
        reducibles = [a for a in AREAS_CINE if req[a] > (1 if activos_posibles >= 5 else 0)]
        if not reducibles:
            break
        a = max(reducibles, key=lambda x: req[x])
        req[a] -= 1
    while sum(req.values()) < max(1, total):
        a = max(AREAS_CINE, key=lambda x: AREA_WEIGHTS[x])
        req[a] += 1
    return req


# =====================================================================
# LECTURA EXCEL Y PREPROCESO
# =====================================================================
def detectar_columna(df: pd.DataFrame, opciones: List[str]) -> Optional[str]:
    normalizadas = {c: sin_acentos(limpiar_columna(c)) for c in list(df.columns)}
    for opcion in opciones:
        op = sin_acentos(opcion)
        for col, col_limpia in normalizadas.items():
            if op in col_limpia:
                return col
    return None


def leer_presupuesto_y_demanda(xls: pd.ExcelFile) -> None:
    global PRESUPUESTO_HORAS_GLOBAL, TOTAL_ASISTENTES_SEMANA, SUELDO_HORA_BASE, DF_DEMANDA
    hoja = next((s for s in xls.sheet_names if sin_acentos(s) == "demanda"), None)
    if not hoja:
        return
    df_dem = pd.read_excel(xls, sheet_name=hoja, header=None)
    DF_DEMANDA = df_dem
    for _, row in df_dem.iterrows():
        etiqueta = sin_acentos(row.iloc[0]) if len(row) > 0 else ""
        valor = row.iloc[1] if len(row) > 1 else None
        try:
            if "semana actual" in etiqueta or "admit" in etiqueta or "asistente" in etiqueta or "demanda" in etiqueta:
                TOTAL_ASISTENTES_SEMANA = int(float(valor))
            elif "presupuesto" in etiqueta or "horas limite" in etiqueta or "ppto" in etiqueta:
                PRESUPUESTO_HORAS_GLOBAL = float(valor)
            elif "pago" in etiqueta or "sueldo" in etiqueta or "costo hora" in etiqueta:
                SUELDO_HORA_BASE = float(valor)
        except Exception:
            pass


def construir_plan_excel_base() -> None:
    global PLAN_EXCEL_BASE
    PLAN_EXCEL_BASE = {d: {h: {} for h in HORAS_BLOQUE} for d in DIAS_SEMANA}
    for e in DB_EMPLEADOS:
        emp_id = str(e["id"])
        area_base = normalizar_area(e.get("area_base", "Lobby"))
        for d in DIAS_SEMANA:
            rango = e["horarios_por_dia"].get(d, "")
            for h in HORAS_BLOQUE:
                PLAN_EXCEL_BASE[d][h][emp_id] = area_base if esta_en_horario(h, rango) else "Descanso General"


def calcular_horas_base_empleado(emp: dict) -> float:
    return sum(calcular_horas_reales(emp["horarios_por_dia"].get(d, "")) for d in DIAS_SEMANA)


def construir_hash_excel() -> str:
    if not ULTIMO_EXCEL_BYTES:
        return "no_excel"
    return hashlib.md5(ULTIMO_EXCEL_BYTES).hexdigest()[:12]


def cache_key_optimizacion() -> Tuple:
    return (
        construir_hash_excel(),
        round(float(PRESUPUESTO_HORAS_GLOBAL), 2),
        int(TOTAL_ASISTENTES_SEMANA),
        tuple(sorted(ASISTENCIAS_REGISTRADAS.items())),
        tuple(sorted((k, tuple(sorted(v))) for k, v in DESCANSOS_MANUALES.items())),
        MAX_DIAS_DESCANSO_DSS,
    )


# =====================================================================
# PLANIFICADOR HÍBRIDO ULTRA RÁPIDO
# =====================================================================
def turno_referencia_empleado(e: dict) -> str:
    """Devuelve el turno más representativo del empleado.
    Esto permite mover descansos entre días sin dejar todo un día vacío.
    El Excel sigue siendo la base, pero los días originalmente vacíos pueden recibir
    el mismo patrón de turno del colaborador si el DSS decide trabajarlo.
    """
    turnos = []
    for d in DIAS_SEMANA:
        r = reparar_mojibake(e.get("horarios_por_dia", {}).get(d, ""))
        if r and "-" in r and calcular_horas_reales(r) > 0:
            turnos.append(r)
    if not turnos:
        return "15:00-22:00"
    # Turno más frecuente; desempata por más horas para evitar turnos inválidos/cortos.
    return max(set(turnos), key=lambda r: (turnos.count(r), calcular_horas_reales(r)))


def rango_operativo_empleado(e: dict, d: str) -> str:
    """Rango que usa el DSS para planear.
    Si el Excel trae turno ese día, se respeta.
    Si el Excel trae descanso/vacío, se usa el turno de referencia para poder mover
    descansos y evitar que todos descansen el mismo día.
    """
    r = reparar_mojibake(e.get("horarios_por_dia", {}).get(d, ""))
    if r and "-" in r and calcular_horas_reales(r) > 0:
        return r
    return e.get("turno_referencia") or turno_referencia_empleado(e)


def slot_day_hours(e: dict, d: str) -> int:
    return int(round(calcular_horas_reales(rango_operativo_empleado(e, d))))


def emp_scheduled_hours(e: dict, d: str) -> List[str]:
    rango = rango_operativo_empleado(e, d)
    return [h for h in HORAS_BLOQUE if esta_en_horario(h, rango)]


def candidate_day_slots() -> List[Tuple[str, str, int, float]]:
    slots = []
    for e in DB_EMPLEADOS:
        emp_id = str(e["id"])
        for d in DIAS_SEMANA:
            hrs = slot_day_hours(e, d)
            if hrs <= 0 or esta_en_falta(emp_id, d) or esta_en_descanso_manual(emp_id, d):
                continue
            valor_demanda = sum(calcular_demanda(d, h) for h in emp_scheduled_hours(e, d))
            # Valor extra para M2 y días fuertes, por Dulcería.
            exp = {normalizar_area(a) for a in e.get("areas_expertis", [])}
            bonus = 1.0 + (0.20 if "Dulcería" in exp else 0) + (0.05 * max(0, int(e.get("nivel_esp", 1)) - 1))
            slots.append((emp_id, d, hrs, valor_demanda * bonus))
    return slots


def solve_work_days_fast() -> Dict[Tuple[str, str], int]:
    """Planificador balanceado y rápido de días trabajados.

    Reglas operativas aplicadas:
    - Cada colaborador descansa máximo 2 días DSS por semana.
    - Los descansos se distribuyen en patrones rotativos, no todos el mismo día.
    - Si el Excel tenía descanso en un día, el DSS puede mover ese descanso y usar
      el turno de referencia del colaborador.
    - Se respeta falta y descanso manual.
    - Si hay personal disponible, cada hora queda con al menos 1 activo.
    - El presupuesto se intenta respetar como límite duro; si el mínimo operativo
      matemáticamente supera el presupuesto, se devuelve el plan mínimo y se alerta
      en KPIs.
    """
    y: Dict[Tuple[str, str], int] = {}
    for e in DB_EMPLEADOS:
        emp_id = str(e["id"])
        for d in DIAS_SEMANA:
            y[(emp_id, d)] = 0

    if not DB_EMPLEADOS:
        return y

    # Patrones rotativos de 2 descansos. Están separados para no concentrar todos
    # los descansos en miércoles/sábado.
    patrones = [
        ("Lunes", "Jueves"),
        ("Martes", "Viernes"),
        ("Miércoles", "Sábado"),
        ("Jueves", "Domingo"),
        ("Viernes", "Lunes"),
        ("Sábado", "Martes"),
        ("Domingo", "Miércoles"),
    ]

    # Orden estable por área/nivel/id para repartir descansos de forma uniforme.
    empleados_orden = sorted(
        DB_EMPLEADOS,
        key=lambda e: (
            normalizar_area(e.get("area_base", "Lobby")),
            -int(e.get("nivel_esp", 1) or 1),
            int(e.get("id", 0)),
        ),
    )

    # 1) Asignación base: trabaja 5 días, descansa 2 días rotativos.
    for idx, e in enumerate(empleados_orden):
        emp_id = str(e["id"])
        descanso_patron = set(patrones[idx % len(patrones)])
        descansos_aplicados = 0
        for d in DIAS_SEMANA:
            if esta_en_falta(emp_id, d) or esta_en_descanso_manual(emp_id, d):
                y[(emp_id, d)] = 0
                continue
            if d in descanso_patron and descansos_aplicados < MAX_DIAS_DESCANSO_DSS:
                y[(emp_id, d)] = 0
                descansos_aplicados += 1
            else:
                y[(emp_id, d)] = 1

    def total_hours_current() -> float:
        return sum(slot_day_hours(e, d) for e in DB_EMPLEADOS for d in DIAS_SEMANA if y.get((str(e["id"]), d), 0) == 1)

    def hour_count(day: str, hour: str) -> int:
        c = 0
        for ee in DB_EMPLEADOS:
            eid = str(ee["id"])
            if y.get((eid, day), 0) == 1 and esta_en_horario(hour, rango_operativo_empleado(ee, day)):
                c += 1
        return c

    def can_rest(emp: dict, day: str) -> bool:
        emp_id = str(emp["id"])
        if y.get((emp_id, day), 0) == 0:
            return False
        # No permitir más de 2 días DSS de descanso por persona; faltas/manuales ya son restricciones externas.
        dss_rest = sum(1 for dd in DIAS_SEMANA if y.get((emp_id, dd), 0) == 0 and not esta_en_falta(emp_id, dd) and not esta_en_descanso_manual(emp_id, dd))
        if dss_rest >= MAX_DIAS_DESCANSO_DSS:
            return False
        # Nunca dejar una hora sin al menos 1 activo si había cobertura.
        for h in emp_scheduled_hours(emp, day):
            if hour_count(day, h) <= MIN_PERSONAL_TOTAL_POR_HORA:
                return False
        return True

    # 2) Reparación de cobertura: si algún día/hora queda vacío, activa al mejor candidato.
    for d in DIAS_SEMANA:
        for h in HORAS_BLOQUE:
            if hour_count(d, h) >= MIN_PERSONAL_TOTAL_POR_HORA:
                continue
            candidatos = []
            for e in DB_EMPLEADOS:
                emp_id = str(e["id"])
                if esta_en_falta(emp_id, d) or esta_en_descanso_manual(emp_id, d):
                    continue
                if not esta_en_horario(h, rango_operativo_empleado(e, d)):
                    continue
                # Preferir quien tenga menos horas actuales y buena polivalencia.
                horas_emp = sum(slot_day_hours(e, dd) for dd in DIAS_SEMANA if y.get((emp_id, dd), 0) == 1)
                poliv = len(e.get("areas_expertis", []))
                score = -horas_emp + poliv * 2 + int(e.get("nivel_esp", 1) or 1)
                candidatos.append((score, e))
            if candidatos:
                _, elegido = max(candidatos, key=lambda x: x[0])
                y[(str(elegido["id"]), d)] = 1

    # 3) Si el presupuesto aún se excede, intenta descansar días adicionales SOLO si no rompe
    # cobertura mínima y sin superar 2 descansos DSS por persona.
    total_h = total_hours_current()
    if total_h > PRESUPUESTO_HORAS_GLOBAL:
        candidates = []
        for e in DB_EMPLEADOS:
            emp_id = str(e["id"])
            exp = {normalizar_area(a) for a in e.get("areas_expertis", [])}
            for d in DIAS_SEMANA:
                if y.get((emp_id, d), 0) != 1:
                    continue
                hrs = slot_day_hours(e, d)
                demanda_val = sum(calcular_demanda(d, h) for h in emp_scheduled_hours(e, d))
                # Bajo impacto = mejor para descansar, pero penaliza descansar certificados M2 en días fuertes.
                proteccion = 300 if "Dulcería" in exp else 0
                score = demanda_val + proteccion
                candidates.append((score, -hrs, e, d, hrs))
        candidates.sort(key=lambda x: (x[0], x[1]))
        for _, _, e, d, hrs in candidates:
            if total_h <= PRESUPUESTO_HORAS_GLOBAL:
                break
            if can_rest(e, d):
                y[(str(e["id"]), d)] = 0
                total_h -= hrs

    return y


def compute_area_targets(dia: str, hora: str, active_count: int) -> Dict[str, int]:
    if active_count <= 0:
        return {a: 0 for a in AREAS_CINE}
    # Base de demanda. Cap al número activo.
    demanda = calcular_demanda(dia, hora)
    target_total = active_count
    req = requerimiento_por_area(demanda, active_count)

    # Si hay suficiente personal, ninguna área queda sola.
    if active_count >= 5:
        for a in AREAS_CINE:
            req[a] = max(1, req.get(a, 0))

    # Rellena el resto con prioridad Dulcería 2025.
    while sum(req.values()) < target_total:
        ratios = {a: (req[a] / max(0.0001, AREA_WEIGHTS[a])) for a in AREAS_CINE}
        a = min(ratios, key=ratios.get)
        req[a] += 1
    while sum(req.values()) > target_total:
        reducibles = [a for a in AREAS_CINE if req[a] > (1 if active_count >= 5 else 0)]
        if not reducibles:
            break
        a = max(reducibles, key=lambda x: req[x] - AREA_WEIGHTS[x] * target_total)
        req[a] -= 1
    return req


def employee_area_score(e: dict, area: str, current_counts: Dict[str, int]) -> float:
    exp = {normalizar_area(a) for a in e.get("areas_expertis", [])}
    nivel = int(e.get("nivel_esp", 1) or 1)
    score = 0.0
    if area in exp:
        score += 120
    else:
        score -= 35
    if area == "Dulcería":
        score += 60 if "Dulcería" in exp else 8
    if area == "Taquilla":
        score += 35 if "Taquilla" in exp else 5
    score += 8 * len(exp)
    score += 5 * nivel
    # Pequeño desempate determinista por nombre/id.
    score += (int(e["id"]) % 17) / 10
    return score


def assign_areas_for_hour(d: str, h: str, empleados_activos: List[dict]) -> Dict[str, str]:
    if not empleados_activos:
        return {}
    targets = compute_area_targets(d, h, len(empleados_activos))
    remaining = empleados_activos[:]
    assigned: Dict[str, str] = {}
    counts = {a: 0 for a in AREAS_CINE}

    # Primero cubre mínimos por área para evitar áreas vacías.
    for area in AREAS_CINE:
        while counts[area] < targets[area] and remaining:
            best = max(remaining, key=lambda e: employee_area_score(e, area, counts))
            assigned[str(best["id"])] = area
            counts[area] += 1
            remaining.remove(best)

    # Si quedó alguien, refuerza Dulcería y áreas con mayor brecha.
    while remaining:
        def need_ratio(a: str) -> float:
            return (targets[a] - counts[a]) + AREA_WEIGHTS[a]
        area = max(AREAS_CINE, key=need_ratio)
        best = max(remaining, key=lambda e: employee_area_score(e, area, counts))
        assigned[str(best["id"])] = area
        counts[area] += 1
        remaining.remove(best)
    return assigned


def asignar_rotaciones(d: str, h: str, assigned: Dict[str, str]) -> Dict[str, str]:
    """Ley Silla/Alimentos sin dejar áreas en cero. No cuenta como activo en estación."""
    rot: Dict[str, str] = {}
    if len(assigned) < 8:
        return rot
    counts = {a: 0 for a in AREAS_CINE}
    for area in assigned.values():
        if area in counts:
            counts[area] += 1

    # Solo sacar de áreas con excedente > 1 para no dejar solas.
    candidates = [emp for emp, area in assigned.items() if counts.get(area, 0) > 1]
    if len(candidates) < 2:
        return rot
    seed = (DIAS_SEMANA.index(d) * 17 + int(h.split(":")[0]))
    candidates = sorted(candidates, key=lambda emp: (int(emp) + seed) % 97)
    for label in ["Ley silla", "Alimentos"]:
        for emp in candidates:
            if emp in rot:
                continue
            area = assigned.get(emp)
            if area in counts and counts[area] > 1:
                rot[emp] = label
                counts[area] -= 1
                break
    return rot


def construir_plan_hibrido() -> Tuple[str, float]:
    global PLAN_OPTIMIZADO, ROTACIONES_PLAN
    inicio = time.perf_counter()
    y = solve_work_days_fast()

    PLAN_OPTIMIZADO = {d: {h: {} for h in HORAS_BLOQUE} for d in DIAS_SEMANA}
    ROTACIONES_PLAN = {d: {h: {} for h in HORAS_BLOQUE} for d in DIAS_SEMANA}

    for d in DIAS_SEMANA:
        for h in HORAS_BLOQUE:
            activos = []
            for e in DB_EMPLEADOS:
                emp_id = str(e["id"])
                if y.get((emp_id, d), 0) == 1 and esta_en_horario(h, rango_operativo_empleado(e, d)):
                    activos.append(e)
            assigned = assign_areas_for_hour(d, h, activos)
            rotations = asignar_rotaciones(d, h, assigned)
            for e in DB_EMPLEADOS:
                emp_id = str(e["id"])
                PLAN_OPTIMIZADO[d][h][emp_id] = assigned.get(emp_id, "Descanso General")
                if emp_id in rotations:
                    ROTACIONES_PLAN[d][h][emp_id] = rotations[emp_id]
    actualizar_cache_horas()
    elapsed = round(time.perf_counter() - inicio, 3)
    return "HYBRID_FAST", elapsed


def actualizar_cache_horas() -> None:
    global PLAN_HORAS_TOTAL, PLAN_HORAS_EMPLEADO, PLAN_HORAS_DIA, PLAN_HORAS_DIA_EMPLEADO
    total = 0.0
    por_emp: Dict[str, float] = {str(e["id"]): 0.0 for e in DB_EMPLEADOS}
    por_dia: Dict[str, float] = {d: 0.0 for d in DIAS_SEMANA}
    por_dia_emp: Dict[str, Dict[str, float]] = {d: {str(e["id"]): 0.0 for e in DB_EMPLEADOS} for d in DIAS_SEMANA}
    for d in DIAS_SEMANA:
        for h in HORAS_BLOQUE:
            for emp_id, area in PLAN_OPTIMIZADO.get(d, {}).get(h, {}).items():
                if normalizar_area(area) in AREAS_CINE:
                    total += 1
                    por_emp[emp_id] = por_emp.get(emp_id, 0.0) + 1
                    por_dia[d] += 1
                    por_dia_emp[d][emp_id] = por_dia_emp[d].get(emp_id, 0.0) + 1
    PLAN_HORAS_TOTAL = total
    PLAN_HORAS_EMPLEADO = por_emp
    PLAN_HORAS_DIA = por_dia
    PLAN_HORAS_DIA_EMPLEADO = por_dia_emp


def calcular_horas_plan(plan: Optional[Dict] = None, empleado_id: Optional[str] = None, dia: Optional[str] = None) -> float:
    if plan is None or plan is PLAN_OPTIMIZADO:
        if empleado_id and dia:
            return PLAN_HORAS_DIA_EMPLEADO.get(dia, {}).get(str(empleado_id), 0.0)
        if empleado_id:
            return PLAN_HORAS_EMPLEADO.get(str(empleado_id), 0.0)
        if dia:
            return PLAN_HORAS_DIA.get(dia, 0.0)
        return PLAN_HORAS_TOTAL

    total = 0.0
    dias = [dia] if dia else list(plan.keys())
    for d in dias:
        for h in plan.get(d, {}):
            for emp_id, area in plan[d][h].items():
                if empleado_id is not None and str(emp_id) != str(empleado_id):
                    continue
                if normalizar_area(area) in AREAS_CINE:
                    total += 1
    return total


@app.post("/optimize-weekly")
def optimize_weekly(req: OptimizationRequest):
    global PLAN_CACHE_KEY, LAST_OPTIMIZATION_META
    if not DB_EMPLEADOS:
        raise HTTPException(status_code=400, detail="Primero sube el Excel en /upload.")

    key = cache_key_optimizacion()
    if PLAN_OPTIMIZADO and PLAN_CACHE_KEY == key:
        return {
            "status": "Optimización cacheada",
            "solver_status": "CACHE_ULTRA_FAST",
            "objetivo": 0,
            "presupuesto_horas": PRESUPUESTO_HORAS_GLOBAL,
            "horas_asignadas_sem": PLAN_HORAS_TOTAL,
            "horas_disponibles_sem": round(PRESUPUESTO_HORAS_GLOBAL - PLAN_HORAS_TOTAL, 2),
            "empleados": len(DB_EMPLEADOS),
            "tiempo_seg": 0.0,
        }

    status, elapsed = construir_plan_hibrido()
    PLAN_CACHE_KEY = key
    LAST_OPTIMIZATION_META = {"status": status, "tiempo_seg": elapsed}
    return {
        "status": "Optimización completada",
        "solver_status": status,
        "objetivo": 0,
        "presupuesto_horas": PRESUPUESTO_HORAS_GLOBAL,
        "horas_asignadas_sem": PLAN_HORAS_TOTAL,
        "horas_disponibles_sem": round(PRESUPUESTO_HORAS_GLOBAL - PLAN_HORAS_TOTAL, 2),
        "empleados": len(DB_EMPLEADOS),
        "tiempo_seg": elapsed,
    }


# =====================================================================
# CARGA DEL EXCEL
# =====================================================================
@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    global DB_EMPLEADOS, ULTIMO_EXCEL_BYTES, PLAN_CACHE_KEY, ASISTENCIAS_REGISTRADAS, DESCANSOS_MANUALES
    try:
        content = await file.read()
        ULTIMO_EXCEL_BYTES = content
        ASISTENCIAS_REGISTRADAS = {}
        DESCANSOS_MANUALES = {}
        PLAN_CACHE_KEY = None

        xls = pd.ExcelFile(io.BytesIO(content))
        leer_presupuesto_y_demanda(xls)
        hoja = next((s for s in xls.sheet_names if "matriz" in sin_acentos(s) or "multi" in sin_acentos(s) or "expert" in sin_acentos(s)), xls.sheet_names[0])
        df_raw = pd.read_excel(xls, sheet_name=hoja, header=None)
        header_row = 0
        for i, row in df_raw.iterrows():
            fila = sin_acentos(" ".join(str(v) for v in row.values if pd.notna(v)))
            if "id" in fila and ("nombre" in fila or "colaborador" in fila):
                header_row = i
                break
        df = pd.read_excel(xls, sheet_name=hoja, header=header_row)
        df.columns = [limpiar_columna(c) for c in df.columns]
        col_id = detectar_columna(df, ["id", "ps", "expediente"])
        col_nombre = detectar_columna(df, ["nombre", "colaborador", "cinepolito"])
        col_nivel = detectar_columna(df, ["nivel especialidad", "nivel", "expertis"])
        col_limite = detectar_columna(df, ["limite", "límite", "maximo", "máximo", "horas max"])
        if not col_id or not col_nombre:
            raise HTTPException(status_code=400, detail="No se detectaron columnas ID y Nombre en el Excel.")

        empleados: List[dict] = []
        for _, row in df.iterrows():
            id_val = row.get(col_id)
            nombre_val = row.get(col_nombre)
            if pd.isna(id_val) or pd.isna(nombre_val):
                continue
            if "control" in sin_acentos(id_val) or "total" in sin_acentos(nombre_val):
                continue
            try:
                emp_id = int(float(id_val))
            except Exception:
                continue

            areas_expertis: List[str] = []
            expertise_binaria: Dict[str, int] = {}
            for col_excel, area in AREA_COLS.items():
                col_real = detectar_columna(df, [col_excel])
                val = row.get(col_real, 0) if col_real else 0
                activo = str(val).strip().upper() in {"1", "1.0", "SI", "SÍ", "TRUE"} or val == 1
                m_key = col_excel.split()[0]
                expertise_binaria[m_key] = 1 if activo else 0
                if activo:
                    areas_expertis.append(area)
            area_base = areas_expertis[0] if areas_expertis else "Lobby"

            horarios = {}
            for d in DIAS_SEMANA:
                col_dia = detectar_columna(df, [f"{d} (E/S)", d])
                valor = row.get(col_dia, "") if col_dia else ""
                valor_str = reparar_mojibake(str(valor).strip()) if pd.notna(valor) else ""
                if sin_acentos(valor_str).upper() in {"NAN", "0", "DESCANSO", "X", ""}:
                    valor_str = ""
                horarios[d] = valor_str
            try:
                nivel = int(float(row.get(col_nivel, 1))) if col_nivel and pd.notna(row.get(col_nivel)) else 1
            except Exception:
                nivel = 1
            base_horas = sum(slot_day_hours({"horarios_por_dia": horarios}, d) for d in DIAS_SEMANA)
            try:
                limite = float(row.get(col_limite)) if col_limite and pd.notna(row.get(col_limite)) else max(MAX_HORAS_DEFAULT, base_horas)
            except Exception:
                limite = max(MAX_HORAS_DEFAULT, base_horas)
            empleados.append({
                "id": emp_id,
                "id_ps": str(emp_id),
                "nombre": reparar_mojibake(str(nombre_val)).upper().strip(),
                "nivel_esp": nivel,
                "areas_expertis": areas_expertis,
                "expertise": expertise_binaria,
                "area_base": area_base,
                "horarios_por_dia": horarios,
                "limite_maximo": limite,
            })
            empleados[-1]["turno_referencia"] = turno_referencia_empleado(empleados[-1])
        if not empleados:
            raise HTTPException(status_code=400, detail="No se encontraron empleados válidos en el Excel.")
        DB_EMPLEADOS = empleados
        construir_plan_excel_base()
        resp = optimize_weekly(OptimizationRequest())
        return {
            "status": "success",
            "total_empleados": len(DB_EMPLEADOS),
            "presupuesto_detectado": PRESUPUESTO_HORAS_GLOBAL,
            "demanda_semana": TOTAL_ASISTENTES_SEMANA,
            "solver": resp,
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al procesar Excel: {str(e)}")


# =====================================================================
# ASISTENCIA Y DESCANSOS
# =====================================================================
@app.post("/attendance")
def registrar_asistencia(req: AttendanceRequest):
    global PLAN_CACHE_KEY
    if not DB_EMPLEADOS:
        raise HTTPException(status_code=400, detail="Primero sube el Excel.")
    dia = normalizar_dia(req.dia)
    if not dia:
        raise HTTPException(status_code=400, detail=f"Día inválido: {reparar_mojibake(req.dia)}")
    emp_id = str(req.empleado_id)
    estado = normalizar_estado_asistencia(req.estado)
    ASISTENCIAS_REGISTRADAS[clave_estado(dia, emp_id)] = estado
    PLAN_CACHE_KEY = None
    optimize_weekly(OptimizationRequest())
    return {"status": "ok", "empleado_id": emp_id, "dia": dia, "estado": estado}


@app.post("/rest-days")
def registrar_descansos(req: RestDaysRequest):
    global PLAN_CACHE_KEY
    if not DB_EMPLEADOS:
        raise HTTPException(status_code=400, detail="Primero sube el Excel.")
    emp_id = str(req.empleado_id)
    dias = [normalizar_dia(d) for d in req.dias]
    dias = [d for d in dias if d]
    if req.modo == "limpiar":
        DESCANSOS_MANUALES.pop(emp_id, None)
    else:
        if not dias:
            raise HTTPException(status_code=400, detail="Selecciona al menos un día válido.")
        actuales = DESCANSOS_MANUALES.setdefault(emp_id, set())
        for d in dias[:2]:
            actuales.add(d)
    PLAN_CACHE_KEY = None
    optimize_weekly(OptimizationRequest())
    return {"status": "ok", "empleado_id": emp_id, "descansos_manuales": sorted(list(DESCANSOS_MANUALES.get(emp_id, set())))}


@app.post("/manual-assignment")
def ajuste_manual(req: ManualAssignmentRequest):
    if not PLAN_OPTIMIZADO:
        raise HTTPException(status_code=400, detail="Primero optimiza el plan.")
    dia = normalizar_dia(req.dia)
    hora = bloque_a_hora(req.hora)
    area = normalizar_area(req.area)
    emp_id = str(req.empleado_id)
    if not dia or hora not in HORAS_BLOQUE or area not in AREAS_CINE:
        raise HTTPException(status_code=400, detail="Día, hora o área inválida.")
    e = next((ee for ee in DB_EMPLEADOS if str(ee["id"]) == emp_id), None)
    if not e or not esta_en_horario(hora, e["horarios_por_dia"].get(dia, "")):
        raise HTTPException(status_code=400, detail="Empleado fuera de horario o no encontrado.")
    PLAN_OPTIMIZADO[dia][hora][emp_id] = area
    actualizar_cache_horas()
    return {"status": "ok", "empleado_id": emp_id, "dia": dia, "hora": hora, "area": area}


# =====================================================================
# KPIs Y DASHBOARD
# =====================================================================
def estado_operativo_empleado(e: dict, dia: str, hora: str, asistencias_extra: Optional[Dict[str, str]] = None) -> Tuple[str, str]:
    emp_id = str(e["id"])
    if esta_en_falta(emp_id, dia, asistencias_extra):
        return "Falta", "Falta registrada en control de lista"
    if esta_en_descanso_manual(emp_id, dia):
        return "Descanso manual", "Descanso registrado por supervisor"
    rot = ROTACIONES_PLAN.get(dia, {}).get(hora, {}).get(emp_id)
    if rot:
        if rot == "Ley silla":
            return "Ley silla", f"Ley silla ({hora}-{sum_minutes(hora, 15)})"
        return "Alimentos", f"Alimentos ({hora}-{sum_minutes(hora, 30)})"
    area_plan = PLAN_OPTIMIZADO.get(dia, {}).get(hora, {}).get(emp_id, "Descanso General")
    if area_plan in AREAS_CINE:
        base = PLAN_EXCEL_BASE.get(dia, {}).get(hora, {}).get(emp_id, "Descanso General")
        if base in AREAS_CINE and base != area_plan:
            return "Reasignado por DSS", "Reasignación generada por DSS"
        return "Operando en piso", "Operando en estación"
    if esta_en_horario(hora, e["horarios_por_dia"].get(dia, "")):
        return "Descanso DSS", "Descanso automático para respetar presupuesto"
    return "Fuera de horario", "No está dentro del horario base"


def sum_minutes(hora: str, mins: int) -> str:
    total = minutes(hora) + mins
    return f"{(total // 60) % 24:02d}:{total % 60:02d}"


def generar_recomendaciones_dss(dia: str, hora: str, faltantes_por_area: Dict[str, int], asistencias_extra: Optional[Dict[str, str]] = None) -> List[dict]:
    recs = []
    for area_obj, faltantes in faltantes_por_area.items():
        if faltantes <= 0:
            continue
        candidatos = []
        for e in DB_EMPLEADOS:
            emp_id = str(e["id"])
            if esta_en_falta(emp_id, dia, asistencias_extra) or esta_en_descanso_manual(emp_id, dia):
                continue
            if ROTACIONES_PLAN.get(dia, {}).get(hora, {}).get(emp_id):
                continue
            if not esta_en_horario(hora, e["horarios_por_dia"].get(dia, "")):
                continue
            area_actual = PLAN_OPTIMIZADO.get(dia, {}).get(hora, {}).get(emp_id, "Descanso General")
            if area_actual == area_obj:
                continue
            horas_restantes = round(float(e.get("limite_maximo", MAX_HORAS_DEFAULT)) - calcular_horas_plan(empleado_id=emp_id), 2)
            exp = {normalizar_area(a) for a in e.get("areas_expertis", [])}
            nivel = int(e.get("nivel_esp", 1) or 1)
            poliv = len(exp)
            score = horas_restantes * 2 + (160 if area_obj in exp else 15) + poliv * 18 + nivel * 12
            if area_obj == "Dulcería" and "Dulcería" in exp:
                score += 85
            candidatos.append({
                "id": emp_id, "empleado": e["nombre"], "area_afectada": area_obj,
                "origen": area_actual, "mover_a": area_obj, "score": round(score, 1),
                "horas_restantes": horas_restantes, "nivel": nivel, "polivalencia": poliv,
                "habilidades": e.get("expertise", {}),
                "motivo": f"{horas_restantes} hrs restantes; {'certificado en ' + area_obj if area_obj in exp else 'polivalente de soporte'}; nivel {nivel}; {poliv} habilidades M1-M5.",
            })
        candidatos.sort(key=lambda c: c["score"], reverse=True)
        recs.extend(candidatos[:max(1, faltantes)])
    return sorted(recs, key=lambda c: c["score"], reverse=True)[:12]


def calcular_axh_por_dia(plan: Dict) -> List[dict]:
    out = []
    for d in DIAS_SEMANA:
        demanda = calcular_demanda(d, None)
        horas = calcular_horas_plan(plan if plan is not PLAN_OPTIMIZADO else None, dia=d)
        out.append({"dia": d, "demanda": demanda, "horas": horas, "axh": round(demanda / horas, 2) if horas > 0 else 0})
    return out


def plan_resumen_empleados() -> List[dict]:
    rows = []
    for e in DB_EMPLEADOS:
        emp_id = str(e["id"])
        dias_info = []
        cambios = 0
        for d in DIAS_SEMANA:
            conteo = {a: 0 for a in AREAS_CINE}
            for h in HORAS_BLOQUE:
                area = PLAN_OPTIMIZADO.get(d, {}).get(h, {}).get(emp_id, "Descanso General")
                if area in AREAS_CINE:
                    conteo[area] += 1
                    if area != PLAN_EXCEL_BASE.get(d, {}).get(h, {}).get(emp_id, "Descanso General"):
                        cambios += 1
            area_dss = max(conteo, key=conteo.get) if sum(conteo.values()) > 0 else "Descanso"
            horario = rango_operativo_empleado(e, d) if area_dss != "Descanso" else "Descanso"
            dias_info.append({"dia": d, "dss": area_dss, "area": area_dss, "horario": horario, "estado": "OK", "horas_dss": sum(conteo.values())})
        rows.append({
            "id": emp_id, "nombre": e["nombre"], "nivel": e.get("nivel_esp", 1),
            "habilidades": e.get("expertise", {}), "area_base": e.get("area_base", "Lobby"),
            "dias": dias_info, "horas_semana": calcular_horas_plan(empleado_id=emp_id), "cambios": cambios,
        })
    return rows


@app.post("/kpis")
def get_kpis(req: KPIRequest = Body(...)):
    if not DB_EMPLEADOS:
        raise HTTPException(status_code=400, detail="Primero sube el Excel.")
    if not PLAN_OPTIMIZADO:
        optimize_weekly(OptimizationRequest())
    dia = normalizar_dia(req.dia)
    if not dia:
        raise HTTPException(status_code=400, detail=f"Día inválido: {reparar_mojibake(req.dia)}")
    hora = bloque_a_hora(req.hora)
    demanda = calcular_demanda(dia, hora)
    demanda_dia = calcular_demanda(dia, None)

    distribucion = {a: 0 for a in AREAS_CINE}
    empleados_det = []
    faltas = []
    descansos = []

    for e in DB_EMPLEADOS:
        emp_id = str(e["id"])
        area = PLAN_OPTIMIZADO.get(dia, {}).get(hora, {}).get(emp_id, "Descanso General")
        estado, detalle = estado_operativo_empleado(e, dia, hora, req.asistencias)
        if estado in {"Operando en piso", "Reasignado por DSS"} and area in AREAS_CINE:
            distribucion[area] += 1
        if estado == "Falta":
            faltas.append({"id": emp_id, "nombre": e["nombre"], "area_afectada": area})
        if estado in {"Descanso manual", "Descanso DSS", "Ley silla", "Alimentos"}:
            descansos.append({"id": emp_id, "nombre": e["nombre"], "estado": estado})
        empleados_det.append({
            "id": e["id"], "id_ps": emp_id, "nombre": e["nombre"],
            "area": area if area in AREAS_CINE else "Ninguna", "area_planificada": area,
            "cambio_area": False, "estado": estado, "detalle_break": detalle,
            "horario": rango_operativo_empleado(e, dia) if area in AREAS_CINE else e["horarios_por_dia"].get(dia, ""), "asistencia": "FALTA" if estado == "Falta" else "PRESENTE",
            "nivel_esp": e.get("nivel_esp", 1), "expertise": e.get("expertise", {}),
            "limite_maximo": e.get("limite_maximo", MAX_HORAS_DEFAULT),
            "horas_usadas": calcular_horas_plan(empleado_id=emp_id),
        })

    activos = sum(distribucion.values())
    req_area = requerimiento_por_area(demanda, activos)
    deficit = {a: max(0, req_area[a] - distribucion[a]) for a in AREAS_CINE}
    exceso = {a: max(0, distribucion[a] - req_area[a]) for a in AREAS_CINE}
    recs = generar_recomendaciones_dss(dia, hora, deficit, req.asistencias)
    axh = round(demanda / activos, 2) if activos > 0 else 0
    cobertura = round((activos / max(1, sum(req_area.values()))) * 100, 1)
    horas = calcular_horas_plan()
    costo = round(horas * SUELDO_HORA_BASE, 2)
    poliv = round(sum(len(e.get("areas_expertis", [])) for e in DB_EMPLEADOS) / max(1, len(DB_EMPLEADOS)), 2)
    empleados_multi = sum(1 for e in DB_EMPLEADOS if len(e.get("areas_expertis", [])) >= 2)
    m2_disponibles = sum(1 for e in DB_EMPLEADOS if "Dulcería" in {normalizar_area(a) for a in e.get("areas_expertis", [])})

    alertas = []
    if deficit["Dulcería"] > 0:
        alertas.append(f"Déficit en Dulcería: faltan {deficit['Dulcería']} colaboradores. Dulcería concentra 80.52% de transacciones 2025.")
    if m2_disponibles < max(1, req_area.get("Dulcería", 0)):
        alertas.append("Capacitación insuficiente M2: no hay suficientes colaboradores certificados para Dulcería.")
    if horas > PRESUPUESTO_HORAS_GLOBAL:
        alertas.append(f"Sobreconsumo de horas: {horas:.0f} / {PRESUPUESTO_HORAS_GLOBAL:.0f}. Se requiere autorizar horas o reducir cobertura.")
    elif horas / max(1, PRESUPUESTO_HORAS_GLOBAL) >= 0.95:
        alertas.append(f"Consumo de horas alto: {horas:.0f} / {PRESUPUESTO_HORAS_GLOBAL:.0f}. Riesgo de exceder presupuesto.")
    for a in AREAS_CINE:
        if distribucion[a] == 0 and activos >= 5:
            alertas.append(f"Área sin cobertura: {a}. Revisar personal capacitado y restricciones de horario.")

    axh_semana = []
    for row in calcular_axh_por_dia(PLAN_OPTIMIZADO):
        axh_semana.append({"dia": row["dia"], "axh_optimo": row["axh"], "axh_base": 0, "horas_optimas": row["horas"], "horas_base": 0, "demanda": row["demanda"]})

    return {
        "dia": dia, "hora": hora, "axh": axh, "axh_meta": AXH_META,
        "desviacion_meta": round(axh - AXH_META, 2), "activos": activos,
        "demanda": demanda, "demanda_dia": demanda_dia, "capacidad": f"{min(100, int(cobertura))}%",
        "cobertura_operativa": cobertura, "costo_semanal": costo, "costo_total_estimado": costo,
        "costo_por_cliente": round(costo / max(1, TOTAL_ASISTENTES_SEMANA), 2),
        "empleados_detallados": empleados_det, "faltas_detectadas": faltas, "descansos_detectados": descansos,
        "alertas": alertas if alertas else ["Operación bajo curvas normales de asistencia."],
        "semaforos": {a: ("ROJO" if deficit[a] > 0 else "VERDE") for a in AREAS_CINE},
        "recomendaciones": recs, "ppto_horas_sem": PRESUPUESTO_HORAS_GLOBAL,
        "horas_asignadas_sem": horas, "horas_excel_base": calcular_horas_plan(PLAN_EXCEL_BASE),
        "horas_disponibles_sem": round(PRESUPUESTO_HORAS_GLOBAL - horas, 2),
        "consumo_horas_pct": round((horas / max(1, PRESUPUESTO_HORAS_GLOBAL)) * 100, 1),
        "axh_semana": axh_semana, "plan_resumen": plan_resumen_empleados(),
        "porcentaje_transacciones_dulceria": PORC_TRANSACCIONES_DULCERIA,
        "disponibles_m2_dulceria": m2_disponibles,
        "tiempo_optimizacion_seg": LAST_OPTIMIZATION_META.get("tiempo_seg", 0),
        "resumen_dss": {
            "polivalencia_promedio": poliv, "empleados_multi_area": empleados_multi,
            "faltas": len(faltas), "descansos": len(descansos), "recomendaciones": len(recs),
            "cambios_vs_excel": 0,
        },
        "analytics": {
            "labels_areas": AREAS_CINE,
            "personal_real": [distribucion[a] for a in AREAS_CINE],
            "personal_planificado": [distribucion[a] for a in AREAS_CINE],
            "personal_excel": [0 for _ in AREAS_CINE],
            "personal_ideal": [req_area[a] for a in AREAS_CINE],
            "deficit_area": [deficit[a] for a in AREAS_CINE],
            "exceso_area": [exceso[a] for a in AREAS_CINE],
            "utilizacion_area": [round((distribucion[a] / max(1, req_area[a])) * 100, 1) for a in AREAS_CINE],
            "transacciones_dulceria_pct": PORC_TRANSACCIONES_DULCERIA * 100,
        },
    }


@app.get("/empleados")
def get_empleados():
    return DB_EMPLEADOS


@app.get("/reporte-semanal")
def obtener_reporte_semanal():
    if not DB_EMPLEADOS:
        raise HTTPException(status_code=400, detail="Primero sube el Excel.")
    return {
        "presupuesto_horas": PRESUPUESTO_HORAS_GLOBAL,
        "horas_optimas": calcular_horas_plan(),
        "axh_semana_optima": calcular_axh_por_dia(PLAN_OPTIMIZADO),
        "plan_resumen": plan_resumen_empleados(),
        "descansos_manuales": {k: sorted(list(v)) for k, v in DESCANSOS_MANUALES.items()},
        "asistencias": ASISTENCIAS_REGISTRADAS,
    }


@app.post("/optimize")
def optimize_schedule(req: OptimizationRequest):
    return optimize_weekly(req)


@app.post("/asignar")
def generar_asignaciones(req: OptimizationRequest):
    return optimize_weekly(req)

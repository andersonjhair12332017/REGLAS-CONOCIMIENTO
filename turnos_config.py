import os
import sys
import json


ARCHIVO_TURNOS = "turnos_config.json"


TURNOS_DEFAULT = [
    {
        "nombre": "Día",
        "hora_inicio": "07:30",
        "hora_fin": "19:30",
        "activo": True,
        "meta_ppt_bodega": 0,
        "meta_ppt_despachos": 0,
        "metas_ppt_bodega_roles": {
            "Montacarguista": 0,
            "Pick and drop": 0,
            "Trilateral": 0
        }
    },
    {
        "nombre": "Noche",
        "hora_inicio": "19:30",
        "hora_fin": "07:30",
        "activo": True,
        "meta_ppt_bodega": 0,
        "meta_ppt_despachos": 0,
        "metas_ppt_bodega_roles": {
            "Montacarguista": 0,
            "Pick and drop": 0,
            "Trilateral": 0
        }
    }
]


def obtener_directorio_app():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)

    return os.path.dirname(os.path.abspath(__file__))


def obtener_ruta_turnos():
    return os.path.join(obtener_directorio_app(), ARCHIVO_TURNOS)


def _normalizar_bool(valor):
    if isinstance(valor, bool):
        return valor

    texto = str(valor).strip().lower()

    if texto in ("true", "1", "si", "sí", "s", "activo", "yes", "y"):
        return True

    if texto in ("false", "0", "no", "n", "inactivo"):
        return False

    return True


def _normalizar_numero(valor, default=0):
    try:
        if valor is None:
            return default

        texto = str(valor).strip().replace(",", ".")

        if texto == "" or texto.lower() == "nan":
            return default

        numero = float(texto)

        if numero < 0:
            return default

        return numero
    except Exception:
        return default


def _normalizar_turno(t):
    turno = dict(t)

    turno["nombre"] = str(turno.get("nombre", "")).strip()
    turno["hora_inicio"] = str(turno.get("hora_inicio", "07:30")).strip()
    turno["hora_fin"] = str(turno.get("hora_fin", "19:30")).strip()
    turno["activo"] = _normalizar_bool(turno.get("activo", True))

    turno["meta_ppt_bodega"] = _normalizar_numero(
        turno.get("meta_ppt_bodega", turno.get("meta_ppd_bodega", 0)),
        default=0
    )

    turno["meta_ppt_despachos"] = _normalizar_numero(
        turno.get("meta_ppt_despachos", turno.get("meta_ppd_despachos", 0)),
        default=0
    )
    
    metas_roles = turno.get("metas_ppt_bodega_roles", {})

    if not isinstance(metas_roles, dict):
        metas_roles = {}

    metas_roles_norm = {}

    for rol, meta in metas_roles.items():
        rol_txt = str(rol).strip()

        if not rol_txt:
            continue

        metas_roles_norm[rol_txt] = _normalizar_numero(meta, default=0)

    turno["metas_ppt_bodega_roles"] = metas_roles_norm   

    return turno


def crear_turnos_si_no_existe():
    ruta = obtener_ruta_turnos()

    if os.path.exists(ruta):
        return ruta

    guardar_turnos(TURNOS_DEFAULT)

    return ruta


def cargar_turnos(solo_activos=False):
    crear_turnos_si_no_existe()

    ruta = obtener_ruta_turnos()

    try:
        with open(ruta, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        data = TURNOS_DEFAULT

    if not isinstance(data, list):
        data = TURNOS_DEFAULT

    turnos = [_normalizar_turno(t) for t in data if str(t.get("nombre", "")).strip()]

    if solo_activos:
        turnos = [t for t in turnos if t.get("activo", True)]

    return turnos


def guardar_turnos(turnos):
    ruta = obtener_ruta_turnos()

    turnos_normalizados = [_normalizar_turno(t) for t in turnos if str(t.get("nombre", "")).strip()]

    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(turnos_normalizados, f, ensure_ascii=False, indent=4)

    return ruta


def obtener_nombres_turnos(turnos):
    return [str(t.get("nombre", "")).strip() for t in turnos if str(t.get("nombre", "")).strip()]


def obtener_turno_por_nombre(turnos, nombre):
    nombre = str(nombre).strip()

    for turno in turnos:
        if str(turno.get("nombre", "")).strip() == nombre:
            return turno

    return None
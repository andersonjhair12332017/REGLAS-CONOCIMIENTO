import PySimpleGUI as sg
from datetime import datetime
import os

from procesamiento_rango import procesar_rango_operacion


from wms_processor import procesar_archivo

from wms_validators import (
    validar_fecha_ui_inmediata,
    validar_excel_y_autoseleccionar_fecha_ui,
    MAX_DIAS_RANGO_EXCEL
)

from historico_sqlite import exportar_reporte_mensual

from turnos_config import (
    cargar_turnos,
    obtener_nombres_turnos,
    obtener_turno_por_nombre,
    guardar_turnos
)

from turnos_admin_ui import abrir_admin_turnos
from maestro_personal_editor import abrir_editor_maestro_personal



# ============================================================
# CONFIGURACIÓN VISUAL
# ============================================================

sg.theme("Reddit")

font_title = ("Segoe UI", 16, "bold")
font_body = ("Segoe UI", 11)
font_small = ("Segoe UI", 10)


# ============================================================
# CARGA INICIAL DE TURNOS
# ============================================================

turnos_disponibles = cargar_turnos(solo_activos=True)
nombres_turnos = obtener_nombres_turnos(turnos_disponibles)

turno_default = turnos_disponibles[0] if turnos_disponibles else {
    "nombre": "Día",
    "hora_inicio": "07:30",
    "hora_fin": "19:30",
    "activo": True,
    "meta_ppt_bodega": 0,
    "meta_ppt_despachos": 0
}


def _texto_metas_turno(turno):
    """
    Devuelve texto informativo de metas PPT del turno seleccionado.
    """
    if not turno:
        return "Meta PPT Bodega: 0 | Meta PPT Despachos: 0"

    meta_bodega = turno.get("meta_ppt_bodega", 0)
    meta_despachos = turno.get("meta_ppt_despachos", 0)

    return f"Meta PPT Bodega: {meta_bodega} | Meta PPT Despachos: {meta_despachos}"


# ============================================================
# LAYOUT
# ============================================================

layout = [
    [
        sg.Text(
            "📦 ANALIZADOR DE PRODUCTIVIDAD WMS",
            font=font_title,
            text_color="#004080",
            justification="center",
            expand_x=True
        )
    ],

    [sg.HorizontalSeparator(color="#cccccc")],

    [
        sg.Frame(
            "⚙️ Configuración de Archivos",
            [
                [
                    sg.Text("Excel Origen:", size=(12, 1), font=font_body),
                    sg.Input(
                        key="-FILE-",
                        expand_x=True,
                        font=font_body,
                        enable_events=True
                    ),
                    sg.FileBrowse("📁 Buscar", font=font_body)
                ],
                [
                    sg.Text("Ruta Destino:", size=(12, 1), font=font_body),
                    sg.Input(
                        key="-FOLDER-",
                        expand_x=True,
                        font=font_body
                    ),
                    sg.FolderBrowse("📁 Ver", font=font_body)
                ]
            ],
            font=("Segoe UI", 11, "bold"),
            expand_x=True,
            pad=((0, 0), (10, 10))
        )
    ],

    [
    sg.Frame(
        "🕒 Parámetros para análisis diario",
        [
            [
                sg.Text("Fecha Operación:", size=(14, 1), font=font_body),
                sg.Input(
                    datetime.now().strftime("%d/%m/%Y"),
                    key="-FECHA-",
                    size=(12, 1),
                    font=font_body,
                    enable_events=True
                ),
                sg.CalendarButton(
                    "📅 Calendario",
                    target="-FECHA-",
                    format="%d/%m/%Y",
                    font=font_body
                )
            ],

            [
                sg.Text(
                    f"Seleccione Excel origen. Máximo permitido: {MAX_DIAS_RANGO_EXCEL} días de rango.",
                    key="-FECHA_ESTADO-",
                    font=("Segoe UI", 9, "bold"),
                    text_color="#555555",
                    expand_x=True
                )
            ],

            [
                sg.Text("Turno:", size=(5, 1), font=font_body),

                sg.Combo(
                    nombres_turnos,
                    default_value=turno_default.get("nombre", ""),
                    key="-TURNO-",
                    readonly=True,
                    enable_events=True,
                    size=(18, 1),
                    font=font_body
                ),

                sg.Button(
                    "⚙️ Administrar Turnos y Metas PPT",
                    key="-ADMIN_TURNOS-",
                    size=(30, 1),
                    font=font_body
                ),

                sg.Button(
                    "👥 Administrar Personal",
                    key="-ADMIN_PERSONAL-",
                    size=(22, 1),
                    font=font_body
                )
            ],

            [
                sg.Text("Inicio:", size=(5, 1), font=font_body),
                sg.Input(
                    turno_default.get("hora_inicio", "07:30"),
                    key="-INI-",
                    size=(8, 1),
                    font=font_body
                ),

                sg.Text("Fin:", size=(4, 1), font=font_body),
                sg.Input(
                    turno_default.get("hora_fin", "19:30"),
                    key="-FIN-",
                    size=(8, 1),
                    font=font_body
                ),

                sg.Text(
                    _texto_metas_turno(turno_default),
                    key="-METAS_TURNO-",
                    font=("Segoe UI", 9, "bold"),
                    text_color="#004080",
                    expand_x=True
                )
            ],
            ],
            font=("Segoe UI", 11, "bold"),
            expand_x=True,
            pad=((0, 0), (0, 10))
        )
    ],
    
    
    [
    sg.Frame(
        "🕒 Parámetros para ánalisis por rango de fecha",
        [
            [
                sg.Text("Fecha Inicio Rango:", size=(17, 1), font=font_body),
                sg.Input(
                    key="-FECHA_INICIO_RANGO-",
                    size=(12, 1),
                    font=font_body
                ),
                sg.CalendarButton(
                    "📅 Calendario",
                    target="-FECHA_INICIO_RANGO-",
                    format="%d/%m/%Y",
                    font=font_body
                ),

                sg.Text("Fecha Fin Rango:", size=(15, 1), font=font_body),
                sg.Input(
                    key="-FECHA_FIN_RANGO-",
                    size=(12, 1),
                    font=font_body
                ),
                sg.CalendarButton(
                    "📅 Calendario",
                    target="-FECHA_FIN_RANGO-",
                    format="%d/%m/%Y",
                    font=font_body
                )
            ]

            ],
            font=("Segoe UI", 11, "bold"),
            expand_x=True,
            pad=((0, 0), (0, 10))
        )
    ],
                                      

    [
        sg.Button(
            "🚀 ANÁLISIS DIARIO",
            key="-PROCESAR-",
            size=(30, 2),
            font=("Segoe UI", 12, "bold"),
            button_color=("white", "#0059b3"),
            pad=((0, 0), (10, 10)),
            expand_x=True
        )
    ],
    
    [
        sg.Button(
            "📆 ANÁLISIS POR RANGO DE FECHA",
            key="-PROCESAR_RANGO-",
            size=(40, 2),
            button_color=("white", "#107C10"),
            font=("Segoe UI", 12, "bold"),
            expand_x=True
        )
    ],

    [
        sg.Button(
            "📊 GENERAR REPORTE MENSUAL",
            key="-MENSUAL-",
            size=(30, 2),
            font=("Segoe UI", 12, "bold"),
            button_color=("white", "#6f42c1"),
            expand_x=True
        )
    ],

    [
        sg.Frame(
            "📊 Progreso de Ejecución",
            [
                [
                    sg.ProgressBar(
                        100,
                        orientation="h",
                        size=(40, 20),
                        key="-PROG-",
                        expand_x=True,
                        bar_color=("#0059b3", "#e6e6e6")
                    ),
                    sg.Text(
                        "0%",
                        key="-PERC-",
                        font=("Segoe UI", 12, "bold"),
                        size=(5, 1),
                        justification="right",
                        text_color="#004080"
                    )
                ],
                [
                    sg.Text("Estado:", font=("Segoe UI", 10, "bold")),
                    sg.Text(
                        "Esperando inicio...",
                        key="-STATUS-",
                        font=font_small,
                        text_color="#555555",
                        expand_x=True
                    )
                ]
            ],
            font=("Segoe UI", 11, "bold"),
            expand_x=True
        )
    ],

    [
        sg.Button(
            "📂 Abrir Reporte Generado",
            key="-OPEN-",
            visible=False,
            font=font_body,
            button_color=("white", "#28a745"),
            expand_x=True
        )
    ]
]


# ============================================================
# VENTANA PRINCIPAL
# ============================================================

window = sg.Window(
    "Productividad WMS v3.0",
    layout,
    finalize=True
)


resultado_path = None
procesando = False

cache_fechas_excel = {
    "ruta": None,
    "fecha_min": None,
    "fecha_max": None,
    "total_validas": 0,
    "dias_rango": 0,
    "bloqueado_por_rango": False,
    "excel_ok": False
}


# ============================================================
# LOOP PRINCIPAL
# ============================================================

while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED:
        break
    
    # ========================================================
    # CAMBIO / CARGA DE EXCEL ORIGEN
    # ========================================================
    if event == "-FILE-":
        window["-FECHA_ESTADO-"].update(
            "⏳ Validando Excel origen...",
            text_color="#004080"
        )
        window.refresh()

        ok_excel, resultado_excel = validar_excel_y_autoseleccionar_fecha_ui(
            window=window,
            ruta_excel=values.get("-FILE-", ""),
            cache_fechas=cache_fechas_excel,
            max_dias_rango=MAX_DIAS_RANGO_EXCEL
        )

        if not ok_excel and resultado_excel.get("bloqueado_por_rango"):
            sg.popup_error(
                resultado_excel["mensaje"],
                font=font_body,
                title="Excel origen bloqueado por rango amplio"
            )

    # ========================================================
    # VALIDACIÓN INMEDIATA DE FECHA
    # ========================================================
    if event == "-FECHA-":
        validar_fecha_ui_inmediata(
            window=window,
            ruta_excel=values.get("-FILE-", ""),
            fecha_str=values.get("-FECHA-", ""),
            cache_fechas=cache_fechas_excel
        )
        
    if event == "-PROCESAR_RANGO-":
        if procesando:
            sg.popup_error(
                "Ya hay un proceso en ejecución.",
                title="Proceso activo"
            )
            continue
    
        try:
            procesando = True
            
            window["-PROCESAR-"].update(disabled=True)
            window["-PROCESAR_RANGO-"].update(disabled=True)
            window["-MENSUAL-"].update(disabled=True)
            window["-ADMIN_TURNOS-"].update(disabled=True)
            window["-ADMIN_PERSONAL-"].update(disabled=True)
            window["-OPEN-"].update(visible=False)
    
            ruta_excel = values.get("-FILE-", "")
            carpeta_destino = values.get("-FOLDER-", "")
    
            fecha_inicio = values.get("-FECHA_INICIO_RANGO-", "")
            fecha_fin = values.get("-FECHA_FIN_RANGO-", "")
    
            if not fecha_inicio:
                fecha_inicio = values.get("-FECHA-", "")
    
            if not fecha_fin:
                fecha_fin = fecha_inicio
    
            turnos_rango = cargar_turnos(solo_activos=True)
    
            resultado = procesar_rango_operacion(
                ruta_excel=ruta_excel,
                carpeta_destino=carpeta_destino,
                fecha_inicio_str=fecha_inicio,
                fecha_fin_str=fecha_fin,
                turnos_activos=turnos_rango,
                window=window
            )
    
            window["-PROG-"].update(100)
            window["-PERC-"].update("100%")
            window["-STATUS-"].update("Procesamiento de rango finalizado.")
            window.refresh()
    
            resultado_path = resultado.get("archivo_mensual", "")

            mensaje = (
                f"Procesamiento mensual finalizado.\n\n"
                f"Procesos totales: {resultado['total_procesos']}\n"
                f"Procesados OK: {resultado['procesados_ok']}\n"
                f"Errores: {len(resultado['errores'])}\n\n"
                f"Excel mensual generado:\n{resultado_path}"
            )

            sg.popup(
                mensaje,
                title="Procesamiento mensual"
            )

            if resultado_path:
                window["-OPEN-"].update(visible=True)
    
        except Exception as e:
            sg.popup_error(
                f"No se pudo procesar el rango.\n\nDetalle:\n{e}",
                title="Error procesamiento rango"
            )
    
            try:
                window["-STATUS-"].update("Error en procesamiento de rango.")
            except Exception:
                pass

            
        finally:
            procesando = False

            window["-PROCESAR-"].update(disabled=False)
            window["-PROCESAR_RANGO-"].update(disabled=False)
            window["-MENSUAL-"].update(disabled=False)
            window["-ADMIN_TURNOS-"].update(disabled=False)
            window["-ADMIN_PERSONAL-"].update(disabled=False)

    # ========================================================
    # CAMBIO DE TURNO
    # ========================================================
    if event == "-TURNO-":
        turno_sel = obtener_turno_por_nombre(
            turnos_disponibles,
            values.get("-TURNO-", "")
        )

        if turno_sel:
            window["-INI-"].update(turno_sel.get("hora_inicio", ""))
            window["-FIN-"].update(turno_sel.get("hora_fin", ""))
            window["-METAS_TURNO-"].update(_texto_metas_turno(turno_sel))

            validar_fecha_ui_inmediata(
                window=window,
                ruta_excel=values.get("-FILE-", ""),
                fecha_str=values.get("-FECHA-", ""),
                cache_fechas=cache_fechas_excel
            )

    # ========================================================
    # ADMINISTRAR TURNOS Y METAS PPT
    # ========================================================
    if event == "-ADMIN_TURNOS-":
        turnos_editables = cargar_turnos(solo_activos=False)

        nuevos_turnos = abrir_admin_turnos(turnos_editables)

        if nuevos_turnos is not None:
            try:
                guardar_turnos(nuevos_turnos)

                turnos_disponibles = cargar_turnos(solo_activos=True)
                nombres_turnos = obtener_nombres_turnos(turnos_disponibles)

                turno_actual = values.get("-TURNO-", "")

                if turno_actual not in nombres_turnos:
                    turno_actual = nombres_turnos[0] if nombres_turnos else ""

                window["-TURNO-"].update(
                    values=nombres_turnos,
                    value=turno_actual
                )

                turno_sel = obtener_turno_por_nombre(
                    turnos_disponibles,
                    turno_actual
                )

                if turno_sel:
                    window["-INI-"].update(turno_sel.get("hora_inicio", ""))
                    window["-FIN-"].update(turno_sel.get("hora_fin", ""))
                    window["-METAS_TURNO-"].update(_texto_metas_turno(turno_sel))
                else:
                    window["-INI-"].update("")
                    window["-FIN-"].update("")
                    window["-METAS_TURNO-"].update(
                        "Meta PPT Bodega: 0 | Meta PPT Despachos: 0"
                    )

                sg.popup(
                    "Turnos y metas PPT actualizados",
                    "La configuración fue guardada correctamente.",
                    font=font_body,
                    title="Turnos y Metas PPT"
                )

            except Exception as e:
                sg.popup_error(
                    f"No se pudieron guardar los turnos:\n{str(e)}",
                    font=font_body,
                    title="Error guardando turnos"
                )

    # ========================================================
    # ADMINISTRAR PERSONAL
    # ========================================================
    if event in ("-ADMIN_PERSONAL-", "-ADMINISTRAR_PERSONAL-", "Administrar Personal"):
        abrir_editor_maestro_personal()

    # ========================================================
    # PROCESAR DATOS
    # ========================================================
    if event == "-PROCESAR-":
        if procesando:
            continue

        if not values.get("-FILE-") or not values.get("-FOLDER-"):
            sg.popup_error(
                "⚠️ Faltan rutas de archivo o carpeta destino",
                font=font_body
            )
            continue

        if cache_fechas_excel.get("bloqueado_por_rango"):
            sg.popup_error(
                "⚠ El Excel origen está bloqueado porque contiene un rango de fechas demasiado amplio.\n\n"
                f"Rango detectado: {cache_fechas_excel.get('dias_rango')} días\n"
                f"Máximo permitido: {MAX_DIAS_RANGO_EXCEL} días\n\n"
                "Carga un Excel filtrado con menos días para evitar saturación del software.",
                font=font_body,
                title="Excel origen no permitido"
            )
            continue

        if not cache_fechas_excel.get("excel_ok"):
            ok_excel, resultado_excel = validar_excel_y_autoseleccionar_fecha_ui(
                window=window,
                ruta_excel=values.get("-FILE-", ""),
                cache_fechas=cache_fechas_excel,
                max_dias_rango=MAX_DIAS_RANGO_EXCEL
            )

            if not ok_excel:
                sg.popup_error(
                    resultado_excel["mensaje"],
                    font=font_body,
                    title="Excel origen no válido"
                )
                continue

        fecha_ok = validar_fecha_ui_inmediata(
            window=window,
            ruta_excel=values.get("-FILE-", ""),
            fecha_str=values.get("-FECHA-", ""),
            cache_fechas=cache_fechas_excel
        )

        if not fecha_ok:
            sg.popup_error(
                "⚠ La fecha de operación seleccionada no es válida para el Excel origen.\n\n"
                "Corrige la fecha antes de procesar.",
                font=font_body,
                title="Fecha no válida"
            )
            continue

        try:
            datetime.strptime(values["-INI-"].strip(), "%H:%M")
            datetime.strptime(values["-FIN-"].strip(), "%H:%M")
        except Exception:
            sg.popup_error(
                "⚠ Las horas del turno no son válidas.\n\n"
                "Usa formato HH:MM, por ejemplo 07:30 o 19:30.",
                font=font_body,
                title="Horas inválidas"
            )
            continue

        turno_sel = obtener_turno_por_nombre(
            turnos_disponibles,
            values.get("-TURNO-", "")
        )

        if not turno_sel:
            sg.popup_error(
                "⚠ No se encontró la configuración del turno seleccionado.\n\n"
                "Verifica la configuración en Administrar Turnos y Metas PPT.",
                font=font_body,
                title="Turno no válido"
            )
            continue

        procesando = True

        window["-PROCESAR-"].update(disabled=True)
        window["-MENSUAL-"].update(disabled=True)
        window["-ADMIN_TURNOS-"].update(disabled=True)
        window["-ADMIN_PERSONAL-"].update(disabled=True)
        window["-OPEN-"].update(visible=False)

        try:
            window["-PERC-"].update("0%")
            window["-PROG-"].update(0)
            window["-STATUS-"].update("Iniciando proceso...")
            window.refresh()

            resultado_path = procesar_archivo(
                values["-FILE-"],
                values["-FOLDER-"],
                values["-INI-"],
                values["-FIN-"],
                values["-FECHA-"],
                window,
                turno_config=turno_sel
            )

            window["-OPEN-"].update(visible=True)

            sg.popup(
                "¡Éxito!",
                f"Archivo generado correctamente en:\n{resultado_path}",
                font=font_body,
                title="Completado"
            )

        except Exception as e:
            sg.popup_error(
                f"Error durante la ejecución:\n{str(e)}",
                font=font_body,
                title="Error Crítico"
            )
            window["-STATUS-"].update(str(e)[:140])

        finally:
            procesando = False

            window["-PROCESAR-"].update(disabled=False)
            window["-MENSUAL-"].update(disabled=False)
            window["-ADMIN_TURNOS-"].update(disabled=False)
            window["-ADMIN_PERSONAL-"].update(disabled=False)

    # ========================================================
    # REPORTE MENSUAL
    # ========================================================
    if event == "-MENSUAL-":
        if not values.get("-FOLDER-"):
            sg.popup_error(
                "⚠️ Selecciona primero la Ruta Destino donde está el histórico.",
                font=font_body,
                title="Ruta requerida"
            )
            continue

        ruta_db = os.path.join(values["-FOLDER-"], "historico_productividad.db")

        if not os.path.exists(ruta_db):
            sg.popup_error(
                "⚠️ No se encontró la base histórica SQLite en la ruta seleccionada.\n\n"
                f"Ruta esperada:\n{ruta_db}\n\n"
                "Primero debes procesar al menos un turno para crear el histórico.",
                font=font_body,
                title="Histórico no encontrado"
            )
            continue

        try:
            fecha_ref = datetime.strptime(values["-FECHA-"].strip(), "%d/%m/%Y")
            anio = fecha_ref.year
            mes = fecha_ref.month

            salida_mensual = exportar_reporte_mensual(
                values["-FOLDER-"],
                anio,
                mes
            )

            sg.popup(
                "Reporte mensual generado",
                f"Archivo creado correctamente:\n{salida_mensual}",
                font=font_body,
                title="Mensual"
            )

            os.startfile(salida_mensual)

        except Exception as e:
            sg.popup_error(
                f"No se pudo generar el reporte mensual:\n{str(e)}",
                font=font_body,
                title="Error mensual"
            )

    # ========================================================
    # ABRIR REPORTE GENERADO
    # ========================================================
    if event == "-OPEN-" and resultado_path:
        os.startfile(resultado_path)


window.close()
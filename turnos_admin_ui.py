import PySimpleGUI as sg

from turnos_config import guardar_turnos

try:
    from personal_config import cargar_maestro_personal
except Exception:
    cargar_maestro_personal = None


def _tabla_turnos(turnos):
    filas = []

    for t in turnos:
        filas.append([
            str(t.get("nombre", "")),
            str(t.get("hora_inicio", "")),
            str(t.get("hora_fin", "")),
            "Sí" if bool(t.get("activo", True)) else "No",
            str(t.get("meta_ppt_bodega", 0)),
            str(t.get("meta_ppt_despachos", 0))
        ])

    return filas


def _parse_meta(valor):
    try:
        texto = str(valor).strip().replace(",", ".")

        if texto == "":
            return 0

        numero = float(texto)

        if numero < 0:
            return 0

        return numero
    except Exception:
        return 0
    
def _obtener_roles_bodega_maestro():
    """
    Obtiene roles únicos de la hoja/maestro Bodega.
    Si no puede cargar el maestro, retorna lista base.
    """

    roles_base = [
        "Pick and Drop",
        "Trilateral",
        "Montacarguista"
    ]

    if cargar_maestro_personal is None:
        return roles_base

    try:
        df = cargar_maestro_personal()
    except Exception:
        return roles_base

    if df is None or df.empty:
        return roles_base

    if "Area" not in df.columns:
        df["Area"] = ""

    if "Rol" not in df.columns:
        df["Rol"] = ""

    df_bodega = df[
        df["Area"].astype(str).str.strip().isin(["Bodega", "Mixto"])
    ].copy()

    roles = (
        df_bodega["Rol"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    roles = sorted(set([r for r in roles if r]))

    if not roles:
        return roles_base

    return roles


def _tabla_metas_roles(turno, roles_disponibles):
    """
    Construye filas para tabla:
    Rol | Meta PPT
    """

    metas = turno.get("metas_ppt_bodega_roles", {})

    if not isinstance(metas, dict):
        metas = {}

    roles = list(roles_disponibles)

    for rol in metas.keys():
        if rol not in roles:
            roles.append(rol)

    filas = []

    for rol in roles:
        filas.append([
            rol,
            str(metas.get(rol, 0))
        ])

    return filas


def _normalizar_metas_roles_desde_tabla(filas):
    """
    Convierte filas de tabla Rol | Meta PPT a dict.
    """

    resultado = {}

    for fila in filas:
        if not isinstance(fila, list) or len(fila) < 2:
            continue

        rol = str(fila[0]).strip()
        meta = _parse_meta(fila[1])

        if not rol:
            continue

        resultado[rol] = meta

    return resultado



def abrir_admin_turnos(turnos_editables):
    turnos = [dict(t) for t in turnos_editables]
    seleccion_index = None
    
    
    roles_bodega_disponibles = _obtener_roles_bodega_maestro()
    metas_roles_actuales = []
    seleccion_meta_rol_index = None

    

    layout = [
        [
            sg.Text(
                "Administrar Turnos y Metas PPT",
                font=("Segoe UI", 14, "bold"),
                justification="center",
                expand_x=True
            )
        ],
        [
            sg.Text(
                "PPT significa Productividad por Turno. Configura aquí las metas por turno.",
                font=("Segoe UI", 10)
            )
        ],
        [sg.HorizontalSeparator()],
        [
            sg.Table(
                values=_tabla_turnos(turnos),
                headings=[
                    "Turno",
                    "Inicio",
                    "Fin",
                    "Activo",
                    "Meta PPT Bodega",
                    "Meta PPT Despachos"
                ],
                key="-TABLA_TURNOS-",
                auto_size_columns=False,
                col_widths=[18, 10, 10, 8, 18, 22],
                justification="left",
                num_rows=8,
                enable_events=True,
                expand_x=True
            )
        ],
        [sg.HorizontalSeparator()],
        [
            sg.Text("Nombre:", size=(12, 1)),
            sg.Input(key="-NOMBRE_TURNO-", size=(22, 1)),
            sg.Text("Inicio:", size=(8, 1)),
            sg.Input(key="-INI_TURNO-", size=(8, 1)),
            sg.Text("Fin:", size=(5, 1)),
            sg.Input(key="-FIN_TURNO-", size=(8, 1))
        ],
        [
            sg.Text("Meta PPT Bodega:", size=(16, 1)),
            sg.Input(key="-META_BODEGA-", size=(12, 1)),
            sg.Text("Meta PPT Despachos:", size=(20, 1)),
            sg.Input(key="-META_DESPACHOS-", size=(12, 1)),
            sg.Checkbox("Activo", key="-ACTIVO_TURNO-", default=True)
        ],
        
        [sg.HorizontalSeparator()],
        [
            sg.Text(
                "Metas PPT Bodega por Rol",
                font=("Segoe UI", 11, "bold")
            )
        ],
        [
            sg.Text(
                "Estas metas aplican según la columna Rol del maestro personal. "
                "Si un rol no tiene meta, se usa la Meta PPT Bodega general.",
                font=("Segoe UI", 9)
            )
        ],
        [
            sg.Table(
                values=[],
                headings=[
                    "Rol",
                    "Meta PPT"
                ],
                key="-TABLA_METAS_ROL-",
                auto_size_columns=False,
                col_widths=[28, 14],
                justification="left",
                num_rows=6,
                enable_events=True,
                expand_x=True
            )
        ],
        [
            sg.Text("Rol:", size=(8, 1)),
            sg.Combo(
                values=roles_bodega_disponibles,
                key="-ROL_META-",
                size=(28, 1)
            ),
            sg.Text("Meta PPT:", size=(10, 1)),
            sg.Input(key="-META_ROL-", size=(12, 1)),
            sg.Button("Agregar/Actualizar rol", key="-GUARDAR_META_ROL-", size=(22, 1)),
            sg.Button("Eliminar rol", key="-ELIMINAR_META_ROL-", size=(14, 1))
        ],
        
        
        [
            sg.Button("🆕 Nuevo", key="-NUEVO_TURNO-", size=(12, 1)),
            sg.Button("💾 Guardar registro", key="-GUARDAR_REG_TURNO-", size=(18, 1)),
            sg.Button("🗑 Eliminar", key="-ELIMINAR_TURNO-", size=(12, 1))
        ],
        [sg.HorizontalSeparator()],
        [
            sg.Button("✅ Guardar todo", key="-GUARDAR_TODO_TURNOS-", size=(16, 1)),
            sg.Button("❌ Cancelar", key="-CANCELAR_TURNOS-", size=(12, 1))
        ]
    ]

    win = sg.Window(
        "Administrar Turnos y Metas PPT",
        layout,
        modal=True,
        finalize=True,
        resizable=False
    )

    def refrescar_tabla():
        win["-TABLA_TURNOS-"].update(values=_tabla_turnos(turnos))
        
    def refrescar_tabla_metas_roles():
        win["-TABLA_METAS_ROL-"].update(values=metas_roles_actuales)


    def cargar_metas_roles_del_turno(turno):
        nonlocal metas_roles_actuales, seleccion_meta_rol_index

        seleccion_meta_rol_index = None
        metas_roles_actuales = _tabla_metas_roles(
            turno,
            roles_bodega_disponibles
        )

        refrescar_tabla_metas_roles()

        win["-ROL_META-"].update("")
        win["-META_ROL-"].update("")
    

    def limpiar_formulario():
        nonlocal seleccion_index, metas_roles_actuales, seleccion_meta_rol_index

        seleccion_index = None
        seleccion_meta_rol_index = None

        win["-NOMBRE_TURNO-"].update("")
        win["-INI_TURNO-"].update("07:30")
        win["-FIN_TURNO-"].update("19:30")
        win["-META_BODEGA-"].update("0")
        win["-META_DESPACHOS-"].update("0")
        win["-ACTIVO_TURNO-"].update(True)

        metas_roles_actuales = _tabla_metas_roles(
            {
                "metas_ppt_bodega_roles": {}
            },
            roles_bodega_disponibles
        )

        refrescar_tabla_metas_roles()

        win["-ROL_META-"].update("")
        win["-META_ROL-"].update("")

    limpiar_formulario()

    resultado = None

    while True:
        event, values = win.read()

        if event in (sg.WIN_CLOSED, "-CANCELAR_TURNOS-"):
            resultado = None
            break

        if event == "-TABLA_TURNOS-":
            seleccion = values.get("-TABLA_TURNOS-", [])

            if seleccion:
                seleccion_index = seleccion[0]
                row = turnos[seleccion_index]

                win["-NOMBRE_TURNO-"].update(str(row.get("nombre", "")))
                win["-INI_TURNO-"].update(str(row.get("hora_inicio", "")))
                win["-FIN_TURNO-"].update(str(row.get("hora_fin", "")))
                win["-META_BODEGA-"].update(str(row.get("meta_ppt_bodega", 0)))
                win["-META_DESPACHOS-"].update(str(row.get("meta_ppt_despachos", 0)))
                win["-ACTIVO_TURNO-"].update(bool(row.get("activo", True)))
                cargar_metas_roles_del_turno(row)

        if event == "-NUEVO_TURNO-":
            limpiar_formulario()
            
        if event == "-TABLA_METAS_ROL-":
            seleccion_meta = values.get("-TABLA_METAS_ROL-", [])

            if seleccion_meta:
                seleccion_meta_rol_index = seleccion_meta[0]

                try:
                    fila_meta = metas_roles_actuales[seleccion_meta_rol_index]
                    win["-ROL_META-"].update(str(fila_meta[0]))
                    win["-META_ROL-"].update(str(fila_meta[1]))
                except Exception:
                    pass


        if event == "-GUARDAR_META_ROL-":
            if seleccion_index is None:
                sg.popup_error(
                    "Primero selecciona un turno de la tabla superior.",
                    title="Turno requerido"
                )
                continue

            rol = str(values.get("-ROL_META-", "")).strip()
            meta = _parse_meta(values.get("-META_ROL-", 0))

            if not rol:
                sg.popup_error(
                    "Debes seleccionar o escribir un rol.",
                    title="Rol requerido"
                )
                continue

            # Asegurar que la lista exista
            if metas_roles_actuales is None:
                metas_roles_actuales = []

            actualizado = False

            for i, fila in enumerate(metas_roles_actuales):
                if not isinstance(fila, list) or len(fila) < 2:
                    continue

                rol_existente = str(fila[0]).strip()

                if rol_existente.lower() == rol.lower():
                    metas_roles_actuales[i] = [rol, str(meta)]
                    actualizado = True
                    break

            if not actualizado:
                metas_roles_actuales.append([rol, str(meta)])

            refrescar_tabla_metas_roles()

            # Si hay un turno seleccionado, actualizarlo inmediatamente en memoria
            if seleccion_index is not None:
                turnos[seleccion_index]["metas_ppt_bodega_roles"] = (
                    _normalizar_metas_roles_desde_tabla(metas_roles_actuales)
                )

            win["-ROL_META-"].update("")
            win["-META_ROL-"].update("")
            


        if event == "-ELIMINAR_META_ROL-":
            if seleccion_index is None:
                sg.popup_error(
                    "Primero selecciona un turno de la tabla superior.",
                    title="Turno requerido"
                )
                continue

        if seleccion_meta_rol_index is None:

            try:
                metas_roles_actuales.pop(seleccion_meta_rol_index)
            except Exception:
                pass

            seleccion_meta_rol_index = None
            refrescar_tabla_metas_roles()

            # Si hay un turno seleccionado, actualizarlo inmediatamente en memoria
            if seleccion_index is not None:
                turnos[seleccion_index]["metas_ppt_bodega_roles"] = (
                    _normalizar_metas_roles_desde_tabla(metas_roles_actuales)
                )

            win["-ROL_META-"].update("")
            win["-META_ROL-"].update("")
            

        if event == "-GUARDAR_REG_TURNO-":
            nombre = str(values.get("-NOMBRE_TURNO-", "")).strip()
            hora_inicio = str(values.get("-INI_TURNO-", "")).strip()
            hora_fin = str(values.get("-FIN_TURNO-", "")).strip()
            meta_bodega = _parse_meta(values.get("-META_BODEGA-", 0))
            meta_despachos = _parse_meta(values.get("-META_DESPACHOS-", 0))
            activo = bool(values.get("-ACTIVO_TURNO-", True))

            if not nombre:
                sg.popup_error("⚠ Debes ingresar el nombre del turno.", title="Dato requerido")
                continue

            if len(hora_inicio) != 5 or ":" not in hora_inicio:
                sg.popup_error("⚠ Hora inicio inválida. Usa formato HH:MM.", title="Hora inválida")
                continue

            if len(hora_fin) != 5 or ":" not in hora_fin:
                sg.popup_error("⚠ Hora fin inválida. Usa formato HH:MM.", title="Hora inválida")
                continue

            nuevo = {
                "nombre": nombre,
                "hora_inicio": hora_inicio,
                "hora_fin": hora_fin,
                "activo": activo,
                "meta_ppt_bodega": meta_bodega,
                "meta_ppt_despachos": meta_despachos,
                "metas_ppt_bodega_roles": _normalizar_metas_roles_desde_tabla(
                    metas_roles_actuales
                )
            }

            if seleccion_index is None:
                nombres_existentes = [str(t.get("nombre", "")).strip() for t in turnos]

                if nombre in nombres_existentes:
                    sg.popup_error("⚠ Ya existe un turno con ese nombre.", title="Turno duplicado")
                    continue

                turnos.append(nuevo)
            else:
                for i, t in enumerate(turnos):
                    if i != seleccion_index and str(t.get("nombre", "")).strip() == nombre:
                        sg.popup_error("⚠ Ya existe otro turno con ese nombre.", title="Turno duplicado")
                        break
                else:
                    turnos[seleccion_index] = nuevo

            refrescar_tabla()
            limpiar_formulario()

        if event == "-ELIMINAR_TURNO-":
            if seleccion_index is None:
                sg.popup_error("⚠ Selecciona un turno para eliminar.", title="Selección requerida")
                continue

            nombre_eliminar = str(turnos[seleccion_index].get("nombre", ""))

            respuesta = sg.popup_yes_no(
                f"¿Deseas eliminar el turno '{nombre_eliminar}'?",
                title="Confirmar eliminación"
            )

            if respuesta == "Yes":
                turnos.pop(seleccion_index)
                refrescar_tabla()
                limpiar_formulario()

        if event == "-GUARDAR_TODO_TURNOS-":
            try:
                # Si hay un turno seleccionado, asegurar que sus metas por rol queden sincronizadas
                if seleccion_index is not None:
                    turnos[seleccion_index]["metas_ppt_bodega_roles"] = (
                        _normalizar_metas_roles_desde_tabla(metas_roles_actuales)
                    )
                
                guardar_turnos(turnos)
                resultado = turnos
                sg.popup("Turnos y metas PPT guardados correctamente.", title="Turnos")
                break
            except Exception as e:
                sg.popup_error(
                    f"No se pudieron guardar los turnos:\n{str(e)}",
                    title="Error guardando turnos"
                )
            

    win.close()
    return resultado
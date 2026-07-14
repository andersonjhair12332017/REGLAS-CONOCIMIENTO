import os
import subprocess
import pandas as pd
import PySimpleGUI as sg

from personal_config import (
    obtener_ruta_maestro_personal,
    crear_maestro_personal_si_no_existe
)


COLUMNAS_VISOR = [
    "ID Operador",
    "Nombre",
    "Rol",
    "Lider",
    "Activo",
    "Observacion"
]

def _df_vacio_visor():
    return pd.DataFrame(columns=COLUMNAS_VISOR)


def _asegurar_maestro_existe():
    """
    Garantiza que maestro_personal.xlsx exista.
    Si no existe, lo crea con hojas Bodega y Despachos.
    """

    ruta = obtener_ruta_maestro_personal()

    if not os.path.exists(ruta):
        crear_maestro_personal_si_no_existe()

    return obtener_ruta_maestro_personal()


def _normalizar_df_visor(df):
    if df is None:
        df = pd.DataFrame(columns=COLUMNAS_VISOR)

    df = df.copy()

    df.columns = [
        str(c).strip()
        for c in df.columns
    ]

    for col in COLUMNAS_VISOR:
        if col not in df.columns:
            df[col] = ""

    df = df[COLUMNAS_VISOR].copy()

    for col in COLUMNAS_VISOR:
        df[col] = df[col].fillna("").astype(str).str.strip()

    return df


def cargar_maestro_para_visualizar():
    """
    Carga maestro_personal.xlsx para visualización.

    Siempre retorna una tupla:
    (df_bodega, df_despachos)

    Nunca debe retornar None.
    """

    try:
        ruta = _asegurar_maestro_existe()
    except Exception as e:
        sg.popup_error(
            f"No se pudo crear o ubicar el maestro de personal.\n\nDetalle:\n{e}",
            title="Error maestro personal"
        )

        return (
            _df_vacio_visor(),
            _df_vacio_visor()
        )

    if not os.path.exists(ruta):
        return (
            _df_vacio_visor(),
            _df_vacio_visor()
        )

    df_bodega = _df_vacio_visor()
    df_despachos = _df_vacio_visor()

    try:
        with pd.ExcelFile(ruta, engine="openpyxl") as xls:
            hojas = xls.sheet_names

            if "Bodega" in hojas:
                try:
                    df_bodega = pd.read_excel(
                        xls,
                        sheet_name="Bodega"
                    )
                except Exception:
                    df_bodega = _df_vacio_visor()

            if "Despachos" in hojas:
                try:
                    df_despachos = pd.read_excel(
                        xls,
                        sheet_name="Despachos"
                    )
                except Exception:
                    df_despachos = _df_vacio_visor()

    except Exception as e:
        sg.popup_error(
            f"No se pudo leer maestro_personal.xlsx.\n\nDetalle:\n{e}",
            title="Error leyendo maestro"
        )

        return (
            _df_vacio_visor(),
            _df_vacio_visor()
        )

    return (
        _normalizar_df_visor(df_bodega),
        _normalizar_df_visor(df_despachos)
    )


def _ordenar_df(df):
    df = _normalizar_df_visor(df)

    if df.empty:
        return df

    return df.sort_values(
        by=["Nombre", "ID Operador"],
        ascending=True
    ).reset_index(drop=True)


def _df_to_table(df):
    df = _normalizar_df_visor(df)

    if df.empty:
        return []

    return df[COLUMNAS_VISOR].values.tolist()


def _abrir_archivo_excel(ruta):
    if not os.path.exists(ruta):
        try:
            crear_maestro_personal_si_no_existe()
            ruta = obtener_ruta_maestro_personal()
        except Exception as e:
            sg.popup_error(
                f"No existe el archivo maestro y no se pudo crear.\n\nRuta:\n{ruta}\n\nDetalle:\n{e}",
                title="Archivo no encontrado"
            )
            return

    if not os.path.exists(ruta):
        sg.popup_error(
            f"No existe el archivo maestro:\n\n{ruta}",
            title="Archivo no encontrado"
        )
        return

    try:
        os.startfile(ruta)
    except AttributeError:
        try:
            subprocess.Popen(["xdg-open", ruta])
        except Exception as e:
            sg.popup_error(
                f"No se pudo abrir el archivo maestro.\n\nDetalle:\n{e}",
                title="Error al abrir Excel"
            )
    except Exception as e:
        sg.popup_error(
            f"No se pudo abrir el archivo maestro.\n\nDetalle:\n{e}",
            title="Error al abrir Excel"
        )


def abrir_editor_maestro_personal():
    """
    Visor del maestro de personal.

    Importante:
    Esta función conserva el mismo nombre para no modificar el evento del botón.
    Pero ya NO permite editar desde el software.
    Solo permite abrir el Excel y recargar.
    """

    ruta_maestro = _asegurar_maestro_existe()

    df_bodega, df_despachos = cargar_maestro_para_visualizar()

    df_bodega = _ordenar_df(df_bodega)
    df_despachos = _ordenar_df(df_despachos)

    layout = [
        [
            sg.Text(
                "Administrar Personal",
                font=("Segoe UI", 15, "bold"),
                justification="center",
                expand_x=True
            )
        ],
        [
            sg.Text(
                "El maestro de personal se edita directamente en Excel. "
                "Desde esta pantalla solo puedes abrir el archivo y recargar la información.",
                font=("Segoe UI", 9)
            )
        ],
        [
            sg.Text(
                f"Ruta maestro: {ruta_maestro}",
                key="-RUTA_MAESTRO-",
                font=("Segoe UI", 8),
                text_color="#666666"
            )
        ],
        [
            sg.Text(
                f"Registros Bodega: {len(df_bodega)}",
                key="-COUNT_BODEGA-",
                font=("Segoe UI", 9, "bold"),
                text_color="#004578"
            ),
            sg.Text("   "),
            sg.Text(
                f"Registros Despachos: {len(df_despachos)}",
                key="-COUNT_DESPACHOS-",
                font=("Segoe UI", 9, "bold"),
                text_color="#004578"
            )
        ],
        [
            sg.TabGroup(
                [
                    [
                        sg.Tab(
                            "Bodega",
                            [
                                [
                                    sg.Table(
                                        values=_df_to_table(df_bodega),
                                        headings=COLUMNAS_VISOR,
                                        key="-TABLA_BODEGA-",
                                        auto_size_columns=False,
                                        col_widths=[15, 35, 18, 25, 45],
                                        justification="left",
                                        num_rows=22,
                                        alternating_row_color="#F2F2F2",
                                        selected_row_colors=("white", "#0078D4"),
                                        expand_x=True,
                                        expand_y=True
                                    )
                                ]
                            ]
                        ),
                        sg.Tab(
                            "Despachos",
                            [
                                [
                                    sg.Table(
                                        values=_df_to_table(df_despachos),
                                        headings=COLUMNAS_VISOR,
                                        key="-TABLA_DESPACHOS-",
                                        auto_size_columns=False,
                                        col_widths=[15, 35, 18, 25, 45],
                                        justification="left",
                                        num_rows=22,
                                        alternating_row_color="#F2F2F2",
                                        selected_row_colors=("white", "#0078D4"),
                                        expand_x=True,
                                        expand_y=True
                                    )
                                ]
                            ]
                        )
                    ]
                ],
                expand_x=True,
                expand_y=True
            )
        ],
        [
            sg.Button(
                "📂 Abrir maestro en Excel",
                key="-ABRIR_EXCEL-",
                size=(24, 1),
                button_color=("white", "#0078D4")
            ),
            sg.Button(
                "🔄 Recargar maestro",
                key="-RECARGAR-",
                size=(20, 1),
                button_color=("white", "#5C2D91")
            ),
            sg.Push(),
            sg.Button(
                "✖ Cerrar",
                key="-CERRAR-",
                size=(14, 1),
                button_color=("white", "#666666")
            )
        ]
    ]

    window = sg.Window(
        "Administrar Personal",
        layout,
        modal=True,
        finalize=True,
        resizable=True,
        size=(1120, 680)
    )

    while True:
        event, values = window.read()

        if event in (sg.WIN_CLOSED, "-CERRAR-"):
            break

        if event == "-ABRIR_EXCEL-":            
            _abrir_archivo_excel(ruta_maestro)
            

        if event == "-RECARGAR-":
            df_bodega, df_despachos = cargar_maestro_para_visualizar()

            df_bodega = _ordenar_df(df_bodega)
            df_despachos = _ordenar_df(df_despachos)

            window["-TABLA_BODEGA-"].update(
                values=_df_to_table(df_bodega)
            )

            window["-TABLA_DESPACHOS-"].update(
                values=_df_to_table(df_despachos)
            )

            window["-COUNT_BODEGA-"].update(
                f"Registros Bodega: {len(df_bodega)}"
            )

            window["-COUNT_DESPACHOS-"].update(
                f"Registros Despachos: {len(df_despachos)}"
            )

            sg.popup(
                "Maestro recargado correctamente.",
                title="Recarga exitosa",
                font=("Segoe UI", 10)
            )

    window.close()
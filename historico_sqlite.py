import sqlite3
import pandas as pd
from datetime import datetime, timedelta
import os


# ==========================================================
# CONFIGURACIÓN BASE
# ==========================================================
DB_NAME = "historico_productividad.db"


def obtener_ruta_db(carpeta_destino):
    """
    Devuelve la ruta completa de la base SQLite.
    La base queda en la misma carpeta donde el usuario guarda reportes.
    """
    return os.path.join(carpeta_destino, DB_NAME)


def calcular_info_turno(fecha_str, ini, fin):
    """
    Calcula Turno, Fecha_Inicio_Turno y Fecha_Fin_Turno.
    Regla:
    - Turno Día: fecha_base 07:30 -> fecha_base 19:30
    - Turno Noche: fecha_base - 1 día 19:30 -> fecha_base 07:30
    """
    fecha_base = datetime.strptime(fecha_str.strip(), "%d/%m/%Y").date()

    h_ini = datetime.strptime(ini, "%H:%M").time()
    h_fin = datetime.strptime(fin, "%H:%M").time()

    cruza = h_ini > h_fin

    if not cruza:
        turno = "DIA"
        fecha_inicio = datetime.combine(fecha_base, h_ini)
        fecha_fin = datetime.combine(fecha_base, h_fin)
    else:
        turno = "NOCHE"
        fecha_inicio = datetime.combine(fecha_base - timedelta(days=1), h_ini)
        fecha_fin = datetime.combine(fecha_base, h_fin)

    id_turno = f"{fecha_base.strftime('%Y-%m-%d')}_{turno}"

    return {
        "Fecha_Base": fecha_base.strftime("%Y-%m-%d"),
        "Turno": turno,
        "FechaHora_Inicio_Turno": fecha_inicio.strftime("%Y-%m-%d %H:%M:%S"),
        "FechaHora_Fin_Turno": fecha_fin.strftime("%Y-%m-%d %H:%M:%S"),
        "ID_Turno": id_turno
    }


# ==========================================================
# CREACIÓN DE TABLA
# ==========================================================
def inicializar_db(carpeta_destino):
    """
    Crea la base de datos y la tabla histórica si no existen.
    """
    ruta_db = obtener_ruta_db(carpeta_destino)

    conn = sqlite3.connect(ruta_db)
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS historico_movimientos (
            ID INTEGER PRIMARY KEY AUTOINCREMENT,

            ID_Turno TEXT NOT NULL,
            Fecha_Base TEXT NOT NULL,
            Turno TEXT NOT NULL,
            FechaHora_Inicio_Turno TEXT,
            FechaHora_Fin_Turno TEXT,

            Area_Movimiento TEXT,

            ID_Operador TEXT,
            Nombre TEXT,

            Movimiento_ID INTEGER,
            Pertenece_a_Pick_Padre INTEGER,
            Tipo_Match TEXT,

            Estado_Cantidad_Hijo TEXT,
            Estado_Cantidad_Padre TEXT,

            Origen TEXT,
            Intermedia TEXT,
            Destino_Final TEXT,

            FH_inicio TEXT,
            FH_fin TEXT,

            Cantidad_Pick REAL,
            Cantidad_Put REAL,
            Diferencia_Hijo REAL,

            Cantidad_Pick_Padre_Total REAL,
            Cantidad_Put_Hijos_Total REAL,
            Diferencia_Padre REAL,

            Items_Distintos INTEGER,
            Lotes_Distintos INTEGER,
            LPs_Distintos INTEGER,
            Picks_Fragmentos INTEGER,
            Puts_Fragmentos INTEGER,

            Fecha_Registro TEXT
        )
    """)

    conn.commit()
    conn.close()

    return ruta_db


# ==========================================================
# NORMALIZAR COLUMNAS PARA HISTÓRICO
# ==========================================================
def preparar_movimientos_para_historico(movimientos_export, fecha_str, ini, fin):
    """
    Recibe la hoja Movimientos Consolidados como DataFrame y la adapta
    para guardarla en SQLite.
    """
    info_turno = calcular_info_turno(fecha_str, ini, fin)

    df = movimientos_export.copy()

    # Asegurar columnas esperadas
    columnas_esperadas = [
        "Movimiento_ID",
        "Pertenece_a_Pick_Padre",
        "Tipo_Match",
        "Estado_Cantidad_Hijo",
        "Estado_Cantidad_Padre",
        "ID Operador",
        "Nombre",
        "Origen",
        "Intermedia",
        "Destino_Final",
        "FH_inicio",
        "FH_fin",
        "Cantidad_Pick",
        "Cantidad_Put",
        "Diferencia_Hijo",
        "Cantidad_Pick_Padre_Total",
        "Cantidad_Put_Hijos_Total",
        "Diferencia_Padre",
        "Items_Distintos",
        "Lotes_Distintos",
        "LPs_Distintos",
        "Picks_Fragmentos",
        "Puts_Fragmentos"
    ]

    for col in columnas_esperadas:
        if col not in df.columns:
            df[col] = pd.NA

    # Determinar área según Tipo_Match/ruta no es suficiente.
    # La forma más estable es usar el origen de los movimientos creados desde Pick_norm,
    # pero como movimientos_export no siempre conserva Pick_norm, dejamos una clasificación
    # base por Destino_Final/Origen y luego la podemos mejorar.
    df["Area_Movimiento"] = "Bodega"

    # Si el destino o intermedia tiene señales típicas de puerta/loading, marcar Despachos.
    destino = df["Destino_Final"].astype(str).str.lower()
    intermedia = df["Intermedia"].astype(str).str.lower()

    mask_despachos = (
        destino.str.contains("door", na=False) |
        destino.str.contains("zfdoor", na=False) |
        intermedia.str.contains("door", na=False) |
        intermedia.str.contains("zfdoor", na=False)
    )

    df.loc[mask_despachos, "Area_Movimiento"] = "Despachos"

    # Metadatos del turno
    df["ID_Turno"] = info_turno["ID_Turno"]
    df["Fecha_Base"] = info_turno["Fecha_Base"]
    df["Turno"] = info_turno["Turno"]
    df["FechaHora_Inicio_Turno"] = info_turno["FechaHora_Inicio_Turno"]
    df["FechaHora_Fin_Turno"] = info_turno["FechaHora_Fin_Turno"]
    df["Fecha_Registro"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Renombrar columnas para SQLite sin espacios
    df = df.rename(columns={
        "ID Operador": "ID_Operador"
    })

    columnas_sql = [
        "ID_Turno",
        "Fecha_Base",
        "Turno",
        "FechaHora_Inicio_Turno",
        "FechaHora_Fin_Turno",
        "Area_Movimiento",
        "ID_Operador",
        "Nombre",
        "Movimiento_ID",
        "Pertenece_a_Pick_Padre",
        "Tipo_Match",
        "Estado_Cantidad_Hijo",
        "Estado_Cantidad_Padre",
        "Origen",
        "Intermedia",
        "Destino_Final",
        "FH_inicio",
        "FH_fin",
        "Cantidad_Pick",
        "Cantidad_Put",
        "Diferencia_Hijo",
        "Cantidad_Pick_Padre_Total",
        "Cantidad_Put_Hijos_Total",
        "Diferencia_Padre",
        "Items_Distintos",
        "Lotes_Distintos",
        "LPs_Distintos",
        "Picks_Fragmentos",
        "Puts_Fragmentos",
        "Fecha_Registro"
    ]

    df = df[columnas_sql].copy()

    return df, info_turno


# ==========================================================
# GUARDAR HISTÓRICO TURNO
# ==========================================================
def guardar_historico_turno(movimientos_export, fecha_str, ini, fin, carpeta_destino, reemplazar=True):
    """
    Guarda los movimientos consolidados del turno en SQLite.

    Si reemplazar=True:
    - Borra primero el histórico existente de ese mismo ID_Turno.
    - Luego inserta la nueva versión.
    """
    ruta_db = inicializar_db(carpeta_destino)

    df_hist, info_turno = preparar_movimientos_para_historico(
        movimientos_export,
        fecha_str,
        ini,
        fin
    )

    conn = sqlite3.connect(ruta_db)

    if reemplazar:
        conn.execute(
            "DELETE FROM historico_movimientos WHERE ID_Turno = ?",
            (info_turno["ID_Turno"],)
        )
        conn.commit()

    df_hist.to_sql(
        "historico_movimientos",
        conn,
        if_exists="append",
        index=False
    )

    conn.commit()
    conn.close()

    return {
        "ruta_db": ruta_db,
        "id_turno": info_turno["ID_Turno"],
        "filas_guardadas": len(df_hist)
    }


# ==========================================================
# CONSULTAS MENSUALES
# ==========================================================
def leer_historico_mes(carpeta_destino, anio, mes):
    """
    Lee todos los movimientos históricos de un mes.
    """
    ruta_db = obtener_ruta_db(carpeta_destino)

    if not os.path.exists(ruta_db):
        raise FileNotFoundError("No existe base histórica SQLite.")

    fecha_ini = f"{anio}-{str(mes).zfill(2)}-01"

    if mes == 12:
        fecha_fin = f"{anio + 1}-01-01"
    else:
        fecha_fin = f"{anio}-{str(mes + 1).zfill(2)}-01"

    query = """
        SELECT *
        FROM historico_movimientos
        WHERE Fecha_Base >= ?
          AND Fecha_Base < ?
    """

    conn = sqlite3.connect(ruta_db)
    df = pd.read_sql_query(query, conn, params=(fecha_ini, fecha_fin))
    conn.close()

    return df


def generar_resumen_mensual(carpeta_destino, anio, mes):
    """
    Devuelve varios DataFrames para construir reporte mensual.
    """
    df = leer_historico_mes(carpeta_destino, anio, mes)

    if df.empty:
        raise ValueError("⚠ No hay histórico para el mes seleccionado.")

    # Movimientos únicos por operador + turno + movimiento
    df_unicos = df.drop_duplicates(
        subset=["ID_Turno", "ID_Operador", "Movimiento_ID"]
    ).copy()

    resumen = pd.DataFrame([{
        "Año": anio,
        "Mes": mes,
        "Turnos_Procesados": df["ID_Turno"].nunique(),
        "Total_Movimientos_Reales": len(df_unicos),
        "Movimientos_Bodega": len(df_unicos[df_unicos["Area_Movimiento"] == "Bodega"]),
        "Movimientos_Despachos": len(df_unicos[df_unicos["Area_Movimiento"] == "Despachos"])
    }])

    ranking_bodega = (
        df_unicos[df_unicos["Area_Movimiento"] == "Bodega"]
        .groupby(["ID_Operador", "Nombre"])["Movimiento_ID"]
        .count()
        .reset_index(name="Movimientos_Reales")
        .sort_values("Movimientos_Reales", ascending=False)
    )

    ranking_despachos = (
        df_unicos[df_unicos["Area_Movimiento"] == "Despachos"]
        .groupby(["ID_Operador", "Nombre"])["Movimiento_ID"]
        .count()
        .reset_index(name="Movimientos_Reales")
        .sort_values("Movimientos_Reales", ascending=False)
    )

    productividad_dia_turno = (
        df_unicos
        .groupby(["Fecha_Base", "Turno", "Area_Movimiento"])["Movimiento_ID"]
        .count()
        .reset_index(name="Movimientos_Reales")
    )

    detalle_turnos = (
        df_unicos
        .groupby(["Fecha_Base", "Turno", "ID_Operador", "Nombre", "Area_Movimiento"])["Movimiento_ID"]
        .count()
        .reset_index(name="Movimientos_Reales")
        .sort_values(["Fecha_Base", "Turno", "Area_Movimiento", "Movimientos_Reales"], ascending=[True, True, True, False])
    )

    return {
        "historico_mes": df,
        "resumen": resumen,
        "ranking_bodega": ranking_bodega,
        "ranking_despachos": ranking_despachos,
        "productividad_dia_turno": productividad_dia_turno,
        "detalle_turnos": detalle_turnos
    }


def exportar_reporte_mensual(carpeta_destino, anio, mes):
    """
    Genera Excel mensual desde SQLite.
    """
    data = generar_resumen_mensual(carpeta_destino, anio, mes)

    salida = os.path.join(
        carpeta_destino,
        f"Productividad_Mensual_{str(mes).zfill(2)}-{anio}.xlsx"
    )

    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        data["resumen"].to_excel(writer, sheet_name="Resumen Mensual", index=False)
        data["ranking_bodega"].to_excel(writer, sheet_name="Ranking Mensual Bodega", index=False)
        data["ranking_despachos"].to_excel(writer, sheet_name="Ranking Mensual Despachos", index=False)
        data["productividad_dia_turno"].to_excel(writer, sheet_name="Productividad Dia Turno", index=False)
        data["detalle_turnos"].to_excel(writer, sheet_name="Detalle Turnos", index=False)
        data["historico_mes"].to_excel(writer, sheet_name="Movimientos Mensuales", index=False)

    return salida

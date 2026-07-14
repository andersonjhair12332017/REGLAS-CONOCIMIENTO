import os
import pandas as pd
import PySimpleGUI as sg
from datetime import datetime, timedelta

from historico_sqlite import guardar_historico_turno
from wms_validators import validar_fecha_y_turno_excel_rapido

from personal_config import (
    cargar_maestro_personal,
    validar_maestro_personal,
    construir_asignacion_operadores_para_turno_df,
    normalizar_nombre,
    normalizar_id_operador    
)

from wms_utils import (
    _ensure_columns,
    _normalize_column_names,
    _parse_fecha_excel,
    _parse_hora_excel,
    _build_fechahora,
    _norm_desc_series,
    _area_auto_por_pick_norm,
    _ranking_desde_movimientos,
    _construir_eventos_despacho_por_fila,
    _construir_ciclos_despacho,
    _calcular_inactividad_despacho,
    _calcular_inactividad_bodega,
    _ranking_despachos_desde_ciclos,
    _generar_analisis_horas_cargue_despachos,
    _generar_analisis_horas_tarea_bodega,
    _aplicar_estilo_excel_resultados
)


def procesar_archivo(
    ruta,
    carpeta,
    ini,
    fin,
    fecha_str,
    window,
    turno_config=None,
    exportar_excel=True,
    retornar_dataframes=False,
    df_pre_filtrado=None,
    validar_fecha_excel=True,
    ignorar_faltantes_maestro=False
):


    def update(p, txt):
        if not isinstance(txt, str):
            txt = str(txt)

        if len(txt) > 140:
            txt = txt[:140] + " ..."

        window["-PROG-"].update(int(p))
        window["-PERC-"].update(f"{int(p)}%")
        window["-STATUS-"].update(txt)
        window.refresh()

    def _crear_movimientos_vacios():
        columnas = [
            "Movimiento_ID",
            "Pertenece_a_Pick_Padre",
            "Tipo_Match",
            "Area_Movimiento",
            "ID Operador",
            "Nombre",
            "Nombre_norm",
            "Origen",
            "Intermedia",
            "Destino_Final",
            "FH_inicio",
            "FH_fin",
            "Cantidad_Pick",
            "Cantidad_Put",
            "Cantidad",
            "Diferencia_Hijo",
            "Estado_Cantidad_Hijo",
            "Cantidad_Pick_Padre_Total",
            "Cantidad_Put_Hijos_Total",
            "Diferencia_Padre",
            "Estado_Cantidad_Padre",
            "Items_Distintos",
            "Lotes_Distintos",
            "LPs_Distintos",
            "Picks_Fragmentos",
            "Puts_Fragmentos",
            "LP",
            "LP_Pick",
            "LP_Put",
            "LP_Diferente",
            "Tiene_LP_Cambiada",
            "LPs_Pick_Distintas",
            "LPs_Put_Distintas",
            "Minutos_Max_Pick_Put",
            "Movimiento_Global_ID",
            "Es_Movimiento_Padre",
            "Es_Hijo_Multidestino",
            "Ruta_Visual",
            "Tiempo_Duracion",
            "Estado_Movimiento",
            "Observacion_Logica",
            "Resumen_Movimiento"
        ]

        return pd.DataFrame(columns=columnas)

    def _limpiar_texto(valor):
        try:
            if pd.isna(valor):
                return ""
        except Exception:
            pass

        return str(valor).strip()

    def _to_float(valor):
        try:
            if pd.isna(valor):
                return 0.0

            return float(valor)
        except Exception:
            return 0.0

    def _sumar_cantidad_por_item_lote(rows):
        resumen = {}

        for row in rows:
            item = _limpiar_texto(row.get("Item Number", ""))
            lote = _limpiar_texto(row.get("Número Lote", ""))
            key = (item, lote)

            resumen[key] = resumen.get(key, 0.0) + _to_float(row.get("Cantidad", 0))

        return {
            k: round(v, 6)
            for k, v in resumen.items()
            if round(v, 6) != 0
        }

    def _crear_movimiento_desde_abierto(mov_abierto, area_movimiento):
        picks = mov_abierto.get("picks", [])
        puts = mov_abierto.get("puts", [])

        if not picks or not puts:
            return None

        fh_inicio = min(pd.to_datetime(r.get("FechaHora"), errors="coerce") for r in picks)
        fh_fin = max(pd.to_datetime(r.get("FechaHora"), errors="coerce") for r in puts)

        origen = _limpiar_texto(mov_abierto.get("origen", ""))
        intermedia = _limpiar_texto(mov_abierto.get("intermedia", ""))

        destinos = []

        for r in puts:
            destino = _limpiar_texto(r.get("A ID Ubicación", ""))

            if destino and destino not in destinos:
                destinos.append(destino)

        if len(destinos) == 1:
            destino_final = destinos[0]
        else:
            destino_final = " | ".join(destinos)

        cantidad_pick = round(sum(_to_float(r.get("Cantidad", 0)) for r in picks), 6)
        cantidad_put = round(sum(_to_float(r.get("Cantidad", 0)) for r in puts), 6)

        detalle_pick = _sumar_cantidad_por_item_lote(picks)
        detalle_put = _sumar_cantidad_por_item_lote(puts)

        if detalle_pick != detalle_put:
            return None

        items = set()
        lotes = set()
        lp_pick_set = set()
        lp_put_set = set()
        lp_general = set()

        for r in picks:
            item = _limpiar_texto(r.get("Item Number", ""))
            lote = _limpiar_texto(r.get("Número Lote", ""))
            lp = _limpiar_texto(r.get("LP", ""))

            if item:
                items.add(item)

            if lote:
                lotes.add(lote)

            if lp:
                lp_pick_set.add(lp)
                lp_general.add(lp)

        for r in puts:
            item = _limpiar_texto(r.get("Item Number", ""))
            lote = _limpiar_texto(r.get("Número Lote", ""))
            lp = _limpiar_texto(r.get("LP", ""))

            if item:
                items.add(item)

            if lote:
                lotes.add(lote)

            if lp:
                lp_put_set.add(lp)
                lp_general.add(lp)

        tiene_lp_cambiada = lp_pick_set != lp_put_set

        picks_fragmentos = len(picks)
        puts_fragmentos = len(puts)

        items_distintos = len(items)
        lotes_distintos = len(lotes)
        lps_distintos = len(lp_general)
        lps_pick_distintas = len(lp_pick_set)
        lps_put_distintas = len(lp_put_set)

        diferencia = round(cantidad_pick - cantidad_put, 6)
        estado_cantidad = "OK" if diferencia == 0 else "REVISION"

        if items_distintos > 1:
            tipo_match = "MULTI_ITEM"
        elif tiene_lp_cambiada:
            tipo_match = "LP_CAMBIADA"
        elif picks_fragmentos > 1 and puts_fragmentos == 1:
            tipo_match = "MERGE"
        elif puts_fragmentos > 1 and picks_fragmentos == 1:
            tipo_match = "SPLIT"
        elif lps_distintos > 1:
            tipo_match = "MULTI_PALLET"
        else:
            tipo_match = "NORMAL"

        nombre_norm = _limpiar_texto(mov_abierto.get("Nombre_norm", ""))
        id_operador = _limpiar_texto(mov_abierto.get("ID Operador", ""))
        nombre = _limpiar_texto(mov_abierto.get("Nombre", ""))

        row_ids = []

        for r in picks + puts:
            row_id = r.get("_row_id_wms", None)

            try:
                if pd.notna(row_id):
                    row_ids.append(int(row_id))
            except Exception:
                pass

        return {
            "Movimiento_ID": pd.NA,
            "Pertenece_a_Pick_Padre": pd.NA,
            "Tipo_Match": tipo_match,
            "Area_Movimiento": area_movimiento,
            "ID Operador": id_operador,
            "Nombre": nombre,
            "Nombre_norm": nombre_norm,
            "Origen": origen,
            "Intermedia": intermedia,
            "Destino_Final": destino_final,
            "FH_inicio": fh_inicio,
            "FH_fin": fh_fin,
            "Cantidad_Pick": cantidad_pick,
            "Cantidad_Put": cantidad_put,
            "Cantidad": cantidad_put,
            "Diferencia_Hijo": diferencia,
            "Estado_Cantidad_Hijo": estado_cantidad,
            "Cantidad_Pick_Padre_Total": cantidad_pick,
            "Cantidad_Put_Hijos_Total": cantidad_put,
            "Diferencia_Padre": diferencia,
            "Estado_Cantidad_Padre": estado_cantidad,
            "Items_Distintos": items_distintos,
            "Lotes_Distintos": lotes_distintos,
            "LPs_Distintos": lps_distintos,
            "Picks_Fragmentos": picks_fragmentos,
            "Puts_Fragmentos": puts_fragmentos,
            "LP": " | ".join(sorted(lp_general)),
            "LP_Pick": " | ".join(sorted(lp_pick_set)),
            "LP_Put": " | ".join(sorted(lp_put_set)),
            "LP_Diferente": bool(tiene_lp_cambiada),
            "Tiene_LP_Cambiada": bool(tiene_lp_cambiada),
            "LPs_Pick_Distintas": lps_pick_distintas,
            "LPs_Put_Distintas": lps_put_distintas,
            "Minutos_Max_Pick_Put": round((fh_fin - fh_inicio).total_seconds() / 60, 2) if pd.notna(fh_inicio) and pd.notna(fh_fin) else 0,
            "_row_ids_wms": row_ids
        }

    def _construir_movimientos_bodega_secuencial(
        df_base,
        asignacion_operadores_area,
        personas_no_maestro_norm,
        nombres_excluidos_maestro
    ):
        """
        Motor limpio de movimientos reales de Bodega.

        Reglas:
        1. Fase 1: LP igual.
        2. Fase 2: LP cambiada.
        3. El PUT es la demanda.
        4. El origen separa movimientos.
        5. El destino separa movimientos.
        6. Tipo válido del PUT se usa como tarea WMS fuerte.
        7. Tipo STORAGE no se usa como tarea fuerte.
        8. Destino igual a origen = NO_PRODUCTIVO_RETORNO_ORIGEN.
        """

        if df_base is None or df_base.empty:
            return _crear_movimientos_vacios(), df_base

        df = df_base.copy()

        pick_descs = {
            "move (pick)",
            "move trailer (pick)",
            "picking (pick)"
        }

        put_descs = {
            "move (put)",
            "move trailer (put)",
            "picking (put)"
        }

        allowed_descs_local = pick_descs.union(put_descs)

        df["Area_Operador_Asignada"] = (
            df["Nombre_norm"]
            .astype(str)
            .map(asignacion_operadores_area)
            .fillna("Excluir")
        )

        df["Es_Operador_Bodega_Maestro"] = df["Area_Operador_Asignada"].isin([
            "Bodega",
            "Mixto"
        ])

        df_calc = df[
            (~df["Nombre_norm"].isin(personas_no_maestro_norm)) &
            (~df["Nombre_norm"].isin(nombres_excluidos_maestro)) &
            (df["Es_Operador_Bodega_Maestro"] == True) &
            (df["Description_norm"].isin(allowed_descs_local))
        ].copy()

        if df_calc.empty:
            return _crear_movimientos_vacios(), df_base

        mask_a_a_local = (
            df_calc["Desde ID Ubicación"].astype(str).str.strip()
            ==
            df_calc["A ID Ubicación"].astype(str).str.strip()
        )

        df_calc = df_calc[~mask_a_a_local].copy()

        if df_calc.empty:
            return _crear_movimientos_vacios(), df_base

        df_calc["FechaHora"] = pd.to_datetime(df_calc["FechaHora"], errors="coerce")
        df_calc = df_calc[df_calc["FechaHora"].notna()].copy()

        if df_calc.empty:
            return _crear_movimientos_vacios(), df_base

        df_calc = df_calc.sort_values([
            "Nombre_norm",
            "ID Operador",
            "FechaHora",
            "_row_id_wms"
        ]).copy()
        
        # ============================================================
        # Familia WMS interna
        # Evita mezclar Move Trailer con Move normal.
        # ============================================================
        
        def _familia_wms_desde_desc(desc):
            desc = _limpiar_texto(desc).lower()
        
            if "move trailer" in desc:
                return "TRAILER"
        
            if "move" in desc or "picking" in desc:
                return "NORMAL"
        
            return "OTRO"
        
        
        df_calc["_Familia_WMS"] = df_calc["Description_norm"].apply(_familia_wms_desde_desc)

        # ============================================================
        # Bloque temporal de flujo WMS
        # Evita mezclar movimientos del mismo operador/item/lote
        # cuando están separados por mucho tiempo.
        # ============================================================
        
        MAX_MINUTOS_SALTO_FLUJO_BODEGA = 45
        
        def _asignar_bloques_tiempo_flujo(df_in):
            df_tmp = df_in.copy()
            df_tmp["_Bloque_Flujo_Tiempo"] = 0
        
            group_cols_tiempo = [
                "Nombre_norm",
                "ID Operador",
                "_Familia_WMS",
                "Item Number",
                "Número Lote"
            ]
        
            for _, idxs in df_tmp.groupby(group_cols_tiempo, dropna=False).groups.items():
                g = df_tmp.loc[list(idxs)].sort_values([
                    "FechaHora",
                    "_row_id_wms"
                ]).copy()
        
                dif_min = (
                    g["FechaHora"]
                    .diff()
                    .dt.total_seconds()
                    .div(60)
                    .fillna(0)
                )
        
                bloque = (dif_min > MAX_MINUTOS_SALTO_FLUJO_BODEGA).cumsum()
        
                df_tmp.loc[g.index, "_Bloque_Flujo_Tiempo"] = bloque.astype(int).values
        
            return df_tmp
        
        df_calc = _asignar_bloques_tiempo_flujo(df_calc)
        
        row_ids_retorno_origen_global = set()
        
        def _flujo_en_tiempo(picks_rows, puts_rows):
            """
            Valida que un flujo PICK/PUT sea cercano en tiempo.
            """

            fechas = []

            for r in picks_rows + puts_rows:
                fh = pd.to_datetime(r.get("FechaHora"), errors="coerce")

                if pd.notna(fh):
                    fechas.append(fh)

            if not fechas:
                return True

            fh_min = min(fechas)
            fh_max = max(fechas)

            duracion_min = (fh_max - fh_min).total_seconds() / 60

            # Ajusta este valor si quieres más o menos tolerancia
            return duracion_min <= 50

        def _row_id_valido(row):
            try:
                row_id = row.get("_row_id_wms", None)

                if pd.isna(row_id):
                    return None

                return int(row_id)
            except Exception:
                return None

        def _row_ids_de_rows(rows):
            ids = []

            for r in rows:
                rid = _row_id_valido(r)

                if rid is not None:
                    ids.append(rid)

            return list(dict.fromkeys(ids))

        def _tipo_es_tarea_wms(valor):
            texto = _limpiar_texto(valor).lower()

            tipos_invalidos = {
                "",
                "nan",
                "none",
                "<na>",
                "storage"
            }

            return texto not in tipos_invalidos
        
        def _fecha_min_max_rows(rows):
            fechas = []

            for r in rows:
                fh = pd.to_datetime(r.get("FechaHora"), errors="coerce")

                if pd.notna(fh):
                    fechas.append(fh)

            if not fechas:
                return None, None

            return min(fechas), max(fechas)


        def _retorno_cronologicamente_valido(picks_rows, puts_rows):
            """
            Un retorno solo es válido si el PUT ocurre después del PICK.
            """

            if not picks_rows or not puts_rows:
                return False

            _, fh_pick_max = _fecha_min_max_rows(picks_rows)
            fh_put_min, _ = _fecha_min_max_rows(puts_rows)

            if fh_pick_max is None or fh_put_min is None:
                return False

            return fh_put_min >= fh_pick_max


        def _existe_put_productivo_posterior(
            picks_rows,
            puts_candidatos,
            origen_pick,
            intermedia_pick
        ):
            """
            Valida si el PICK tiene un PUT productivo compatible posterior.

            Si existe, el PICK NO debe marcarse como retorno.
            """

            if not picks_rows or puts_candidatos is None or puts_candidatos.empty:
                return False

            cantidad_pick = round(
                sum(_to_float(r.get("Cantidad", 0)) for r in picks_rows),
                6
            )

            if cantidad_pick <= 0:
                return False

            _, fh_pick_max = _fecha_min_max_rows(picks_rows)

            if fh_pick_max is None:
                return False

            item_set = {
                _limpiar_texto(r.get("Item Number", ""))
                for r in picks_rows
            }

            lote_set = {
                _limpiar_texto(r.get("Número Lote", ""))
                for r in picks_rows
            }

            lp_set = {
                _limpiar_texto(r.get("LP", ""))
                for r in picks_rows
            }

            candidatos = puts_candidatos.copy()
            candidatos["FechaHora"] = pd.to_datetime(
                candidatos["FechaHora"],
                errors="coerce"
            )

            candidatos = candidatos[
                candidatos["FechaHora"].notna()
            ].copy()

            if candidatos.empty:
                return False

            candidatos = candidatos[
                (candidatos["FechaHora"] >= fh_pick_max) &
                (candidatos["Desde ID Ubicación"].astype(str).str.strip() == intermedia_pick) &
                (candidatos["A ID Ubicación"].astype(str).str.strip().str.upper() != origen_pick.upper()) &
                (candidatos["Item Number"].astype(str).str.strip().isin(item_set)) &
                (candidatos["Número Lote"].astype(str).str.strip().isin(lote_set)) &
                (candidatos["LP"].astype(str).str.strip().isin(lp_set))
            ].copy()

            if candidatos.empty:
                return False

            cantidad_put = round(
                candidatos["Cantidad"].apply(_to_float).sum(),
                6
            )

            return cantidad_put >= cantidad_pick
        
        
        

        def _resumen_item_lote(rows, incluir_lp=False):
            resumen = {}

            for r in rows:
                item = _limpiar_texto(r.get("Item Number", ""))
                lote = _limpiar_texto(r.get("Número Lote", ""))

                if incluir_lp:
                    lp = _limpiar_texto(r.get("LP", ""))
                    key = (item, lote, lp)
                else:
                    key = (item, lote)

                resumen[key] = resumen.get(key, 0.0) + _to_float(r.get("Cantidad", 0))

            return {
                k: round(v, 6)
                for k, v in resumen.items()
                if round(v, 6) != 0
            }
        
        def _hay_lp_diferente_o_mixta(picks_rows, puts_rows):
            """
            Valida si existe cambio de LP total o parcial entre PICK y PUT.

            Casos válidos:
            - LP Pick != LP Put
            - Algunos PUT conservan LP y otros cambian LP
            """

            lp_pick_set = {
                _limpiar_texto(r.get("LP", ""))
                for r in picks_rows
                if _limpiar_texto(r.get("LP", ""))
            }

            lp_put_set = {
                _limpiar_texto(r.get("LP", ""))
                for r in puts_rows
                if _limpiar_texto(r.get("LP", ""))
            }

            if not lp_pick_set or not lp_put_set:
                return False

            return lp_pick_set != lp_put_set
        
        


        

        def _seleccionar_picks_exactos(candidatos, cantidad_objetivo, fecha_put):
            if candidatos is None or candidatos.empty:
                return []

            try:
                cantidad_objetivo = round(float(cantidad_objetivo), 6)
            except Exception:
                return []

            if cantidad_objetivo <= 0:
                return []

            c = candidatos.copy()
            c["FechaHora"] = pd.to_datetime(c["FechaHora"], errors="coerce")
            fecha_put = pd.to_datetime(fecha_put, errors="coerce")

            if pd.isna(fecha_put):
                c["_dist_tiempo"] = 0
            else:
                c["_dist_tiempo"] = (
                    c["FechaHora"] - fecha_put
                ).dt.total_seconds().abs().fillna(999999999)

            c = c.sort_values([
                "_dist_tiempo",
                "FechaHora",
                "_row_id_wms"
            ]).copy()

            seleccionados = []
            acumulado = 0.0

            for _, r in c.iterrows():
                qty = round(_to_float(r.get("Cantidad", 0)), 6)

                if qty <= 0:
                    continue

                if round(acumulado + qty, 6) > cantidad_objetivo:
                    continue

                seleccionados.append(r.to_dict())
                acumulado = round(acumulado + qty, 6)

                if acumulado == cantidad_objetivo:
                    break

            if round(acumulado, 6) == cantidad_objetivo:
                return seleccionados

            return []

        def _crear_picks_parciales_desde_saldo(picks_rows, cantidad_objetivo):
            """
            Crea picks parciales internos para balancear movimientos hijos.
            No modifica filas originales del Excel.
            """

            try:
                cantidad_objetivo = round(float(cantidad_objetivo), 6)
            except Exception:
                return []

            if cantidad_objetivo <= 0:
                return []

            resultado = []
            saldo_pendiente = cantidad_objetivo

            def _orden_pick(row):
                fecha = pd.to_datetime(row.get("FechaHora"), errors="coerce")

                if pd.isna(fecha):
                    fecha = pd.Timestamp.max

                try:
                    row_id = int(row.get("_row_id_wms", 0))
                except Exception:
                    row_id = 0

                return fecha, row_id

            picks_ordenados = sorted(
                picks_rows,
                key=_orden_pick
            )

            for pick in picks_ordenados:
                if saldo_pendiente <= 0:
                    break

                cantidad_pick = round(_to_float(pick.get("Cantidad", 0)), 6)

                if cantidad_pick <= 0:
                    continue

                cantidad_usar = min(cantidad_pick, saldo_pendiente)
                cantidad_usar = round(cantidad_usar, 6)

                pick_parcial = dict(pick)
                pick_parcial["Cantidad"] = cantidad_usar

                resultado.append(pick_parcial)

                saldo_pendiente = round(saldo_pendiente - cantidad_usar, 6)

            if saldo_pendiente == 0:
                return resultado

            return []

        def _buscar_combinacion_exacta_grupos(grupos, cantidad_objetivo):
            escala = 1000

            try:
                objetivo = int(round(float(cantidad_objetivo) * escala))
            except Exception:
                return []

            if objetivo <= 0:
                return []

            cantidades = []

            for g in grupos:
                try:
                    cantidades.append(int(round(float(g.get("cantidad", 0)) * escala)))
                except Exception:
                    cantidades.append(0)

            dp = {0: []}

            for idx, cantidad in enumerate(cantidades):
                if cantidad <= 0:
                    continue

                nuevos = dict(dp)

                for suma_actual, indices_actuales in dp.items():
                    nueva_suma = suma_actual + cantidad

                    if nueva_suma > objetivo:
                        continue

                    if nueva_suma not in nuevos:
                        nuevos[nueva_suma] = indices_actuales + [idx]

                    if nueva_suma == objetivo:
                        return nuevos[nueva_suma]

                dp = nuevos

            return dp.get(objetivo, [])

        def _area_bodega_desde_row(row):
            area_asignada = _limpiar_texto(
                row.get("Area_Operador_Asignada", "Excluir")
            )

            if area_asignada in ["Bodega", "Mixto"]:
                return "Bodega"

            return "Bodega"

        def _crear_movimiento_desde_match(
            picks_rows,
            puts_rows,
            area_movimiento,
            forzar_lp_cambiada=False
        ):
            if not picks_rows or not puts_rows:
                return None

            origenes = []

            for r in picks_rows:
                origen = _limpiar_texto(r.get("Desde ID Ubicación", ""))

                if origen and origen not in origenes:
                    origenes.append(origen)

            if len(origenes) != 1:
                return None

            origen = origenes[0]

            intermedias_pick = []

            for r in picks_rows:
                intermedia = _limpiar_texto(r.get("A ID Ubicación", ""))

                if intermedia and intermedia not in intermedias_pick:
                    intermedias_pick.append(intermedia)

            if len(intermedias_pick) != 1:
                return None

            intermedia = intermedias_pick[0]

            destinos_put = []

            for r in puts_rows:
                destino = _limpiar_texto(r.get("A ID Ubicación", ""))

                if destino and destino not in destinos_put:
                    destinos_put.append(destino)

            if len(destinos_put) != 1:
                return None

            destino_final = destinos_put[0]

            # ============================================================
            # VALIDACIÓN CORRECTA DE RETORNO
            # ============================================================
            
            if _limpiar_texto(destino_final).upper() == _limpiar_texto(origen).upper():

                # ====================================================
                # RETORNO VÁLIDO SOLO SI:
                # 1. Hay PUT real.
                # 2. El flujo está en tiempo.
                # 3. El PUT ocurre después del PICK.
                # ====================================================

                put_ids = _row_ids_de_rows(puts_rows)

                if not put_ids:
                    return None

                if not _flujo_en_tiempo(picks_rows, puts_rows):
                    return None

                if not _retorno_cronologicamente_valido(picks_rows, puts_rows):
                    return None

                for rid in _row_ids_de_rows(picks_rows + puts_rows):
                    try:
                        row_ids_retorno_origen_global.add(int(rid))
                    except Exception:
                        pass

                return None

            resumen_pick = _resumen_item_lote(picks_rows, incluir_lp=False)
            resumen_put = _resumen_item_lote(puts_rows, incluir_lp=False)

            if resumen_pick != resumen_put:
                return None

            mov_abierto = {
                "Nombre_norm": picks_rows[0].get("Nombre_norm", ""),
                "ID Operador": picks_rows[0].get("ID Operador", ""),
                "Nombre": picks_rows[0].get("Nombre", ""),
                "origen": origen,
                "intermedia": intermedia,
                "picks": picks_rows,
                "puts": puts_rows,
                "area_movimiento": area_movimiento
            }

            mov = _crear_movimiento_desde_abierto(
                mov_abierto,
                area_movimiento
            )

            if mov is None:
                return None

            if forzar_lp_cambiada:
                mov["Tipo_Match"] = "LP_CAMBIADA"
                mov["Tiene_LP_Cambiada"] = True
                mov["LP_Diferente"] = True

            return mov


        def _procesar_por_put_agrupado(
            df_fase,
            permitir_lp_cambiada,
            row_ids_usados_globales,
            usar_tipo_fuerte=False
        ):
            movimientos_fase = []
            row_ids_usados_fase = set()

            if df_fase is None or df_fase.empty:
                return movimientos_fase, row_ids_usados_fase

            df_work = df_fase[
                ~df_fase["_row_id_wms"].isin(row_ids_usados_globales)
            ].copy()

            if df_work.empty:
                return movimientos_fase, row_ids_usados_fase

            df_work["Tipo_Texto"] = df_work["Tipo"].apply(_limpiar_texto)

            for _, g_operador in df_work.groupby(
                ["Nombre_norm", "ID Operador", "Nombre"],
                dropna=False,
                sort=False
            ):
                g_operador = g_operador.sort_values([
                    "FechaHora",
                    "_row_id_wms"
                ]).copy()

                picks = g_operador[
                    g_operador["Description_norm"].isin(pick_descs)
                ].copy()

                puts = g_operador[
                    g_operador["Description_norm"].isin(put_descs)
                ].copy()

                if picks.empty or puts.empty:
                    continue
                
                # ============================================================
                # PRIORIDAD 0: GLOBAL MATCH UNIDESTINO
                # Corrige casos donde el motor parte un flujo balanceado
                # en movimientos pequeños.
                # ============================================================

                if permitir_lp_cambiada:
                    group_cols_global = [
                        "_Familia_WMS",
                        "_Bloque_Flujo_Tiempo",
                        "Item Number",
                        "Número Lote",
                        "Desde ID Ubicación",
                        "A ID Ubicación"
                    ]
                else:
                    group_cols_global = [
                        "_Familia_WMS",
                        "_Bloque_Flujo_Tiempo",
                        "Item Number",
                        "Número Lote",
                        "Desde ID Ubicación",
                        "A ID Ubicación",
                        "LP"
                    ]

                picks_global_disponibles = picks[
                    (~picks["_row_id_wms"].isin(row_ids_usados_globales)) &
                    (~picks["_row_id_wms"].isin(row_ids_usados_fase))
                ].copy()

                puts_global_disponibles = puts[
                    (~puts["_row_id_wms"].isin(row_ids_usados_globales)) &
                    (~puts["_row_id_wms"].isin(row_ids_usados_fase))
                ].copy()

                if not picks_global_disponibles.empty and not puts_global_disponibles.empty:

                    for _, picks_full in picks_global_disponibles.groupby(
                        group_cols_global,
                        dropna=False,
                        sort=False
                    ):
                        if picks_full.empty:
                            continue

                        familia_global = _limpiar_texto(
                            picks_full["_Familia_WMS"].iloc[0]
                        )

                        bloque_global = picks_full["_Bloque_Flujo_Tiempo"].iloc[0]

                        item_global = _limpiar_texto(
                            picks_full["Item Number"].iloc[0]
                        )

                        lote_global = _limpiar_texto(
                            picks_full["Número Lote"].iloc[0]
                        )

                        origen_global = _limpiar_texto(
                            picks_full["Desde ID Ubicación"].iloc[0]
                        )

                        intermedia_global = _limpiar_texto(
                            picks_full["A ID Ubicación"].iloc[0]
                        )

                        if not origen_global or not intermedia_global:
                            continue

                        candidatos_put_global = puts_global_disponibles[
                            (puts_global_disponibles["_Familia_WMS"].astype(str).str.strip() == familia_global) &
                            (puts_global_disponibles["_Bloque_Flujo_Tiempo"] == bloque_global) &
                            (puts_global_disponibles["Item Number"].astype(str).str.strip() == item_global) &
                            (puts_global_disponibles["Número Lote"].astype(str).str.strip() == lote_global) &
                            (puts_global_disponibles["Desde ID Ubicación"].astype(str).str.strip() == intermedia_global)
                        ].copy()

                        if candidatos_put_global.empty:
                            continue

                        if not permitir_lp_cambiada:
                            lp_global = _limpiar_texto(
                                picks_full["LP"].iloc[0]
                            )

                            candidatos_put_global = candidatos_put_global[
                                candidatos_put_global["LP"].astype(str).str.strip() == lp_global
                            ].copy()

                        if candidatos_put_global.empty:
                            continue

                        destinos_globales = (
                            candidatos_put_global["A ID Ubicación"]
                            .astype(str)
                            .str.strip()
                            .replace("", pd.NA)
                            .dropna()
                            .unique()
                        )

                        # Este filtro solo aplica cuando todos los PUT van al mismo destino.
                        # Los multidestino los maneja _procesar_padre_hijos_por_destino.
                        if len(destinos_globales) != 1:
                            continue

                        destino_global = _limpiar_texto(destinos_globales[0])

                        # Si vuelve al origen, no se maneja aquí.
                        if destino_global.upper() == origen_global.upper():
                            continue

                        picks_rows = [
                            r.to_dict()
                            for _, r in picks_full.iterrows()
                        ]

                        candidatos_put_global = candidatos_put_global.sort_values([
                            "FechaHora",
                            "_row_id_wms"
                        ]).copy()

                        puts_rows = [
                            r.to_dict()
                            for _, r in candidatos_put_global.iterrows()
                        ]

                        if not picks_rows or not puts_rows:
                            continue

                        total_pick_global = round(
                            sum(_to_float(r.get("Cantidad", 0)) for r in picks_rows),
                            6
                        )

                        total_put_global = round(
                            sum(_to_float(r.get("Cantidad", 0)) for r in puts_rows),
                            6
                        )

                        if total_pick_global != total_put_global:
                            continue

                        resumen_pick_global = _resumen_item_lote(
                            picks_rows,
                            incluir_lp=False
                        )

                        resumen_put_global = _resumen_item_lote(
                            puts_rows,
                            incluir_lp=False
                        )

                        if resumen_pick_global != resumen_put_global:
                            continue

                        if not permitir_lp_cambiada:
                            resumen_pick_lp_global = _resumen_item_lote(
                                picks_rows,
                                incluir_lp=True
                            )

                            resumen_put_lp_global = _resumen_item_lote(
                                puts_rows,
                                incluir_lp=True
                            )

                            if resumen_pick_lp_global != resumen_put_lp_global:
                                continue

                        if not _flujo_en_tiempo(picks_rows, puts_rows):
                            continue

                        mov_global = _crear_movimiento_desde_match(
                            picks_rows=picks_rows,
                            puts_rows=puts_rows,
                            area_movimiento=_area_bodega_desde_row(picks_rows[0]),
                            forzar_lp_cambiada=permitir_lp_cambiada
                        )

                        if mov_global is None:
                            continue

                        row_ids_global = _row_ids_de_rows(
                            picks_rows + puts_rows
                        )

                        mov_global["_row_ids_wms"] = row_ids_global
                        mov_global["Tipo_Match"] = "GLOBAL_MATCH"

                        movimientos_fase.append(mov_global)

                        for rid in row_ids_global:
                            row_ids_usados_fase.add(rid)
                            
                
                
                
                # ============================================================
                # PRIORIDAD 0: GLOBAL MATCH POR FLUJO COMPLETO
                # Se ejecuta antes de hacer matches parciales por PUT.
                # ============================================================

                if permitir_lp_cambiada:
                    group_global_cols = [
                        "_Familia_WMS",
                        "_Bloque_Flujo_Tiempo",
                        "Item Number",
                        "Número Lote",
                        "Desde ID Ubicación",
                        "A ID Ubicación"
                    ]
                else:
                    group_global_cols = [
                        "_Familia_WMS",
                        "_Bloque_Flujo_Tiempo",
                        "Item Number",
                        "Número Lote",
                        "Desde ID Ubicación",
                        "A ID Ubicación",
                        "LP"
                    ]

                for _, picks_full in picks.groupby(
                    group_global_cols,
                    dropna=False,
                    sort=False
                ):
                    if picks_full.empty:
                        continue

                    pick_global_ids = _row_ids_de_rows([
                        r.to_dict()
                        for _, r in picks_full.iterrows()
                    ])

                    if not pick_global_ids:
                        continue

                    if any(
                        rid in row_ids_usados_globales or rid in row_ids_usados_fase
                        for rid in pick_global_ids
                    ):
                        continue

                    familia_global = _limpiar_texto(
                        picks_full["_Familia_WMS"].iloc[0]
                    )

                    bloque_global = picks_full["_Bloque_Flujo_Tiempo"].iloc[0]

                    item_global = _limpiar_texto(
                        picks_full["Item Number"].iloc[0]
                    )

                    lote_global = _limpiar_texto(
                        picks_full["Número Lote"].iloc[0]
                    )

                    origen_global = _limpiar_texto(
                        picks_full["Desde ID Ubicación"].iloc[0]
                    )

                    intermedia_global = _limpiar_texto(
                        picks_full["A ID Ubicación"].iloc[0]
                    )

                    if not origen_global or not intermedia_global:
                        continue

                    candidatos_put_global = puts[
                        (~puts["_row_id_wms"].isin(row_ids_usados_globales)) &
                        (~puts["_row_id_wms"].isin(row_ids_usados_fase)) &
                        (puts["_Familia_WMS"].astype(str).str.strip() == familia_global) &
                        (puts["_Bloque_Flujo_Tiempo"] == bloque_global) &
                        (puts["Item Number"].astype(str).str.strip() == item_global) &
                        (puts["Número Lote"].astype(str).str.strip() == lote_global) &
                        (puts["Desde ID Ubicación"].astype(str).str.strip() == intermedia_global)
                    ].copy()

                    if candidatos_put_global.empty:
                        continue

                    if not permitir_lp_cambiada:
                        lp_global = _limpiar_texto(
                            picks_full["LP"].iloc[0]
                        )

                        candidatos_put_global = candidatos_put_global[
                            candidatos_put_global["LP"].astype(str).str.strip() == lp_global
                        ].copy()

                    if candidatos_put_global.empty:
                        continue

                    destinos_globales = (
                        candidatos_put_global["A ID Ubicación"]
                        .astype(str)
                        .str.strip()
                        .replace("", pd.NA)
                        .dropna()
                        .unique()
                    )

                    # Esta capa solo resuelve flujos de un solo destino.
                    # Los multidestino los maneja _procesar_padre_hijos_por_destino.
                    if len(destinos_globales) != 1:
                        continue

                    destino_global = _limpiar_texto(destinos_globales[0])

                    # No resolver retornos al origen aquí.
                    if destino_global.upper() == origen_global.upper():
                        continue

                    picks_rows = [
                        r.to_dict()
                        for _, r in picks_full.iterrows()
                        if r["_row_id_wms"] not in row_ids_usados_globales
                        and r["_row_id_wms"] not in row_ids_usados_fase
                    ]

                    candidatos_put_global = candidatos_put_global.sort_values([
                        "FechaHora",
                        "_row_id_wms"
                    ]).copy()

                    puts_rows = [
                        r.to_dict()
                        for _, r in candidatos_put_global.iterrows()
                    ]

                    if not picks_rows or not puts_rows:
                        continue

                    total_pick_global = round(
                        sum(_to_float(r.get("Cantidad", 0)) for r in picks_rows),
                        6
                    )

                    total_put_global = round(
                        sum(_to_float(r.get("Cantidad", 0)) for r in puts_rows),
                        6
                    )

                    if total_pick_global != total_put_global:
                        continue

                    if not _flujo_en_tiempo(picks_rows, puts_rows):
                        continue

                    mov_global = _crear_movimiento_desde_match(
                        picks_rows=picks_rows,
                        puts_rows=puts_rows,
                        area_movimiento=_area_bodega_desde_row(picks_rows[0]),
                        forzar_lp_cambiada=permitir_lp_cambiada
                    )

                    if mov_global is None:
                        continue

                    row_ids_global = _row_ids_de_rows(
                        picks_rows + puts_rows
                    )

                    mov_global["_row_ids_wms"] = row_ids_global
                    mov_global["Tipo_Match"] = "GLOBAL_MATCH"

                    movimientos_fase.append(mov_global)

                    for rid in row_ids_global:
                        row_ids_usados_fase.add(rid)
                
                                

                
                

                if usar_tipo_fuerte:
                    puts = puts[
                        puts["Tipo_Texto"].apply(_tipo_es_tarea_wms)
                    ].copy()

                    if puts.empty:
                        continue
                    


                

                picks_disponibles = picks[
                    (~picks["_row_id_wms"].isin(row_ids_usados_globales)) &
                    (~picks["_row_id_wms"].isin(row_ids_usados_fase))
                ].copy()

                puts_disponibles = puts[
                    (~puts["_row_id_wms"].isin(row_ids_usados_globales)) &
                    (~puts["_row_id_wms"].isin(row_ids_usados_fase))
                ].copy()
                            
            
                if picks_disponibles.empty or puts_disponibles.empty:
                    continue

                group_put_cols = [
                    "Nombre_norm",
                    "ID Operador",
                    "Nombre",
                    "_Familia_WMS",
                    "_Bloque_Flujo_Tiempo",
                    "Item Number",
                    "Número Lote",
                    "Desde ID Ubicación",
                    "A ID Ubicación"
                ]

                if usar_tipo_fuerte:
                    group_put_cols.append("Tipo_Texto")

                if not permitir_lp_cambiada:
                    group_put_cols.append("LP")

                puts_disponibles["_row_dict"] = puts_disponibles.apply(
                    lambda r: r.to_dict(),
                    axis=1
                )

                put_batches = (
                    puts_disponibles
                    .groupby(group_put_cols, dropna=False, sort=False)
                    .agg(
                        Cantidad_Put_Total=("Cantidad", "sum"),
                        FH_put=("FechaHora", "max"),
                        Puts_Rows=("_row_dict", list)
                    )
                    .reset_index()
                    .sort_values("FH_put")
                )

                for _, put_batch in put_batches.iterrows():
                    puts_rows = put_batch.get("Puts_Rows", [])

                    if not isinstance(puts_rows, list) or not puts_rows:
                        continue

                    put_ids = _row_ids_de_rows(puts_rows)

                    if not put_ids:
                        continue

                    if any(
                        rid in row_ids_usados_globales or rid in row_ids_usados_fase
                        for rid in put_ids
                    ):
                        continue

                    nombre_norm = _limpiar_texto(put_batch.get("Nombre_norm", ""))
                    id_operador = _limpiar_texto(put_batch.get("ID Operador", ""))
                    nombre = _limpiar_texto(put_batch.get("Nombre", ""))
                    item_put = _limpiar_texto(put_batch.get("Item Number", ""))
                    lote_put = _limpiar_texto(put_batch.get("Número Lote", ""))
                    intermedia_put = _limpiar_texto(put_batch.get("Desde ID Ubicación", ""))
                    
                    bloque_put = put_batch.get("_Bloque_Flujo_Tiempo", None)

                    cantidad_put_total = round(
                        _to_float(put_batch.get("Cantidad_Put_Total", 0)),
                        6
                    )

                    fecha_put = pd.to_datetime(
                        put_batch.get("FH_put"),
                        errors="coerce"
                    )

                    if cantidad_put_total <= 0:
                        continue

                    if not intermedia_put:
                        continue
                    
                    familia_put = _limpiar_texto(put_batch.get("_Familia_WMS", ""))

                    candidatos_base = picks_disponibles[
                        (~picks_disponibles["_row_id_wms"].isin(row_ids_usados_globales)) &
                        (~picks_disponibles["_row_id_wms"].isin(row_ids_usados_fase)) &
                        (picks_disponibles["Nombre_norm"].astype(str).str.strip() == nombre_norm) &
                        (picks_disponibles["ID Operador"].astype(str).str.strip() == id_operador) &
                        (picks_disponibles["Nombre"].astype(str).str.strip() == nombre) &
                        (picks_disponibles["_Familia_WMS"].astype(str).str.strip() == familia_put) &
                        (picks_disponibles["_Bloque_Flujo_Tiempo"] == bloque_put) &
                        (picks_disponibles["A ID Ubicación"].astype(str).str.strip() == intermedia_put) &
                        (picks_disponibles["Item Number"].astype(str).str.strip() == item_put) &
                        (picks_disponibles["Número Lote"].astype(str).str.strip() == lote_put)
                    ].copy()

                    if candidatos_base.empty:
                        continue

                    if permitir_lp_cambiada:
                        lp_puts_set = {
                            _limpiar_texto(r.get("LP", ""))
                            for r in puts_rows
                            if _limpiar_texto(r.get("LP", ""))
                        }

                        candidatos_base = candidatos_base[
                            ~candidatos_base["LP"].astype(str).str.strip().isin(lp_puts_set)
                        ].copy()
                    else:
                        lp_put = _limpiar_texto(put_batch.get("LP", ""))

                        candidatos_base = candidatos_base[
                            candidatos_base["LP"].astype(str).str.strip() == lp_put
                        ].copy()

                    if candidatos_base.empty:
                        continue

                    mejor_match = None

                    for origen_pick, g_origen in candidatos_base.groupby(
                        candidatos_base["Desde ID Ubicación"].astype(str).str.strip(),
                        sort=False
                    ):
                        if not origen_pick:
                            continue

                        picks_seleccionados = _seleccionar_picks_exactos(
                            candidatos=g_origen,
                            cantidad_objetivo=cantidad_put_total,
                            fecha_put=fecha_put
                        )

                        if not picks_seleccionados:
                            continue

                        resumen_pick = _resumen_item_lote(
                            picks_seleccionados,
                            incluir_lp=False
                        )

                        resumen_put = _resumen_item_lote(
                            puts_rows,
                            incluir_lp=False
                        )

                        if resumen_pick != resumen_put:
                            continue

                        if not permitir_lp_cambiada:
                            resumen_pick_lp = _resumen_item_lote(
                                picks_seleccionados,
                                incluir_lp=True
                            )

                            resumen_put_lp = _resumen_item_lote(
                                puts_rows,
                                incluir_lp=True
                            )

                            if resumen_pick_lp != resumen_put_lp:
                                continue

                        fechas_pick = [
                            pd.to_datetime(r.get("FechaHora"), errors="coerce")
                            for r in picks_seleccionados
                        ]

                        fechas_pick_validas = [
                            f for f in fechas_pick
                            if pd.notna(f)
                        ]

                        if fechas_pick_validas and pd.notna(fecha_put):
                            ultima_fecha_pick = max(fechas_pick_validas)
                            distancia = abs((fecha_put - ultima_fecha_pick).total_seconds())
                        else:
                            distancia = 999999999

                        candidato_match = {
                            "picks": picks_seleccionados,
                            "puts": puts_rows,
                            "distancia": distancia
                        }

                        if mejor_match is None:
                            mejor_match = candidato_match
                        else:
                            if candidato_match["distancia"] < mejor_match["distancia"]:
                                mejor_match = candidato_match

                    if mejor_match is None:
                        continue

                    picks_match = mejor_match["picks"]
                    puts_match = mejor_match["puts"]

                    area_movimiento = _area_bodega_desde_row(puts_match[0])

                    mov_creado = _crear_movimiento_desde_match(
                        picks_rows=picks_match,
                        puts_rows=puts_match,
                        area_movimiento=area_movimiento,
                        forzar_lp_cambiada=permitir_lp_cambiada
                    )

                    if mov_creado is None:
                        continue

                    row_ids_mov = _row_ids_de_rows(picks_match + puts_match)
                    mov_creado["_row_ids_wms"] = row_ids_mov

                    movimientos_fase.append(mov_creado)

                    for rid in row_ids_mov:
                        row_ids_usados_fase.add(rid)

            return movimientos_fase, row_ids_usados_fase
        
        def _procesar_split_m2m(
            df_fase,
            permitir_lp_cambiada,
            row_ids_usados_globales
        ):
            movimientos_fase = []
            row_ids_usados_fase = set()

            if df_fase is None or df_fase.empty:
                return movimientos_fase, row_ids_usados_fase

            df_work = df_fase[
                ~df_fase["_row_id_wms"].isin(row_ids_usados_globales)
            ].copy()

            if df_work.empty:
                return movimientos_fase, row_ids_usados_fase

            for _, g_operador in df_work.groupby(
                ["Nombre_norm", "ID Operador", "Nombre"],
                dropna=False,
                sort=False
            ):
                g_operador = g_operador.sort_values([
                    "FechaHora",
                    "_row_id_wms"
                ]).copy()

                picks = g_operador[
                    g_operador["Description_norm"].isin(pick_descs)
                ].copy()

                puts = g_operador[
                    g_operador["Description_norm"].isin(put_descs)
                ].copy()

                if picks.empty or puts.empty:
                    continue
                
                
                
                
                
                
                

                picks = picks[
                    (~picks["_row_id_wms"].isin(row_ids_usados_globales)) &
                    (~picks["_row_id_wms"].isin(row_ids_usados_fase))
                ].copy()

                puts = puts[
                    (~puts["_row_id_wms"].isin(row_ids_usados_globales)) &
                    (~puts["_row_id_wms"].isin(row_ids_usados_fase))
                ].copy()

                if picks.empty or puts.empty:
                    continue

                if permitir_lp_cambiada:
                    group_pick_cols = [
                        "Nombre_norm",
                        "ID Operador",
                        "Nombre",
                        "_Familia_WMS",
                        "_Bloque_Flujo_Tiempo",
                        "Item Number",
                        "Número Lote",
                        "Desde ID Ubicación",
                        "A ID Ubicación"
                    ]
                else:
                    group_pick_cols = [
                        "Nombre_norm",
                        "ID Operador",
                        "Nombre",
                        "_Familia_WMS",
                        "_Bloque_Flujo_Tiempo",
                        "Item Number",
                        "Número Lote",
                        "Desde ID Ubicación",
                        "A ID Ubicación",
                        "LP"
                    ]

                picks["_row_dict"] = picks.apply(
                    lambda r: r.to_dict(),
                    axis=1
                )

                pick_batches = (
                    picks
                    .groupby(group_pick_cols, dropna=False, sort=False)
                    .agg(
                        Cantidad_Pick_Total=("Cantidad", "sum"),
                        FH_pick=("FechaHora", "min"),
                        Picks_Rows=("_row_dict", list)
                    )
                    .reset_index()
                    .sort_values("FH_pick")
                )

                for _, pick_batch in pick_batches.iterrows():
                    picks_rows = pick_batch.get("Picks_Rows", [])

                    if not isinstance(picks_rows, list) or not picks_rows:
                        continue

                    pick_ids = _row_ids_de_rows(picks_rows)

                    if not pick_ids:
                        continue

                    if any(
                        rid in row_ids_usados_globales or rid in row_ids_usados_fase
                        for rid in pick_ids
                    ):
                        continue

                    nombre_norm = _limpiar_texto(pick_batch.get("Nombre_norm", ""))
                    id_operador = _limpiar_texto(pick_batch.get("ID Operador", ""))
                    nombre = _limpiar_texto(pick_batch.get("Nombre", ""))
                    item_pick = _limpiar_texto(pick_batch.get("Item Number", ""))
                    lote_pick = _limpiar_texto(pick_batch.get("Número Lote", ""))
                    origen_pick = _limpiar_texto(pick_batch.get("Desde ID Ubicación", ""))
                    intermedia_pick = _limpiar_texto(pick_batch.get("A ID Ubicación", ""))
                    
                    familia_pick = _limpiar_texto(pick_batch.get("_Familia_WMS", ""))
                    bloque_pick = pick_batch.get("_Bloque_Flujo_Tiempo", None)

                    cantidad_pick_total = round(
                        _to_float(pick_batch.get("Cantidad_Pick_Total", 0)),
                        6
                    )

                    if cantidad_pick_total <= 0:
                        continue

                    if not origen_pick or not intermedia_pick:
                        continue

                    candidatos_put = puts[
                        (~puts["_row_id_wms"].isin(row_ids_usados_globales)) &
                        (~puts["_row_id_wms"].isin(row_ids_usados_fase)) &
                        (puts["Nombre_norm"].astype(str).str.strip() == nombre_norm) &
                        (puts["ID Operador"].astype(str).str.strip() == id_operador) &
                        (puts["Nombre"].astype(str).str.strip() == nombre) &
                        (puts["_Familia_WMS"].astype(str).str.strip() == familia_pick) &
                        (puts["_Bloque_Flujo_Tiempo"] == bloque_pick) &
                        (puts["Item Number"].astype(str).str.strip() == item_pick) &
                        (puts["Número Lote"].astype(str).str.strip() == lote_pick) &
                        (puts["Desde ID Ubicación"].astype(str).str.strip() == intermedia_pick)
                    ].copy()

                    if candidatos_put.empty:
                        continue

                    if permitir_lp_cambiada:
                        # Fase 2:
                        # NO excluir PUT con LP igual, porque puede existir LP mixta parcial.
                        # Ejemplo:
                        # PICK LP ZF13095204
                        # PUT 7 con ZF13095204 + PUT 5 con C023367429
                        candidatos_put = candidatos_put.copy()
                    else:
                        lp_pick = _limpiar_texto(pick_batch.get("LP", ""))

                        candidatos_put = candidatos_put[
                            candidatos_put["LP"].astype(str).str.strip() == lp_pick
                        ].copy()

                    if candidatos_put.empty:
                        continue

                    candidatos_put["_row_dict"] = candidatos_put.apply(
                        lambda r: r.to_dict(),
                        axis=1
                    )

                    put_groups_df = (
                        candidatos_put
                        .groupby(["A ID Ubicación"], dropna=False, sort=False)
                        .agg(
                            Cantidad_Put_Total=("Cantidad", "sum"),
                            FH_put=("FechaHora", "max"),
                            Puts_Rows=("_row_dict", list)
                        )
                        .reset_index()
                        .sort_values("FH_put")
                    )

                    grupos_put = []

                    for _, g_put in put_groups_df.iterrows():
                        destino = _limpiar_texto(g_put.get("A ID Ubicación", ""))
                        cantidad = round(_to_float(g_put.get("Cantidad_Put_Total", 0)), 6)
                        puts_rows = g_put.get("Puts_Rows", [])
                        fh_put = pd.to_datetime(g_put.get("FH_put"), errors="coerce")

                        if not destino or cantidad <= 0:
                            continue

                        if not isinstance(puts_rows, list) or not puts_rows:
                            continue

                        put_ids = _row_ids_de_rows(puts_rows)

                        if not put_ids:
                            continue

                        if any(
                            rid in row_ids_usados_globales or rid in row_ids_usados_fase
                            for rid in put_ids
                        ):
                            continue

                        grupos_put.append({
                            "destino": destino,
                            "cantidad": cantidad,
                            "puts_rows": puts_rows,
                            "fh_put": fh_put,
                            "put_ids": put_ids
                        })

                    if len(grupos_put) <= 1:
                        continue

                    indices_seleccionados = _buscar_combinacion_exacta_grupos(
                        grupos=grupos_put,
                        cantidad_objetivo=cantidad_pick_total
                    )

                    if not indices_seleccionados:
                        continue

                    grupos_seleccionados = [
                        grupos_put[i]
                        for i in indices_seleccionados
                    ]

                    grupos_productivos = []
                    grupos_retorno_origen = []

                    for g in grupos_seleccionados:
                        destino_g = _limpiar_texto(g.get("destino", ""))

                        if destino_g.upper() == origen_pick.upper():
                            grupos_retorno_origen.append(g)
                        else:
                            grupos_productivos.append(g)

                    if not grupos_productivos:
                        for rid in pick_ids:
                            row_ids_retorno_origen_global.add(rid)

                        for g in grupos_retorno_origen:
                            for rid in g.get("put_ids", []):
                                row_ids_retorno_origen_global.add(rid)

                        continue

                    cantidad_productiva = round(
                        sum(g["cantidad"] for g in grupos_productivos),
                        6
                    )

                    cantidad_retorno = round(
                        sum(g["cantidad"] for g in grupos_retorno_origen),
                        6
                    )

                    if round(cantidad_productiva + cantidad_retorno, 6) != cantidad_pick_total:
                        continue

                    puts_todos = []

                    for g in grupos_productivos:
                        puts_todos.extend(g["puts_rows"])

                    for g in grupos_retorno_origen:
                        puts_todos.extend(g["puts_rows"])

                    resumen_pick = _resumen_item_lote(
                        picks_rows,
                        incluir_lp=False
                    )

                    resumen_put = _resumen_item_lote(
                        puts_todos,
                        incluir_lp=False
                    )

                    if resumen_pick != resumen_put:
                        continue
                    
                    if permitir_lp_cambiada:
                        if not _hay_lp_diferente_o_mixta(picks_rows, puts_todos):
                            continue

                    if not permitir_lp_cambiada:
                        resumen_pick_lp = _resumen_item_lote(
                            picks_rows,
                            incluir_lp=True
                        )

                        resumen_put_lp = _resumen_item_lote(
                            puts_todos,
                            incluir_lp=True
                        )

                        if resumen_pick_lp != resumen_put_lp:
                            continue

                    area_movimiento = _area_bodega_desde_row(picks_rows[0])

                    grupo_split_key = (
                        "M2M|"
                        + "|".join(str(x) for x in pick_ids)
                        + "|"
                        + id_operador
                        + "|"
                        + item_pick
                        + "|"
                        + lote_pick
                        + "|"
                        + origen_pick
                        + "|"
                        + intermedia_pick
                    )

                    fechas_pick_validas = [
                        pd.to_datetime(r.get("FechaHora"), errors="coerce")
                        for r in picks_rows
                        if pd.notna(pd.to_datetime(r.get("FechaHora"), errors="coerce"))
                    ]

                    if fechas_pick_validas:
                        fh_inicio = min(fechas_pick_validas)
                    else:
                        fh_inicio = pd.NaT

                    fechas_put_validas = [
                        g["fh_put"]
                        for g in grupos_productivos + grupos_retorno_origen
                        if pd.notna(g["fh_put"])
                    ]

                    if fechas_put_validas:
                        fh_fin = max(fechas_put_validas)
                    else:
                        fh_fin = fh_inicio

                    lp_pick_set = {
                        _limpiar_texto(r.get("LP", ""))
                        for r in picks_rows
                        if _limpiar_texto(r.get("LP", ""))
                    }

                    lp_put_set = {
                        _limpiar_texto(r.get("LP", ""))
                        for r in puts_todos
                        if _limpiar_texto(r.get("LP", ""))
                    }

                    mov_padre = {
                        "Movimiento_ID": pd.NA,
                        "Pertenece_a_Pick_Padre": pd.NA,
                        "Tipo_Match": "SPLIT",
                        "Area_Movimiento": area_movimiento,
                        "ID Operador": id_operador,
                        "Nombre": nombre,
                        "Nombre_norm": nombre_norm,
                        "Origen": origen_pick,
                        "Intermedia": intermedia_pick,
                        "Destino_Final": " | ".join([g["destino"] for g in grupos_productivos]),
                        "FH_inicio": fh_inicio,
                        "FH_fin": fh_fin,
                        "Cantidad_Pick": cantidad_pick_total,
                        "Cantidad_Put": cantidad_productiva,
                        "Cantidad": cantidad_productiva,
                        "Diferencia_Hijo": 0,
                        "Estado_Cantidad_Hijo": "OK",
                        "Cantidad_Pick_Padre_Total": cantidad_pick_total,
                        "Cantidad_Put_Hijos_Total": cantidad_productiva,
                        "Diferencia_Padre": 0,
                        "Estado_Cantidad_Padre": "OK",
                        "Items_Distintos": 1,
                        "Lotes_Distintos": 1,
                        "LPs_Distintos": len(lp_pick_set.union(lp_put_set)),
                        "Picks_Fragmentos": len(picks_rows),
                        "Puts_Fragmentos": len(puts_todos),
                        "LP": " | ".join(sorted(lp_pick_set.union(lp_put_set))),
                        "LP_Pick": " | ".join(sorted(lp_pick_set)),
                        "LP_Put": " | ".join(sorted(lp_put_set)),
                        "LP_Diferente": bool(permitir_lp_cambiada),
                        "Tiene_LP_Cambiada": bool(permitir_lp_cambiada),
                        "LPs_Pick_Distintas": len(lp_pick_set),
                        "LPs_Put_Distintas": len(lp_put_set),
                        "Minutos_Max_Pick_Put": round((fh_fin - fh_inicio).total_seconds() / 60, 2) if pd.notna(fh_inicio) and pd.notna(fh_fin) else 0,
                        "_row_ids_wms": pick_ids,
                        "_Grupo_Split_Key": grupo_split_key,
                        "_Orden_Split": 0,
                        "_Es_Padre_Split": True
                    }

                    movimientos_fase.append(mov_padre)

                    orden_hijo = 1
                    picks_rows_base = [dict(r) for r in picks_rows]

                    for g in grupos_productivos:
                        cantidad_hijo = round(g["cantidad"], 6)

                        picks_parciales = _crear_picks_parciales_desde_saldo(
                            picks_rows_base,
                            cantidad_hijo
                        )

                        if not picks_parciales:
                            continue

                        mov_hijo = _crear_movimiento_desde_match(
                            picks_rows=picks_parciales,
                            puts_rows=g["puts_rows"],
                            area_movimiento=area_movimiento,
                            forzar_lp_cambiada=permitir_lp_cambiada
                        )

                        if mov_hijo is None:
                            continue

                        mov_hijo["Tipo_Match"] = "SPLIT"
                        mov_hijo["Cantidad_Pick_Padre_Total"] = cantidad_pick_total
                        mov_hijo["Cantidad_Put_Hijos_Total"] = cantidad_productiva
                        mov_hijo["_row_ids_wms"] = g["put_ids"]
                        mov_hijo["_Grupo_Split_Key"] = grupo_split_key
                        mov_hijo["_Orden_Split"] = orden_hijo
                        mov_hijo["_Es_Padre_Split"] = False

                        movimientos_fase.append(mov_hijo)
                        orden_hijo += 1

                    for rid in pick_ids:
                        row_ids_usados_fase.add(rid)

                    for g in grupos_productivos:
                        for rid in g.get("put_ids", []):
                            row_ids_usados_fase.add(rid)

                    for g in grupos_retorno_origen:
                        for rid in g.get("put_ids", []):
                            row_ids_usados_fase.add(rid)
                            row_ids_retorno_origen_global.add(rid)

            return movimientos_fase, row_ids_usados_fase
        
                    
        def _procesar_por_origen_y_retorno(
            df_fase,
            permitir_lp_cambiada,
            row_ids_usados_globales
        ):
            """
            Capa final de recuperación.

            Detecta:
            - MOV_REAL separado por origen.
            - NO_PRODUCTIVO_RETORNO_ORIGEN cuando destino == origen.
            - No mezcla Move Trailer con Move normal gracias a _Familia_WMS.
            """

            movimientos_fase = []
            row_ids_usados_fase = set()

            if df_fase is None or df_fase.empty:
                return movimientos_fase, row_ids_usados_fase

            df_work = df_fase[
                ~df_fase["_row_id_wms"].isin(row_ids_usados_globales)
            ].copy()

            if df_work.empty:
                return movimientos_fase, row_ids_usados_fase

            for _, g_operador in df_work.groupby(
                ["Nombre_norm", "ID Operador", "Nombre"],
                dropna=False,
                sort=False
            ):
                g_operador = g_operador.sort_values([
                    "FechaHora",
                    "_row_id_wms"
                ]).copy()

                picks = g_operador[
                    g_operador["Description_norm"].isin(pick_descs)
                ].copy()

                puts = g_operador[
                    g_operador["Description_norm"].isin(put_descs)
                ].copy()

                if picks.empty or puts.empty:
                    continue

                picks = picks[
                    (~picks["_row_id_wms"].isin(row_ids_usados_globales)) &
                    (~picks["_row_id_wms"].isin(row_ids_usados_fase))
                ].copy()

                puts = puts[
                    (~puts["_row_id_wms"].isin(row_ids_usados_globales)) &
                    (~puts["_row_id_wms"].isin(row_ids_usados_fase))
                ].copy()

                if picks.empty or puts.empty:
                    continue
                    
                    
            if permitir_lp_cambiada:
                    group_pick_cols = [
                        "Nombre_norm",
                        "ID Operador",
                        "Nombre",
                        "_Familia_WMS",
                        "_Bloque_Flujo_Tiempo",
                        "Item Number",
                        "Número Lote",
                        "Desde ID Ubicación",
                        "A ID Ubicación"
                    ]
            else:
                    group_pick_cols = [
                        "Nombre_norm",
                        "ID Operador",
                        "Nombre",
                        "_Familia_WMS",
                        "_Bloque_Flujo_Tiempo",
                        "Item Number",
                        "Número Lote",
                        "Desde ID Ubicación",
                        "A ID Ubicación",
                        "LP"
                    ]

            picks["_row_dict"] = picks.apply(
                lambda r: r.to_dict(),
                axis=1
            )

            pick_batches = (
                picks
                .groupby(group_pick_cols, dropna=False, sort=False)
                .agg(
                    Cantidad_Pick_Total=("Cantidad", "sum"),
                    FH_pick=("FechaHora", "min"),
                    Picks_Rows=("_row_dict", list)
                    )
                .reset_index()
                .sort_values("FH_pick")
            )

            for _, pick_batch in pick_batches.iterrows():
                picks_rows = pick_batch.get("Picks_Rows", [])

                if not isinstance(picks_rows, list) or not picks_rows:
                    continue

                pick_ids = _row_ids_de_rows(picks_rows)

                if not pick_ids:
                        continue

                if any(
                    rid in row_ids_usados_globales or rid in row_ids_usados_fase
                    for rid in pick_ids
                ):
                    continue

                nombre_norm = _limpiar_texto(pick_batch.get("Nombre_norm", ""))
                id_operador = _limpiar_texto(pick_batch.get("ID Operador", ""))
                nombre = _limpiar_texto(pick_batch.get("Nombre", ""))
                familia_pick = _limpiar_texto(pick_batch.get("_Familia_WMS", ""))
                bloque_pick = pick_batch.get("_Bloque_Flujo_Tiempo", None)
                item_pick = _limpiar_texto(pick_batch.get("Item Number", ""))
                lote_pick = _limpiar_texto(pick_batch.get("Número Lote", ""))
                origen_pick = _limpiar_texto(pick_batch.get("Desde ID Ubicación", ""))
                intermedia_pick = _limpiar_texto(pick_batch.get("A ID Ubicación", ""))

                cantidad_pick_total = round(
                    _to_float(pick_batch.get("Cantidad_Pick_Total", 0)),
                    6
                )

                if cantidad_pick_total <= 0:
                    continue

                if not origen_pick or not intermedia_pick:
                    continue
                    
                candidatos_put = puts[
                    (~puts["_row_id_wms"].isin(row_ids_usados_globales)) &
                    (~puts["_row_id_wms"].isin(row_ids_usados_fase)) &
                    (puts["Nombre_norm"].astype(str).str.strip() == nombre_norm) &
                    (puts["ID Operador"].astype(str).str.strip() == id_operador) &
                    (puts["Nombre"].astype(str).str.strip() == nombre) &
                    (puts["_Familia_WMS"].astype(str).str.strip() == familia_pick) &
                    (puts["_Bloque_Flujo_Tiempo"] == bloque_pick) &
                    (puts["Item Number"].astype(str).str.strip() == item_pick) &
                    (puts["Número Lote"].astype(str).str.strip() == lote_pick) &
                    (puts["Desde ID Ubicación"].astype(str).str.strip() == intermedia_pick)
                ].copy()

                if candidatos_put.empty:
                    continue

                if permitir_lp_cambiada:
                    # Fase 2:
                    # NO excluir PUT con LP igual, porque puede existir LP mixta parcial.
                    # Ejemplo:
                    # PICK LP ZF13095204
                    # PUT 7 con ZF13095204 + PUT 5 con C023367429
                    candidatos_put = candidatos_put.copy()
                else:
                    lp_pick = _limpiar_texto(pick_batch.get("LP", ""))

                    candidatos_put = candidatos_put[
                        candidatos_put["LP"].astype(str).str.strip() == lp_pick
                    ].copy()

                if candidatos_put.empty:
                    continue

                candidatos_put["_row_dict"] = candidatos_put.apply(
                    lambda r: r.to_dict(),
                    axis=1
                )

                put_groups = (
                    candidatos_put
                    .groupby(["A ID Ubicación"], dropna=False, sort=False)
                    .agg(
                        Cantidad_Put_Total=("Cantidad", "sum"),
                        FH_put=("FechaHora", "max"),
                        Puts_Rows=("_row_dict", list)
                    )
                    .reset_index()
                    .sort_values("FH_put")
                )

                grupos_productivos = []
                grupos_retorno = []

                for _, put_group in put_groups.iterrows():
                    destino = _limpiar_texto(put_group.get("A ID Ubicación", ""))
                    cantidad_put = round(
                        _to_float(put_group.get("Cantidad_Put_Total", 0)),
                        6
                    )
                    puts_rows = put_group.get("Puts_Rows", [])
                    fh_put = pd.to_datetime(
                        put_group.get("FH_put"),
                        errors="coerce"
                    )

                    if not destino or cantidad_put <= 0:
                        continue

                    if not isinstance(puts_rows, list) or not puts_rows:
                        continue

                    put_ids = _row_ids_de_rows(puts_rows)

                    if not put_ids:
                            continue

                    if any(
                        rid in row_ids_usados_globales or rid in row_ids_usados_fase
                        for rid in put_ids
                    ):
                        continue

                    grupo = {
                        "destino": destino,
                        "cantidad": cantidad_put,
                        "puts_rows": puts_rows,
                        "put_ids": put_ids,
                        "fh_put": fh_put
                    }

                    if destino.upper() == origen_pick.upper():
                        grupos_retorno.append(grupo)
                    else:
                        grupos_productivos.append(grupo)
                        
                    
                    cantidad_productiva = round(
                    sum(g["cantidad"] for g in grupos_productivos),
                    6
                )

                cantidad_retorno = round(
                    sum(g["cantidad"] for g in grupos_retorno),
                    6
                )

                if round(cantidad_productiva + cantidad_retorno, 6) != cantidad_pick_total:               
                    continue
                
                # ============================================================
                # PRIORIDAD PRODUCTIVA SOBRE RETORNO
                # ============================================================

                puts_productivos_validacion = []

                for g in grupos_productivos:
                    puts_productivos_validacion.extend(g.get("puts_rows", []))

                if puts_productivos_validacion:
                    # Si hay PUT productivo compatible, NO marcar como retorno.
                    if _flujo_en_tiempo(picks_rows, puts_productivos_validacion):
                        pass

                # ============================================================
                # VALIDAR RETORNOS CRONOLÓGICOS
                # El PUT retorno debe ocurrir después del PICK.
                # ============================================================

                grupos_retorno_validos = []

                for g in grupos_retorno:
                    puts_retorno_tmp = g.get("puts_rows", [])

                    if not puts_retorno_tmp:
                        continue

                    if _retorno_cronologicamente_valido(picks_rows, puts_retorno_tmp):
                        grupos_retorno_validos.append(g)

                grupos_retorno = grupos_retorno_validos
                

                # Caso 1: todo regresó al origen.
                # Solo se marca si existen retornos cronológicamente válidos.
                if cantidad_productiva <= 0:

                    if not grupos_retorno:
                        continue

                    for rid in pick_ids:
                        try:
                            rid = int(rid)
                            row_ids_retorno_origen_global.add(rid)
                            row_ids_usados_fase.add(rid)
                        except Exception:
                            pass

                    for g in grupos_retorno:
                        for rid in g.get("put_ids", []):
                            try:
                                rid = int(rid)
                                row_ids_retorno_origen_global.add(rid)
                                row_ids_usados_fase.add(rid)
                            except Exception:
                                pass

                    continue
                
                puts_todos_validacion = []

                for g in grupos_productivos:
                    puts_todos_validacion.extend(g.get("puts_rows", []))

                for g in grupos_retorno:
                    puts_todos_validacion.extend(g.get("puts_rows", []))

                if permitir_lp_cambiada:
                    if not _hay_lp_diferente_o_mixta(picks_rows, puts_todos_validacion):
                        continue

                area_movimiento = _area_bodega_desde_row(picks_rows[0])

                for g in grupos_productivos:
                    cantidad_hijo = round(g["cantidad"], 6)

                    picks_parciales = _crear_picks_parciales_desde_saldo(
                        picks_rows,
                        cantidad_hijo
                    )

                    if not picks_parciales:
                        continue

                    mov_creado = _crear_movimiento_desde_match(
                        picks_rows=picks_parciales,
                        puts_rows=g["puts_rows"],
                        area_movimiento=area_movimiento,
                        forzar_lp_cambiada=permitir_lp_cambiada
                    )

                    if mov_creado is None:
                        continue

                    mov_creado["Tipo_Match"] = "NORMAL"
                    mov_creado["_row_ids_wms"] = _row_ids_de_rows(
                        picks_parciales + g["puts_rows"]
                    )

                    movimientos_fase.append(mov_creado)

                    for rid in mov_creado["_row_ids_wms"]:
                        try:
                            row_ids_usados_fase.add(int(rid))
                        except Exception:
                            pass

                # Marcar retornos como no productivos.
                for g in grupos_retorno:
                    for rid in g.get("put_ids", []):
                        try:
                            rid = int(rid)
                            row_ids_retorno_origen_global.add(rid)
                            row_ids_usados_fase.add(rid)
                        except Exception:
                            pass

            return movimientos_fase, row_ids_usados_fase

        def _reconstruir_movimientos_por_balance(filas):
            """
            Reconstruye movimientos permitiendo desorden PICK/PUT.

            - No importa el orden en que vengan las filas.
            - Se basa en balance acumulado.
            """

            buffer_picks = []
            buffer_puts = []

            suma_pick = 0.0
            suma_put = 0.0

            resultados = []

            for row in filas:
                desc = _limpiar_texto(row.get("Description_norm", ""))
                qty = round(_to_float(row.get("Cantidad", 0)), 6)

                if qty <= 0:
                    continue

                es_pick = desc in pick_descs
                es_put = desc in put_descs

                if not es_pick and not es_put:
                    continue

                # --------------------------------------
                # Separar por tipo (NO importa el orden)
                # --------------------------------------

                if es_pick:
                    buffer_picks.append(row)
                    suma_pick += qty

                if es_put:
                    buffer_puts.append(row)
                    suma_put += qty

                # --------------------------------------
                # Evaluar balance
                # --------------------------------------
                if (
                    suma_pick > 0
                    and suma_put > 0
                    and round(suma_pick, 6) == round(suma_put, 6)
                ):

                    # ====================================================
                    # FILTRO EXTRA 1:
                    # Validar que la intermedia del PICK sea la misma
                    # desde donde sale el PUT.
                    #
                    # PICK: Origen -> Intermedia
                    # PUT:  Intermedia -> Destino
                    # ====================================================

                    intermedias_pick = {
                        _limpiar_texto(r.get("A ID Ubicación", ""))
                        for r in buffer_picks
                        if _limpiar_texto(r.get("A ID Ubicación", ""))
                    }

                    intermedias_put = {
                        _limpiar_texto(r.get("Desde ID Ubicación", ""))
                        for r in buffer_puts
                        if _limpiar_texto(r.get("Desde ID Ubicación", ""))
                    }

                    if intermedias_pick != intermedias_put:
                        continue

                    # ====================================================
                    # FILTRO EXTRA 2:
                    # Anti-fraude / precisión.
                    # Evita mezclar demasiados orígenes diferentes dentro
                    # de un mismo flujo reconstruido.
                    # ====================================================

                    origenes_pick = {
                        _limpiar_texto(r.get("Desde ID Ubicación", ""))
                        for r in buffer_picks
                        if _limpiar_texto(r.get("Desde ID Ubicación", ""))
                    }

                    if len(origenes_pick) > 3:
                        continue

                    resultados.append((list(buffer_picks), list(buffer_puts)))

                    # Reset buffers
                    buffer_picks = []
                    buffer_puts = []
                    suma_pick = 0.0
                    suma_put = 0.0

            return resultados

        
        def _procesar_balance_secuencial(
            df_fase,
            row_ids_usados_globales
        ):
            movimientos = []
            usados = set()

            if df_fase is None or df_fase.empty:
                return movimientos, usados

            for _, g in df_fase.groupby(
                ["Nombre_norm", "ID Operador", "_Familia_WMS", "_Bloque_Flujo_Tiempo", "Item Number", "Número Lote"],
                dropna=False,
                sort=False
            ):

                g = g.sort_values(["FechaHora", "_row_id_wms"]).copy()

                filas = [
                    r.to_dict()
                    for _, r in g.iterrows()
                    if r["_row_id_wms"] not in row_ids_usados_globales
                ]

                if not filas:
                    continue

                matches = _reconstruir_movimientos_por_balance(filas)

                for picks_rows, puts_rows in matches:

                    area = _area_bodega_desde_row(picks_rows[0])

                    mov = _crear_movimiento_desde_match(
                        picks_rows,
                        puts_rows,
                        area
                    )

                    if mov is None:
                        continue

                    row_ids = _row_ids_de_rows(picks_rows + puts_rows)
                    mov["_row_ids_wms"] = row_ids
                    mov["Tipo_Match"] = "BALANCE_SECUENCIAL"

                    movimientos.append(mov)

                    for rid in row_ids:
                        usados.add(rid)

            return movimientos, usados
        
        def _procesar_balance_lp_mixta_por_bloque(
            df_fase,
            permitir_lp_cambiada,
            row_ids_usados_globales
        ):
            """
            Capa final para recuperar movimientos reales con LP mixta.

            Caso típico:
            - Algunos PUT tienen la misma LP del PICK.
            - Otros PUT tienen LP diferente.
            - El balance total PICK/PUT sí cuadra.
            """

            movimientos_fase = []
            row_ids_usados_fase = set()

            # Esta capa solo debe correr en fase 2.
            # Primero dejamos que la fase LP igual resuelva lo normal.
            if not permitir_lp_cambiada:
                return movimientos_fase, row_ids_usados_fase

            if df_fase is None or df_fase.empty:
                return movimientos_fase, row_ids_usados_fase

            df_work = df_fase[
                ~df_fase["_row_id_wms"].isin(row_ids_usados_globales)
            ].copy()

            if df_work.empty:
                return movimientos_fase, row_ids_usados_fase

            for _, g_operador in df_work.groupby(
                ["Nombre_norm", "ID Operador", "Nombre"],
                dropna=False,
                sort=False
            ):
                g_operador = g_operador.sort_values([
                    "FechaHora",
                    "_row_id_wms"
                ]).copy()

                picks = g_operador[
                    g_operador["Description_norm"].isin(pick_descs)
                ].copy()

                puts = g_operador[
                    g_operador["Description_norm"].isin(put_descs)
                ].copy()

                if picks.empty or puts.empty:
                    continue

                picks = picks[
                    (~picks["_row_id_wms"].isin(row_ids_usados_globales)) &
                    (~picks["_row_id_wms"].isin(row_ids_usados_fase))
                ].copy()

                puts = puts[
                    (~puts["_row_id_wms"].isin(row_ids_usados_globales)) &
                    (~puts["_row_id_wms"].isin(row_ids_usados_fase))
                ].copy()

                if picks.empty or puts.empty:
                    continue

                group_pick_cols = [
                    "Nombre_norm",
                    "ID Operador",
                    "Nombre",
                    "_Familia_WMS",
                    "_Bloque_Flujo_Tiempo",
                    "Item Number",
                    "Número Lote",
                    "Desde ID Ubicación",
                    "A ID Ubicación"
                ]

                picks["_row_dict"] = picks.apply(
                    lambda r: r.to_dict(),
                    axis=1
                )

                pick_batches = (
                    picks
                    .groupby(group_pick_cols, dropna=False, sort=False)
                    .agg(
                        Cantidad_Pick_Total=("Cantidad", "sum"),
                        FH_pick=("FechaHora", "min"),
                        Picks_Rows=("_row_dict", list)
                    )
                    .reset_index()
                    .sort_values("FH_pick")
                )
                
            for _, pick_batch in pick_batches.iterrows():
                    picks_rows = pick_batch.get("Picks_Rows", [])

                    if not isinstance(picks_rows, list) or not picks_rows:
                        continue

                    pick_ids = _row_ids_de_rows(picks_rows)

                    if not pick_ids:
                        continue

                    if any(
                        rid in row_ids_usados_globales or rid in row_ids_usados_fase
                        for rid in pick_ids
                    ):
                        continue

                    nombre_norm = _limpiar_texto(pick_batch.get("Nombre_norm", ""))
                    id_operador = _limpiar_texto(pick_batch.get("ID Operador", ""))
                    nombre = _limpiar_texto(pick_batch.get("Nombre", ""))
                    familia_pick = _limpiar_texto(pick_batch.get("_Familia_WMS", ""))
                    bloque_pick = pick_batch.get("_Bloque_Flujo_Tiempo", None)
                    item_pick = _limpiar_texto(pick_batch.get("Item Number", ""))
                    lote_pick = _limpiar_texto(pick_batch.get("Número Lote", ""))
                    origen_pick = _limpiar_texto(pick_batch.get("Desde ID Ubicación", ""))
                    intermedia_pick = _limpiar_texto(pick_batch.get("A ID Ubicación", ""))

                    cantidad_pick_total = round(
                        _to_float(pick_batch.get("Cantidad_Pick_Total", 0)),
                        6
                    )

                    if cantidad_pick_total <= 0:
                        continue

                    candidatos_put = puts[
                        (~puts["_row_id_wms"].isin(row_ids_usados_globales)) &
                        (~puts["_row_id_wms"].isin(row_ids_usados_fase)) &
                        (puts["Nombre_norm"].astype(str).str.strip() == nombre_norm) &
                        (puts["ID Operador"].astype(str).str.strip() == id_operador) &
                        (puts["Nombre"].astype(str).str.strip() == nombre) &
                        (puts["_Familia_WMS"].astype(str).str.strip() == familia_pick) &
                        (puts["_Bloque_Flujo_Tiempo"] == bloque_pick) &
                        (puts["Item Number"].astype(str).str.strip() == item_pick) &
                        (puts["Número Lote"].astype(str).str.strip() == lote_pick) &
                        (puts["Desde ID Ubicación"].astype(str).str.strip() == intermedia_pick)
                    ].copy()

                    if candidatos_put.empty:
                        continue

                    candidatos_put["_row_dict"] = candidatos_put.apply(
                        lambda r: r.to_dict(),
                        axis=1
                    )

                    put_groups_df = (
                        candidatos_put
                        .groupby(["A ID Ubicación"], dropna=False, sort=False)
                        .agg(
                            Cantidad_Put_Total=("Cantidad", "sum"),
                            FH_put=("FechaHora", "max"),
                            Puts_Rows=("_row_dict", list)
                        )
                        .reset_index()
                        .sort_values("FH_put")
                    )

                    grupos_put = []

                    for _, g_put in put_groups_df.iterrows():
                        destino = _limpiar_texto(g_put.get("A ID Ubicación", ""))
                        cantidad = round(_to_float(g_put.get("Cantidad_Put_Total", 0)), 6)
                        puts_rows = g_put.get("Puts_Rows", [])
                        fh_put = pd.to_datetime(g_put.get("FH_put"), errors="coerce")

                        if not destino or cantidad <= 0:
                            continue

                        if not isinstance(puts_rows, list) or not puts_rows:
                            continue

                        put_ids = _row_ids_de_rows(puts_rows)

                        if not put_ids:
                            continue

                        grupos_put.append({
                            "destino": destino,
                            "cantidad": cantidad,
                            "puts_rows": puts_rows,
                            "fh_put": fh_put,
                            "put_ids": put_ids
                        })

                    if not grupos_put:
                        continue
                    
                    cantidad_put_total = round(
                        sum(g["cantidad"] for g in grupos_put),
                        6
                    )

                    if cantidad_put_total != cantidad_pick_total:
                        continue

                    puts_todos = []

                    for g in grupos_put:
                        puts_todos.extend(g["puts_rows"])

                    resumen_pick = _resumen_item_lote(
                        picks_rows,
                        incluir_lp=False
                    )

                    resumen_put = _resumen_item_lote(
                        puts_todos,
                        incluir_lp=False
                    )

                    if resumen_pick != resumen_put:
                        continue
                    
                    if permitir_lp_cambiada:
                        if not _hay_lp_diferente_o_mixta(picks_rows, puts_todos):
                            continue

                    lp_pick_set = {
                        _limpiar_texto(r.get("LP", ""))
                        for r in picks_rows
                        if _limpiar_texto(r.get("LP", ""))
                    }

                    lp_put_set = {
                        _limpiar_texto(r.get("LP", ""))
                        for r in puts_todos
                        if _limpiar_texto(r.get("LP", ""))
                    }

                    if lp_pick_set == lp_put_set:
                        continue

                    grupos_productivos = []
                    grupos_retorno = []

                    for g in grupos_put:
                        destino_g = _limpiar_texto(g.get("destino", ""))

                        if destino_g.upper() == origen_pick.upper():
                            grupos_retorno.append(g)
                        else:
                            grupos_productivos.append(g)

                    if not grupos_productivos:
                        for rid in pick_ids:
                            row_ids_retorno_origen_global.add(rid)
                            row_ids_usados_fase.add(rid)

                        for g in grupos_retorno:
                            for rid in g.get("put_ids", []):
                                row_ids_retorno_origen_global.add(rid)
                                row_ids_usados_fase.add(rid)

                        continue

                    area_movimiento = _area_bodega_desde_row(picks_rows[0])

                    grupo_split_key = (
                        "LP_MIXTA|"
                        + "|".join(str(x) for x in pick_ids)
                        + "|"
                        + id_operador
                        + "|"
                        + item_pick
                        + "|"
                        + lote_pick
                        + "|"
                        + origen_pick
                        + "|"
                        + intermedia_pick
                    )

                    mov_padre = {
                        "Movimiento_ID": pd.NA,
                        "Pertenece_a_Pick_Padre": pd.NA,
                        "Tipo_Match": "LP_CAMBIADA",
                        "Area_Movimiento": area_movimiento,
                        "ID Operador": id_operador,
                        "Nombre": nombre,
                        "Nombre_norm": nombre_norm,
                        "Origen": origen_pick,
                        "Intermedia": intermedia_pick,
                        "Destino_Final": " | ".join([g["destino"] for g in grupos_productivos]),
                        "FH_inicio": min(pd.to_datetime(r.get("FechaHora"), errors="coerce") for r in picks_rows),
                        "FH_fin": max(g["fh_put"] for g in grupos_put if pd.notna(g["fh_put"])),
                        "Cantidad_Pick": cantidad_pick_total,
                        "Cantidad_Put": cantidad_put_total,
                        "Cantidad": cantidad_put_total,
                        "Diferencia_Hijo": 0,
                        "Estado_Cantidad_Hijo": "OK",
                        "Cantidad_Pick_Padre_Total": cantidad_pick_total,
                        "Cantidad_Put_Hijos_Total": cantidad_put_total,
                        "Diferencia_Padre": 0,
                        "Estado_Cantidad_Padre": "OK",
                        "Items_Distintos": 1,
                        "Lotes_Distintos": 1,
                        "LPs_Distintos": len(lp_pick_set.union(lp_put_set)),
                        "Picks_Fragmentos": len(picks_rows),
                        "Puts_Fragmentos": len(puts_todos),
                        "LP": " | ".join(sorted(lp_pick_set.union(lp_put_set))),
                        "LP_Pick": " | ".join(sorted(lp_pick_set)),
                        "LP_Put": " | ".join(sorted(lp_put_set)),
                        "LP_Diferente": True,
                        "Tiene_LP_Cambiada": True,
                        "LPs_Pick_Distintas": len(lp_pick_set),
                        "LPs_Put_Distintas": len(lp_put_set),
                        "Minutos_Max_Pick_Put": 0,
                        "_row_ids_wms": pick_ids,
                        "_Grupo_Split_Key": grupo_split_key,
                        "_Orden_Split": 0,
                        "_Es_Padre_Split": True
                    }

                    movimientos_fase.append(mov_padre)

                    orden_hijo = 1

                    for g in grupos_productivos:
                        cantidad_hijo = round(g["cantidad"], 6)

                        picks_parciales = _crear_picks_parciales_desde_saldo(
                            picks_rows,
                            cantidad_hijo
                        )

                        if not picks_parciales:
                            continue

                        mov_hijo = _crear_movimiento_desde_match(
                            picks_rows=picks_parciales,
                            puts_rows=g["puts_rows"],
                            area_movimiento=area_movimiento,
                            forzar_lp_cambiada=True
                        )

                        if mov_hijo is None:
                            continue

                        mov_hijo["Tipo_Match"] = "LP_CAMBIADA"
                        mov_hijo["_row_ids_wms"] = g["put_ids"]
                        mov_hijo["_Grupo_Split_Key"] = grupo_split_key
                        mov_hijo["_Orden_Split"] = orden_hijo
                        mov_hijo["_Es_Padre_Split"] = False

                        movimientos_fase.append(mov_hijo)
                        orden_hijo += 1

                    for rid in pick_ids:
                        row_ids_usados_fase.add(rid)

                    for g in grupos_productivos:
                        for rid in g.get("put_ids", []):
                            row_ids_usados_fase.add(rid)

                    for g in grupos_retorno:
                        for rid in g.get("put_ids", []):
                            row_ids_retorno_origen_global.add(rid)
                            row_ids_usados_fase.add(rid)

            return movimientos_fase, row_ids_usados_fase
        
        def _procesar_padre_hijos_por_destino(
            df_fase,
            permitir_lp_cambiada,
            row_ids_usados_globales
        ):
            """
            Capa prioritaria para movimientos multidestino.

            Regla:
            - Un flujo padre se define por operador + familia + bloque tiempo
              + item/lote + origen + intermedia + LP en fase 1.
            - Dentro del padre, cada destino final genera un movimiento hijo.
            - Si el mismo destino aparece en varios PUT, se consolida en un solo hijo.
            """

            movimientos_fase = []
            row_ids_usados_fase = set()

            if df_fase is None or df_fase.empty:
                return movimientos_fase, row_ids_usados_fase

            df_work = df_fase[
                ~df_fase["_row_id_wms"].isin(row_ids_usados_globales)
            ].copy()

            if df_work.empty:
                return movimientos_fase, row_ids_usados_fase

            for _, g_operador in df_work.groupby(
                ["Nombre_norm", "ID Operador", "Nombre"],
                dropna=False,
                sort=False
            ):
                g_operador = g_operador.sort_values([
                    "FechaHora",
                    "_row_id_wms"
                ]).copy()

                picks = g_operador[
                    g_operador["Description_norm"].isin(pick_descs)
                ].copy()

                puts = g_operador[
                    g_operador["Description_norm"].isin(put_descs)
                ].copy()

                if picks.empty or puts.empty:
                    continue

                picks = picks[
                    (~picks["_row_id_wms"].isin(row_ids_usados_globales)) &
                    (~picks["_row_id_wms"].isin(row_ids_usados_fase))
                ].copy()

                puts = puts[
                    (~puts["_row_id_wms"].isin(row_ids_usados_globales)) &
                    (~puts["_row_id_wms"].isin(row_ids_usados_fase))
                ].copy()

                if picks.empty or puts.empty:
                    continue

                if permitir_lp_cambiada:
                    group_pick_cols = [
                        "Nombre_norm",
                        "ID Operador",
                        "Nombre",
                        "_Familia_WMS",
                        "_Bloque_Flujo_Tiempo",
                        "Item Number",
                        "Número Lote",
                        "Desde ID Ubicación",
                        "A ID Ubicación"
                    ]
                else:
                    group_pick_cols = [
                        "Nombre_norm",
                        "ID Operador",
                        "Nombre",
                        "_Familia_WMS",
                        "_Bloque_Flujo_Tiempo",
                        "Item Number",
                        "Número Lote",
                        "Desde ID Ubicación",
                        "A ID Ubicación",
                        "LP"
                    ]

                picks["_row_dict"] = picks.apply(
                    lambda r: r.to_dict(),
                    axis=1
                )

                pick_batches = (
                    picks
                    .groupby(group_pick_cols, dropna=False, sort=False)
                    .agg(
                        Cantidad_Pick_Total=("Cantidad", "sum"),
                        FH_pick=("FechaHora", "min"),
                        Picks_Rows=("_row_dict", list)
                    )
                    .reset_index()
                    .sort_values("FH_pick")
                )

                for _, pick_batch in pick_batches.iterrows():
                    picks_rows = pick_batch.get("Picks_Rows", [])

                    if not isinstance(picks_rows, list) or not picks_rows:
                        continue

                    pick_ids = _row_ids_de_rows(picks_rows)

                    if not pick_ids:
                        continue

                    if any(
                        rid in row_ids_usados_globales or rid in row_ids_usados_fase
                        for rid in pick_ids
                    ):
                        continue

                    nombre_norm = _limpiar_texto(pick_batch.get("Nombre_norm", ""))
                    id_operador = _limpiar_texto(pick_batch.get("ID Operador", ""))
                    nombre = _limpiar_texto(pick_batch.get("Nombre", ""))
                    familia_pick = _limpiar_texto(pick_batch.get("_Familia_WMS", ""))
                    bloque_pick = pick_batch.get("_Bloque_Flujo_Tiempo", None)
                    item_pick = _limpiar_texto(pick_batch.get("Item Number", ""))
                    lote_pick = _limpiar_texto(pick_batch.get("Número Lote", ""))
                    origen_pick = _limpiar_texto(pick_batch.get("Desde ID Ubicación", ""))
                    intermedia_pick = _limpiar_texto(pick_batch.get("A ID Ubicación", ""))

                    cantidad_pick_total = round(
                        _to_float(pick_batch.get("Cantidad_Pick_Total", 0)),
                        6
                    )

                    if cantidad_pick_total <= 0:
                        continue

                    if not origen_pick or not intermedia_pick:
                        continue

                    candidatos_put = puts[
                        (~puts["_row_id_wms"].isin(row_ids_usados_globales)) &
                        (~puts["_row_id_wms"].isin(row_ids_usados_fase)) &
                        (puts["Nombre_norm"].astype(str).str.strip() == nombre_norm) &
                        (puts["ID Operador"].astype(str).str.strip() == id_operador) &
                        (puts["Nombre"].astype(str).str.strip() == nombre) &
                        (puts["_Familia_WMS"].astype(str).str.strip() == familia_pick) &
                        (puts["_Bloque_Flujo_Tiempo"] == bloque_pick) &
                        (puts["Item Number"].astype(str).str.strip() == item_pick) &
                        (puts["Número Lote"].astype(str).str.strip() == lote_pick) &
                        (puts["Desde ID Ubicación"].astype(str).str.strip() == intermedia_pick)
                    ].copy()

                    if candidatos_put.empty:
                        continue

                    if not permitir_lp_cambiada:
                        lp_pick = _limpiar_texto(pick_batch.get("LP", ""))

                        candidatos_put = candidatos_put[
                            candidatos_put["LP"].astype(str).str.strip() == lp_pick
                        ].copy()

                    if candidatos_put.empty:
                        continue

                    candidatos_put["_row_dict"] = candidatos_put.apply(
                        lambda r: r.to_dict(),
                        axis=1
                    )

                    put_groups_df = (
                        candidatos_put
                        .groupby(["A ID Ubicación"], dropna=False, sort=False)
                        .agg(
                            Cantidad_Put_Total=("Cantidad", "sum"),
                            FH_put=("FechaHora", "max"),
                            Puts_Rows=("_row_dict", list)
                        )
                        .reset_index()
                        .sort_values("FH_put")
                    )

                    grupos_productivos = []
                    grupos_retorno = []

                    for _, g_put in put_groups_df.iterrows():
                        destino = _limpiar_texto(g_put.get("A ID Ubicación", ""))
                        cantidad = round(
                            _to_float(g_put.get("Cantidad_Put_Total", 0)),
                            6
                        )
                        puts_rows = g_put.get("Puts_Rows", [])
                        fh_put = pd.to_datetime(g_put.get("FH_put"), errors="coerce")

                        if not destino or cantidad <= 0:
                            continue

                        if not isinstance(puts_rows, list) or not puts_rows:
                            continue

                        put_ids = _row_ids_de_rows(puts_rows)

                        if not put_ids:
                            continue

                        if any(
                            rid in row_ids_usados_globales or rid in row_ids_usados_fase
                            for rid in put_ids
                        ):
                            continue

                        grupo = {
                            "destino": destino,
                            "cantidad": cantidad,
                            "puts_rows": puts_rows,
                            "put_ids": put_ids,
                            "fh_put": fh_put
                        }

                        if destino.upper() == origen_pick.upper():
                            grupos_retorno.append(grupo)
                        else:
                            grupos_productivos.append(grupo)

                    if len(grupos_productivos) <= 1:
                        # Esta capa solo debe resolver padre con varios destinos.
                        continue

                    cantidad_productiva = round(
                        sum(g["cantidad"] for g in grupos_productivos),
                        6
                    )

                    cantidad_retorno = round(
                        sum(g["cantidad"] for g in grupos_retorno),
                        6
                    )

                    cantidad_put_total = round(
                        cantidad_productiva + cantidad_retorno,
                        6
                    )

                    if cantidad_put_total != cantidad_pick_total:
                        continue

                    puts_todos = []

                    for g in grupos_productivos:
                        puts_todos.extend(g["puts_rows"])

                    for g in grupos_retorno:
                        puts_todos.extend(g["puts_rows"])

                    resumen_pick = _resumen_item_lote(
                        picks_rows,
                        incluir_lp=False
                    )

                    resumen_put = _resumen_item_lote(
                        puts_todos,
                        incluir_lp=False
                    )

                    if resumen_pick != resumen_put:
                        continue

                    if not permitir_lp_cambiada:
                        resumen_pick_lp = _resumen_item_lote(
                            picks_rows,
                            incluir_lp=True
                        )

                        resumen_put_lp = _resumen_item_lote(
                            puts_todos,
                            incluir_lp=True
                        )

                        if resumen_pick_lp != resumen_put_lp:
                            continue

                    if permitir_lp_cambiada:
                        if not _hay_lp_diferente_o_mixta(picks_rows, puts_todos):
                            continue

                    area_movimiento = _area_bodega_desde_row(picks_rows[0])

                    fechas_pick_validas = [
                        pd.to_datetime(r.get("FechaHora"), errors="coerce")
                        for r in picks_rows
                        if pd.notna(pd.to_datetime(r.get("FechaHora"), errors="coerce"))
                    ]

                    fechas_put_validas = [
                        g["fh_put"]
                        for g in grupos_productivos + grupos_retorno
                        if pd.notna(g["fh_put"])
                    ]

                    if fechas_pick_validas:
                        fh_inicio = min(fechas_pick_validas)
                    else:
                        fh_inicio = pd.NaT

                    if fechas_put_validas:
                        fh_fin = max(fechas_put_validas)
                    else:
                        fh_fin = fh_inicio

                    lp_pick_set = {
                        _limpiar_texto(r.get("LP", ""))
                        for r in picks_rows
                        if _limpiar_texto(r.get("LP", ""))
                    }

                    lp_put_set = {
                        _limpiar_texto(r.get("LP", ""))
                        for r in puts_todos
                        if _limpiar_texto(r.get("LP", ""))
                    }

                    grupo_split_key = (
                        "PADRE_DESTINO|"
                        + "|".join(str(x) for x in pick_ids)
                        + "|"
                        + id_operador
                        + "|"
                        + item_pick
                        + "|"
                        + lote_pick
                        + "|"
                        + origen_pick
                        + "|"
                        + intermedia_pick
                    )

                    mov_padre = {
                        "Movimiento_ID": pd.NA,
                        "Pertenece_a_Pick_Padre": pd.NA,
                        "Tipo_Match": "MULTIDESTINO_PADRE",
                        "Area_Movimiento": area_movimiento,
                        "ID Operador": id_operador,
                        "Nombre": nombre,
                        "Nombre_norm": nombre_norm,
                        "Origen": origen_pick,
                        "Intermedia": intermedia_pick,
                        "Destino_Final": " | ".join([g["destino"] for g in grupos_productivos]),
                        "FH_inicio": fh_inicio,
                        "FH_fin": fh_fin,
                        "Cantidad_Pick": cantidad_pick_total,
                        "Cantidad_Put": cantidad_productiva,
                        "Cantidad": cantidad_productiva,
                        "Diferencia_Hijo": 0,
                        "Estado_Cantidad_Hijo": "OK",
                        "Cantidad_Pick_Padre_Total": cantidad_pick_total,
                        "Cantidad_Put_Hijos_Total": cantidad_productiva,
                        "Diferencia_Padre": round(cantidad_pick_total - cantidad_productiva - cantidad_retorno, 6),
                        "Estado_Cantidad_Padre": "OK",
                        "Items_Distintos": 1,
                        "Lotes_Distintos": 1,
                        "LPs_Distintos": len(lp_pick_set.union(lp_put_set)),
                        "Picks_Fragmentos": len(picks_rows),
                        "Puts_Fragmentos": len(puts_todos),
                        "LP": " | ".join(sorted(lp_pick_set.union(lp_put_set))),
                        "LP_Pick": " | ".join(sorted(lp_pick_set)),
                        "LP_Put": " | ".join(sorted(lp_put_set)),
                        "LP_Diferente": bool(lp_pick_set != lp_put_set),
                        "Tiene_LP_Cambiada": bool(lp_pick_set != lp_put_set),
                        "LPs_Pick_Distintas": len(lp_pick_set),
                        "LPs_Put_Distintas": len(lp_put_set),
                        "Minutos_Max_Pick_Put": round((fh_fin - fh_inicio).total_seconds() / 60, 2) if pd.notna(fh_inicio) and pd.notna(fh_fin) else 0,
                        "_row_ids_wms": pick_ids,
                        "_Grupo_Split_Key": grupo_split_key,
                        "_Orden_Split": 0,
                        "_Es_Padre_Split": True
                    }

                    movimientos_fase.append(mov_padre)

                    orden_hijo = 1

                    for g in grupos_productivos:
                        cantidad_hijo = round(g["cantidad"], 6)

                        picks_parciales = _crear_picks_parciales_desde_saldo(
                            picks_rows,
                            cantidad_hijo
                        )

                        if not picks_parciales:
                            continue

                        mov_hijo = _crear_movimiento_desde_match(
                            picks_rows=picks_parciales,
                            puts_rows=g["puts_rows"],
                            area_movimiento=area_movimiento,
                            forzar_lp_cambiada=permitir_lp_cambiada
                        )

                        if mov_hijo is None:
                            continue

                        mov_hijo["Tipo_Match"] = "MULTIDESTINO_HIJO"
                        mov_hijo["Cantidad_Pick_Padre_Total"] = cantidad_pick_total
                        mov_hijo["Cantidad_Put_Hijos_Total"] = cantidad_productiva
                        mov_hijo["_row_ids_wms"] = g["put_ids"]
                        mov_hijo["_Grupo_Split_Key"] = grupo_split_key
                        mov_hijo["_Orden_Split"] = orden_hijo
                        mov_hijo["_Es_Padre_Split"] = False

                        movimientos_fase.append(mov_hijo)
                        orden_hijo += 1

                    for rid in pick_ids:
                        row_ids_usados_fase.add(rid)

                    for g in grupos_productivos:
                        for rid in g.get("put_ids", []):
                            row_ids_usados_fase.add(rid)

                    for g in grupos_retorno:
                        for rid in g.get("put_ids", []):
                            row_ids_retorno_origen_global.add(rid)
                            row_ids_usados_fase.add(rid)

            return movimientos_fase, row_ids_usados_fase
        
        def _procesar_merge_multi_lp_picking(
            df_fase,
            permitir_lp_cambiada,
            row_ids_usados_globales
        ):
            """
            Recupera movimientos Picking donde varios PICK de distintas LP/orígenes
            consolidan en un solo PUT con otra LP.

            Caso típico:
            PICK 1+1+1+1+1 desde varios orígenes -> Intermedia
            PUT 5 desde Intermedia -> Destino
            """

            movimientos_fase = []
            row_ids_usados_fase = set()

            # Esta capa debe correr principalmente en fase LP cambiada.
            if not permitir_lp_cambiada:
                return movimientos_fase, row_ids_usados_fase

            if df_fase is None or df_fase.empty:
                return movimientos_fase, row_ids_usados_fase

            df_work = df_fase[
                ~df_fase["_row_id_wms"].isin(row_ids_usados_globales)
            ].copy()

            if df_work.empty:
                return movimientos_fase, row_ids_usados_fase

            for _, g_operador in df_work.groupby(
                ["Nombre_norm", "ID Operador", "Nombre"],
                dropna=False,
                sort=False
            ):
                g_operador = g_operador.sort_values([
                    "FechaHora",
                    "_row_id_wms"
                ]).copy()

                picks = g_operador[
                    g_operador["Description_norm"].isin(pick_descs)
                ].copy()

                puts = g_operador[
                    g_operador["Description_norm"].isin(put_descs)
                ].copy()

                if picks.empty or puts.empty:
                    continue

                picks = picks[
                    (~picks["_row_id_wms"].isin(row_ids_usados_globales)) &
                    (~picks["_row_id_wms"].isin(row_ids_usados_fase))
                ].copy()

                puts = puts[
                    (~puts["_row_id_wms"].isin(row_ids_usados_globales)) &
                    (~puts["_row_id_wms"].isin(row_ids_usados_fase))
                ].copy()

                if picks.empty or puts.empty:
                    continue

                group_pick_cols = [
                    "Nombre_norm",
                    "ID Operador",
                    "Nombre",
                    "_Familia_WMS",
                    "_Bloque_Flujo_Tiempo",
                    "Item Number",
                    "Número Lote",
                    "A ID Ubicación"
                ]

                picks["_row_dict"] = picks.apply(
                    lambda r: r.to_dict(),
                    axis=1
                )

                pick_batches = (
                    picks
                    .groupby(group_pick_cols, dropna=False, sort=False)
                    .agg(
                        Cantidad_Pick_Total=("Cantidad", "sum"),
                        FH_pick=("FechaHora", "min"),
                        Picks_Rows=("_row_dict", list)
                    )
                    .reset_index()
                    .sort_values("FH_pick")
                )

                for _, pick_batch in pick_batches.iterrows():
                    picks_rows = pick_batch.get("Picks_Rows", [])

                    if not isinstance(picks_rows, list) or not picks_rows:
                        continue

                    pick_ids = _row_ids_de_rows(picks_rows)

                    if not pick_ids:
                        continue

                    if any(
                        rid in row_ids_usados_globales or rid in row_ids_usados_fase
                        for rid in pick_ids
                    ):
                        continue

                    nombre_norm = _limpiar_texto(pick_batch.get("Nombre_norm", ""))
                    id_operador = _limpiar_texto(pick_batch.get("ID Operador", ""))
                    nombre = _limpiar_texto(pick_batch.get("Nombre", ""))
                    familia_pick = _limpiar_texto(pick_batch.get("_Familia_WMS", ""))
                    bloque_pick = pick_batch.get("_Bloque_Flujo_Tiempo", None)
                    item_pick = _limpiar_texto(pick_batch.get("Item Number", ""))
                    lote_pick = _limpiar_texto(pick_batch.get("Número Lote", ""))
                    intermedia_pick = _limpiar_texto(pick_batch.get("A ID Ubicación", ""))

                    if not intermedia_pick:
                        continue

                    cantidad_pick_total = round(
                        _to_float(pick_batch.get("Cantidad_Pick_Total", 0)),
                        6
                    )

                    if cantidad_pick_total <= 0:
                        continue

                    candidatos_put = puts[
                        (~puts["_row_id_wms"].isin(row_ids_usados_globales)) &
                        (~puts["_row_id_wms"].isin(row_ids_usados_fase)) &
                        (puts["Nombre_norm"].astype(str).str.strip() == nombre_norm) &
                        (puts["ID Operador"].astype(str).str.strip() == id_operador) &
                        (puts["Nombre"].astype(str).str.strip() == nombre) &
                        (puts["_Familia_WMS"].astype(str).str.strip() == familia_pick) &
                        (puts["_Bloque_Flujo_Tiempo"] == bloque_pick) &
                        (puts["Item Number"].astype(str).str.strip() == item_pick) &
                        (puts["Número Lote"].astype(str).str.strip() == lote_pick) &
                        (puts["Desde ID Ubicación"].astype(str).str.strip() == intermedia_pick)
                    ].copy()

                    if candidatos_put.empty:
                        continue

                    destinos = (
                        candidatos_put["A ID Ubicación"]
                        .astype(str)
                        .str.strip()
                        .replace("", pd.NA)
                        .dropna()
                        .unique()
                    )

                    # Esta capa solo resuelve consolidación a un destino.
                    if len(destinos) != 1:
                        continue

                    destino_final = _limpiar_texto(destinos[0])

                    if not destino_final:
                        continue

                    origenes_pick = [
                        _limpiar_texto(r.get("Desde ID Ubicación", ""))
                        for r in picks_rows
                        if _limpiar_texto(r.get("Desde ID Ubicación", ""))
                    ]

                    origenes_pick = list(dict.fromkeys(origenes_pick))

                    if not origenes_pick:
                        continue

                    # Si el destino vuelve a alguno de los orígenes, no resolver aquí.
                    if destino_final.upper() in {x.upper() for x in origenes_pick}:
                        continue

                    candidatos_put = candidatos_put.sort_values([
                        "FechaHora",
                        "_row_id_wms"
                    ]).copy()

                    puts_rows = [
                        r.to_dict()
                        for _, r in candidatos_put.iterrows()
                    ]

                    put_ids = _row_ids_de_rows(puts_rows)

                    if not put_ids:
                        continue

                    if any(
                        rid in row_ids_usados_globales or rid in row_ids_usados_fase
                        for rid in put_ids
                    ):
                        continue

                    cantidad_put_total = round(
                        sum(_to_float(r.get("Cantidad", 0)) for r in puts_rows),
                        6
                    )

                    if cantidad_pick_total != cantidad_put_total:
                        continue

                    resumen_pick = _resumen_item_lote(
                        picks_rows,
                        incluir_lp=False
                    )

                    resumen_put = _resumen_item_lote(
                        puts_rows,
                        incluir_lp=False
                    )

                    if resumen_pick != resumen_put:
                        continue

                    lp_pick_set = {
                        _limpiar_texto(r.get("LP", ""))
                        for r in picks_rows
                        if _limpiar_texto(r.get("LP", ""))
                    }

                    lp_put_set = {
                        _limpiar_texto(r.get("LP", ""))
                        for r in puts_rows
                        if _limpiar_texto(r.get("LP", ""))
                    }

                    # Debe existir cambio/consolidación de LP.
                    if not lp_pick_set or not lp_put_set:
                        continue

                    if lp_pick_set == lp_put_set:
                        continue

                    if not _flujo_en_tiempo(picks_rows, puts_rows):
                        continue

                    area_movimiento = _area_bodega_desde_row(picks_rows[0])

                    mov_abierto = {
                        "Nombre_norm": nombre_norm,
                        "ID Operador": id_operador,
                        "Nombre": nombre,
                        "origen": " | ".join(origenes_pick),
                        "intermedia": intermedia_pick,
                        "picks": picks_rows,
                        "puts": puts_rows,
                        "area_movimiento": area_movimiento
                    }

                    mov_creado = _crear_movimiento_desde_abierto(
                        mov_abierto,
                        area_movimiento
                    )

                    if mov_creado is None:
                        continue

                    row_ids_mov = _row_ids_de_rows(
                        picks_rows + puts_rows
                    )

                    mov_creado["_row_ids_wms"] = row_ids_mov
                    mov_creado["Tipo_Match"] = "MERGE_MULTI_LP"
                    mov_creado["Tiene_LP_Cambiada"] = True
                    mov_creado["LP_Diferente"] = True

                    movimientos_fase.append(mov_creado)

                    for rid in row_ids_mov:
                        row_ids_usados_fase.add(rid)

            return movimientos_fase, row_ids_usados_fase
        
        def _procesar_retorno_parcial_con_productivo(
            df_fase,
            permitir_lp_cambiada,
            row_ids_usados_globales
        ):
            """
            Detecta casos donde un PICK se divide en:
            - una parte que retorna al origen
            - una parte que sí va a destino productivo

            Ejemplo:
            PICK 117: Origen -> Intermedia
            PUT 86: Intermedia -> Origen        = retorno no productivo
            PUT 31: Intermedia -> Destino real  = movimiento productivo

            Resultado:
            - Los PUT retorno se marcan como NO_PRODUCTIVO_RETORNO_ORIGEN.
            - Los PUT productivos se convierten en movimiento real.
            - El PICK NO se marca como retorno completo.
            """

            movimientos_fase = []
            row_ids_usados_fase = set()

            if df_fase is None or df_fase.empty:
                return movimientos_fase, row_ids_usados_fase

            df_work = df_fase[
                ~df_fase["_row_id_wms"].isin(row_ids_usados_globales)
            ].copy()

            if df_work.empty:
                return movimientos_fase, row_ids_usados_fase
            
            # ============================================================
            # FILTRO ESTRICTO PARA RETORNO PARCIAL PRODUCTIVO
            # Este caso solo aplica para Move(Pick) / Move(Put) tipo STORAGE.
            # Evita mezclar Picking(pick) / Picking(put) dentro del mismo flujo.
            # ============================================================

            df_work["_Desc_RP"] = (
                df_work["Description_norm"]
                .astype(str)
                .str.strip()
                .str.lower()
            )

            df_work["_Tipo_RP"] = (
                df_work["Tipo"]
                .apply(_limpiar_texto)
                .astype(str)
                .str.strip()
                .str.upper()
            )

            def _clase_retorno_parcial(desc):
                desc = _limpiar_texto(desc).lower()

                if desc in ["move (pick)", "move (put)"]:
                    return "MOVE"

                if desc in ["move trailer (pick)", "move trailer (put)"]:
                    return "TRAILER"

                return "NO_APLICA"

            df_work["_Clase_RP"] = df_work["_Desc_RP"].apply(
                _clase_retorno_parcial
            )

            df_work = df_work[
                (df_work["_Clase_RP"] != "NO_APLICA") &
                (df_work["_Tipo_RP"] == "STORAGE")
            ].copy()

            if df_work.empty:
                return movimientos_fase, row_ids_usados_fase
            

            for _, g_operador in df_work.groupby(
                ["Nombre_norm", "ID Operador", "Nombre"],
                dropna=False,
                sort=False
            ):
                g_operador = g_operador.sort_values([
                    "FechaHora",
                    "_row_id_wms"
                ]).copy()

                picks = g_operador[
                    g_operador["_Desc_RP"].isin([
                        "move (pick)",
                        "move trailer (pick)"
                    ])
                ].copy()

                puts = g_operador[
                    g_operador["_Desc_RP"].isin([
                        "move (put)",
                        "move trailer (put)"
                    ])
                ].copy()

                if picks.empty or puts.empty:
                    continue

                picks = picks[
                    (~picks["_row_id_wms"].isin(row_ids_usados_globales)) &
                    (~picks["_row_id_wms"].isin(row_ids_usados_fase))
                ].copy()

                puts = puts[
                    (~puts["_row_id_wms"].isin(row_ids_usados_globales)) &
                    (~puts["_row_id_wms"].isin(row_ids_usados_fase))
                ].copy()

                if picks.empty or puts.empty:
                    continue

                if permitir_lp_cambiada:
                    group_pick_cols = [
                        "Nombre_norm",
                        "ID Operador",
                        "Nombre",
                        "_Familia_WMS",
                        "_Bloque_Flujo_Tiempo",
                        "Item Number",
                        "Número Lote",
                        "Desde ID Ubicación",
                        "A ID Ubicación"
                    ]
                else:
                    group_pick_cols = [
                        "Nombre_norm",
                        "ID Operador",
                        "Nombre",
                        "_Familia_WMS",
                        "_Bloque_Flujo_Tiempo",                        
                        "_Clase_RP",
                        "_Tipo_RP",
                        "Item Number",
                        "Número Lote",
                        "Desde ID Ubicación",
                        "A ID Ubicación",
                        "LP"
                    ]

                picks["_row_dict"] = picks.apply(
                    lambda r: r.to_dict(),
                    axis=1
                )

                pick_batches = (
                    picks
                    .groupby(group_pick_cols, dropna=False, sort=False)
                    .agg(
                        Cantidad_Pick_Total=("Cantidad", "sum"),
                        FH_pick=("FechaHora", "min"),
                        Picks_Rows=("_row_dict", list)
                    )
                    .reset_index()
                    .sort_values("FH_pick")
                )

                for _, pick_batch in pick_batches.iterrows():
                    picks_rows = pick_batch.get("Picks_Rows", [])

                    if not isinstance(picks_rows, list) or not picks_rows:
                        continue

                    pick_ids = _row_ids_de_rows(picks_rows)

                    if not pick_ids:
                        continue

                    if any(
                        rid in row_ids_usados_globales or rid in row_ids_usados_fase
                        for rid in pick_ids
                    ):
                        continue

                    nombre_norm = _limpiar_texto(pick_batch.get("Nombre_norm", ""))
                    id_operador = _limpiar_texto(pick_batch.get("ID Operador", ""))
                    nombre = _limpiar_texto(pick_batch.get("Nombre", ""))
                    familia_pick = _limpiar_texto(pick_batch.get("_Familia_WMS", ""))
                    bloque_pick = pick_batch.get("_Bloque_Flujo_Tiempo", None)
                    clase_pick = _limpiar_texto(pick_batch.get("_Clase_RP", ""))
                    tipo_pick = _limpiar_texto(pick_batch.get("_Tipo_RP", ""))
                    item_pick = _limpiar_texto(pick_batch.get("Item Number", ""))
                    lote_pick = _limpiar_texto(pick_batch.get("Número Lote", ""))
                    origen_pick = _limpiar_texto(pick_batch.get("Desde ID Ubicación", ""))
                    intermedia_pick = _limpiar_texto(pick_batch.get("A ID Ubicación", ""))

                    cantidad_pick_total = round(
                        _to_float(pick_batch.get("Cantidad_Pick_Total", 0)),
                        6
                    )

                    if cantidad_pick_total <= 0:
                        continue

                    if not origen_pick or not intermedia_pick:
                        continue

                    candidatos_put = puts[
                        (~puts["_row_id_wms"].isin(row_ids_usados_globales)) &
                        (~puts["_row_id_wms"].isin(row_ids_usados_fase)) &
                        (puts["Nombre_norm"].astype(str).str.strip() == nombre_norm) &
                        (puts["ID Operador"].astype(str).str.strip() == id_operador) &
                        (puts["Nombre"].astype(str).str.strip() == nombre) &
                        (puts["_Familia_WMS"].astype(str).str.strip() == familia_pick) &
                        (puts["_Bloque_Flujo_Tiempo"] == bloque_pick) &                        
                        (puts["_Clase_RP"].astype(str).str.strip() == clase_pick) &
                        (puts["_Tipo_RP"].astype(str).str.strip() == tipo_pick) &
                        (puts["Item Number"].astype(str).str.strip() == item_pick) &
                        (puts["Número Lote"].astype(str).str.strip() == lote_pick) &
                        (puts["Desde ID Ubicación"].astype(str).str.strip() == intermedia_pick)
                    ].copy()

                    if candidatos_put.empty:
                        continue

                    if not permitir_lp_cambiada:
                        lp_pick = _limpiar_texto(pick_batch.get("LP", ""))

                        candidatos_put = candidatos_put[
                            candidatos_put["LP"].astype(str).str.strip() == lp_pick
                        ].copy()

                    if candidatos_put.empty:
                        continue

                    candidatos_put["_row_dict"] = candidatos_put.apply(
                        lambda r: r.to_dict(),
                        axis=1
                    )

                    put_groups = (
                        candidatos_put
                        .groupby(["A ID Ubicación"], dropna=False, sort=False)
                        .agg(
                            Cantidad_Put_Total=("Cantidad", "sum"),
                            FH_put=("FechaHora", "max"),
                            Puts_Rows=("_row_dict", list)
                        )
                        .reset_index()
                        .sort_values("FH_put")
                    )

                    grupos_productivos = []
                    grupos_retorno = []

                    for _, put_group in put_groups.iterrows():
                        destino = _limpiar_texto(put_group.get("A ID Ubicación", ""))
                        cantidad_put = round(
                            _to_float(put_group.get("Cantidad_Put_Total", 0)),
                            6
                        )
                        puts_rows = put_group.get("Puts_Rows", [])
                        fh_put = pd.to_datetime(
                            put_group.get("FH_put"),
                            errors="coerce"
                        )

                        if not destino or cantidad_put <= 0:
                            continue

                        if not isinstance(puts_rows, list) or not puts_rows:
                            continue

                        put_ids = _row_ids_de_rows(puts_rows)

                        if not put_ids:
                            continue

                        if any(
                            rid in row_ids_usados_globales or rid in row_ids_usados_fase
                            for rid in put_ids
                        ):
                            continue

                        grupo = {
                            "destino": destino,
                            "cantidad": cantidad_put,
                            "puts_rows": puts_rows,
                            "put_ids": put_ids,
                            "fh_put": fh_put
                        }

                        if destino.upper() == origen_pick.upper():
                            grupos_retorno.append(grupo)
                        else:
                            grupos_productivos.append(grupo)

                    # Debe existir retorno y también productividad real.
                    if not grupos_retorno or not grupos_productivos:
                        continue

                    cantidad_retorno = round(
                        sum(g["cantidad"] for g in grupos_retorno),
                        6
                    )

                    cantidad_productiva = round(
                        sum(g["cantidad"] for g in grupos_productivos),
                        6
                    )

                    if round(cantidad_retorno + cantidad_productiva, 6) != cantidad_pick_total:
                        continue

                    puts_todos = []

                    for g in grupos_retorno:
                        puts_todos.extend(g["puts_rows"])

                    for g in grupos_productivos:
                        puts_todos.extend(g["puts_rows"])

                    resumen_pick = _resumen_item_lote(
                        picks_rows,
                        incluir_lp=False
                    )

                    resumen_put = _resumen_item_lote(
                        puts_todos,
                        incluir_lp=False
                    )

                    if resumen_pick != resumen_put:
                        continue

                    if not permitir_lp_cambiada:
                        resumen_pick_lp = _resumen_item_lote(
                            picks_rows,
                            incluir_lp=True
                        )

                        resumen_put_lp = _resumen_item_lote(
                            puts_todos,
                            incluir_lp=True
                        )

                        if resumen_pick_lp != resumen_put_lp:
                            continue

                    # El retorno debe ser cronológicamente válido.
                    puts_retorno_todos = []

                    for g in grupos_retorno:
                        puts_retorno_todos.extend(g["puts_rows"])

                    if not _retorno_cronologicamente_valido(
                        picks_rows,
                        puts_retorno_todos
                    ):
                        continue

                    if not _flujo_en_tiempo(picks_rows, puts_todos):
                        continue

                    area_movimiento = _area_bodega_desde_row(picks_rows[0])

                    # Crear movimiento productivo por cada destino productivo.
                    for g in grupos_productivos:
                        cantidad_hijo = round(g["cantidad"], 6)

                        picks_parciales = _crear_picks_parciales_desde_saldo(
                            picks_rows,
                            cantidad_hijo
                        )

                        if not picks_parciales:
                            continue

                        mov_creado = _crear_movimiento_desde_match(
                            picks_rows=picks_parciales,
                            puts_rows=g["puts_rows"],
                            area_movimiento=area_movimiento,
                            forzar_lp_cambiada=permitir_lp_cambiada
                        )

                        if mov_creado is None:
                            continue

                        row_ids_mov = _row_ids_de_rows(
                            picks_parciales + g["puts_rows"]
                        )

                        mov_creado["_row_ids_wms"] = row_ids_mov
                        mov_creado["Tipo_Match"] = "RETORNO_PARCIAL_PRODUCTIVO"

                        movimientos_fase.append(mov_creado)

                        for rid in row_ids_mov:
                            row_ids_usados_fase.add(rid)

                    # Marcar SOLO los PUT de retorno como no productivos.
                    # El PICK no debe marcarse como retorno completo.
                    for g in grupos_retorno:
                        for rid in g.get("put_ids", []):
                            try:
                                rid = int(rid)
                                row_ids_retorno_origen_global.add(rid)
                                row_ids_usados_fase.add(rid)
                            except Exception:
                                pass

                    # Bloquear el PICK para que no sea tomado luego como retorno total.
                    for rid in pick_ids:
                        try:
                            row_ids_usados_fase.add(int(rid))
                        except Exception:
                            pass

            return movimientos_fase, row_ids_usados_fase
        
        def _procesar_picking_lp_cambiada_con_move_retorno(
            df_fase,
            row_ids_usados_globales
        ):
            """
            Detecta el patrón:
        
            1) Movimiento productivo real:
               Picking(pick): almacenamiento -> intermedia
               Picking(put):  intermedia -> destino
               con cambio de LP.
        
            2) Movimiento Move posterior no productivo:
               Move(Pick): destino -> intermedia
               Move(Put):  intermedia -> mismo destino
        
            Resultado:
            - Picking queda como MOV_REAL.
            - Move queda como RETORNO_ORIGEN / NO_PRODUCTIVO_RETORNO_ORIGEN.
            """
        
            movimientos_fase = []
            row_ids_usados_fase = set()
        
            if df_fase is None or df_fase.empty:
                return movimientos_fase, row_ids_usados_fase
        
            df_work = df_fase[
                ~df_fase["_row_id_wms"].isin(row_ids_usados_globales)
            ].copy()
        
            if df_work.empty:
                return movimientos_fase, row_ids_usados_fase
        
            df_work["FechaHora"] = pd.to_datetime(
                df_work["FechaHora"],
                errors="coerce"
            )
        
            df_work = df_work[df_work["FechaHora"].notna()].copy()
        
            if df_work.empty:
                return movimientos_fase, row_ids_usados_fase
        
            for _, g_operador in df_work.groupby(
                ["Nombre_norm", "ID Operador", "Nombre"],
                dropna=False,
                sort=False
            ):
                g_operador = g_operador.sort_values([
                    "FechaHora",
                    "_row_id_wms"
                ]).copy()
        
                picking_picks = g_operador[
                    g_operador["Description_norm"].astype(str).str.strip().eq("picking (pick)")
                ].copy()
        
                picking_puts = g_operador[
                    g_operador["Description_norm"].astype(str).str.strip().eq("picking (put)")
                ].copy()
        
                if picking_picks.empty or picking_puts.empty:
                    continue
        
                group_pick_cols = [
                    "Nombre_norm",
                    "ID Operador",
                    "Nombre",
                    "_Familia_WMS",
                    "_Bloque_Flujo_Tiempo",
                    "Item Number",
                    "Número Lote",
                    "Desde ID Ubicación",
                    "A ID Ubicación"
                ]
        
                picking_picks["_row_dict"] = picking_picks.apply(
                    lambda r: r.to_dict(),
                    axis=1
                )
        
                pick_batches = (
                    picking_picks
                    .groupby(group_pick_cols, dropna=False, sort=False)
                    .agg(
                        Cantidad_Pick_Total=("Cantidad", "sum"),
                        FH_pick=("FechaHora", "min"),
                        Picks_Rows=("_row_dict", list)
                    )
                    .reset_index()
                    .sort_values("FH_pick")
                )
        
                for _, pick_batch in pick_batches.iterrows():
                    picks_rows = pick_batch.get("Picks_Rows", [])
        
                    if not isinstance(picks_rows, list) or not picks_rows:
                        continue
        
                    pick_ids = _row_ids_de_rows(picks_rows)
        
                    if not pick_ids:
                        continue
        
                    if any(
                        rid in row_ids_usados_globales or rid in row_ids_usados_fase
                        for rid in pick_ids
                    ):
                        continue
        
                    nombre_norm = _limpiar_texto(pick_batch.get("Nombre_norm", ""))
                    id_operador = _limpiar_texto(pick_batch.get("ID Operador", ""))
                    nombre = _limpiar_texto(pick_batch.get("Nombre", ""))
                    familia_pick = _limpiar_texto(pick_batch.get("_Familia_WMS", ""))
                    bloque_pick = pick_batch.get("_Bloque_Flujo_Tiempo", None)
                    item_pick = _limpiar_texto(pick_batch.get("Item Number", ""))
                    lote_pick = _limpiar_texto(pick_batch.get("Número Lote", ""))
                    origen_pick = _limpiar_texto(pick_batch.get("Desde ID Ubicación", ""))
                    intermedia_pick = _limpiar_texto(pick_batch.get("A ID Ubicación", ""))
        
                    if not origen_pick or not intermedia_pick:
                        continue
        
                    cantidad_pick_total = round(
                        _to_float(pick_batch.get("Cantidad_Pick_Total", 0)),
                        6
                    )
        
                    if cantidad_pick_total <= 0:
                        continue
        
                    candidatos_put = picking_puts[
                        (~picking_puts["_row_id_wms"].isin(row_ids_usados_globales)) &
                        (~picking_puts["_row_id_wms"].isin(row_ids_usados_fase)) &
                        (picking_puts["Nombre_norm"].astype(str).str.strip() == nombre_norm) &
                        (picking_puts["ID Operador"].astype(str).str.strip() == id_operador) &
                        (picking_puts["Nombre"].astype(str).str.strip() == nombre) &
                        (picking_puts["_Familia_WMS"].astype(str).str.strip() == familia_pick) &
                        (picking_puts["_Bloque_Flujo_Tiempo"] == bloque_pick) &
                        (picking_puts["Item Number"].astype(str).str.strip() == item_pick) &
                        (picking_puts["Número Lote"].astype(str).str.strip() == lote_pick) &
                        (picking_puts["Desde ID Ubicación"].astype(str).str.strip() == intermedia_pick)
                    ].copy()
        
                    if candidatos_put.empty:
                        continue
        
                    candidatos_put = candidatos_put.sort_values([
                        "FechaHora",
                        "_row_id_wms"
                    ]).copy()
        
                    destinos_put = (
                        candidatos_put["A ID Ubicación"]
                        .astype(str)
                        .str.strip()
                        .replace("", pd.NA)
                        .dropna()
                        .unique()
                    )
        
                    if len(destinos_put) != 1:
                        continue
        
                    destino_productivo = _limpiar_texto(destinos_put[0])
        
                    if not destino_productivo:
                        continue
        
                    if destino_productivo.upper() == origen_pick.upper():
                        continue
        
                    puts_rows = [
                        r.to_dict()
                        for _, r in candidatos_put.iterrows()
                    ]
        
                    put_ids = _row_ids_de_rows(puts_rows)
        
                    if not put_ids:
                        continue
        
                    cantidad_put_total = round(
                        sum(_to_float(r.get("Cantidad", 0)) for r in puts_rows),
                        6
                    )
        
                    if cantidad_put_total != cantidad_pick_total:
                        continue
        
                    resumen_pick = _resumen_item_lote(
                        picks_rows,
                        incluir_lp=False
                    )
        
                    resumen_put = _resumen_item_lote(
                        puts_rows,
                        incluir_lp=False
                    )
        
                    if resumen_pick != resumen_put:
                        continue
        
                    lp_pick_set = {
                        _limpiar_texto(r.get("LP", ""))
                        for r in picks_rows
                        if _limpiar_texto(r.get("LP", ""))
                    }
        
                    lp_put_set = {
                        _limpiar_texto(r.get("LP", ""))
                        for r in puts_rows
                        if _limpiar_texto(r.get("LP", ""))
                    }
        
                    if not lp_pick_set or not lp_put_set:
                        continue
        
                    # Debe existir cambio de LP.
                    if lp_pick_set == lp_put_set:
                        continue
        
                    if not _flujo_en_tiempo(picks_rows, puts_rows):
                        continue
        
                    area_movimiento = _area_bodega_desde_row(picks_rows[0])
        
                    mov_creado = _crear_movimiento_desde_match(
                        picks_rows=picks_rows,
                        puts_rows=puts_rows,
                        area_movimiento=area_movimiento,
                        forzar_lp_cambiada=True
                    )
        
                    if mov_creado is None:
                        continue
        
                    row_ids_mov = _row_ids_de_rows(
                        picks_rows + puts_rows
                    )
        
                    mov_creado["_row_ids_wms"] = row_ids_mov
                    mov_creado["Tipo_Match"] = "PICKING_LP_CAMBIADA"
                    mov_creado["LP_Diferente"] = True
                    mov_creado["Tiene_LP_Cambiada"] = True
        
                    movimientos_fase.append(mov_creado)
        
                    for rid in row_ids_mov:
                        row_ids_usados_fase.add(rid)
        
                    # ====================================================
                    # Buscar Move(Pick)/Move(Put) que retornan desde
                    # el destino productivo hacia la misma intermedia
                    # y vuelven al destino productivo.
                    # ====================================================
        
                    lp_put_principal = ""
        
                    if lp_put_set:
                        lp_put_principal = sorted(lp_put_set)[0]
        
                    moves_relacionados = g_operador[
                        (~g_operador["_row_id_wms"].isin(row_ids_usados_globales)) &
                        (~g_operador["_row_id_wms"].isin(row_ids_usados_fase)) &
                        (g_operador["Description_norm"].astype(str).str.strip().isin([
                            "move (pick)",
                            "move (put)"
                        ])) &
                        (g_operador["Item Number"].astype(str).str.strip() == item_pick) &
                        (g_operador["Número Lote"].astype(str).str.strip() == lote_pick) &
                        (g_operador["LP"].astype(str).str.strip() == lp_put_principal)
                    ].copy()
        
                    if moves_relacionados.empty:
                        continue
        
                    moves_relacionados["Cantidad_num_tmp"] = (
                        moves_relacionados["Cantidad"].apply(_to_float)
                    )
        
                    move_picks_retorno = moves_relacionados[
                        (moves_relacionados["Description_norm"].astype(str).str.strip() == "move (pick)") &
                        (moves_relacionados["Desde ID Ubicación"].astype(str).str.strip() == destino_productivo) &
                        (moves_relacionados["A ID Ubicación"].astype(str).str.strip() == intermedia_pick) &
                        (moves_relacionados["Cantidad_num_tmp"].round(6) == cantidad_pick_total)
                    ].copy()
        
                    move_puts_retorno = moves_relacionados[
                        (moves_relacionados["Description_norm"].astype(str).str.strip() == "move (put)") &
                        (moves_relacionados["Desde ID Ubicación"].astype(str).str.strip() == intermedia_pick) &
                        (moves_relacionados["A ID Ubicación"].astype(str).str.strip() == destino_productivo) &
                        (moves_relacionados["Cantidad_num_tmp"].round(6) == cantidad_put_total)
                    ].copy()
        
                    if move_picks_retorno.empty or move_puts_retorno.empty:
                        continue
        
                    move_picks_rows = [
                        r.to_dict()
                        for _, r in move_picks_retorno.iterrows()
                    ]
        
                    move_puts_rows = [
                        r.to_dict()
                        for _, r in move_puts_retorno.iterrows()
                    ]
        
                    # Validar que el retorno Move también sea cronológicamente válido.
                    if not _retorno_cronologicamente_valido(
                        move_picks_rows,
                        move_puts_rows
                    ):
                        continue
        
                    retorno_ids = _row_ids_de_rows(
                        move_picks_rows + move_puts_rows
                    )
        
                    for rid in retorno_ids:
                        try:
                            rid = int(rid)
                            row_ids_retorno_origen_global.add(rid)
                            row_ids_usados_fase.add(rid)
                        except Exception:
                            pass
        
            return movimientos_fase, row_ids_usados_fase

        
                                    
        
        
        def _procesar_fase(df_fase, permitir_lp_cambiada, row_ids_usados_globales):
            """
            Procesa una fase completa.

            Orden:
            1. Match fuerte por Tipo/Tarea WMS.
            2. Match por demanda PUT agrupada.
            3. Split many-to-many por saldo.
            4. Capa final por origen y retorno.
            """
            
            # ====================================================
            # 0. Capa prioritaria: padre con hijos por destino
            # Evita que el motor separe en varios movimientos
            # lo que pertenece al mismo flujo padre.
            # ====================================================

            movimientos_padre_destino, usados_padre_destino = _procesar_padre_hijos_por_destino(
                df_fase=df_fase,
                permitir_lp_cambiada=permitir_lp_cambiada,
                row_ids_usados_globales=row_ids_usados_globales
            )

            row_ids_ocupados = set(row_ids_usados_globales)
            row_ids_ocupados.update(usados_padre_destino)
            
            # ====================================================
            # 0.1 Picking LP cambiada + Move retorno posterior
            # ====================================================
            
            movimientos_picking_lp_retorno, usados_picking_lp_retorno = _procesar_picking_lp_cambiada_con_move_retorno(
                df_fase=df_fase,
                row_ids_usados_globales=row_ids_ocupados
            )
            
            row_ids_ocupados.update(usados_picking_lp_retorno)
            
            # ====================================================
            # 0.1 Retorno parcial con movimiento productivo
            # Ejemplo:
            # PICK 117
            # PUT 86 retorna al origen
            # PUT 31 va a destino productivo
            # ====================================================

            movimientos_retorno_parcial, usados_retorno_parcial = _procesar_retorno_parcial_con_productivo(
                df_fase=df_fase,
                permitir_lp_cambiada=permitir_lp_cambiada,
                row_ids_usados_globales=row_ids_ocupados
            )

            row_ids_ocupados.update(usados_retorno_parcial)
            
            
            # ====================================================
            # 0.1 Capa Picking MERGE multi-LP
            # Varios PICK de distintas LP/orígenes consolidan
            # en un solo PUT con otra LP.
            # ====================================================

            movimientos_merge_multi_lp, usados_merge_multi_lp = _procesar_merge_multi_lp_picking(
                df_fase=df_fase,
                permitir_lp_cambiada=permitir_lp_cambiada,
                row_ids_usados_globales=row_ids_ocupados
            )

            row_ids_ocupados.update(usados_merge_multi_lp)
            
            

            # ====================================================
            # 1. Match fuerte por Tipo WMS
            # ====================================================

            movimientos_tipo, usados_tipo = _procesar_por_put_agrupado(
                df_fase=df_fase,
                permitir_lp_cambiada=permitir_lp_cambiada,
                row_ids_usados_globales=row_ids_ocupados,
                usar_tipo_fuerte=True
            )

            row_ids_ocupados.update(usados_tipo)

            # ====================================================
            # 2. Match por demanda PUT agrupada
            # ====================================================

            movimientos_demanda, usados_demanda = _procesar_por_put_agrupado(
                df_fase=df_fase,
                permitir_lp_cambiada=permitir_lp_cambiada,
                row_ids_usados_globales=row_ids_ocupados,
                usar_tipo_fuerte=False
            )

            row_ids_ocupados.update(usados_demanda)

            # ====================================================
            # 3. Split many-to-many por saldo
            # ====================================================

            movimientos_m2m, usados_m2m = _procesar_split_m2m(
                df_fase=df_fase,
                permitir_lp_cambiada=permitir_lp_cambiada,
                row_ids_usados_globales=row_ids_ocupados
            )

            row_ids_ocupados.update(usados_m2m)
            
            movimientos_balance, usados_balance = _procesar_balance_secuencial(
                df_fase=df_fase,
                row_ids_usados_globales=row_ids_ocupados
            )
            
            row_ids_ocupados.update(usados_balance)

            # ====================================================
            # 4. Capa final por origen y retorno
            # ====================================================

            movimientos_origen, usados_origen = _procesar_por_origen_y_retorno(
                df_fase=df_fase,
                permitir_lp_cambiada=permitir_lp_cambiada,
                row_ids_usados_globales=row_ids_ocupados
            )
            
            row_ids_ocupados.update(usados_origen)

            movimientos_lp_mixta, usados_lp_mixta = _procesar_balance_lp_mixta_por_bloque(
                df_fase=df_fase,
                permitir_lp_cambiada=permitir_lp_cambiada,
                row_ids_usados_globales=row_ids_ocupados
            )

            # ====================================================
            # Consolidar usados
            # ====================================================

            usados_total = set()
            usados_total.update(usados_padre_destino)
            usados_total.update(usados_picking_lp_retorno)
            usados_total.update(usados_retorno_parcial)
            usados_total.update(usados_merge_multi_lp)
            usados_total.update(usados_tipo)
            usados_total.update(usados_demanda)
            usados_total.update(usados_m2m)
            usados_total.update(usados_origen)
            usados_total.update(usados_balance)
            usados_total.update(usados_lp_mixta)
            
            # ====================================================
            # Consolidar movimientos
            # ====================================================
            
            movimientos_total = (
                movimientos_padre_destino
                + movimientos_picking_lp_retorno
                + movimientos_retorno_parcial
                + movimientos_merge_multi_lp
                + movimientos_tipo
                + movimientos_demanda
                + movimientos_m2m
                + movimientos_origen
                + movimientos_balance
                + movimientos_lp_mixta
            )
                        
            return movimientos_total, usados_total
        
        
        
        
        def _corregir_fragmentacion_unidestino_subbloque_local(
            movimientos_df,
            df_calc_base,
            row_ids_retorno_origen_global
        ):
            """
            Corrige fragmentaciones unidestino por sub-bloque local.

            Caso esperado:
            PICK 1 + 2 + 2 = PUT 1 + 2 + 2
            mismo operador, item, lote, LP, origen, intermedia y destino.
            """

            if df_calc_base is None or df_calc_base.empty:
                return movimientos_df

            df_work = df_calc_base.copy()

            if "_row_id_wms" not in df_work.columns:
                return movimientos_df

            df_work["FechaHora"] = pd.to_datetime(
                df_work["FechaHora"],
                errors="coerce"
            )

            df_work = df_work[df_work["FechaHora"].notna()].copy()

            if df_work.empty:
                return movimientos_df

            df_work = df_work[
                df_work["Description_norm"].isin(pick_descs.union(put_descs))
            ].copy()

            if df_work.empty:
                return movimientos_df

            MAX_GAP_SUBBLOQUE_MIN = 2
            df_work["_SubBloque_Local"] = 0

            group_cols_local = [
                "Nombre_norm",
                "ID Operador",
                "_Familia_WMS",
                "Item Number",
                "Número Lote",
                "LP"
            ]

            for _, idxs in df_work.groupby(
                group_cols_local,
                dropna=False
            ).groups.items():

                g = df_work.loc[list(idxs)].sort_values([
                    "FechaHora",
                    "_row_id_wms"
                ]).copy()

                dif_min = (
                    g["FechaHora"]
                    .diff()
                    .dt.total_seconds()
                    .div(60)
                    .fillna(0)
                )

                sub_bloque = (dif_min > MAX_GAP_SUBBLOQUE_MIN).cumsum()

                df_work.loc[g.index, "_SubBloque_Local"] = (
                    sub_bloque.astype(int).values
                )

            if movimientos_df is not None and not movimientos_df.empty:
                movimientos_existentes = movimientos_df.to_dict("records")
            else:
                movimientos_existentes = []

            movimientos_nuevos = []
            row_ids_corregidos = set()

            group_cols_revision = [
                "Nombre_norm",
                "ID Operador",
                "Nombre",
                "_Familia_WMS",
                "_SubBloque_Local",
                "Item Number",
                "Número Lote",
                "LP"
            ]

            for _, g_local in df_work.groupby(
                group_cols_revision,
                dropna=False,
                sort=False
            ):
                g_local = g_local.sort_values([
                    "FechaHora",
                    "_row_id_wms"
                ]).copy()

                picks_local = g_local[
                    g_local["Description_norm"].isin(pick_descs)
                ].copy()

                puts_local = g_local[
                    g_local["Description_norm"].isin(put_descs)
                ].copy()

                if picks_local.empty or puts_local.empty:
                    continue

                for (origen, intermedia), g_pick in picks_local.groupby(
                    [
                        picks_local["Desde ID Ubicación"].astype(str).str.strip(),
                        picks_local["A ID Ubicación"].astype(str).str.strip()
                    ],
                    sort=False
                ):
                    origen = _limpiar_texto(origen)
                    intermedia = _limpiar_texto(intermedia)

                    if not origen or not intermedia:
                        continue

                    candidate_puts = puts_local[
                        puts_local["Desde ID Ubicación"]
                        .astype(str)
                        .str.strip()
                        == intermedia
                    ].copy()

                    if candidate_puts.empty:
                        continue

                    destinos = (
                        candidate_puts["A ID Ubicación"]
                        .astype(str)
                        .str.strip()
                        .replace("", pd.NA)
                        .dropna()
                        .unique()
                    )

                    if len(destinos) != 1:
                        continue

                    destino = _limpiar_texto(destinos[0])

                    if not destino:
                        continue

                    if destino.upper() == origen.upper():
                        continue

                    picks_rows = [
                        r.to_dict()
                        for _, r in g_pick.iterrows()
                    ]

                    puts_rows = [
                        r.to_dict()
                        for _, r in candidate_puts.iterrows()
                    ]

                    row_ids_global = _row_ids_de_rows(
                        picks_rows + puts_rows
                    )

                    if not row_ids_global:
                        continue

                    if any(
                        int(rid) in row_ids_corregidos
                        for rid in row_ids_global
                    ):
                        continue

                    total_pick = round(
                        sum(_to_float(r.get("Cantidad", 0)) for r in picks_rows),
                        6
                    )

                    total_put = round(
                        sum(_to_float(r.get("Cantidad", 0)) for r in puts_rows),
                        6
                    )

                    if total_pick != total_put:
                        continue

                    resumen_pick = _resumen_item_lote(
                        picks_rows,
                        incluir_lp=False
                    )

                    resumen_put = _resumen_item_lote(
                        puts_rows,
                        incluir_lp=False
                    )

                    if resumen_pick != resumen_put:
                        continue

                    resumen_pick_lp = _resumen_item_lote(
                        picks_rows,
                        incluir_lp=True
                    )

                    resumen_put_lp = _resumen_item_lote(
                        puts_rows,
                        incluir_lp=True
                    )

                    if resumen_pick_lp != resumen_put_lp:
                        continue

                    if not _flujo_en_tiempo(picks_rows, puts_rows):
                        continue

                    mov_global = _crear_movimiento_desde_match(
                        picks_rows=picks_rows,
                        puts_rows=puts_rows,
                        area_movimiento=_area_bodega_desde_row(picks_rows[0]),
                        forzar_lp_cambiada=False
                    )

                    if mov_global is None:
                        continue

                    mov_global["_row_ids_wms"] = row_ids_global
                    mov_global["Tipo_Match"] = "GLOBAL_MATCH_CORREGIDO"
                    mov_global["_Grupo_Split_Key"] = ""
                    mov_global["_Orden_Split"] = 999
                    mov_global["_Es_Padre_Split"] = False

                    movimientos_nuevos.append(mov_global)

                    for rid in row_ids_global:
                        try:
                            rid = int(rid)
                            row_ids_corregidos.add(rid)
                            row_ids_retorno_origen_global.discard(rid)
                        except Exception:
                            pass

            if not movimientos_nuevos:
                return movimientos_df

            movimientos_limpios = []

            for mov in movimientos_existentes:
                row_ids_mov = mov.get("_row_ids_wms", [])

                if not isinstance(row_ids_mov, list):
                    movimientos_limpios.append(mov)
                    continue

                intersecta = False

                for rid in row_ids_mov:
                    try:
                        if int(rid) in row_ids_corregidos:
                            intersecta = True
                            break
                    except Exception:
                        pass

                if intersecta:
                    continue

                movimientos_limpios.append(mov)

            movimientos_limpios.extend(movimientos_nuevos)

            return pd.DataFrame(movimientos_limpios)
        

                        
        
        
        
        
        
        def _corregir_fragmentacion_unidestino_postproceso(
            movimientos_lista,
            df_calc_base,
            row_ids_retorno_origen_global
        ):
            """
            Corrige casos donde el motor partió un flujo balanceado en varios
            movimientos parciales.

            Caso típico:
            PICK 1 + 2 + 2 = 5
            PUT  1 + 2 + 2 = 5
            Mismo origen, intermedia, destino, item, lote, LP y bloque temporal.

            Resultado:
            - Elimina movimientos parciales existentes que usen esas filas.
            - Quita retornos falsos asociados a esas filas.
            - Crea un único movimiento GLOBAL_MATCH_CORREGIDO.
            """

            if not movimientos_lista:
                return movimientos_lista

            if df_calc_base is None or df_calc_base.empty:
                return movimientos_lista

            df_work = df_calc_base.copy()
            df_work = df_work.sort_values([
                "Nombre_norm",
                "ID Operador",
                "FechaHora",
                "_row_id_wms"
            ]).copy()

            movimientos_corregidos = []
            row_ids_corregidos = set()

            picks_all = df_work[
                df_work["Description_norm"].isin(pick_descs)
            ].copy()

            puts_all = df_work[
                df_work["Description_norm"].isin(put_descs)
            ].copy()

            if picks_all.empty or puts_all.empty:
                return movimientos_lista

            group_pick_cols = [
                "Nombre_norm",
                "ID Operador",
                "Nombre",
                "_Familia_WMS",
                "_Bloque_Flujo_Tiempo",
                "Item Number",
                "Número Lote",
                "Desde ID Ubicación",
                "A ID Ubicación",
                "LP"
            ]

            picks_all["_row_dict"] = picks_all.apply(
                lambda r: r.to_dict(),
                axis=1
            )

            pick_batches = (
                picks_all
                .groupby(group_pick_cols, dropna=False, sort=False)
                .agg(
                    Cantidad_Pick_Total=("Cantidad", "sum"),
                    FH_pick=("FechaHora", "min"),
                    Picks_Rows=("_row_dict", list)
                )
                .reset_index()
                .sort_values("FH_pick")
            )

            for _, pick_batch in pick_batches.iterrows():
                picks_rows = pick_batch.get("Picks_Rows", [])

                if not isinstance(picks_rows, list) or not picks_rows:
                    continue

                pick_ids = _row_ids_de_rows(picks_rows)

                if not pick_ids:
                    continue

                if any(rid in row_ids_corregidos for rid in pick_ids):
                    continue

                nombre_norm = _limpiar_texto(pick_batch.get("Nombre_norm", ""))
                id_operador = _limpiar_texto(pick_batch.get("ID Operador", ""))
                nombre = _limpiar_texto(pick_batch.get("Nombre", ""))
                familia = _limpiar_texto(pick_batch.get("_Familia_WMS", ""))
                bloque = pick_batch.get("_Bloque_Flujo_Tiempo", None)
                item = _limpiar_texto(pick_batch.get("Item Number", ""))
                lote = _limpiar_texto(pick_batch.get("Número Lote", ""))
                origen = _limpiar_texto(pick_batch.get("Desde ID Ubicación", ""))
                intermedia = _limpiar_texto(pick_batch.get("A ID Ubicación", ""))
                lp = _limpiar_texto(pick_batch.get("LP", ""))

                if not origen or not intermedia:
                    continue

                candidatos_put = puts_all[
                    (puts_all["Nombre_norm"].astype(str).str.strip() == nombre_norm) &
                    (puts_all["ID Operador"].astype(str).str.strip() == id_operador) &
                    (puts_all["Nombre"].astype(str).str.strip() == nombre) &
                    (puts_all["_Familia_WMS"].astype(str).str.strip() == familia) &
                    (puts_all["_Bloque_Flujo_Tiempo"] == bloque) &
                    (puts_all["Item Number"].astype(str).str.strip() == item) &
                    (puts_all["Número Lote"].astype(str).str.strip() == lote) &
                    (puts_all["Desde ID Ubicación"].astype(str).str.strip() == intermedia) &
                    (puts_all["LP"].astype(str).str.strip() == lp)
                ].copy()

                if candidatos_put.empty:
                    continue

                destinos = (
                    candidatos_put["A ID Ubicación"]
                    .astype(str)
                    .str.strip()
                    .replace("", pd.NA)
                    .dropna()
                    .unique()
                )

                # Solo corrige casos de destino único.
                # Multidestino lo maneja la lógica padre/hijos.
                if len(destinos) != 1:
                    continue

                destino = _limpiar_texto(destinos[0])

                if not destino:
                    continue

                # No corregir retornos reales aquí.
                if destino.upper() == origen.upper():
                    continue

                candidatos_put = candidatos_put.sort_values([
                    "FechaHora",
                    "_row_id_wms"
                ]).copy()

                puts_rows = [
                    r.to_dict()
                    for _, r in candidatos_put.iterrows()
                ]

                put_ids = _row_ids_de_rows(puts_rows)

                if not put_ids:
                    continue

                if any(rid in row_ids_corregidos for rid in put_ids):
                    continue

                total_pick = round(
                    sum(_to_float(r.get("Cantidad", 0)) for r in picks_rows),
                    6
                )

                total_put = round(
                    sum(_to_float(r.get("Cantidad", 0)) for r in puts_rows),
                    6
                )

                if total_pick != total_put:
                    continue

                resumen_pick = _resumen_item_lote(
                    picks_rows,
                    incluir_lp=False
                )

                resumen_put = _resumen_item_lote(
                    puts_rows,
                    incluir_lp=False
                )

                if resumen_pick != resumen_put:
                    continue

                resumen_pick_lp = _resumen_item_lote(
                    picks_rows,
                    incluir_lp=True
                )

                resumen_put_lp = _resumen_item_lote(
                    puts_rows,
                    incluir_lp=True
                )

                if resumen_pick_lp != resumen_put_lp:
                    continue

                if not _flujo_en_tiempo(picks_rows, puts_rows):
                    continue

                mov_global = _crear_movimiento_desde_match(
                    picks_rows=picks_rows,
                    puts_rows=puts_rows,
                    area_movimiento=_area_bodega_desde_row(picks_rows[0]),
                    forzar_lp_cambiada=False
                )

                if mov_global is None:
                    continue

                row_ids_global = _row_ids_de_rows(
                    picks_rows + puts_rows
                )

                if not row_ids_global:
                    continue

                mov_global["_row_ids_wms"] = row_ids_global
                mov_global["Tipo_Match"] = "GLOBAL_MATCH_CORREGIDO"

                movimientos_corregidos.append(mov_global)

                for rid in row_ids_global:
                    row_ids_corregidos.add(rid)

                # Quitar retornos falsos relacionados con este bloque.
                for rid in row_ids_global:
                    try:
                        row_ids_retorno_origen_global.discard(int(rid))
                    except Exception:
                        pass

            if not movimientos_corregidos:
                return movimientos_lista

            movimientos_limpios = []

            for mov in movimientos_lista:
                row_ids_mov = mov.get("_row_ids_wms", [])

                if not isinstance(row_ids_mov, list):
                    movimientos_limpios.append(mov)
                    continue

                tiene_interseccion = False

                for rid in row_ids_mov:
                    try:
                        if int(rid) in row_ids_corregidos:
                            tiene_interseccion = True
                            break
                    except Exception:
                        pass

                if tiene_interseccion:
                    continue

                movimientos_limpios.append(mov)

            movimientos_limpios.extend(movimientos_corregidos)

            return movimientos_limpios
        
        def _corregir_fragmentacion_unidestino_final(
            movimientos_df,
            df_calc_base,
            row_ids_retorno_origen_global
        ):
            """
            Corrección final para bloques balanceados de un solo destino.

            Caso:
            PICK 1+2+2 = PUT 1+2+2
            mismo operador, item, lote, LP, origen, intermedia y destino único.

            Resultado:
            - Elimina movimientos parciales que usen esas filas.
            - Quita retornos falsos.
            - Crea un único GLOBAL_MATCH_CORREGIDO.
            """

            if df_calc_base is None or df_calc_base.empty:
                return movimientos_df

            df_work = df_calc_base.copy()

            picks_all = df_work[
                df_work["Description_norm"].isin(pick_descs)
            ].copy()

            puts_all = df_work[
                df_work["Description_norm"].isin(put_descs)
            ].copy()

            if picks_all.empty or puts_all.empty:
                return movimientos_df

            movimientos_existentes = []

            if movimientos_df is not None and not movimientos_df.empty:
                movimientos_existentes = movimientos_df.to_dict("records")

            movimientos_nuevos = []
            row_ids_corregidos = set()

            group_cols = [
                "Nombre_norm",
                "ID Operador",
                "Nombre",
                "_Familia_WMS",
                "_Bloque_Flujo_Tiempo",
                "Item Number",
                "Número Lote",
                "Desde ID Ubicación",
                "A ID Ubicación",
                "LP"
            ]

            picks_all["_row_dict"] = picks_all.apply(
                lambda r: r.to_dict(),
                axis=1
            )

            pick_batches = (
                picks_all
                .groupby(group_cols, dropna=False, sort=False)
                .agg(
                    Cantidad_Pick_Total=("Cantidad", "sum"),
                    FH_pick=("FechaHora", "min"),
                    Picks_Rows=("_row_dict", list)
                )
                .reset_index()
                .sort_values("FH_pick")
            )

            for _, pick_batch in pick_batches.iterrows():
                picks_rows = pick_batch.get("Picks_Rows", [])

                if not isinstance(picks_rows, list) or not picks_rows:
                    continue

                pick_ids = _row_ids_de_rows(picks_rows)

                if not pick_ids:
                    continue

                if any(int(rid) in row_ids_corregidos for rid in pick_ids):
                    continue

                nombre_norm = _limpiar_texto(pick_batch.get("Nombre_norm", ""))
                id_operador = _limpiar_texto(pick_batch.get("ID Operador", ""))
                nombre = _limpiar_texto(pick_batch.get("Nombre", ""))
                familia = _limpiar_texto(pick_batch.get("_Familia_WMS", ""))
                bloque = pick_batch.get("_Bloque_Flujo_Tiempo", None)
                item = _limpiar_texto(pick_batch.get("Item Number", ""))
                lote = _limpiar_texto(pick_batch.get("Número Lote", ""))
                origen = _limpiar_texto(pick_batch.get("Desde ID Ubicación", ""))
                intermedia = _limpiar_texto(pick_batch.get("A ID Ubicación", ""))
                lp = _limpiar_texto(pick_batch.get("LP", ""))

                if not origen or not intermedia:
                    continue

                candidatos_put = puts_all[
                    (puts_all["Nombre_norm"].astype(str).str.strip() == nombre_norm) &
                    (puts_all["ID Operador"].astype(str).str.strip() == id_operador) &
                    (puts_all["Nombre"].astype(str).str.strip() == nombre) &
                    (puts_all["_Familia_WMS"].astype(str).str.strip() == familia) &
                    (puts_all["_Bloque_Flujo_Tiempo"] == bloque) &
                    (puts_all["Item Number"].astype(str).str.strip() == item) &
                    (puts_all["Número Lote"].astype(str).str.strip() == lote) &
                    (puts_all["Desde ID Ubicación"].astype(str).str.strip() == intermedia) &
                    (puts_all["LP"].astype(str).str.strip() == lp)
                ].copy()

                if candidatos_put.empty:
                    continue

                destinos = (
                    candidatos_put["A ID Ubicación"]
                    .astype(str)
                    .str.strip()
                    .replace("", pd.NA)
                    .dropna()
                    .unique()
                )

                if len(destinos) != 1:
                    continue

                destino = _limpiar_texto(destinos[0])

                if not destino:
                    continue

                if destino.upper() == origen.upper():
                    continue

                candidatos_put = candidatos_put.sort_values([
                    "FechaHora",
                    "_row_id_wms"
                ]).copy()

                puts_rows = [
                    r.to_dict()
                    for _, r in candidatos_put.iterrows()
                ]

                put_ids = _row_ids_de_rows(puts_rows)

                if not put_ids:
                    continue

                if any(int(rid) in row_ids_corregidos for rid in put_ids):
                    continue

                total_pick = round(
                    sum(_to_float(r.get("Cantidad", 0)) for r in picks_rows),
                    6
                )

                total_put = round(
                    sum(_to_float(r.get("Cantidad", 0)) for r in puts_rows),
                    6
                )

                if total_pick != total_put:
                    continue

                resumen_pick = _resumen_item_lote(
                    picks_rows,
                    incluir_lp=False
                )

                resumen_put = _resumen_item_lote(
                    puts_rows,
                    incluir_lp=False
                )

                if resumen_pick != resumen_put:
                    continue

                resumen_pick_lp = _resumen_item_lote(
                    picks_rows,
                    incluir_lp=True
                )

                resumen_put_lp = _resumen_item_lote(
                    puts_rows,
                    incluir_lp=True
                )

                if resumen_pick_lp != resumen_put_lp:
                    continue

                if not _flujo_en_tiempo(picks_rows, puts_rows):
                    continue

                mov_global = _crear_movimiento_desde_match(
                    picks_rows=picks_rows,
                    puts_rows=puts_rows,
                    area_movimiento=_area_bodega_desde_row(picks_rows[0]),
                    forzar_lp_cambiada=False
                )

                if mov_global is None:
                    continue

                row_ids_global = _row_ids_de_rows(
                    picks_rows + puts_rows
                )

                if not row_ids_global:
                    continue

                mov_global["_row_ids_wms"] = row_ids_global
                mov_global["Tipo_Match"] = "GLOBAL_MATCH_CORREGIDO"
                mov_global["_Grupo_Split_Key"] = ""
                mov_global["_Orden_Split"] = 999
                mov_global["_Es_Padre_Split"] = False

                movimientos_nuevos.append(mov_global)

                for rid in row_ids_global:
                    try:
                        row_ids_corregidos.add(int(rid))
                        row_ids_retorno_origen_global.discard(int(rid))
                    except Exception:
                        pass

            if not movimientos_nuevos:
                return movimientos_df

            movimientos_limpios = []

            for mov in movimientos_existentes:
                row_ids_mov = mov.get("_row_ids_wms", [])

                if not isinstance(row_ids_mov, list):
                    movimientos_limpios.append(mov)
                    continue

                intersecta = False

                for rid in row_ids_mov:
                    try:
                        if int(rid) in row_ids_corregidos:
                            intersecta = True
                            break
                    except Exception:
                        pass

                if intersecta:
                    continue

                movimientos_limpios.append(mov)

            movimientos_limpios.extend(movimientos_nuevos)

            return pd.DataFrame(movimientos_limpios)
        
        
        def _limpiar_retornos_y_movimientos_huerfanos(
            movimientos_df,
            df_calc_base,
            row_ids_retorno_origen_global
        ):
            """
            Limpieza final:
            1. Un row_id usado por un movimiento real no puede quedar como retorno.
            2. Un movimiento NORMAL/GLOBAL/REINTENTO no puede quedar con solo PUT o solo PICK.
            """

            if movimientos_df is None or movimientos_df.empty:
                return movimientos_df

            if df_calc_base is None or df_calc_base.empty:
                return movimientos_df

            df_lookup = df_calc_base.set_index("_row_id_wms", drop=False)

            row_ids_mov_real = set()

            for _, mov in movimientos_df.iterrows():
                row_ids = mov.get("_row_ids_wms", [])

                if not isinstance(row_ids, list):
                    continue

                for rid in row_ids:
                    try:
                        row_ids_mov_real.add(int(rid))
                    except Exception:
                        pass

            # Si una fila ya pertenece a un movimiento real, no puede ser retorno.
            for rid in list(row_ids_retorno_origen_global):
                try:
                    rid_int = int(rid)

                    if rid_int in row_ids_mov_real:
                        row_ids_retorno_origen_global.discard(rid_int)
                except Exception:
                    pass

            movimientos_limpios = []

            tipos_exigen_pick_put = {
                "NORMAL",
                "GLOBAL_MATCH",
                "GLOBAL_MATCH_CORREGIDO",
                "REINTENTO_FINAL"
            }

            for _, mov in movimientos_df.iterrows():
                tipo_match = _limpiar_texto(
                    mov.get("Tipo_Match", "")
                ).upper()

                row_ids = mov.get("_row_ids_wms", [])

                if not isinstance(row_ids, list):
                    movimientos_limpios.append(mov.to_dict())
                    continue

                if tipo_match not in tipos_exigen_pick_put:
                    movimientos_limpios.append(mov.to_dict())
                    continue

                tiene_pick = False
                tiene_put = False

                for rid in row_ids:
                    try:
                        rid_int = int(rid)
                    except Exception:
                        continue

                    if rid_int not in df_lookup.index:
                        continue

                    desc = _limpiar_texto(
                        df_lookup.loc[rid_int, "Description_norm"]
                    )

                    if desc in pick_descs:
                        tiene_pick = True

                    if desc in put_descs:
                        tiene_put = True

                # Si NORMAL/GLOBAL tiene solo PUT o solo PICK, no es confiable.
                if not tiene_pick or not tiene_put:
                    continue

                movimientos_limpios.append(mov.to_dict())

            if not movimientos_limpios:
                return _crear_movimientos_vacios()

            return pd.DataFrame(movimientos_limpios)
        
        

        row_ids_usados_globales = set()

        movimientos_fase_1, usados_fase_1 = _procesar_fase(
            df_fase=df_calc,
            permitir_lp_cambiada=False,
            row_ids_usados_globales=row_ids_usados_globales
        )

        row_ids_usados_globales.update(usados_fase_1)

        df_fase_2 = df_calc[
            ~df_calc["_row_id_wms"].isin(row_ids_usados_globales)
        ].copy()

        movimientos_fase_2, usados_fase_2 = _procesar_fase(
            df_fase=df_fase_2,
            permitir_lp_cambiada=True,
            row_ids_usados_globales=row_ids_usados_globales
        )

        row_ids_usados_globales.update(usados_fase_2)

        movimientos_lista_total = movimientos_fase_1 + movimientos_fase_2
        
        # ============================================================
        # REINTENTO FINAL DE MATCH (CRÍTICO)
        # Evita PICK sueltos y PUT huérfanos
        # ============================================================
        
        movimientos_extra = []
        usados_extra = set()
        
        df_restantes = df_calc[
            ~df_calc["_row_id_wms"].isin(row_ids_usados_globales)
        ].copy()
        
        for _, g in df_restantes.groupby(
            ["Nombre_norm", "ID Operador", "_Familia_WMS", "_Bloque_Flujo_Tiempo", "Item Number", "Número Lote"],
            dropna=False,
            sort=False
        ):
        
            g = g.sort_values(["FechaHora", "_row_id_wms"]).copy()
        
            picks = g[g["Description_norm"].isin(pick_descs)].copy()
            puts = g[g["Description_norm"].isin(put_descs)].copy()
        
            if picks.empty or puts.empty:
                continue
        
            for _, pick_row in picks.iterrows():
        
                rid_pick = _row_id_valido(pick_row)
        
                if rid_pick in row_ids_usados_globales:
                    continue
        
                candidatos_put = puts[
                    (puts["LP"] == pick_row["LP"]) &
                    (puts["Desde ID Ubicación"] == pick_row["A ID Ubicación"])
                ].copy()
        
                if candidatos_put.empty:
                    continue
        
                fh_pick = pd.to_datetime(pick_row["FechaHora"], errors="coerce")
                candidatos_put["FechaHora"] = pd.to_datetime(candidatos_put["FechaHora"], errors="coerce")
        
                candidatos_put["_dist"] = (
                    candidatos_put["FechaHora"] - fh_pick
                ).dt.total_seconds().abs()
        
                candidatos_put = candidatos_put.sort_values("_dist")
        
                for _, put_row in candidatos_put.iterrows():
        
                    rid_put = _row_id_valido(put_row)
        
                    if rid_put in row_ids_usados_globales:
                        continue
        
                    mov = _crear_movimiento_desde_match(
                        picks_rows=[pick_row.to_dict()],
                        puts_rows=[put_row.to_dict()],
                        area_movimiento=_area_bodega_desde_row(pick_row),
                        forzar_lp_cambiada=False
                    )
        
                    if mov is None:
                        continue
        
                    row_ids = _row_ids_de_rows([pick_row.to_dict(), put_row.to_dict()])
        
                    mov["_row_ids_wms"] = row_ids
                    mov["Tipo_Match"] = "REINTENTO_FINAL"
        
                    movimientos_extra.append(mov)
        
                    for rid in row_ids:
                        usados_extra.add(rid)
        
                    break
        
        row_ids_usados_globales.update(usados_extra)
        movimientos_lista_total.extend(movimientos_extra)
        
        
        # ============================================================
        # CORRECCIÓN DEFINITIVA DE FRAGMENTACIÓN UNIDESTINO
        # Corrige casos como:
        # PICK 1+2+2 = PUT 1+2+2 hacia el mismo destino.
        # ============================================================

        movimientos_lista_total = _corregir_fragmentacion_unidestino_postproceso(
            movimientos_lista=movimientos_lista_total,
            df_calc_base=df_calc,
            row_ids_retorno_origen_global=row_ids_retorno_origen_global
        )        
 
        
        
        df_result = df_base.copy()

        if not movimientos_lista_total:
            for row_id in row_ids_retorno_origen_global:
                try:
                    row_id = int(row_id)
                except Exception:
                    continue

                if row_id in df_result.index:
                    df_result.loc[row_id, "Estado"] = "NO_PRODUCTIVO_RETORNO_ORIGEN"
                    df_result.loc[row_id, "Tipo_Match"] = "RETORNO_ORIGEN"
                    df_result.loc[row_id, "Area_Movimiento"] = "Bodega"
                    df_result.loc[row_id, "Movimiento_ID"] = pd.NA
                    df_result.loc[row_id, "Pertenece_a_Pick_Padre"] = pd.NA

            return _crear_movimientos_vacios(), df_result

        movimientos = pd.DataFrame(movimientos_lista_total)
        
        # ============================================================
        # LIMPIEZA FINAL DE RETORNOS FALSOS Y MOVIMIENTOS HUÉRFANOS
        # ============================================================

        movimientos = _limpiar_retornos_y_movimientos_huerfanos(
            movimientos_df=movimientos,
            df_calc_base=df_calc,
            row_ids_retorno_origen_global=row_ids_retorno_origen_global
        )
        
        
        # ============================================================
        # CORRECCIÓN FINAL POR SUB-BLOQUE LOCAL UNIDESTINO
        # Corrige casos PICK 1+2+2 = PUT 1+2+2 en el mismo destino.
        # ============================================================

        movimientos = _corregir_fragmentacion_unidestino_subbloque_local(
            movimientos_df=movimientos,
            df_calc_base=df_calc,
            row_ids_retorno_origen_global=row_ids_retorno_origen_global
        )
        
        
        # ============================================================
        # CORRECCIÓN FINAL DE FRAGMENTACIÓN UNIDESTINO
        # Debe ejecutarse antes de asignar Movimiento_ID finales.
        # ============================================================

        movimientos = _corregir_fragmentacion_unidestino_final(
            movimientos_df=movimientos,
            df_calc_base=df_calc,
            row_ids_retorno_origen_global=row_ids_retorno_origen_global
        )
        
        
        

        if "_Grupo_Split_Key" not in movimientos.columns:
            movimientos["_Grupo_Split_Key"] = ""

        if "_Orden_Split" not in movimientos.columns:
            movimientos["_Orden_Split"] = 999

        if "_Es_Padre_Split" not in movimientos.columns:
            movimientos["_Es_Padre_Split"] = False

        movimientos["_Grupo_Split_Key"] = movimientos["_Grupo_Split_Key"].fillna("")
        movimientos["_Orden_Split"] = pd.to_numeric(
            movimientos["_Orden_Split"],
            errors="coerce"
        ).fillna(999)
        
        # ============================================================
        # Overrides por fila para trazabilidad.
        # Permite que los PICK padre absorbidos mantengan
        # Tipo_Match = MULTIDESTINO_PADRE en la hoja Trazabilidad,
        # sin crear un movimiento adicional en consolidados.
        # ============================================================

        row_ids_tipo_match_override = {}
        
        
        # ============================================================
        # AJUSTE MULTIDESTINO:
        # El PICK padre NO debe contar como movimiento adicional.
        #
        # Regla:
        # - Se elimina MULTIDESTINO_PADRE del consolidado.
        # - Sus row_ids PICK se asignan al primer MULTIDESTINO_HIJO.
        # - El primer hijo queda como ancla lógica del grupo.
        # - Todos los hijos compartirán Pertenece_a_Pick_Padre
        #   con el Movimiento_ID del primer hijo.
        # ============================================================

        def _absorber_padre_multidestino_en_primer_hijo(movs_df):
            if movs_df is None or movs_df.empty:
                return movs_df

            columnas_requeridas = {
                "Tipo_Match",
                "_Grupo_Split_Key",
                "_Orden_Split",
                "_Es_Padre_Split",
                "_row_ids_wms"
            }

            if not columnas_requeridas.issubset(set(movs_df.columns)):
                return movs_df

            df_tmp = movs_df.copy()
            indices_padre_a_eliminar = set()

            for grupo_key, g in df_tmp.groupby("_Grupo_Split_Key", dropna=False):
                grupo_key_limpio = _limpiar_texto(grupo_key)

                if not grupo_key_limpio:
                    continue

                tipos = g["Tipo_Match"].astype(str).str.upper().str.strip()

                padres = g[
                    tipos.eq("MULTIDESTINO_PADRE")
                ].copy()

                hijos = g[
                    tipos.eq("MULTIDESTINO_HIJO")
                ].copy()

                if padres.empty or hijos.empty:
                    continue

                hijos = hijos.sort_values([
                    "_Orden_Split",
                    "FH_inicio",
                    "Destino_Final"
                ]).copy()

                primer_hijo_idx = hijos.index[0]

                row_ids_padre = []

                for _, padre_row in padres.iterrows():
                    ids_tmp = padre_row.get("_row_ids_wms", [])

                    if isinstance(ids_tmp, list):
                        for rid in ids_tmp:
                            try:
                                rid_int = int(rid)
                                row_ids_padre.append(rid_int)

                                # CLAVE:
                                # Aunque el padre se absorbe en el primer hijo,
                                # sus filas PICK deben conservar Tipo_Match padre
                                # en Trazabilidad.
                                row_ids_tipo_match_override[rid_int] = "MULTIDESTINO_PADRE"
                            except Exception:
                                pass

                    indices_padre_a_eliminar.add(padre_row.name)

                row_ids_hijo = df_tmp.at[primer_hijo_idx, "_row_ids_wms"]

                if not isinstance(row_ids_hijo, list):
                    row_ids_hijo = []

                row_ids_hijo_limpios = []

                for rid in row_ids_hijo:
                    try:
                        row_ids_hijo_limpios.append(int(rid))
                    except Exception:
                        pass

                row_ids_unificados = list(
                    dict.fromkeys(row_ids_padre + row_ids_hijo_limpios)
                )

                # El primer hijo carga también los row_ids del PICK padre.
                df_tmp.at[primer_hijo_idx, "_row_ids_wms"] = row_ids_unificados

                # El primer hijo será el ancla lógica del grupo para Pertenece_a_Pick_Padre.
                df_tmp.at[primer_hijo_idx, "_Es_Padre_Split"] = True
                df_tmp.at[primer_hijo_idx, "_Orden_Split"] = 0

                # Importante:
                # En Movimientos Consolidados sigue siendo HIJO para no inflar productividad.
                df_tmp.at[primer_hijo_idx, "Tipo_Match"] = "MULTIDESTINO_HIJO"

                otros_hijos_idx = [
                    idx for idx in hijos.index
                    if idx != primer_hijo_idx
                ]

                for idx in otros_hijos_idx:
                    df_tmp.at[idx, "_Es_Padre_Split"] = False

            if indices_padre_a_eliminar:
                df_tmp = df_tmp.drop(
                    index=list(indices_padre_a_eliminar),
                    errors="ignore"
                ).copy()

            return df_tmp.reset_index(drop=True)


        movimientos = _absorber_padre_multidestino_en_primer_hijo(movimientos)
        
        
        

        movimientos["_Origen_norm_tmp"] = (
            movimientos["Origen"].astype(str).str.strip().str.upper()
        )

        movimientos["_Destino_norm_tmp"] = (
            movimientos["Destino_Final"].astype(str).str.strip().str.upper()
        )

        mask_retorno_directo = (
            movimientos["_Origen_norm_tmp"].ne("") &
            movimientos["_Destino_norm_tmp"].ne("") &
            movimientos["_Origen_norm_tmp"].eq(movimientos["_Destino_norm_tmp"])
        )

        if mask_retorno_directo.any():
            movimientos_retorno = movimientos[mask_retorno_directo].copy()

            for _, mov_retorno in movimientos_retorno.iterrows():
                row_ids_tmp = mov_retorno.get("_row_ids_wms", [])

                if not isinstance(row_ids_tmp, list):
                    continue

                for rid in row_ids_tmp:
                    try:
                        row_ids_retorno_origen_global.add(int(rid))
                    except Exception:
                        pass

            movimientos = movimientos[~mask_retorno_directo].copy()

        movimientos = movimientos.drop(
            columns=[
                "_Origen_norm_tmp",
                "_Destino_norm_tmp"
            ],
            errors="ignore"
        )

        if movimientos.empty:
            for row_id in row_ids_retorno_origen_global:
                try:
                    row_id = int(row_id)
                except Exception:
                    continue

                if row_id in df_result.index:
                    df_result.loc[row_id, "Estado"] = "NO_PRODUCTIVO_RETORNO_ORIGEN"
                    df_result.loc[row_id, "Tipo_Match"] = "RETORNO_ORIGEN"
                    df_result.loc[row_id, "Area_Movimiento"] = "Bodega"
                    df_result.loc[row_id, "Movimiento_ID"] = pd.NA
                    df_result.loc[row_id, "Pertenece_a_Pick_Padre"] = pd.NA

            return _crear_movimientos_vacios(), df_result

        movimientos = movimientos.sort_values([
            "Nombre_norm",
            "FH_inicio",
            "_Grupo_Split_Key",
            "_Orden_Split",
            "Origen",
            "Intermedia",
            "Destino_Final"
        ]).reset_index(drop=True)

        movimientos["Movimiento_ID"] = (
            movimientos.groupby("Nombre_norm").cumcount() + 1
        )

        movimientos["Pertenece_a_Pick_Padre"] = movimientos["Movimiento_ID"]

        for grupo_key, g in movimientos.groupby("_Grupo_Split_Key", dropna=False):
            grupo_key = _limpiar_texto(grupo_key)

            if not grupo_key:
                continue

            padres = g[g["_Es_Padre_Split"] == True]

            if padres.empty:
                continue

            padre_id = padres["Movimiento_ID"].iloc[0]

            movimientos.loc[
                movimientos["_Grupo_Split_Key"].astype(str).str.strip() == grupo_key,
                "Pertenece_a_Pick_Padre"
            ] = padre_id

        # ============================================================
        # MAPEO ROBUSTO DE MOVIMIENTOS HACIA TRAZABILIDAD
        # Usa _row_id_wms y no el índice del DataFrame.
        # Esto evita que los movimientos se cuenten pero no aparezcan
        # en la hoja Trazabilidad.
        # ============================================================

        row_ids_usados_en_movimientos = set()

        def _normalizar_lista_row_ids(valor):
            if isinstance(valor, list):
                salida = []

                for x in valor:
                    try:
                        salida.append(int(x))
                    except Exception:
                        pass

                return list(dict.fromkeys(salida))

            if pd.isna(valor):
                return []

            # Por seguridad, si algún row_ids quedó como string tipo "1|2|3"
            texto = str(valor).strip()

            if not texto:
                return []

            for sep in ["|", ",", ";"]:
                texto = texto.replace(sep, " ")

            salida = []

            for parte in texto.split():
                try:
                    salida.append(int(float(parte)))
                except Exception:
                    pass

            return list(dict.fromkeys(salida))


        for _, mov in movimientos.iterrows():
            row_ids = _normalizar_lista_row_ids(
                mov.get("_row_ids_wms", [])
            )

            if not row_ids:
                continue

            for row_id in row_ids:
                row_ids_usados_en_movimientos.add(row_id)

                tipo_match_fila = mov.get("Tipo_Match", pd.NA)

                try:
                    if "row_ids_tipo_match_override" in locals():
                        tipo_match_fila = row_ids_tipo_match_override.get(
                            int(row_id),
                            tipo_match_fila
                        )
                except Exception:
                    pass

                if "_row_id_wms" in df_result.columns:
                    mask_row = df_result["_row_id_wms"].astype("Int64") == int(row_id)
                else:
                    mask_row = df_result.index == int(row_id)

                if not mask_row.any():
                    continue

                df_result.loc[mask_row, "Movimiento_ID"] = mov["Movimiento_ID"]
                df_result.loc[mask_row, "Tipo_Match"] = tipo_match_fila
                df_result.loc[mask_row, "Pertenece_a_Pick_Padre"] = mov["Pertenece_a_Pick_Padre"]
                df_result.loc[mask_row, "Area_Movimiento"] = mov["Area_Movimiento"]

        # Si una fila ya pertenece a un movimiento real, no puede quedar como retorno.
        for rid in list(row_ids_retorno_origen_global):
            try:
                rid_int = int(rid)

                if rid_int in row_ids_usados_en_movimientos:
                    row_ids_retorno_origen_global.discard(rid_int)
            except Exception:
                pass
            

                    

        for row_id in row_ids_retorno_origen_global:
            try:
                row_id = int(row_id)
            except Exception:
                continue

            if row_id in df_result.index:
                df_result.loc[row_id, "Estado"] = "NO_PRODUCTIVO_RETORNO_ORIGEN"
                df_result.loc[row_id, "Tipo_Match"] = "RETORNO_ORIGEN"
                df_result.loc[row_id, "Area_Movimiento"] = "Bodega"
                df_result.loc[row_id, "Movimiento_ID"] = pd.NA
                df_result.loc[row_id, "Pertenece_a_Pick_Padre"] = pd.NA

        movimientos = movimientos.drop(
            columns=[
                "_row_ids_wms",
                "_Grupo_Split_Key",
                "_Orden_Split",
                "_Es_Padre_Split"
            ],
            errors="ignore"
        )

        movimientos = _ensure_columns(
            movimientos,
            _crear_movimientos_vacios().columns.tolist()
        )

        return movimientos, df_result


    # ============================================================
    # VALIDACIONES INICIALES
    # ============================================================

    if not ruta or not str(ruta).strip():
        raise ValueError(
            "⚠ No se encontró la ruta del Excel origen.\n\n"
            "Selecciona un archivo antes de procesar."
        )

    ruta = str(ruta).strip()

    if not os.path.exists(ruta):
        raise ValueError(
            "⚠ El archivo Excel origen no existe en la ruta seleccionada.\n\n"
            f"Ruta detectada:\n{ruta}"
        )

    if not carpeta or not str(carpeta).strip():
        raise ValueError(
            "⚠ No se encontró la carpeta destino.\n\n"
            "Selecciona una carpeta antes de procesar."
        )

    carpeta = str(carpeta).strip()

    if not os.path.exists(carpeta):
        raise ValueError(
            "⚠ La carpeta destino no existe.\n\n"
            f"Ruta detectada:\n{carpeta}"
        )

    if turno_config is None:
        turno_config = {}

    meta_ppt_bodega = float(turno_config.get("meta_ppt_bodega", 0) or 0)
    meta_ppt_despachos = float(turno_config.get("meta_ppt_despachos", 0) or 0)

    if df_pre_filtrado is None and validar_fecha_excel:
        fecha_valida, mensaje_fecha, fecha_min_excel, fecha_max_excel, registros_turno = validar_fecha_y_turno_excel_rapido(
            ruta_excel=ruta,
            fecha_seleccionada=fecha_str,
            hora_inicio=ini,
            hora_fin=fin,
            columna_fecha="Fecha",
            columna_tiempo="Tiempo"
        )

        if not fecha_valida:
            update(0, "Fecha u hora de turno fuera de rango. Proceso detenido.")
            raise ValueError(mensaje_fecha)

        update(
            5,
            f"Fecha y turno validados correctamente. Registros detectados: {registros_turno}"
        )
    else:
        registros_turno = 0 if df_pre_filtrado is None else len(df_pre_filtrado)

        update(
            5,
            f"Turno recibido prefiltrado en memoria. Registros detectados: {registros_turno}"
        )

    # ============================================================
    # CARGA Y NORMALIZACIÓN DEL EXCEL ORIGEN
    # ============================================================

    update(8, "📥 Cargando y normalizando datos...")

    if df_pre_filtrado is not None:
        df0 = df_pre_filtrado.copy()
        df0 = _normalize_column_names(df0)
    else:
        df0 = pd.read_excel(ruta, engine="openpyxl")
        df0 = _normalize_column_names(df0)

    if df0 is None:
        raise ValueError(
            "Error interno: _normalize_column_names devolvió None.\n\n"
            "Revisa wms_utils.py y confirma que la función termine con return df."
        )

    try:
        fecha_base = datetime.strptime(fecha_str.strip(), "%d/%m/%Y").date()
    except Exception:
        raise ValueError("⚠ Fecha inválida. Usa formato dd/mm/aaaa")

    if "Fecha" not in df0.columns:
        raise ValueError("⚠ El archivo no contiene la columna Fecha")

    df0["Fecha_dt_tmp"] = df0["Fecha"].apply(_parse_fecha_excel)

    fecha_min = df0["Fecha_dt_tmp"].min()
    fecha_max = df0["Fecha_dt_tmp"].max()

    if pd.isna(fecha_min) or pd.isna(fecha_max):
        raise ValueError("⚠ No se pudieron leer fechas válidas en el archivo")

    if fecha_base < fecha_min - timedelta(days=1) or fecha_base > fecha_max + timedelta(days=1):
        raise ValueError(
            f"⚠ Fecha fuera de rango.\n"
            f"Rango permitido según archivo: {fecha_min} a {fecha_max}"
        )

    cols_base = [
        "Item Number",
        "Número Lote",
        "Código Transacción",
        "Description",
        "ID Bodega",
        "Desde ID Ubicación",
        "A ID Ubicación",
        "LP",
        "ID Operador",
        "Nombre",
        "Fecha",
        "Tiempo",
        "Cantidad",
        "Tipo"
    ]

    df0 = _ensure_columns(df0, cols_base)
    df0 = df0[cols_base].copy()

    df0["Description_norm"] = _norm_desc_series(df0["Description"])
    df0["Nombre_norm"] = df0["Nombre"].apply(normalizar_nombre)

    # ============================================================
    # FILTRO POR TURNO
    # ============================================================

    # ============================================================
    # FILTRO POR TURNO
    # ============================================================

    if df_pre_filtrado is None:
        update(12, "🧠 Filtrando por turno...")

        h_ini = datetime.strptime(ini, "%H:%M").time()
        h_fin = datetime.strptime(fin, "%H:%M").time()

        df0["Fecha_dt"] = df0["Fecha"].apply(_parse_fecha_excel)
        df0["Hora_dt"] = df0["Tiempo"].apply(_parse_hora_excel)

        cruza = h_ini > h_fin

        total_antes_turno = len(df0)
        filas_sin_fecha_valida = df0["Fecha_dt"].isna().sum()
        filas_sin_hora_valida = df0["Hora_dt"].isna().sum()

        if filas_sin_fecha_valida > 0 or filas_sin_hora_valida > 0:
            update(
                12,
                f"Filtrando turno... Total: {total_antes_turno} | "
                f"Sin fecha válida: {filas_sin_fecha_valida} | "
                f"Sin hora válida: {filas_sin_hora_valida}"
            )

        if not cruza:
            mask_turno = (
                (df0["Fecha_dt"] == fecha_base) &
                (df0["Hora_dt"].notna()) &
                (df0["Hora_dt"] >= h_ini) &
                (df0["Hora_dt"] <= h_fin)
            )
        else:
            fecha_turno = fecha_base - timedelta(days=1)

            cond_noche = df0["Hora_dt"] >= h_ini
            cond_madrugada = df0["Hora_dt"] <= h_fin

            fecha_turno_col = df0["Fecha_dt"].copy()
            fecha_turno_col[cond_madrugada] = (
                fecha_turno_col[cond_madrugada] - timedelta(days=1)
            )

            mask_turno = (
                (fecha_turno_col == fecha_turno) &
                (df0["Hora_dt"].notna()) &
                (cond_noche | cond_madrugada)
            )

        df0 = df0[mask_turno].copy()
        df0 = df0.reset_index(drop=True)
        df0["_row_id_wms"] = df0.index

    else:
        update(12, "🧠 Usando turno prefiltrado en memoria...")

        df0 = df0.copy()
        df0 = df0.reset_index(drop=True)
        df0["_row_id_wms"] = df0.index

        if "Fecha_dt" not in df0.columns:
            df0["Fecha_dt"] = df0["Fecha"].apply(_parse_fecha_excel)

        if "Hora_dt" not in df0.columns:
            df0["Hora_dt"] = df0["Tiempo"].apply(_parse_hora_excel)

    # ============================================================
    # MAESTRO DE PERSONAL
    # ============================================================

    update(16, "👥 Validando maestro de personal...")

    df_maestro = cargar_maestro_personal()

    update(
        16,
        f"👥 Maestro cargado: {len(df_maestro)} registros"
    )

    ok_maestro, mensaje_maestro = validar_maestro_personal(df_maestro)

    if not ok_maestro:
        raise ValueError(mensaje_maestro)
        
        
    # ============================================================
    # PERSONAL DETECTADO EN EL TURNO
    # Se usa para cruzar operadores del WMS contra maestro_personal.xlsx
    # ============================================================

    if "ID Operador" not in df0.columns:
        df0["ID Operador"] = ""

    if "Nombre" not in df0.columns:
        df0["Nombre"] = ""

    if "Nombre_norm" not in df0.columns:
        df0["Nombre_norm"] = df0["Nombre"].apply(normalizar_nombre)

    personal_turno = (
        df0[
            ["ID Operador", "Nombre", "Nombre_norm"]
        ]
        .copy()
    )

    personal_turno["ID Operador"] = (
        personal_turno["ID Operador"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    personal_turno["Nombre"] = (
        personal_turno["Nombre"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    personal_turno["Nombre_norm"] = (
        personal_turno["Nombre_norm"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    personal_turno = personal_turno[
        (personal_turno["ID Operador"].astype(str).str.strip() != "") |
        (personal_turno["Nombre_norm"].astype(str).str.strip() != "")
    ].copy()

    personal_turno = personal_turno.drop_duplicates(
        subset=["ID Operador", "Nombre_norm"]
    ).reset_index(drop=True)

    # ============================================================
    # ASIGNACIÓN DE OPERADORES CONTRA MAESTRO
    # ============================================================

    resultado_asignacion = construir_asignacion_operadores_para_turno_df(
        df_turno=personal_turno,
        df_maestro=df_maestro
    )

    # ============================================================
    # MAPA ROBUSTO DE ROLES Y LÍDERES DESDE MAESTRO PERSONAL
    # ============================================================

    df_maestro_roles = df_maestro.copy()

    columnas_maestro_roles = [
        "ID Operador",
        "Nombre",
        "Nombre_norm",
        "Area",
        "Rol",
        "Lider",
        "Activo"
    ]

    for col in columnas_maestro_roles:
        if col not in df_maestro_roles.columns:
            df_maestro_roles[col] = ""

    df_maestro_roles["ID_Operador_norm"] = (
        df_maestro_roles["ID Operador"]
        .apply(normalizar_id_operador)
    )

    df_maestro_roles["Nombre_norm"] = df_maestro_roles.apply(
        lambda r: str(r.get("Nombre_norm", "")).strip()
        if str(r.get("Nombre_norm", "")).strip()
        else normalizar_nombre(r.get("Nombre", "")),
        axis=1
    )

    df_maestro_roles["Rol"] = (
        df_maestro_roles["Rol"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df_maestro_roles["Lider"] = (
        df_maestro_roles["Lider"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df_maestro_roles["Area"] = (
        df_maestro_roles["Area"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    def _activo_bool_maestro(valor):
        texto = str(valor).strip().lower()

        if texto in ["false", "falso", "0", "no", "n", "inactivo"]:
            return False

        return True

    df_maestro_roles["Activo_bool"] = df_maestro_roles["Activo"].apply(
        _activo_bool_maestro
    )

    # Prioridad: Bodega/Mixto activos
    df_roles_bodega = df_maestro_roles[
        (df_maestro_roles["Activo_bool"] == True) &
        (df_maestro_roles["Area"].isin(["Bodega", "Mixto"]))
    ].copy()

    # Fallback: si no detecta área, usar todo el maestro activo
    if df_roles_bodega.empty:
        df_roles_bodega = df_maestro_roles[
            df_maestro_roles["Activo_bool"] == True
        ].copy()

    rol_por_id_operador = (
        df_roles_bodega[
            df_roles_bodega["ID_Operador_norm"].astype(str).str.strip().ne("")
        ]
        .drop_duplicates("ID_Operador_norm")
        .set_index("ID_Operador_norm")["Rol"]
        .to_dict()
    )

    rol_por_nombre_norm = (
        df_roles_bodega[
            df_roles_bodega["Nombre_norm"].astype(str).str.strip().ne("")
        ]
        .drop_duplicates("Nombre_norm")
        .set_index("Nombre_norm")["Rol"]
        .to_dict()
    )

    lider_por_id_operador = (
        df_roles_bodega[
            df_roles_bodega["ID_Operador_norm"].astype(str).str.strip().ne("")
        ]
        .drop_duplicates("ID_Operador_norm")
        .set_index("ID_Operador_norm")["Lider"]
        .to_dict()
    )

    lider_por_nombre_norm = (
        df_roles_bodega[
            df_roles_bodega["Nombre_norm"].astype(str).str.strip().ne("")
        ]
        .drop_duplicates("Nombre_norm")
        .set_index("Nombre_norm")["Lider"]
        .to_dict()
    )

    def _rol_desde_operador(id_operador, nombre_norm, nombre=""):
        id_norm = normalizar_id_operador(id_operador)
        nombre_norm = str(nombre_norm).strip()

        if not nombre_norm and nombre:
            nombre_norm = normalizar_nombre(nombre)

        rol = rol_por_id_operador.get(id_norm, "")

        if rol:
            return rol

        rol = rol_por_nombre_norm.get(nombre_norm, "")

        if rol:
            return rol

        return ""

    def _lider_desde_operador(id_operador, nombre_norm, nombre=""):
        id_norm = normalizar_id_operador(id_operador)
        nombre_norm = str(nombre_norm).strip()

        if not nombre_norm and nombre:
            nombre_norm = normalizar_nombre(nombre)

        lider = lider_por_id_operador.get(id_norm, "")

        if lider:
            return lider

        lider = lider_por_nombre_norm.get(nombre_norm, "")

        if lider:
            return lider

        return ""

    # ============================================================
    # RESULTADOS DE ASIGNACIÓN
    # ============================================================

    asignacion_operadores_area = resultado_asignacion.get("asignacion", {})
    personas_no_maestro_norm = set(resultado_asignacion.get("faltantes", []))
    personas_no_maestro_raw = resultado_asignacion.get("faltantes_raw", [])

    # ============================================================
    # NOMBRES EXCLUIDOS DESDE MAESTRO
    # IMPORTANTE:
    # Esta variable debe existir SIEMPRE antes de cualquier uso posterior.
    # ============================================================

    nombres_excluidos_maestro = {
        str(nombre_norm)
        for nombre_norm, area in asignacion_operadores_area.items()
        if str(area).strip() == "Excluir"
    }

    # ============================================================
    # PERSONAL NO REGISTRADO EN MAESTRO
    # En proceso mensual/rango no debe preguntar, solo ignorar.
    # En proceso diario sí pregunta.
    # ============================================================

    if personas_no_maestro_norm:
        mensaje_faltantes = (
            "Se encontraron personas en el turno que no existen en el maestro de personal.\n\n"
            "Personal faltante:\n- " + "\n- ".join(personas_no_maestro_raw) + "\n\n"
            "Si continúas, esas personas serán ignoradas del análisis."
        )

        if ignorar_faltantes_maestro:
            update(
                17,
                f"⚠ Personal no registrado en maestro ignorado automáticamente: {len(personas_no_maestro_norm)}"
            )
        else:
            respuesta = sg.popup_yes_no(
                mensaje_faltantes + "\n\n¿Deseas continuar e ignorar esas personas?",
                title="Personal no registrado en maestro",
                font=("Segoe UI", 10)
            )

            if respuesta != "Yes":
                raise ValueError(
                    "⚠ Proceso cancelado por el usuario.\n\n"
                    "Debes actualizar el maestro de personal o aceptar continuar ignorando "
                    "las personas no registradas."
                )

    # ============================================================
    # NOMBRES A IGNORAR EN EL ANÁLISIS
    # Incluye:
    # - Personal excluido desde maestro
    # - Personal no registrado en maestro
    #
    # IMPORTANTE:
    # Esta variable es usada por Despachos y Bodega.
    # Debe existir siempre, también en procesamiento mensual/rango.
    # ============================================================

    if personas_no_maestro_norm is None:
        personas_no_maestro_norm = set()

    if nombres_excluidos_maestro is None:
        nombres_excluidos_maestro = set()

    nombres_ignorar = set(nombres_excluidos_maestro).union(
        set(personas_no_maestro_norm)
    )

    # ============================================================
    # ÁREA ASIGNADA DEL OPERADOR
    # ============================================================

    df0["Area_Operador_Asignada"] = (
        df0["Nombre_norm"]
        .astype(str)
        .map(asignacion_operadores_area)
        .fillna("Excluir")
    )

    df0["Es_Operador_Bodega_Maestro"] = df0["Area_Operador_Asignada"].isin([
        "Bodega",
        "Mixto"
    ])

    # ============================================================
    # ASIGNAR ROL Y LÍDER DESDE MAESTRO PERSONAL
    # ============================================================

    if "Rol" not in df0.columns:
        df0["Rol"] = ""

    if "Lider" not in df0.columns:
        df0["Lider"] = ""

    df0["Rol"] = df0.apply(
        lambda r: _rol_desde_operador(
            r.get("ID Operador", ""),
            r.get("Nombre_norm", ""),
            r.get("Nombre", "")
        ),
        axis=1
    )

    df0["Lider"] = df0.apply(
        lambda r: _lider_desde_operador(
            r.get("ID Operador", ""),
            r.get("Nombre_norm", ""),
            r.get("Nombre", "")
        ),
        axis=1
    )

    df0["Rol"] = (
        df0["Rol"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df0["Lider"] = (
        df0["Lider"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # ============================================================
    # DIAGNÓSTICO DE OPERADORES DE BODEGA SIN ROL / SIN LÍDER
    # ============================================================

    operadores_sin_rol = (
        df0[
            (df0["Es_Operador_Bodega_Maestro"] == True) &
            (df0["Rol"].astype(str).str.strip() == "")
        ][["ID Operador", "Nombre", "Nombre_norm"]]
        .drop_duplicates()
        .copy()
    )

    operadores_sin_lider = (
        df0[
            (df0["Es_Operador_Bodega_Maestro"] == True) &
            (df0["Lider"].astype(str).str.strip() == "")
        ][["ID Operador", "Nombre", "Nombre_norm"]]
        .drop_duplicates()
        .copy()
    )

    if not operadores_sin_rol.empty:
        update(
            18,
            f"⚠ Operadores de Bodega sin Rol detectados: {len(operadores_sin_rol)}"
        )

    if not operadores_sin_lider.empty:
        update(
            18,
            f"⚠ Operadores de Bodega sin Líder detectados: {len(operadores_sin_lider)}"
        )

    # ============================================================
    # FECHAHORA
    # ============================================================

    update(20, "⏱ Construyendo FechaHora...")

    df0["FechaHora"] = _build_fechahora(df0)
    df0 = df0[df0["FechaHora"].notna()].copy()

    if df0.empty:
        raise ValueError("⚠ No quedaron registros con FechaHora válida.")

    df0["Fecha"] = df0["Fecha"].apply(_parse_fecha_excel)

    # ============================================================
    # DESPACHOS
    # ============================================================

    update(24, "🚛 Construyendo eventos y ciclos de Despachos...")
    
    if "nombres_ignorar" not in locals():
        nombres_ignorar = set(nombres_excluidos_maestro).union(
            set(personas_no_maestro_norm)
        )
    

    eventos_despacho = _construir_eventos_despacho_por_fila(
        df_base=df0,
        asignacion_operadores_area=asignacion_operadores_area,
        nombres_ignorar=nombres_ignorar
    )

    ciclos_despacho = _construir_ciclos_despacho(
        eventos_despacho,
        max_gap_min=60
    )

    inactividad_despacho = _calcular_inactividad_despacho(
        ciclos_despacho=ciclos_despacho,
        fecha_str=fecha_str,
        ini=ini,
        fin=fin,
        min_inactividad_reportable=5
    )

    # ============================================================
    # BODEGA
    # ============================================================

    update(50, "🔗 Construyendo movimientos reales de Bodega...")

    df0["Estado"] = "NO_APLICA"
    df0["Movimiento_ID"] = pd.NA
    df0["Tipo_Match"] = pd.NA
    df0["Pertenece_a_Pick_Padre"] = pd.NA
    df0["Area_Movimiento"] = pd.NA

    mask_personal_no_maestro = df0["Nombre_norm"].isin(personas_no_maestro_norm)
    mask_personal_excluido = df0["Nombre_norm"].isin(nombres_excluidos_maestro)

    df0.loc[mask_personal_no_maestro, "Estado"] = "PERSONAL_NO_MAESTRO"
    df0.loc[mask_personal_no_maestro, "Area_Movimiento"] = "No maestro"

    df0.loc[mask_personal_excluido, "Estado"] = "PERSONAL_EXCLUIDO"
    df0.loc[mask_personal_excluido, "Area_Movimiento"] = "Excluido"

    mask_crossdock = df0["Description_norm"].str.contains("crossdock", na=False)

    if "Código Transacción" in df0.columns:
        cod_num = pd.to_numeric(df0["Código Transacción"], errors="coerce")
        mask_crossdock = mask_crossdock | (cod_num == 311)

    df0.loc[mask_crossdock, "Estado"] = "CROSSDOCK_IGNORADO"

    allowed_descs = {
        "move (pick)",
        "move (put)",
        "move trailer (pick)",
        "move trailer (put)",
        "picking (pick)",
        "picking (put)"
    }

    mask_a_a = (
        df0["Desde ID Ubicación"].astype(str).str.strip()
        ==
        df0["A ID Ubicación"].astype(str).str.strip()
    )

    df0.loc[
        mask_a_a & df0["Description_norm"].isin(allowed_descs),
        "Estado"
    ] = "MOV_INVALIDO_A_A"

    movimientos, df0 = _construir_movimientos_bodega_secuencial(
        df_base=df0,
        asignacion_operadores_area=asignacion_operadores_area,
        personas_no_maestro_norm=personas_no_maestro_norm,
        nombres_excluidos_maestro=nombres_excluidos_maestro
    )
    
    # ============================================================
    # AGREGAR ROL Y LÍDER A MOVIMIENTOS CONSOLIDADOS
    # ============================================================
    
    if movimientos is not None and not movimientos.empty:
    
        if "Nombre_norm" not in movimientos.columns:
            movimientos["Nombre_norm"] = movimientos["Nombre"].apply(
                normalizar_nombre
            )
    
        if "Rol" not in movimientos.columns:
            movimientos["Rol"] = ""
    
        if "Lider" not in movimientos.columns:
            movimientos["Lider"] = ""
    
        movimientos["Rol"] = movimientos.apply(
            lambda r: _rol_desde_operador(
                r.get("ID Operador", ""),
                r.get("Nombre_norm", ""),
                r.get("Nombre", "")
            ),
            axis=1
        )
    
        movimientos["Lider"] = movimientos.apply(
            lambda r: _lider_desde_operador(
                r.get("ID Operador", ""),
                r.get("Nombre_norm", ""),
                r.get("Nombre", "")
            ),
            axis=1
        )
    
        movimientos["Rol"] = (
            movimientos["Rol"]
            .fillna("")
            .astype(str)
            .str.strip()
        )
    
        movimientos["Lider"] = (
            movimientos["Lider"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    # ============================================================
    # Estados que NO deben ser sobrescritos como SIN_MATCH
    # ============================================================

    estados_no_sobrescribir = {
        "NO_PRODUCTIVO_RETORNO_ORIGEN",
        "MOV_INVALIDO_A_A",
        "CROSSDOCK_IGNORADO",
        "PERSONAL_EXCLUIDO",
        "PERSONAL_NO_MAESTRO",
        "DUPLICADO_WMS_IGNORADO"
    }

    # ============================================================
    # MÁSCARAS PARA MARCAR SIN_MATCH
    # ============================================================

    if "Estado" not in df0.columns:
        df0["Estado"] = ""

    if "Movimiento_ID" not in df0.columns:
        df0["Movimiento_ID"] = pd.NA

    if "Area_Movimiento" not in df0.columns:
        df0["Area_Movimiento"] = ""

    # ------------------------------------------------------------
    # Movimiento_ID válido:
    # fila que ya fue identificada como movimiento real.
    # ------------------------------------------------------------

    movimiento_id_valido = (
        df0["Movimiento_ID"]
        .notna() &
        df0["Movimiento_ID"]
        .astype(str)
        .str.strip()
        .ne("") &
        df0["Movimiento_ID"]
        .astype(str)
        .str.strip()
        .str.lower()
        .ne("nan") &
        df0["Movimiento_ID"]
        .astype(str)
        .str.strip()
        .str.lower()
        .ne("<na>")
    )

    # ------------------------------------------------------------
    # Filas candidatas a análisis de Bodega.
    # Solo deben marcarse como SIN_MATCH las filas de operadores
    # Bodega/Mixto que no fueron identificadas como movimiento
    # y que no tienen un estado protegido.
    # ------------------------------------------------------------

    if "Area_Operador_Asignada" in df0.columns:
        mask_calc_any = (
            df0["Area_Operador_Asignada"]
            .fillna("")
            .astype(str)
            .str.strip()
            .isin(["Bodega", "Mixto"])
        )
    else:
        mask_calc_any = pd.Series(True, index=df0.index)

    # ------------------------------------------------------------
    # Reforzar que solo afecte Bodega o filas sin área de movimiento.
    # Así evitamos pisar Despachos u otras clasificaciones.
    # ------------------------------------------------------------

    if "Area_Movimiento" in df0.columns:
        mask_area_movimiento_bodega = (
            df0["Area_Movimiento"]
            .fillna("")
            .astype(str)
            .str.strip()
            .isin(["", "Bodega"])
        )

        mask_calc_any = mask_calc_any & mask_area_movimiento_bodega

    # ------------------------------------------------------------
    # Estados protegidos que no deben ser cambiados a SIN_MATCH.
    # ------------------------------------------------------------

    mask_no_sobrescribir = (
        df0["Estado"]
        .fillna("")
        .astype(str)
        .str.strip()
        .isin(estados_no_sobrescribir)
    )

    df0.loc[
        mask_calc_any & ~movimiento_id_valido & ~mask_no_sobrescribir,
        "Estado"
    ] = "SIN_MATCH"

    df0.loc[
        mask_calc_any & ~movimiento_id_valido & ~mask_no_sobrescribir,
        "Area_Movimiento"
    ] = "Bodega"


    # ============================================================
    # ORGANIZAR MOVIMIENTOS CONSOLIDADOS
    # ============================================================

    update(80, "📘 Organizando Movimientos Consolidados...")

    movimientos = _ensure_columns(
        movimientos,
        _crear_movimientos_vacios().columns.tolist()
    )

    if not movimientos.empty:
        movimientos = movimientos.sort_values([
            "Area_Movimiento",
            "Nombre",
            "FH_inicio",
            "Movimiento_ID"
        ]).reset_index(drop=True)

        movimientos["Movimiento_Global_ID"] = (
            "MOV-" + (movimientos.index + 1).astype(str).str.zfill(5)
        )

        movimientos["Es_Movimiento_Padre"] = (
            movimientos["Movimiento_ID"].astype(str)
            ==
            movimientos["Pertenece_a_Pick_Padre"].astype(str)
        )

        movimientos["Es_Hijo_Multidestino"] = False

        movimientos["Ruta_Visual"] = (
            movimientos["Origen"].astype(str).str.strip()
            + " → "
            + movimientos["Intermedia"].astype(str).str.strip()
            + " → "
            + movimientos["Destino_Final"].astype(str).str.strip()
        )

        movimientos["FH_inicio"] = pd.to_datetime(movimientos["FH_inicio"], errors="coerce")
        movimientos["FH_fin"] = pd.to_datetime(movimientos["FH_fin"], errors="coerce")

        movimientos["Tiempo_Duracion"] = (
            (movimientos["FH_fin"] - movimientos["FH_inicio"])
            .dt.total_seconds()
            .div(60)
            .round(2)
        )

        movimientos["Estado_Movimiento"] = "OK"

        movimientos.loc[
            movimientos["Estado_Cantidad_Hijo"].astype(str).str.upper() != "OK",
            "Estado_Movimiento"
        ] = "REVISAR_CANTIDAD"

        movimientos["Observacion_Logica"] = "Movimiento normal Pick → Put."

        movimientos.loc[
            movimientos["Tipo_Match"].astype(str).str.upper() == "MERGE",
            "Observacion_Logica"
        ] = "Movimiento generado por múltiples PICK que consolidan en un PUT único."

        movimientos.loc[
            movimientos["Tipo_Match"].astype(str).str.upper() == "SPLIT",
            "Observacion_Logica"
        ] = "Movimiento generado por un PICK que se distribuye en varios PUT."

        movimientos.loc[
            movimientos["Tipo_Match"].astype(str).str.upper() == "MULTI_ITEM",
            "Observacion_Logica"
        ] = "Movimiento físico consolidado con varios ítems/lotes en la misma ruta."

        movimientos.loc[
            movimientos["Tipo_Match"].astype(str).str.upper() == "LP_CAMBIADA",
            "Observacion_Logica"
        ] = "Movimiento real con cambio de LP entre Pick y Put."

        movimientos["Resumen_Movimiento"] = (
            "Movimiento real de "
            + movimientos["Area_Movimiento"].astype(str)
            + " realizado por "
            + movimientos["Nombre"].astype(str)
            + " desde "
            + movimientos["Origen"].astype(str)
            + " hacia "
            + movimientos["Destino_Final"].astype(str)
            + ". Tipo: "
            + movimientos["Tipo_Match"].astype(str)
            + "."
        )
    else:
        movimientos = _crear_movimientos_vacios()

    # ============================================================
    # RANKINGS Y ANÁLISIS HORARIO
    # ============================================================

    update(84, "📊 Generando rankings y análisis horario...")

    ranking_bodega = _ranking_desde_movimientos(movimientos, "Bodega")

    ranking_despachos = _ranking_despachos_desde_ciclos(
        eventos_despacho=eventos_despacho,
        ciclos_despacho=ciclos_despacho,
        inactividad_despacho=inactividad_despacho
    )

    analisis_horas_despachos = _generar_analisis_horas_cargue_despachos(
        eventos_despacho=eventos_despacho,
        fecha_str=fecha_str,
        ini=ini,
        fin=fin,
        meta_ppt_despachos=meta_ppt_despachos
    )

    analisis_horas_bodega = _generar_analisis_horas_tarea_bodega(
        movimientos=movimientos,
        fecha_str=fecha_str,
        ini=ini,
        fin=fin,
        meta_ppt_bodega=meta_ppt_bodega
    )
    
    
    # ============================================================
    # AGREGAR COLUMNA ROL A ANALISIS MOVIMIENTOS BODEGA
    # ============================================================

    def _limpiar_rol(valor):
        texto = str(valor).strip()

        if texto.lower() in ["", "nan", "none", "<na>"]:
            return ""

        return texto


    def _construir_mapa_roles_bodega(df_maestro, movimientos):
        """
        Construye mapa de roles priorizando:
        1. Movimientos consolidados, si ya tienen Rol.
        2. Maestro personal por Nombre_norm.
        3. Maestro personal por Nombre.
        """

        mapa = {}

        # ============================================================
        # 1. Prioridad: movimientos consolidados
        # ============================================================

        if movimientos is not None and not movimientos.empty:
            mov_tmp = movimientos.copy()

            if "Rol" in mov_tmp.columns:
                if "Nombre_norm" not in mov_tmp.columns and "Nombre" in mov_tmp.columns:
                    mov_tmp["Nombre_norm"] = mov_tmp["Nombre"].apply(normalizar_nombre)

                for _, row in mov_tmp.iterrows():
                    nombre_norm = str(row.get("Nombre_norm", "")).strip()
                    rol = _limpiar_rol(row.get("Rol", ""))

                    if nombre_norm and rol:
                        mapa[nombre_norm] = rol

        # ============================================================
        # 2. Fallback: maestro personal
        # ============================================================

        if df_maestro is not None and not df_maestro.empty:
            maestro_tmp = df_maestro.copy()

            if "Rol" not in maestro_tmp.columns:
                maestro_tmp["Rol"] = ""

            if "Nombre_norm" not in maestro_tmp.columns and "Nombre" in maestro_tmp.columns:
                maestro_tmp["Nombre_norm"] = maestro_tmp["Nombre"].apply(normalizar_nombre)

            for _, row in maestro_tmp.iterrows():
                nombre_norm = str(row.get("Nombre_norm", "")).strip()
                rol = _limpiar_rol(row.get("Rol", ""))

                if nombre_norm and rol and nombre_norm not in mapa:
                    mapa[nombre_norm] = rol

        return mapa
    
    # ============================================================
    # APLICAR META PPT BODEGA POR ROL CONFIGURADA EN TURNO
    # ============================================================

    def _aplicar_meta_ppt_bodega_por_rol_desde_turno(
        df_analisis,
        turno_config,
        meta_general_bodega
    ):
        if df_analisis is None or df_analisis.empty:
            return df_analisis

        df_tmp = df_analisis.copy()

        metas_roles = {}

        if turno_config:
            metas_roles = turno_config.get("metas_ppt_bodega_roles", {})

        if not isinstance(metas_roles, dict):
            metas_roles = {}

        # Normalizar mapa de metas por rol
        metas_roles_norm = {}

        for rol, meta in metas_roles.items():
            rol_norm = normalizar_nombre(rol)

            if not rol_norm:
                continue

            try:
                metas_roles_norm[rol_norm] = float(meta)
            except Exception:
                metas_roles_norm[rol_norm] = 0

        col_nombre = None
        col_rol = None
        col_total = None
        col_meta = None
        col_pct = None

        for col in df_tmp.columns:
            col_norm = str(col).strip().lower()

            if col_norm in ["nombre", "nombres", "operador"]:
                col_nombre = col

            elif col_norm == "rol":
                col_rol = col

            elif col_norm in ["total movimiento", "total movimientos", "total"]:
                col_total = col

            elif col_norm in ["meta ppt", "meta"]:
                col_meta = col

            elif col_norm in ["% productividad", "productividad", "% cumplimiento"]:
                col_pct = col

        if col_total is None:
            return df_tmp

        if col_meta is None:
            df_tmp["Meta PPT"] = meta_general_bodega
            col_meta = "Meta PPT"

        if col_pct is None:
            df_tmp["% Productividad"] = 0
            col_pct = "% Productividad"

        if col_rol is None:
            # Si no hay rol, se mantiene meta general
            return df_tmp

        def _meta_por_fila(row):
            nombre_val = ""

            if col_nombre is not None:
                nombre_val = str(row.get(col_nombre, "")).strip()

            if nombre_val.upper() == "TOTAL":
                return ""

            rol_val = str(row.get(col_rol, "")).strip()
            rol_norm = normalizar_nombre(rol_val)

            if rol_norm in metas_roles_norm and metas_roles_norm[rol_norm] > 0:
                return metas_roles_norm[rol_norm]

            return float(meta_general_bodega or 0)

        df_tmp[col_meta] = df_tmp.apply(
            _meta_por_fila,
            axis=1
        )

        def _pct_por_fila(row):
            nombre_val = ""

            if col_nombre is not None:
                nombre_val = str(row.get(col_nombre, "")).strip()

            if nombre_val.upper() == "TOTAL":
                return ""

            try:
                total = float(row.get(col_total, 0) or 0)
            except Exception:
                total = 0

            try:
                meta = float(row.get(col_meta, 0) or 0)
            except Exception:
                meta = 0

            if meta <= 0:
                return 0

            return round(total / meta, 4)

        df_tmp[col_pct] = df_tmp.apply(
            _pct_por_fila,
            axis=1
        )

        return df_tmp
    
    analisis_horas_bodega = _aplicar_meta_ppt_bodega_por_rol_desde_turno(
        df_analisis=analisis_horas_bodega,
        turno_config=turno_config,
        meta_general_bodega=meta_ppt_bodega
    )
    

    def _agregar_rol_a_analisis_bodega(df_analisis, df_maestro, movimientos):
        if df_analisis is None or df_analisis.empty:
            return df_analisis

        df_tmp = df_analisis.copy()

        # Eliminar Rol previo si ya existe para recalcularlo bien
        columnas_originales = df_tmp.columns.tolist()

        for col in columnas_originales:
            if str(col).strip().lower() == "rol":
                df_tmp = df_tmp.drop(columns=[col])
                break

        # Buscar columna de nombres sin importar mayúsculas/minúsculas
        columna_nombre = None

        for col in df_tmp.columns:
            col_norm = str(col).strip().lower()

            if col_norm in ["nombre", "nombres", "operador"]:
                columna_nombre = col
                break

        if columna_nombre is None:
            # Si no encuentra columna nombre, deja Rol vacío al inicio
            df_tmp.insert(0, "Rol", "")
            return df_tmp

        mapa_roles = _construir_mapa_roles_bodega(
            df_maestro=df_maestro,
            movimientos=movimientos
        )

        def _rol_desde_nombre_analisis(nombre):
            nombre_norm = normalizar_nombre(nombre)
            return mapa_roles.get(nombre_norm, "")

        df_tmp["Rol"] = df_tmp[columna_nombre].apply(_rol_desde_nombre_analisis)
        
        def _lider_desde_nombre_analisis(nombre):
            nombre_norm = normalizar_nombre(nombre)
            return lider_por_nombre_norm.get(nombre_norm, "")

        df_tmp["Lider"] = df_tmp[columna_nombre].apply(_lider_desde_nombre_analisis)
        

        # Reordenar: NOMBRES | Rol | resto
        columnas = df_tmp.columns.tolist()

        for col_extra in ["Rol", "Lider"]:
            if col_extra in columnas:
                columnas.remove(col_extra)

        idx_nombre = columnas.index(columna_nombre)
        columnas.insert(idx_nombre + 1, "Rol")
        columnas.insert(idx_nombre + 2, "Lider")

        df_tmp = df_tmp[columnas].copy()

        return df_tmp


    analisis_horas_bodega = _agregar_rol_a_analisis_bodega(
        df_analisis=analisis_horas_bodega,
        df_maestro=df_maestro,
        movimientos=movimientos
    )
    
    

    inactividad_bodega = _calcular_inactividad_bodega(
        movimientos=movimientos,
        fecha_str=fecha_str,
        ini=ini,
        fin=fin,
        min_inactividad_reportable=25
    )

    # ============================================================
    # EXPORTAR EXCEL
    # ============================================================

    update(92, "💾 Exportando Excel...")

    columnas_orden = [
        "Item Number",
        "Número Lote",
        "Código Transacción",
        "Description",
        "ID Bodega",
        "Desde ID Ubicación",
        "A ID Ubicación",
        "LP",
        "ID Operador",
        "Nombre",
        "Fecha",
        "Tiempo",
        "Cantidad",
        "Tipo",
        "Estado",
        "Movimiento_ID",
        "Tipo_Match",
        "Pertenece_a_Pick_Padre",
        "Area_Movimiento"
    ]

    df0 = _ensure_columns(df0, columnas_orden)
    df_export = df0[columnas_orden].copy()

    if "Fecha" in df_export.columns:
        df_export["Fecha"] = pd.to_datetime(
            df_export["Fecha"],
            errors="coerce"
        ).dt.date

    turno_nombre = str(
        turno_config.get("nombre", "Turno") if turno_config else "Turno"
    ).strip()

    turno_nombre = (
        turno_nombre
        .replace(" ", "_")
        .replace("/", "-")
        .replace("\\", "-")
        .replace(":", "-")
    )

    turno_nombre = str(
        turno_config.get("nombre", "Turno") if turno_config else "Turno"
    ).strip()

    turno_nombre = (
        turno_nombre
        .replace(" ", "_")
        .replace("/", "-")
        .replace("\\", "-")
        .replace(":", "-")
    )

    nombre_salida = f"Productividad_{fecha_str.replace('/','-')}_{turno_nombre}_RESULTADO.xlsx"
    salida = os.path.join(carpeta, nombre_salida)

    sheet_trazabilidad = "Trazabilidad"
    sheet_movimientos = "Movimientos Consolidados"
    sheet_analisis_bodega = "Analisis movimientos Bodega"
    sheet_inactividad_bodega = "Inactividad Bodega"
    sheet_analisis_despachos = "Analisis cargues Despachos"
    sheet_inactividad_despacho = "Inactividad Despacho"

    columnas_movimientos_orden = [
        "Movimiento_Global_ID",
        "Movimiento_ID",
        "Pertenece_a_Pick_Padre",
        "Es_Movimiento_Padre",
        "Es_Hijo_Multidestino",
        "Tipo_Match",
        "Area_Movimiento",
        "ID Operador",
        "Nombre",
        "Rol",
        "Lider",
        "Origen",
        "Intermedia",
        "Destino_Final",
        "Ruta_Visual",
        "FH_inicio",
        "FH_fin",
        "Tiempo_Duracion",
        "Cantidad",
        "Resumen_Movimiento",
        "Observacion_Logica"
    ]

    movimientos_export = _ensure_columns(
        movimientos.copy(),
        columnas_movimientos_orden
    )

    movimientos_export = movimientos_export[columnas_movimientos_orden].copy()

    if exportar_excel:
        with pd.ExcelWriter(salida, engine="openpyxl") as writer:
            df_export.to_excel(
                writer,
                sheet_name=sheet_trazabilidad,
                index=False
            )

            movimientos_export.to_excel(
                writer,
                sheet_name=sheet_movimientos,
                index=False
            )

            inactividad_despacho.to_excel(
                writer,
                sheet_name=sheet_inactividad_despacho,
                index=False
            )

            inactividad_bodega.to_excel(
                writer,
                sheet_name=sheet_inactividad_bodega,
                index=False
            )

            analisis_horas_despachos.to_excel(
                writer,
                sheet_name=sheet_analisis_despachos,
                index=False
            )

            analisis_horas_bodega.to_excel(
                writer,
                sheet_name=sheet_analisis_bodega,
                index=False
            )

            _aplicar_estilo_excel_resultados(writer)
    else:
        salida = ""

    # ============================================================
    # HISTÓRICO SQLITE
    # ============================================================

    update(96, "🗄️ Guardando histórico SQLite...")

    info_hist = guardar_historico_turno(
        movimientos,
        fecha_str,
        ini,
        fin,
        carpeta,
        reemplazar=True
    )

    update(
        100,
        f"✅ Finalizado | Histórico SQLite: {info_hist['filas_guardadas']} movimientos"
    )

    if retornar_dataframes:
        return {
            "salida": salida,
            "trazabilidad": df_export.copy(),
            "movimientos": movimientos_export.copy(),
            "inactividad_bodega": inactividad_bodega.copy(),
            "inactividad_despacho": inactividad_despacho.copy(),
            "analisis_horas_bodega": analisis_horas_bodega.copy(),
            "analisis_horas_despachos": analisis_horas_despachos.copy(),
            "ranking_bodega": ranking_bodega.copy(),
            "ranking_despachos": ranking_despachos.copy()
        }

    return salida
# 05 - Reglas del Motor de Bodega

Reconstruir movimientos reales PICK/PUT. Descripciones: move (pick), move (put), move trailer (pick), move trailer (put), picking (pick), picking (put).

Fases: LP igual y LP cambiada/mixta. Movimiento válido: PICK+PUT, cantidades balanceadas, item/lote coinciden, ruta Origen→Intermedia→Destino, máximo 50 minutos, filas no reutilizadas. Separar bloques con saltos >45 minutos.

Prioridad: multidestino, Picking LP cambiada con retorno, retorno parcial productivo, merge multi-LP, Tipo fuerte, demanda PUT, split m2m, balance secuencial, origen/retorno, LP mixta, reintento final, correcciones.

Tipos: NORMAL, GLOBAL_MATCH, GLOBAL_MATCH_CORREGIDO, BALANCE_SECUENCIAL, SPLIT, MERGE, MULTI_ITEM, MULTI_PALLET, LP_CAMBIADA, MERGE_MULTI_LP, MULTIDESTINO_PADRE, MULTIDESTINO_HIJO, RETORNO_PARCIAL_PRODUCTIVO, PICKING_LP_CAMBIADA, REINTENTO_FINAL, RETORNO_ORIGEN.

Destino igual a Origen no cuenta productividad. Usar NO_PRODUCTIVO_RETORNO_ORIGEN según reglas de tiempo y ausencia de PUT productivo posterior.

# 06 - Reglas de Despachos e Inactividad

Despachos usa: loading (pick), loading (put), unload/unpick (pick), unload/unpick (put).

Clasificación: CARGUE_PICK_REAL, CARGUE_PUT_REAL, DEVOLUCION_PICK_REAL, DEVOLUCION_PUT_FINAL_REAL, DEVOLUCION_PUT_ESPEJO_SISTEMA. Puts espejo no cuentan productividad. LP no es llave principal.

Ciclos cargue: operador + puerta + 60 min. Ciclos devolución: operador + item + lote + 60 min.

Inactividad: Despachos mínimo 5 minutos sobre ciclos completos. Bodega mínimo 25 minutos sobre movimientos reales.

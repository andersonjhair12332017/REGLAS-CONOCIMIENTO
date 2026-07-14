# 03 - Reglas de Validación de Entradas

Validar Excel WMS, columnas, fechas, horas, rango máximo 31 días y existencia de registros dentro del turno.

Columnas base: Item Number, Número Lote, Código Transacción, Description, ID Bodega, Desde ID Ubicación, A ID Ubicación, LP, ID Operador, Nombre, Fecha, Tiempo, Cantidad, Tipo.

Variantes: Codigo Transaccion → Código Transacción; Numero Lote → Número Lote; Descripcion → Description; Desde Ubicacion → Desde ID Ubicación; A Ubicacion → A ID Ubicación.

Fechas WMS prioritariamente MM/DD/YYYY. Usuario usa dd/mm/aaaa. Turno Nocturno cruza fecha si hora_inicio > hora_fin.

Con turnos actuales: Mañana 07:00-14:00, Tarde 14:00-21:00, Noche 21:00-07:00. Noche cruza medianoche y la fecha seleccionada corresponde al cierre del turno.

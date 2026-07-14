# 07 - Reglas de Análisis Diario, Rango y Mensual

Diario: solicitar Excel, fecha dd/mm/aaaa y turno. Si no indica turno, preguntar Mañana, Tarde o Noche.

Rango/mensual: máximo 31 días. Leer Excel una sola vez. Si no indica turnos, procesar turnos activos actuales: Mañana, Tarde y Noche. No procesar Completo porque está inactivo salvo petición explícita.

Log: Fecha, Turno, Estado, Segundos, Filas_Turno, Filas_Trazabilidad, Filas_Movimientos, Filas_Analisis_Bodega, Error. Estados: OK, SIN_DATOS_TURNO, ERROR.

No crear archivos intermedios ni carpeta destino local. Entregar Excel descargable en conversación.

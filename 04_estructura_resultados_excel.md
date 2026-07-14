# 04 - Estructura de Resultados Excel

## Entrega
El agente debe entregar el Excel resultado como archivo descargable en la conversación. No debe pedir carpeta local.

## Excel diario
Nombre sugerido:
Productividad_DD-MM-AAAA_Turno_RESULTADO.xlsx

Hojas:
1. Trazabilidad
2. Movimientos Consolidados
3. Analisis movimientos Bodega
4. Inactividad Bodega
5. Analisis cargues Despachos
6. Inactividad Despacho

## Trazabilidad
Columnas:
- Item Number
- Número Lote
- Código Transacción
- Description
- ID Bodega
- Desde ID Ubicación
- A ID Ubicación
- LP
- ID Operador
- Nombre
- Fecha
- Tiempo
- Cantidad
- Tipo
- Estado
- Movimiento_ID
- Tipo_Match
- Pertenece_a_Pick_Padre
- Area_Movimiento

## Movimientos Consolidados
Columnas:
- Movimiento_Global_ID
- Movimiento_ID
- Pertenece_a_Pick_Padre
- Es_Movimiento_Padre
- Es_Hijo_Multidestino
- Tipo_Match
- Area_Movimiento
- ID Operador
- Nombre
- Rol
- Lider
- Origen
- Intermedia
- Destino_Final
- Ruta_Visual
- FH_inicio
- FH_fin
- Tiempo_Duracion
- Cantidad
- Resumen_Movimiento
- Observacion_Logica

## Excel mensual/rango
Nombre sugerido:
Productividad_MENSUAL_DD-MM-AAAA_a_DD-MM-AAAA_RESULTADO.xlsx

Hojas:
1. Resumen Mensual Operadores
2. Resumen Mensual Roles
3. Resumen Mensual Turnos
4. Resumen Estados
5. Estados Trazabilidad
6. Log Procesamiento
7. Trazabilidad
8. Movimientos Consolidados
9. Analisis movimientos Bodega
10. Inactividad Bodega
11. Analisis cargues Despachos
12. Inactividad Despacho

Cada hoja consolidada debe incluir al inicio:
- Fecha_Proceso
- Turno_Proceso
- Archivo_Proceso

## Resumen Mensual Operadores
Agrupar por ID Operador, Nombre y Rol. Calcular:
- Turnos_Procesados
- Movimientos
- Cantidad_Total

## Resumen Mensual Roles
Agrupar por Rol. Calcular:
- Operadores
- Movimientos
- Cantidad_Total

## Resumen Mensual Turnos
Agrupar por Fecha_Proceso y Turno_Proceso. Calcular:
- Operadores
- Movimientos
- Cantidad_Total

## Resumen Estados
Agrupar por Lider, Rol y Nombre. Mostrar:
- Nombre
- Rol
- Total
- %

Fórmula:
% = Total operador / máximo Total dentro del mismo Lider + Rol.

No contar MULTIDESTINO_PADRE. Si faltan datos usar SIN LIDER, SIN ROL o SIN NOMBRE. Esta hoja no debe tener autofiltros.

## Estados Trazabilidad
Conteo por Estado y Registros.

## Formato
Encabezados azules, texto blanco, bordes, congelar fila superior, ajustar anchos, porcentajes formateados, barras visuales de productividad, horas cero en rojo, tiempos en [h]:mm, cuadrículas ocultas. No crear tablas internas de Excel si pueden causar reparación del archivo.

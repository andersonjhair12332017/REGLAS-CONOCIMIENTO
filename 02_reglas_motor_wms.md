# 02 - Reglas del Motor WMS

## Propósito
Este documento contiene las reglas detalladas del motor de análisis WMS para reconstruir movimientos reales de Bodega y Despachos.

## Principio fundamental
Una fila del WMS no equivale necesariamente a un movimiento físico completo. El agente debe reconstruir flujos usando relaciones entre PICK y PUT, cantidades, item, lote, LP, ubicación, operador y tiempo.

## Columnas WMS base
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

## Normalización
Normalizar Description a minúsculas sin espacios redundantes. Normalizar fechas del WMS prioritariamente como MM/DD/YYYY. Construir FechaHora con Fecha + Tiempo.

## Bodega: descripciones válidas
- move (pick)
- move (put)
- move trailer (pick)
- move trailer (put)
- picking (pick)
- picking (put)

## Familias WMS
- TRAILER: move trailer.
- NORMAL: move o picking.
No mezclar TRAILER con NORMAL.

## Bloque temporal
Separar flujos cuando haya salto mayor a 45 minutos entre registros del mismo operador, familia WMS, item y lote.

## Validación de movimiento real
Un movimiento real requiere:
- al menos un PICK y un PUT;
- cantidad PICK = cantidad PUT;
- resumen item/lote igual;
- en fase LP igual: item/lote/LP igual;
- ruta lógica Origen → Intermedia → Destino;
- destino productivo diferente del origen;
- duración máxima aproximada de 50 minutos;
- filas no usadas en otro movimiento.

## Fases
1. LP igual: exige misma LP.
2. LP cambiada/mixta: permite LP distinta si item, lote, operador, intermedia, cantidad y tiempo coinciden.

## Orden de capas
1. Multidestino padre/hijo.
2. Picking LP cambiada con Move retorno posterior.
3. Retorno parcial productivo.
4. Merge multi-LP.
5. Match por Tipo WMS fuerte.
6. Demanda PUT agrupada.
7. Split many-to-many.
8. Balance secuencial.
9. Recuperación por origen/retorno.
10. LP mixta por bloque.
11. Reintento final.
12. Corrección fragmentación unidestino.

## Tipos de match
- NORMAL: un flujo simple pick→put balanceado.
- GLOBAL_MATCH: flujo completo balanceado detectado antes de fragmentar.
- GLOBAL_MATCH_CORREGIDO: corrección de fragmentación unidestino.
- BALANCE_SECUENCIAL: reconstrucción por balance acumulado.
- SPLIT: un pick se reparte en varios puts.
- MERGE: varios picks alimentan un put.
- MULTI_ITEM: varios ítems/lotes en la misma ruta.
- MULTI_PALLET: varias LP en el flujo.
- LP_CAMBIADA: LP Pick y LP Put difieren.
- MERGE_MULTI_LP: varios picks de diferentes LP/orígenes consolidan en un put con otra LP.
- MULTIDESTINO_PADRE: flujo padre hacia varios destinos.
- MULTIDESTINO_HIJO: destino productivo individual del multidestino.
- RETORNO_PARCIAL_PRODUCTIVO: una parte retorna y otra sí produce.
- PICKING_LP_CAMBIADA: picking productivo con cambio de LP.
- REINTENTO_FINAL: emparejamiento simple final de picks/puts huérfanos.
- RETORNO_ORIGEN: retorno no productivo.

## Retornos
Si Destino_Final = Origen, no cuenta como productividad. Marcar NO_PRODUCTIVO_RETORNO_ORIGEN solo si PUT ocurre después del PICK, está dentro del tiempo válido y no hay PUT productivo posterior compatible.

## Fragmentación
Si varios PICK y PUT parciales realmente forman un único flujo unidestino balanceado, consolidar como GLOBAL_MATCH_CORREGIDO y eliminar parciales previos.

## Estados protegidos
- NO_PRODUCTIVO_RETORNO_ORIGEN
- MOV_INVALIDO_A_A
- CROSSDOCK_IGNORADO
- PERSONAL_EXCLUIDO
- PERSONAL_NO_MAESTRO
- DUPLICADO_WMS_IGNORADO
No sobrescribir estos estados como SIN_MATCH.

## SIN_MATCH
Marcar SIN_MATCH solo para registros candidatos de Bodega/Mixto que no fueron identificados como movimiento y no tienen estado protegido.

## Despachos
Eventos válidos:
- loading (pick) → CARGUE_PICK_REAL
- loading (put) → CARGUE_PUT_REAL
- unload/unpick (pick) → DEVOLUCION_PICK_REAL
- unload/unpick (put) → DEVOLUCION_PUT_FINAL_REAL o DEVOLUCION_PUT_ESPEJO_SISTEMA

Los puts espejo no cuentan productividad. LP no es llave principal de relación en Despachos.

## Ciclos despacho
Cargue: agrupar por operador + puerta relación + ventana máxima 60 minutos.
Devolución: agrupar por operador + item + lote + ventana máxima 60 minutos.

## Inactividad
Despachos: mínimo reportable 5 minutos sobre ciclos completos.
Bodega: mínimo reportable 25 minutos sobre movimientos consolidados.

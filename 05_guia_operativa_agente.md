# 05 - Guía Operativa del Agente

## Flujo análisis diario
1. Recibir Excel WMS.
2. Validar columnas, fechas y horas.
3. Solicitar fecha dd/mm/aaaa.
4. Solicitar turno o usar turno activo seleccionado.
5. Validar que haya registros en el turno.
6. Cargar maestro personal.
7. Cruzar operadores por ID/nombre.
8. Aplicar reglas de Bodega y Despachos.
9. Calcular productividad, inactividad y análisis horario.
10. Generar Excel diario.
11. Entregar archivo descargable.

## Flujo análisis rango/mensual
1. Recibir Excel WMS.
2. Solicitar fecha inicio y fecha fin.
3. Validar máximo 31 días.
4. Leer Excel una sola vez.
5. Filtrar cada fecha y turno en memoria.
6. Procesar todos los turnos activos.
7. No detenerse por operadores no encontrados; reportarlos.
8. Si un turno falla, registrar ERROR y continuar.
9. Consolidar hojas.
10. Generar Excel mensual/rango.
11. Entregar archivo descargable.

## Respuesta final esperada
Análisis finalizado correctamente.

Archivo resultado:
[archivo descargable]

Resumen:
- Tipo de análisis
- Fecha o rango procesado
- Turnos procesados
- Procesos OK
- Turnos sin datos
- Errores
- Movimientos consolidados
- Registros SIN_MATCH
- Operadores no encontrados en maestro

## Límites
No inventar datos. Si falta columna crítica, detener e informar. Si no puede generar archivo por limitación técnica, explicarlo. Si no hay maestro, solicitarlo o advertir que no se asignarán roles/líderes. Si el rango supera 31 días, pedir dividir o cargar archivo filtrado. Si un movimiento es dudoso, usar SIN_MATCH o revisión requerida.

## Frases de ejemplo
- Analiza este Excel WMS para el turno Día del 10/06/2026.
- Procesa este archivo WMS para el rango del 01/06/2026 al 30/06/2026.
- Genera la productividad mensual de Bodega y Despachos.
- Valida este Excel WMS y dime qué fechas contiene.
- Identifica los operadores que no están en el maestro personal.
- Explícame por qué estos movimientos quedaron como SIN_MATCH.
- Genera el Resumen Estados por líder y rol.
- Calcula la productividad por operador, rol y líder.
- Analiza la inactividad de Bodega y Despachos.
- Revisa si hay movimientos con LP cambiada, multidestino o retorno al origen.

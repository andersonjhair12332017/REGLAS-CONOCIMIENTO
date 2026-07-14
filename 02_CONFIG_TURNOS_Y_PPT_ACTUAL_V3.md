# 02 - Configuración Actual Real de Turnos y Metas PPT V3

## Fuente
Este archivo se construyó a partir del `turnos_config.json` compartido por el usuario.

## Turnos activos actuales

### Mañana
- Nombre original detectado en archivo: `MaĆ±ana`.
- Nombre normalizado recomendado: `Mañana`.
- Aliases reconocidos: Mañana, Manana, MaĆ±ana, mañana, manana, turno mañana.
- hora_inicio: 07:00
- hora_fin: 14:00
- activo: true
- meta_ppt_bodega: 175.0
- meta_ppt_despachos: 1050.0

Metas PPT Bodega por rol:
- Montacarguita: 100.0
- Montacarguista: 100.0
- Pick and drop: 130.0
- Pick and Drop: 130.0
- Trilateral: 175.0

### Tarde
- hora_inicio: 14:00
- hora_fin: 21:00
- activo: true
- meta_ppt_bodega: 175.0
- meta_ppt_despachos: 1050.0

Metas PPT Bodega por rol:
- Montacarguita: 100.0
- Montacarguista: 100.0
- Pick and drop: 130.0
- Pick and Drop: 130.0
- Trilateral: 175.0

### Noche
- hora_inicio: 21:00
- hora_fin: 07:00
- activo: true
- meta_ppt_bodega: 250.0
- meta_ppt_despachos: 1500.0

Metas PPT Bodega por rol:
- Montacarguita: 120.0
- Montacarguista: 120.0
- Pick and drop: 155.0
- Pick and Drop: 155.0
- Trilateral: 200.0

## Turno inactivo

### Completo
- hora_inicio: 07:00
- hora_fin: 06:59
- activo: false
- meta_ppt_bodega: 600.0
- meta_ppt_despachos: 3150.0

Metas PPT Bodega por rol:
- Montacarguita: 320.0
- Montacarguista: 320.0
- Pick and drop: 415.0
- Pick and Drop: 415.0
- Trilateral: 550.0

## Reglas obligatorias

1. En análisis diario, si el usuario no indica turno, preguntar si desea Mañana, Tarde o Noche.
2. En análisis por rango o mensual, si el usuario no indica turno, procesar todos los turnos activos: Mañana, Tarde y Noche.
3. No procesar Completo salvo que el usuario lo pida explícitamente, porque está inactivo.
4. Si el usuario escribe MaĆ±ana, Manana o Mañana, debe interpretarse como Mañana.
5. Si el rol del maestro aparece como Montacarguista, usar la misma meta de Montacarguita, porque el archivo tiene `Montacarguita` pero el rol operativo esperado suele ser `Montacarguista`.
6. Si el rol aparece como Pick and Drop o Pick and drop, tratarlos como equivalentes.
7. Si se carga un nuevo `turnos_config.json`, ese archivo reemplaza esta configuración.

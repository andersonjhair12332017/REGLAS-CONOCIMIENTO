# Guía Para Usar Los Archivos .py Como Fuente De Conocimiento Del Agente WMS

## 1. Objetivo

Esta guía indica cómo debe usar el agente los archivos `.py` del software de productividad WMS como fuente principal de conocimiento técnico y operativo.

La idea es que el agente no dependa únicamente de documentos resumidos, sino que pueda consultar los archivos originales del software para entender:

- validaciones,
- turnos,
- maestro de personal,
- lógica WMS,
- movimientos Bodega,
- Despachos,
- inactividad,
- análisis mensual,
- estructura del Excel resultado,
- histórico,
- interfaz y flujo operativo.

El agente debe tratar los archivos `.py` como documentación técnica fuente. Si existe conflicto entre un resumen y un archivo `.py`, debe priorizar el archivo `.py`, especialmente el módulo que implementa directamente la función solicitada.

---

## 2. Archivos fuente que debe cargar el agente

Subir como conocimiento todos estos archivos `.py`:

1. `app.py`
2. `wms_processor.py`
3. `procesamiento_rango.py`
4. `wms_utils.py`
5. `personal_config.py`
6. `turnos_config.py`
7. `turnos_admin_ui.py`
8. `maestro_personal_editor.py`
9. `wms_validators.py`
10. `historico_sqlite.py`

También subir:

11. `turnos_config.json`
12. `maestro_personal.xlsx`, si la plataforma lo permite como conocimiento o archivo de referencia.
13. Documentos guía `.md` generados para el agente.

---

## 3. Regla de prioridad entre archivos

Cuando el agente deba responder o ejecutar una tarea, debe consultar los archivos en este orden según el tema:

### Análisis diario WMS
Prioridad:

1. `wms_processor.py`
2. `wms_utils.py`
3. `personal_config.py`
4. `turnos_config.py`
5. `wms_validators.py`
6. `app.py`

### Análisis por rango o mensual desde Excel WMS
Prioridad:

1. `procesamiento_rango.py`
2. `wms_processor.py`
3. `wms_utils.py`
4. `turnos_config.py`
5. `personal_config.py`
6. `app.py`

### Turnos y metas PPT
Prioridad:

1. `turnos_config.json`
2. `turnos_config.py`
3. `turnos_admin_ui.py`
4. `app.py`

### Maestro de personal, roles y líderes
Prioridad:

1. `personal_config.py`
2. `maestro_personal_editor.py`
3. `wms_processor.py`
4. `wms_utils.py`

### Despachos, ciclos e inactividad
Prioridad:

1. `wms_utils.py`
2. `wms_processor.py`
3. `procesamiento_rango.py`

### Validación de Excel, fechas y turnos
Prioridad:

1. `wms_validators.py`
2. `wms_utils.py`
3. `app.py`
4. `procesamiento_rango.py`

### Histórico SQLite y reporte mensual desde histórico
Prioridad:

1. `historico_sqlite.py`
2. `app.py`

### Flujo visual del software
Prioridad:

1. `app.py`
2. `turnos_admin_ui.py`
3. `maestro_personal_editor.py`

---

## 4. Descripción de cada archivo

### `app.py`

Es el orquestador visual del software.

Define:

- carga inicial de turnos activos,
- selección de Excel WMS,
- selección de fecha,
- selección de turno,
- botones principales,
- análisis diario,
- análisis por rango,
- reporte mensual desde histórico,
- administración de turnos,
- administración de personal,
- validación inmediata del Excel y fecha,
- apertura del reporte generado.

El agente debe usar este archivo para entender el flujo general que veía el usuario en la aplicación de escritorio.

Acciones principales identificadas:

- ANÁLISIS DIARIO,
- ANÁLISIS POR RANGO DE FECHA,
- GENERAR REPORTE MENSUAL,
- Administrar Turnos y Metas PPT,
- Administrar Personal,
- Abrir Reporte Generado.

Adaptación para agente:

- No pedir carpeta destino.
- Entregar archivo descargable en conversación.
- Mantener la misma lógica de validaciones y flujos.

---

### `wms_processor.py`

Es el motor principal del análisis diario.

Debe usarse como fuente principal para:

- construcción de movimientos reales de Bodega,
- procesamiento Pick/Put,
- LP igual,
- LP cambiada,
- LP mixta,
- split,
- merge,
- multidestino,
- retorno parcial productivo,
- retorno al origen,
- balance secuencial,
- reintento final,
- correcciones de fragmentación,
- trazabilidad,
- asignación de Movimiento_ID,
- asignación de Rol y Lider,
- exportación de Excel diario,
- guardado histórico.

Regla para el agente:

Cuando el usuario pregunte por una clasificación de movimiento, por qué algo fue SIN_MATCH, por qué un movimiento fue LP_CAMBIADA, MULTIDESTINO, SPLIT, MERGE, RETORNO_ORIGEN o GLOBAL_MATCH_CORREGIDO, el agente debe consultar este archivo.

---

### `procesamiento_rango.py`

Es el motor de análisis por rango o mensual desde el Excel WMS.

Define:

- límite máximo de 31 días,
- lectura única del Excel,
- normalización única,
- filtrado de cada fecha/turno en memoria,
- uso de `procesar_archivo` con `exportar_excel=False`,
- consolidación de hojas,
- log de procesamiento,
- Resumen Mensual Operadores,
- Resumen Mensual Roles,
- Resumen Mensual Turnos,
- Resumen Estados,
- Estados Trazabilidad,
- Excel mensual final.

Regla para el agente:

Para análisis mensual/rango, el agente no debe procesar cada turno como archivo aislado ni generar archivos intermedios. Debe consolidar todo en un único resultado descargable.

---

### `wms_utils.py`

Contiene utilidades base.

Define:

- asegurar columnas,
- normalizar encabezados,
- parsear fechas,
- parsear horas,
- construir FechaHora,
- normalizar Description,
- segmentar por tiempo,
- rankings,
- eventos de Despachos,
- ciclos de Despachos,
- inactividad de Bodega,
- inactividad de Despachos,
- análisis horario,
- formato Excel.

Regla para el agente:

Cuando la pregunta sea sobre fechas, horas, Despachos, ciclos, inactividad, análisis horario, rankings o formato del Excel, debe consultar este archivo.

---

### `personal_config.py`

Contiene la lógica del maestro de personal.

Define:

- ruta de `maestro_personal.xlsx`,
- columnas obligatorias,
- áreas válidas,
- normalización de nombres,
- normalización de ID Operador,
- carga del maestro,
- validación del maestro,
- matching por ID,
- matching por nombre exacto,
- matching por tokens,
- matching fuzzy,
- construcción de asignación de operadores por turno.

Regla para el agente:

El agente debe usar este archivo para resolver cualquier duda sobre operadores, roles, líderes, áreas, activos, excluidos, faltantes y maestro personal.

---

### `turnos_config.py`

Contiene la estructura y persistencia de turnos.

Define:

- archivo `turnos_config.json`,
- turnos por defecto,
- carga de turnos,
- guardado de turnos,
- normalización de activo,
- normalización de metas PPT,
- metas por rol.

Regla para el agente:

El agente debe usar este archivo junto con `turnos_config.json` para entender cómo se deben interpretar los turnos y metas PPT.

---

### `turnos_admin_ui.py`

Contiene la lógica visual para administrar turnos y metas PPT.

Define:

- tabla de turnos,
- tabla de metas por rol,
- creación de turno,
- edición de turno,
- eliminación de turno,
- activación/inactivación,
- metas PPT por Bodega y Despachos,
- metas PPT Bodega por rol,
- obtención de roles desde maestro.

Regla para el agente:

Debe usarse para entender cómo el usuario administra turnos y cómo se derivan roles disponibles desde maestro.

---

### `maestro_personal_editor.py`

Contiene el visor del maestro de personal.

Define:

- apertura de `maestro_personal.xlsx`,
- visualización de hojas Bodega y Despachos,
- recarga del maestro,
- conteo de registros,
- edición manual externa en Excel.

Regla para el agente:

El agente no debe modificar automáticamente el maestro salvo instrucción explícita. Debe tratarlo como fuente externa editable por el usuario.

---

### `wms_validators.py`

Contiene validaciones rápidas previas.

Define:

- interpretación de fechas WMS como MM/DD/YYYY,
- validación rápida de fecha y turno,
- análisis rápido del rango de fechas del Excel,
- bloqueo por rango mayor a 31 días,
- autoselección de fecha máxima,
- validación inmediata de fecha contra cache.

Regla para el agente:

Antes de procesar, debe aplicar estas validaciones para evitar análisis con fecha fuera de rango o sin registros de turno.

---

### `historico_sqlite.py`

Contiene histórico local SQLite.

Define:

- base `historico_productividad.db`,
- tabla `historico_movimientos`,
- guardado por turno,
- reemplazo de turno reprocesado,
- consulta mensual,
- ranking mensual,
- reporte mensual desde histórico.

Regla para el agente:

Si el agente maneja histórico, debe usar este archivo para entender cómo construir reportes desde histórico. Si no tiene capacidad de SQLite, debe informar que esta función requiere herramienta o backend conectado.

---

## 5. Cómo debe usar el agente los archivos `.py`

El agente debe:

1. Consultar el archivo correspondiente según la intención del usuario.
2. No inventar reglas si existen en código.
3. Priorizar el módulo que implementa la función.
4. Usar los documentos `.md` solo como guía resumida.
5. Usar los `.py` como fuente técnica principal.
6. Responder en lenguaje operativo, no como explicación de código salvo que el usuario lo pida.
7. Si necesita generar Excel, seguir las estructuras definidas por `wms_processor.py`, `procesamiento_rango.py` y `wms_utils.py`.

---

## 6. Guía específica para turnos reales

El archivo `turnos_config.json` compartido por el usuario contiene:

- Mañana: 07:00 a 14:00, activo.
- Tarde: 14:00 a 21:00, activo.
- Noche: 21:00 a 07:00, activo.
- Completo: 07:00 a 06:59, inactivo.

El agente debe usar esos turnos antes que los defaults de `turnos_config.py`.

Si el usuario no indica turno en análisis mensual/rango, debe procesar Mañana, Tarde y Noche.

No debe procesar Completo salvo que el usuario lo pida explícitamente.

---

## 7. Qué subir como conocimiento

Subir estos archivos `.py` completos:

```text
app.py
wms_processor.py
procesamiento_rango.py
wms_utils.py
personal_config.py
turnos_config.py
turnos_admin_ui.py
maestro_personal_editor.py
wms_validators.py
historico_sqlite.py
```

Subir también:

```text
turnos_config.json
maestro_personal.xlsx
```

Si `maestro_personal.xlsx` no se puede subir como conocimiento útil, permitir que el usuario lo cargue durante la conversación.

---

## 8. Frase de instrucción para el agente

Agregar en instrucciones del agente:

```text
Cuando necesites reglas detalladas, consulta los archivos .py cargados como fuente principal. Usa los documentos .md como guía, pero prioriza el código fuente del módulo correspondiente. Para turnos y metas PPT, consulta primero turnos_config.json. Para motor WMS, consulta wms_processor.py. Para rango/mensual, consulta procesamiento_rango.py. Para maestro, consulta personal_config.py. Para Despachos, fechas, horas, inactividad y formato Excel, consulta wms_utils.py. Para validaciones previas, consulta wms_validators.py.
```

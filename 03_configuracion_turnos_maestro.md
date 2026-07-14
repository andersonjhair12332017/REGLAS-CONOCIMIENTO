# 03 - Configuración de Turnos, Metas PPT y Maestro Personal

## Turnos
Cada turno debe contener:
- nombre
- hora_inicio
- hora_fin
- activo
- meta_ppt_bodega
- meta_ppt_despachos
- metas_ppt_bodega_roles

Si no existe configuración, usar:
- Día: 07:30 a 19:30
- Noche: 19:30 a 07:30

PPT significa Productividad por Turno.

## Activo/inactivo
Valores activos: true, 1, si, sí, s, activo, yes, y.
Valores inactivos: false, 0, no, n, inactivo.
Si no es reconocible, asumir activo.

## Metas PPT
Normalizar metas como números positivos. Vacío, inválido o negativo = 0.
Priorizar meta_ppt_bodega y meta_ppt_despachos. Si aparecen meta_ppd_bodega o meta_ppd_despachos, tratarlos como compatibilidad antigua.

## Metas por rol
Para Bodega, usar metas_ppt_bodega_roles. Roles base:
- Montacarguista
- Pick and Drop
- Trilateral

Si el rol tiene meta específica, usarla. Si no, usar meta_ppt_bodega general.

## Maestro personal
Archivo esperado: maestro_personal.xlsx.
Columnas mínimas:
- ID Operador
- Nombre
- Area
- Rol
- Lider
- Activo
- Observacion

Áreas válidas:
- Bodega
- Despachos
- Mixto
- Excluir

## Normalización de nombres
Eliminar tildes, convertir a minúsculas, quitar caracteres especiales y reducir espacios múltiples.

## Normalización ID Operador
Convertir a texto limpio en mayúsculas. Valores nan, none, <na> o vacío se tratan como vacío.

## Variantes de columnas maestro
- ID_Operador, ID operador, Id Operador, id_operador → ID Operador
- nombre, nombres → Nombre
- area, área → Area
- rol, cargo, funcion, función → Rol
- lider, líder, jefe, supervisor, leader → Lider
- observacion, observación, obs, comentario → Observacion

## Validación maestro
El maestro no debe estar vacío, debe tener columnas obligatorias, no debe tener nombres duplicados normalizados y no debe tener áreas inválidas.

## Matching contra maestro
Prioridad:
1. ID Operador normalizado exacto.
2. Nombre normalizado exacto.
3. Tokens del nombre en distinto orden.
4. Subconjunto de tokens con coincidencia alta.
5. Fuzzy con score alto y sin ambigüedad.

Si operador está inactivo, tratar como Excluir. Si no aparece, marcar PERSONAL_NO_MAESTRO y excluir de productividad.

En análisis mensual/rango, no detenerse por faltantes del maestro; continuar y reportarlos.

## Mixto
Operador Mixto puede contar en Bodega o Despachos según el tipo de operación.

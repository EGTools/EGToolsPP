# Aplicar compatibilidad

**Ubicación**: pestaña `EGTools++` de la cinta → grupo `EGTools` → menú desplegable del botón dividido

Cuando un libro creado en un Excel moderno se abre en versiones antiguas de Excel,
las funciones nuevas aparecen como tokens rotos del tipo `_xlfn.XLOOKUP` y producen
errores `#NAME?`. Este botón convierte de una sola vez, en todas las fórmulas del
libro activo, los tokens de funciones modernas/nativas a los nombres de función
compatibles de EGTools, de modo que las fórmulas sigan calculándose igual en
versiones antiguas de Excel.

## Procedimiento

1. **Confirmación de guardado** — Como las fórmulas se modifican directamente, solo
   se ejecuta sobre un libro guardado. Si no está guardado, se pregunta si desea
   guardarlo antes de continuar (si rechaza o cancela, se aborta).
2. **Análisis previo** — Se recorren en modo de solo lectura las fórmulas de todas
   las hojas y los nombres definidos, y se muestra un resumen con el número de
   coincidencias por función objetivo. Si no hay nada que convertir, se informa y se
   termina.
3. **Copia de seguridad** — Si elige continuar, justo antes de la conversión se
   guarda automáticamente un archivo de copia de seguridad (en la misma carpeta que
   el libro, `nombreoriginal_egbak_<marca de tiempo>.<extensión>`). Si la copia de
   seguridad falla, no se convierte nada.
4. **Conversión y resultado** — Se muestra un resumen con el número de ubicaciones
   convertidas y la ruta del archivo de copia de seguridad.

## Reglas de conversión

| Objetivo | Conversión |
|---|---|
| Funciones de compatibilidad generales | `_xlfn.FUNCIÓN` / `_xlfn._xlws.FUNCIÓN` / `EG.FUNCIÓN` → `FUNCIÓN` (nombre de reemplazo directo de EGTools). Los nombres en conflicto SORT/FILTER/LET se convierten en `xSORT`/`xFILTER`/`xLET` |
| IMAGE | `_xlfn.IMAGE` / `EG.IMAGE` → `IMAGE` |
| GROUPBY / PIVOTBY | Cambia el nombre de la función a `GROUPBY`/`PIVOTBY` y añade comillas al argumento del agregador (`SUM` o `_xleta.SUM` → `"SUM"`; solo para los 16 agregadores con compatibilidad nativa — con LAMBDA u otros, la fórmula no se toca). También se convierte aunque le sigan argumentos opcionales (el orden de argumentos es idéntico al nativo) |
| `_xlpm.*` (parámetros de LET/LAMBDA) | Prefijo interno con el que el Excel moderno almacena los **nombres de parámetro** de LET/LAMBDA. En versiones antiguas aparece como `=_xlfn.LET(_xlpm.x, …)` y, como Excel se niega a reescribir cualquier fórmula que contenga `_xlpm.`, el prefijo se elimina (`_xlpm.x` → `x`). En el resumen de conversión se contabiliza como `_xlpm.*` |
| `_xll.*` / `_xludf.*` (prefijos de almacenamiento de UDF de complementos) | Elimina los prefijos que quedan en libros abiertos en un entorno donde las funciones EGTools no estaban registradas (sobre todo en los **nombres definidos**) y vuelve a vincular la llamada con la función EGTools registrada (`_xll.xSORT(…)` → `xSORT(…)`). Las funciones de otros complementos no se tocan. En el resumen de conversión se contabiliza como `_xll.*` |

El interior de los literales de cadena y las partes de otros identificadores no se
modifican (la sustitución se hace por identificador completo).

## Notas

- Para deshacer la conversión, ejecute [Restaurar funciones nativas](restore-native.md) o abra el archivo de copia de seguridad.
- La lista de funciones objetivo de la conversión se gestiona en una única fuente
  junto con la lista de registro de funciones del complemento, por lo que siempre
  coinciden.

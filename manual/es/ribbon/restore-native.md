# Restaurar funciones nativas

**Ubicación**: pestaña `EGTools++` de la cinta → grupo `EGTools` → menú desplegable del botón dividido

Cuando un libro convertido a las funciones compatibles de EGTools se lleva de vuelta
a un Excel moderno, este botón restaura de una sola vez, en todas las fórmulas del
libro activo, los nombres de función de EGTools (`EG.*`, etc.) a las funciones
nativas integradas. Tras la restauración, esas fórmulas se calculan sin necesidad
del complemento.

## Procedimiento

Sigue el mismo procedimiento de seguridad que [Aplicar compatibilidad](apply-compat.md):
confirmación de guardado → análisis previo (resumen del número de coincidencias por
función objetivo) → copia de seguridad automática
(`nombreoriginal_egbak_<marca de tiempo>.<extensión>`) → conversión → resumen del resultado.

## Reglas de conversión

| Objetivo | Conversión |
|---|---|
| Funciones de compatibilidad generales | `EG.FUNCIÓN` / `xFUNCIÓN` (xSORT/xFILTER/xLET) → `FUNCIÓN` (nativa). Las celdas que aún contienen el nombre de reemplazo directo usado en versiones antiguas también se vuelven a tokenizar para forzar que las calcule la función nativa |
| IMAGE | **Según el origen**: si el primer argumento es un origen `https://`, se mantiene la `IMAGE` nativa; en los demás casos (ruta local, `http://`) se convierte a `EG.IMAGE`. Los argumentos de referencia o fórmula se juzgan por **el valor evaluado en el momento de la conversión** (si la evaluación falla, da error o no es texto, se usa `EG.IMAGE` por seguridad). La IMAGE nativa solo admite https, por lo que si tras la conversión el valor referenciado cambia a una ruta local, esa celda mostrará `#VALUE!` — `EG.IMAGE`, en cambio, funciona con cualquier origen |
| GROUPBY / PIVOTBY | Devuelve el nombre de la función al nativo y quita las comillas del argumento del agregador (`"SUM"` → `SUM`; solo para los 16 agregadores con compatibilidad nativa — los agregadores ampliados de EGTools, como TEXTJOIN, los rechaza la función nativa, así que esas fórmulas no se tocan). También se convierte aunque le sigan argumentos opcionales |
| `_xlpm.*` (parámetros de LET/LAMBDA) | Elimina el prefijo de almacenamiento de parámetros que queda en las fórmulas que pasaron por versiones antiguas (`xLET(_xlpm.x, …)` → `LET(x, …)`). Excel se niega a escribir fórmulas que contengan `_xlpm.`, así que sin esta eliminación la celda no se puede restaurar. En el resumen de conversión se contabiliza como `_xlpm.*` |
| `@` (intersección implícita) | La `@` que el Excel moderno antepone al abrir archivos de versiones antiguas se elimina también cuando está **justo delante de un nombre de función que se restaura a nativo** (`TAKE(@xSORT(…),3)` → `TAKE(SORT(…),3)`). Si no se eliminara, `@SORT` colapsaría la matriz en un único valor y el resultado sería incorrecto. La `@` delante de referencias de celda o de funciones que no son objeto de conversión se conserva. En el resumen de conversión se contabiliza como `@ (intersección implícita)` |
| `_xll.*` / `_xludf.*` (prefijos de almacenamiento de UDF de complementos) | Elimina los prefijos que quedan como rastro de haber abierto el libro en un entorno donde las funciones del complemento no estaban registradas (`_xll.xSORT(…)` → `SORT(…)`). En particular, corrige el problema por el que estos prefijos quedaban en las fórmulas de los **nombres definidos** e impedían su conversión. Solo se eliminan los prefijos delante de funciones EGTools; las funciones de otros complementos no se tocan. En el resumen de conversión se contabiliza como `_xll.*` |
| Funciones de implantación aún no generalizada (IMPORTTEXT/IMPORTCSV) | Las funciones ya introducidas pero todavía no distribuidas de forma general (solo canales de vista previa de 365) se **comprueban antes de restaurar para verificar si la función nativa existe realmente en este Excel**, y solo se convierten a nativa si la comprobación es positiva. Si no se confirma, se mantiene la implementación de EGTools, igual que con IMAGE (`IMPORTTEXT` → `EG.IMPORTTEXT`). Las funciones ya distribuidas de forma general se convierten siempre sin esta comprobación |

## Notas

- Si el Excel en uso no dispone de la función nativa correspondiente (versiones
  antiguas), las fórmulas restauradas darán `#NAME?`; use este botón al volver a un
  Excel moderno.
- Para convertir de nuevo el libro para versiones antiguas, ejecute
  [Aplicar compatibilidad](apply-compat.md).

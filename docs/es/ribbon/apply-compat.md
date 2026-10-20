# Aplicar compatibilidad

**Ubicación**: pestaña `EGTools` de la cinta → grupo `EGTools` → menú desplegable del botón dividido

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
| `_xlfn._ARRAY` / `_xlfn._ARRAYROW` (envoltorios de matriz en celda) | Funciones internas que la beta de Microsoft 365 guarda cuando encierra un resultado desbordado en una sola celda. En hosts sin la función nativa dan `#NAME?`, así que se reescriben como `ARRAYTOTEXT(…,1)`, que devuelve los mismos valores en notación de texto: `_xlfn._ARRAYROW(a)` → `a`, `_xlfn._ARRAYROW(a,b,…)` → `HSTACK(a,b,…)`, `_xlfn._ARRAY(r)` → `ARRAYTOTEXT(r,1)`, `_xlfn._ARRAY(r1,r2,…)` → `ARRAYTOTEXT(VSTACK(r1,r2,…),1)`. Ejemplo: `=_xlfn._ARRAY(_xlfn._ARRAYROW(FILTER(A:A,B:B=1)))` → `=ARRAYTOTEXT(FILTER(A:A,B:B=1),1)`. El texto resultante coincide con la notación estricta de [Celdas de lista/matriz → texto](list-to-text.md) (`{"a",1;"b",2}`), por lo que HAS/HASANY/HASALL/FLATTEN lo leen tal cual. Se procesa de forma recursiva de dentro hacia fuera; un `_ARRAY` anidado se convierte en un elemento de texto entrecomillado dentro del texto exterior. En el resumen de conversión se contabiliza como `_ARRAY`/`_ARRAYROW` |
| Funciones de implantación escalonada (HAS/HASANY/HASALL/FLATTEN, IMPORTTEXT/IMPORTCSV) | Si este Excel realmente no dispone de la función nativa (se comprueba en tiempo de ejecución), `_xlfn.HAS` / `HAS` sin prefijo se convierten al **nombre EG. registrado en este host**, p. ej. `EG.HAS` (evita `#NAME?` en los canales no beta de Microsoft 365). En versiones antiguas (2024 y anteriores) se mantiene, como hasta ahora, el nombre sin prefijo (`HAS`). Un libro convertido a nombres EG. pasa a nombres sin prefijo si se vuelve a convertir más tarde en una versión antigua |

El interior de los literales de cadena y las partes de otros identificadores no se
modifican (la sustitución se hace por identificador completo).

## Notas

- Para deshacer la conversión, ejecute [Restaurar funciones nativas](restore-native.md) o abra el archivo de copia de seguridad.
- La lista de funciones objetivo de la conversión se gestiona en una única fuente
  junto con la lista de registro de funciones del complemento, por lo que siempre
  coinciden.
- Las propias celdas de lista/matriz creadas en la beta no las trata este comando. Para que
  las fórmulas como HAS que hacen referencia a una lista en celda no den `#VALUE!`,
  **ejecute primero [Celdas de lista/matriz → texto](list-to-text.md)** y después este comando.

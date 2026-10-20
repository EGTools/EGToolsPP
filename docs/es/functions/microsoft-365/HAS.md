# HAS

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Microsoft 365 (novedad de 2026)

Devuelve VERDADERO si el valor aparece en cualquier parte de una lista o matriz (texto de lista, rango o constante matricial).

## Sintaxis

```
=HAS(matriz, valor)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| matriz | Obligatorio | texto de lista, rango o matriz donde buscar |
| valor | Obligatorio | valor a buscar (elemento completo, sin distinguir mayúsculas) |

## Devuelve

Devuelve TRUE/FALSE (escalar); si se indica una matriz en valor, se evalúa elemento a elemento y se derrama. Si el propio argumento matriz es un error, devuelve ese error tal cual; si se omite valor, devuelve #VALUE!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=HAS("Mon, Tue, Wed","tue")` | TRUE | coincidencia del elemento completo, sin distinguir mayúsculas |
| `=HAS("Ann, Bob","An")` | FALSE | no hay coincidencia parcial |
| `=HAS("{""a"",{""b"",""c""}}","c")` | TRUE | texto de matriz anidada |
| `=HAS("1, 2","1")` | FALSE | el texto "1" ≠ el número 1 |

## Notas

- matriz acepta texto de lista (Mon, Tue, Wed o la forma de constante matricial como {"a",{"b","c"}}), un rango o una constante matricial, y se recorre de forma recursiva incluyendo los niveles anidados. Los elementos de una lista separada por comas o dentro de {…} se leen como número si tienen forma numérica, como lógico si son TRUE/FALSE y como error si son nombres de error como #N/A; en cambio, un texto único sin comas ni llaves no se clasifica y queda como un solo elemento de texto (una celda normal con "1" es el texto "1"). Si el texto no se puede analizar (p. ej. comillas sin cerrar), se trata como un único elemento sin producir error.
- La comparación es del elemento completo sin distinguir mayúsculas en texto ("Ann"≠"Anna") y numérica en números; el texto "1" y el número 1 son distintos. Los elementos de error dentro de la lista nunca coinciden. matriz no se eleva — un rango se toma como una única matriz completa, así que para obtener un resultado por fila use la referencia [@columna] de una tabla o BYROW.
- Las listas en celdas de la beta de Microsoft 365 solo se ven como #UNKNOWN! (#VALUE! en Excel 2016) en Excel antiguo, por lo que conviene convertirlas primero con la cinta [EGTools] › [Celdas de lista/matriz → texto]. Funciones relacionadas: HASANY, HASALL, FLATTEN
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `HAS` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.HAS`.

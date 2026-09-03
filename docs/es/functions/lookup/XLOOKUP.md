# XLOOKUP

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2021

Busca una coincidencia en un rango o matriz y devuelve el elemento correspondiente de un segundo rango o matriz.

## Sintaxis

```
=XLOOKUP(valor_buscado, matriz_buscada, matriz_devuelta, [si_no_encontrado], [modo_coincidencia], [modo_busqueda])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| valor_buscado | Obligatorio | el valor a buscar |
| matriz_buscada | Obligatorio | el rango/matriz donde buscar |
| matriz_devuelta | Obligatorio | el rango/matriz a devolver |
| si_no_encontrado | Opcional | valor si no hay coincidencia (opcional) |
| modo_coincidencia | Opcional | 0 exacto (pred.), -1 menor, 1 mayor, 3 regex |
| modo_busqueda | Opcional | 1 de inicio a fin (pred.), -1 de fin a inicio |

## Devuelve

Devuelve como matriz la fila (búsqueda vertical) o la columna (búsqueda horizontal) coincidente del rango de retorno; en Excel con matrices dinámicas se derrama. Devuelve #N/A si se omite un rango, la matriz está vacía o no hay coincidencia y no se indica un valor alternativo; #REF! si la posición coincidente queda fuera del rango de retorno; #VALUE! en caso de error interno.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=XLOOKUP("b",{"a";"b";"c"},{10;20;30})` | 20 | Búsqueda con coincidencia exacta |
| `=XLOOKUP(2,{1;2;3},{10,11;20,21;30,31})` | {20,21} | La fila coincidente completa se derrama |
| `=XLOOKUP(9,{1;2;3},{10;20;30},"none")` | none | Valor alternativo si no hay coincidencia |
| `=XLOOKUP({2;3},{1;2;3},{"a";"b";"c"})` | {b;c} | Valor buscado en matriz → búsqueda por elementos |
| `=XLOOKUP("^B",{"apple";"Banana";"cherry"},{1;2;3},"none",3)` | 2 | Coincidencia regex (match_mode 3) |
| `=XLOOKUP("(?i)^b",{"apple";"Banana";"cherry"},{1;2;3},"none",3)` | 2 | El prefijo (?i) ignora mayúsculas/minúsculas |

## Notas

- No se admiten match_mode 2 (caracteres comodín) ni search_mode 2/-2 (búsqueda binaria).
- match_mode 3 (regex) interpreta el valor buscado como un patrón de expresión regular y encuentra los elementos cuyo texto contiene una coincidencia del patrón (coincidencia parcial, al estilo de REGEXTEST) — el mismo número de modo que la novedad de 2024 del XLOOKUP nativo de Microsoft 365. De forma predeterminada distingue mayúsculas y minúsculas; anteponga (?i) al patrón para ignorarlas (mismo uso que la función nativa). Las celdas que no son texto nunca coinciden; un valor buscado que no sea texto, una regex no válida o la combinación con search_mode 2/-2 (búsqueda binaria) devuelve #VALUE!. Si no hay coincidencia, se sigue aplicando si_no_encontrado.
- La sintaxis de expresiones regulares es la de ECMAScript de std::wregex (puede diferir en algunos aspectos del PCRE2 de la versión nativa de 365).
- Si lookup_array tiene una sola columna se realiza una búsqueda vertical; en caso contrario se busca horizontalmente en la primera fila.
- Si el valor buscado es una matriz, cada elemento se busca por separado y el resultado se derrama con la misma forma que el valor buscado; si el rango devuelto tiene varias columnas, el resultado de cada elemento se reduce a su primer valor (igual que la función nativa), y un valor buscado con error devuelve ese error.
- Funciones relacionadas: XMATCH, FILTER.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `XLOOKUP` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.XLOOKUP`.

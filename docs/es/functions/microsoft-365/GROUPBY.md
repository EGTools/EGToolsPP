# GROUPBY

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Microsoft 365

Agrupa filas por clave(s) y agrega valores. El agregador es texto: SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF + extensiones (parámetro en relative_to): TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC).

## Sintaxis

```
=GROUPBY(campos_fila, valores, funcion, [encabezados], [prof_totales], [orden], [matriz_filtro], [relacion_campos], [relativo_a])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| campos_fila | Obligatorio | columna(s) clave |
| valores | Obligatorio | columna(s) de valores a agregar |
| funcion | Obligatorio | texto agregador: uno de SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF + extensiones (parámetro en relative_to): TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) |
| encabezados | Opcional | 0 sin/ocultar, 1 con/ocultar, 2 sin/mostrar generados, 3 con/mostrar (auto por defecto) |
| prof_totales | Opcional | 0 ninguno, 1 total (por defecto), 2 total+subtotales; negativo = arriba |
| orden | Opcional | número(s) de columna de salida, negativo descendente |
| matriz_filtro | Opcional | VERDADERO/FALSO por fila |
| relacion_campos | Opcional | aceptado por compatibilidad, ignorado |
| relativo_a | Opcional | parámetro del agregador extendido: delimitador de TEXTJOIN, k de LARGE/SMALL, p de PERCENTILE (0-1), QUARTILE 0-4 |

## Devuelve

Devuelve, como matriz derramada, una matriz 2D formada por la columna de claves de grupo y las columnas de valores agregados. Si faltan argumentos obligatorios, el número de filas de claves y valores no coincide, o el agregador o los valores de opción no son válidos, devuelve #VALUE!; si ninguna fila pasa el filtro, devuelve #N/A; y en AVERAGE·PERCENTOF las celdas con denominador 0 dan #DIV/0!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=GROUPBY({"a";"b";"a"},{10;20;30},"SUM")` | {"a",40;"b",20;"Total",60} | Suma por clave y fila de total general |
| `=GROUPBY({"a";"b";"a"},{10;20;30},"COUNT",0,0)` | {"a",2;"b",1} | Solo el recuento, sin totales |

## Notas

- El argumento function se especifica como texto en lugar de la lambda de la función nativa; los agregadores admitidos son los 16 SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF (los mismos que la función nativa), y como extensión de EGTools también se admiten TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) - el parámetro de los agregadores extendidos (separador de TEXTJOIN, k de LARGE/SMALL, p de PERCENTILE, 0~4 de QUARTILE) se pasa mediante el argumento relative_to (cualquier otro nombre devuelve #VALUE!). El argumento field_relationship se acepta pero se ignora.
- sort_order mantiene el orden jerárquico de los campos clave: un número de campo clave solo cambia la dirección de ese campo, y un número de columna de valores ordena por el valor agregado en lugar del último campo clave (los empates se resuelven por esa clave en orden ascendente).
- Las etiquetas de total y subtotal siguen el idioma de la interfaz (UI en coreano: 합계/총합계; UI en inglés: Total/Grand Total). La etiqueta de una fila de subtotal es el valor de la primera clave.
- Si se omite field_headers, la primera fila se reconoce automáticamente como encabezado cuando toda ella es texto y debajo hay valores no textuales.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `GROUPBY` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.GROUPBY`.

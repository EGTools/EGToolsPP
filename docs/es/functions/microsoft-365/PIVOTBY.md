# PIVOTBY

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Microsoft 365

Dinamiza datos por claves de fila y columna. El agregador es texto (SUM, COUNT…).

## Sintaxis

```
=PIVOTBY(campos_fila, campos_col, valores, funcion, [encabezados], [prof_total_fila], [orden_filas], [prof_total_col], [orden_cols], [matriz_filtro], [relativo_a])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| campos_fila | Obligatorio | columna(s) clave de fila |
| campos_col | Obligatorio | columna(s) clave de columna |
| valores | Obligatorio | columna(s) de valores a agregar |
| funcion | Obligatorio | texto agregador: uno de SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF + extensiones (parámetro en relative_to): TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) |
| encabezados | Opcional | 0 sin/ocultar, 1 con/ocultar, 2 sin/mostrar, 3 con/mostrar (auto) |
| prof_total_fila | Opcional | 0 ninguno, 1 total (por defecto), 2 total+subtotales; negativo = arriba |
| orden_filas | Opcional | orden de filas: número de columna, negativo descendente |
| prof_total_col | Opcional | 0 ninguno, 1 columna total (por defecto), 2 columnas total+subtotal; negativo = izquierda |
| orden_cols | Opcional | ±número(s) de campo, negativo descendente |
| matriz_filtro | Opcional | VERDADERO/FALSO por fila |
| relativo_a | Opcional | base de PERCENTOF 0-2 (0 total por defecto); para agregadores extendidos el parámetro: delimitador de TEXTJOIN, k de LARGE/SMALL, p de PERCENTILE (0-1), QUARTILE 0-4 |

## Devuelve

Devuelve, como matriz 2D derramada, una tabla cruzada de claves de fila × claves de columna con una fila de encabezado por cada campo de columna y, de forma predeterminada, fila/columna de totales; row_fields, col_fields y values admiten varias columnas cada uno, igual que la función nativa. Un agregador o valor de opción no válido, o un relative_to de PERCENTOF fuera de 0~2, devuelve #VALUE!; si ninguna fila pasa el filtro, devuelve #N/A; y las celdas de intersección sin datos quedan como cadenas vacías, también en las filas/columnas de subtotales.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=PIVOTBY({"a";"b";"a"},{"x";"x";"y"},{10;20;30},"SUM")` | {"","x","y","Total";"a",10,30,40;"b",20,"",20;"Total",30,30,60} | Pivote de sumas por fila × columna |
| `=PIVOTBY({"a";"a";"b"},{"x","p";"x","q";"y","p"},{10;20;30},"SUM")` | {"","x","x","y","Total";"","p","q","p","";"a",10,20,"",30;"b","","",30,30;"Total",10,20,30,60} | Dos campos de columna → dos filas de encabezado |

## Notas

- El argumento function se especifica como texto en lugar de la lambda de la función nativa; los agregadores admitidos son los 16 SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF (los mismos que la función nativa), y como extensión de EGTools también se admiten TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) - el parámetro de los agregadores extendidos (separador de TEXTJOIN, k de LARGE/SMALL, p de PERCENTILE, 0~4 de QUARTILE) se pasa mediante el argumento relative_to (cualquier otro nombre devuelve #VALUE!).
- row_fields, col_fields y values admiten varias columnas cada uno (igual que la función nativa): cada campo de columna añade una fila de encabezado y bajo cada grupo de columnas se anida una columna de valores por cada columna de values. Incluso con dos o más columnas de valores, las celdas de datos de las columnas de total/subtotal se rellenan agregando por cada columna de valores - la función nativa deja estas celdas en blanco, por lo que solo esta parte difiere de la salida nativa (diferencia intencionada). Con field_headers 2 o 3 (modo visible), el bloque de encabezado es: una fila con los nombres de los campos de columna unidos por ", ", las filas de claves de columna y una fila con los nombres de los campos de fila y de los valores.
- El valor 2 de row_total_depth/col_total_depth añade filas/columnas de subtotal por cada bloque del primer campo de fila/columna (ese eje necesita 2+ campos; si no, actúa como 1); los valores negativos colocan los totales arriba/a la izquierda. row_sort_order/col_sort_order aceptan ±número(s) de campo (también como matriz); la ordenación mantiene siempre el orden jerárquico de los campos y solo cambia la dirección del campo indicado.
- En esta implementación, el relative_to de PERCENTOF solo admite 0 (respecto al total de columna, predeterminado) / 1 (respecto al total de fila) / 2 (respecto al total general); 3 y 4 (respecto al padre) devuelven #VALUE!. Las etiquetas de la fila/columna de totales siguen el idioma de la interfaz (UI en coreano: 합계/총합계; UI en inglés: Total/Grand Total).
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `PIVOTBY` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.PIVOTBY`.

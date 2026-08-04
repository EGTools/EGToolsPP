# PIVOTBY

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2024 / Microsoft 365

Dinamiza datos por claves de fila y columna. El agregador es texto (SUM, COUNT…).

## Sintaxis

```
=PIVOTBY(campos_fila, campos_col, valores, funcion, [encabezados], [prof_total_fila], [orden_filas], [prof_total_col], [orden_cols], [matriz_filtro], [relativo_a])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| campos_fila | Obligatorio | columna clave de fila |
| campos_col | Obligatorio | columna clave de columna |
| valores | Obligatorio | valores a agregar |
| funcion | Obligatorio | texto agregador: uno de SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF + extensiones (parámetro en relative_to): TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) |
| encabezados | Opcional | 0 sin/ocultar, 1 con/ocultar, 2 sin/mostrar, 3 con/mostrar (auto) |
| prof_total_fila | Opcional | 0 ninguno, 1 total (por defecto); negativo = arriba |
| orden_filas | Opcional | orden de filas: número de columna, negativo descendente |
| prof_total_col | Opcional | 0 ninguno, 1 columna total (por defecto); negativo = izquierda |
| orden_cols | Opcional | 1 ascendente, -1 descendente |
| matriz_filtro | Opcional | VERDADERO/FALSO por fila |
| relativo_a | Opcional | base de PERCENTOF 0-2 (0 total por defecto); para agregadores extendidos el parámetro: delimitador de TEXTJOIN, k de LARGE/SMALL, p de PERCENTILE (0-1), QUARTILE 0-4 |

## Devuelve

Devuelve, como matriz 2D derramada, una tabla cruzada de claves de fila × claves de columna que incluye la fila de encabezados de las claves de columna. Si row_fields, col_fields o values tienen más de 1 columna cada uno, o relative_to está fuera de 0~2, devuelve #VALUE!; si ninguna fila pasa el filtro, devuelve #N/A; y las celdas de intersección sin datos quedan como cadenas vacías.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=PIVOTBY({"a";"b";"a"},{"x";"x";"y"},{10;20;30},"SUM")` | {"","x","y","Total";"a",10,30,40;"b",20,"",20;"Total",30,30,60} | Pivote de sumas por fila × columna |

## Notas

- row_fields, col_fields y values admiten solo 1 columna cada uno (no se admite el anidamiento de varias columnas). Los agregadores admitidos son los 16 SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF (los mismos que la función nativa), y como extensión de EGTools también se admiten TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) - el parámetro de los agregadores extendidos (separador de TEXTJOIN, k de LARGE/SMALL, p de PERCENTILE, 0~4 de QUARTILE) se pasa mediante el argumento relative_to (cualquier otro nombre devuelve #VALUE!).
- En esta implementación, el relative_to de PERCENTOF solo admite 0 (respecto al total de columna, predeterminado) / 1 (respecto al total de fila) / 2 (respecto al total general); 3 y 4 (respecto al padre) devuelven #VALUE!.
- Las etiquetas de la fila/columna de totales siguen el idioma de la interfaz (UI en coreano: 합계; UI en inglés: Total).
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `PIVOTBY` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.PIVOTBY`.

# SORTBY

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2021

Ordena una matriz según valores de otras matrices.

## Sintaxis

```
=SORTBY(matriz, por_matriz1, [orden1], ...)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| matriz | Obligatorio | datos a ordenar |
| por_matriz1 | Obligatorio | clave (alineada a filas) |
| orden1 | Opcional | 1 asc (pred.), -1 desc |
| … | Opcional | Los argumentos anteriores se pueden repetir |

## Devuelve

Devuelve una matriz con las filas reordenadas según el orden de los valores de las matrices de claves; se derrama. Devuelve #VALUE! si no existe ninguna matriz de claves válida con el mismo número de filas que array.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=SORTBY({"a";"b";"c"},{3;1;2})` | {b;c;a} | Ordenación por matriz de claves |
| `=SORTBY({1;2;3},{2;2;1},1,{9;1;5},-1)` | {3;1;2} | Segunda clave en orden descendente |

## Notas

- Las matrices de claves deben ser vectores verticales (de una columna) con el mismo número de filas que array; solo se admite la ordenación de filas (no se admite ordenar por columnas).
- Las matrices de claves con un número de filas distinto se ignoran silenciosamente.
- Función relacionada: SORT.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `SORTBY` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.SORTBY`.

# CHOOSECOLS

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2024 / Microsoft 365

Devuelve las columnas indicadas de una matriz.

## Sintaxis

```
=CHOOSECOLS(matriz, núm_col1, ...)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| matriz | Obligatorio | la matriz |
| núm_col1 | Obligatorio | columna a devolver (neg = desde el final) |
| … | Opcional | Los argumentos anteriores se pueden repetir |

## Devuelve

Devuelve una matriz formada por las columnas seleccionadas; en las versiones de Excel que admiten matrices dinámicas el resultado se derrama. Devuelve #VALUE! si la matriz está vacía, un número de columna es 0 o está fuera de rango, o no se especifica ningún número de columna.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=CHOOSECOLS({1,2,3;4,5,6},1,3)` | {1,3;4,6} | selecciona las columnas 1 y 3 |
| `=CHOOSECOLS({1,2,3;4,5,6},-1)` | {3;6} | selecciona la última columna |
| `=CHOOSECOLS({1,2;3,4},5)` | #VALUE! | número de columna fuera de rango |

## Notas

- Los números de columna negativos se cuentan desde el final.
- Para seleccionar filas, use CHOOSEROWS.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `CHOOSECOLS` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.CHOOSECOLS`.

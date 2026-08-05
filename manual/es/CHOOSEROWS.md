# CHOOSEROWS

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2024

Devuelve las filas indicadas de una matriz.

## Sintaxis

```
=CHOOSEROWS(matriz, núm_fila1, ...)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| matriz | Obligatorio | la matriz |
| núm_fila1 | Obligatorio | fila a devolver (neg = desde el final) |
| … | Opcional | Los argumentos anteriores se pueden repetir |

## Devuelve

Devuelve una matriz formada por las filas seleccionadas; en las versiones de Excel que admiten matrices dinámicas el resultado se derrama. Devuelve #VALUE! si la matriz está vacía, un número de fila es 0 o está fuera de rango, o no se especifica ningún número de fila.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=CHOOSEROWS({1,2;3,4;5,6},1,-1)` | {1,2;5,6} | primera y última fila |
| `=CHOOSEROWS({1,2;3,4},3)` | #VALUE! | número de fila fuera de rango |

## Notas

- Los números de fila negativos se cuentan desde el final.
- Para seleccionar columnas, use CHOOSECOLS.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `CHOOSEROWS` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.CHOOSEROWS`.

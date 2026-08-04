# TAKE

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2024 / Microsoft 365

Devuelve filas/columnas del principio o final de una matriz.

## Sintaxis

```
=TAKE(matriz, filas, [columnas])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| matriz | Obligatorio | la matriz |
| filas | Obligatorio | filas a tomar (neg = desde el final) |
| columnas | Opcional | columnas a tomar (opcional) |

## Devuelve

Devuelve una matriz con el número indicado de filas y columnas tomadas desde el principio (positivo) o el final (negativo) de la matriz; en las versiones de Excel que admiten matrices dinámicas el resultado se derrama. Devuelve #VALUE! si la matriz está vacía o el número de filas o columnas que se van a tomar es 0.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=TAKE({1,2,3;4,5,6;7,8,9},2)` | {1,2,3;4,5,6} | primeras 2 filas |
| `=TAKE({1,2,3;4,5,6;7,8,9},-1,2)` | {7,8} | primeras 2 columnas de la última fila |
| `=TAKE({1,2;3,4},0)` | #VALUE! | error con 0 filas |

## Notas

- Si el número de filas o columnas es negativo, se toman desde el final.
- Si el resultado es una matriz vacía, devuelve #VALUE! en lugar del #CALC! nativo.
- Función relacionada: DROP
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `TAKE` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.TAKE`.

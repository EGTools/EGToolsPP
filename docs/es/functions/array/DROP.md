# DROP

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2024

Quita filas/columnas del principio o final de una matriz.

## Sintaxis

```
=DROP(matriz, filas, [columnas])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| matriz | Obligatorio | la matriz |
| filas | Obligatorio | filas a quitar (neg = desde el final) |
| columnas | Opcional | columnas a quitar (opcional) |

## Devuelve

Devuelve la matriz restante tras eliminar el número indicado de filas y columnas; en las versiones de Excel que admiten matrices dinámicas el resultado se derrama. Devuelve #VALUE! si la matriz está vacía o no queda ninguna fila o columna después de la eliminación.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=DROP({1,2,3;4,5,6;7,8,9},1)` | {4,5,6;7,8,9} | elimina la primera fila |
| `=DROP({1,2,3;4,5,6;7,8,9},1,-1)` | {4,5;7,8} | elimina la primera fila y la última columna |
| `=DROP({1,2},5)` | #VALUE! | error si se elimina todo |

## Notas

- Si el número de filas o columnas es negativo, se eliminan desde el final.
- Si el resultado es una matriz vacía, devuelve #VALUE! en lugar del #CALC! nativo.
- Una matriz en filas·columnas devuelve una matriz con la forma de ese argumento, donde cada elemento es el primer valor del resultado de ese elemento (reducción al primer valor, igual que la función nativa).
- Función relacionada: TAKE
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `DROP` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.DROP`.

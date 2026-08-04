# SORT

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2021

Ordena una matriz.

## Sintaxis

```
=SORT(matriz, [índice_orden], [orden], [por_columna])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| matriz | Obligatorio | datos a ordenar |
| índice_orden | Opcional | columna (o fila) por la que ordenar (pred. 1) |
| orden | Opcional | 1 ascendente (pred.), -1 descendente |
| por_columna | Opcional | VERDADERO para ordenar columnas |

## Devuelve

Devuelve una matriz del mismo tamaño ordenada según la columna (o fila) indicada; se derrama. Devuelve #VALUE! si sort_index queda fuera de la matriz o se produce un error interno.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=SORT({3;1;2})` | {1;2;3} | Orden ascendente |
| `=SORT({1,3;2,1;3,2},2)` | {2,1;3,2;1,3} | Orden según la columna 2 |
| `=SORT({5;3;9},1,-1)` | {9;5;3} | Orden descendente |

## Notas

- Si sort_order es negativo se ordena de forma descendente; cualquier otro valor (incluido 0) se trata como ascendente (la función nativa solo admite 1/-1).
- Las claves iguales conservan su orden original (ordenación estable).
- Funciones relacionadas: SORTBY, UNIQUE.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `SORT` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.SORT`. En algunas versiones antiguas se registra como `xSORT` por un conflicto de palabra reservada.

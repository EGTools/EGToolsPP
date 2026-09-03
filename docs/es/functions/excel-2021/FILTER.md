# FILTER

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2021

> ⚠️ En algunas versiones antiguas se registra como `xFILTER` por un conflicto de palabra reservada.

Filtra una matriz conservando filas/columnas donde incluir es VERDADERO.

## Sintaxis

```
=FILTER(matriz, incluir, [si_vacío])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| matriz | Obligatorio | datos a filtrar |
| incluir | Obligatorio | matriz booleana (filas o columnas) |
| si_vacío | Opcional | valor si no hay coincidencias (opcional) |

## Devuelve

Devuelve una matriz que conserva solo las filas o columnas cuya condición es verdadera; se derrama. Devuelve #VALUE! si el tamaño de include no coincide con el número de filas o columnas de array, y #N/A si no queda ningún elemento y se omite if_empty.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=FILTER({1;2;3;4},{1;0;1;0})` | {1;3} | Filtrado de filas |
| `=FILTER({1,2;3,4;5,6},{0;1;1})` | {3,4;5,6} | Filtro de filas con varias columnas |
| `=FILTER({1;2},{0;0},"none")` | none | Valor alternativo si no hay resultados |

## Notas

- Los valores de texto y vacíos de include se tratan como FALSE (la función nativa devuelve #VALUE!).
- Si no hay resultados y se omite if_empty, devuelve #N/A en lugar de #CALC!.
- Si include contiene un valor de error, todo el resultado es ese error.
- Funciones relacionadas: SORT, UNIQUE.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `FILTER` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.FILTER`.

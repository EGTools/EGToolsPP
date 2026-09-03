# EXPAND

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2024

Expande una matriz al tamaño dado, rellenando con un valor.

## Sintaxis

```
=EXPAND(matriz, filas, [columnas], [rellenar_con])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| matriz | Obligatorio | la matriz |
| filas | Obligatorio | filas totales |
| columnas | Opcional | columnas totales (opcional) |
| rellenar_con | Opcional | valor de relleno (pred. #N/A) |

## Devuelve

Devuelve la matriz expandida al tamaño indicado de filas × columnas; en las versiones de Excel que admiten matrices dinámicas el resultado se derrama. Las posiciones nuevas se rellenan con el valor de relleno (predeterminado #N/A). Devuelve #VALUE! si la matriz está vacía o el tamaño indicado es menor que el original.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=EXPAND({1,2;3,4},3,3,0)` | {1,2,0;3,4,0;0,0,0} | expande a 3×3 y rellena con 0 |
| `=EXPAND({1,2},1,4)` | {1,2,#N/A,#N/A} | valor de relleno predeterminado #N/A |
| `=EXPAND({1,2;3,4},1,1)` | #VALUE! | no se permite reducir |

## Notas

- Si se omiten las filas o las columnas, se conserva el tamaño original.
- Si se omite el valor de relleno, se rellena con #N/A.
- Una matriz en filas·columnas·valor de relleno devuelve una matriz con la forma de ese argumento, donde cada elemento es el primer valor del resultado de ese elemento (reducción al primer valor, igual que la función nativa).
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `EXPAND` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.EXPAND`.

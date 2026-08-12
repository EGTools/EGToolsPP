# WRAPROWS

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2024

Ajusta un vector en una matriz 2D, con un número fijo de valores por fila.

## Sintaxis

```
=WRAPROWS(vector, núm_ajuste, [rellenar_con])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| vector | Obligatorio | el vector a ajustar |
| núm_ajuste | Obligatorio | valores por fila |
| rellenar_con | Opcional | valor de relleno (pred. #N/A) |

## Devuelve

Devuelve una matriz bidimensional con los valores del vector dispuestos horizontalmente, con wrap_count valores por fila; en las versiones de Excel que admiten matrices dinámicas, el resultado se derrama. Las posiciones que faltan en la última fila se rellenan con pad_with (predeterminado #N/A). Devuelve #VALUE! si el vector está vacío o wrap_count es 0 o menor.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=WRAPROWS({1,2,3,4,5},2)` | {1,2;3,4;5,#N/A} | Dos valores por fila |
| `=WRAPROWS({1,2,3,4,5},2,0)` | {1,2;3,4;5,0} | Relleno con 0 |

## Notas

- También se acepta una matriz bidimensional como entrada, que se aplana por filas.
- Una matriz en recuento de ajuste·valor de relleno devuelve una matriz con la forma de ese argumento, donde cada elemento es el primer valor del resultado de ese elemento (reducción al primer valor, igual que la función nativa).
- Función relacionada: WRAPCOLS
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `WRAPROWS` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.WRAPROWS`.

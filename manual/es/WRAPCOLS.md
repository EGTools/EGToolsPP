# WRAPCOLS

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2024

Ajusta un vector en una matriz 2D, con un número fijo de valores por columna.

## Sintaxis

```
=WRAPCOLS(vector, núm_ajuste, [rellenar_con])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| vector | Obligatorio | el vector a ajustar |
| núm_ajuste | Obligatorio | valores por columna |
| rellenar_con | Opcional | valor de relleno (pred. #N/A) |

## Devuelve

Devuelve una matriz bidimensional con los valores del vector dispuestos verticalmente, con wrap_count valores por columna; en las versiones de Excel que admiten matrices dinámicas, el resultado se derrama. Las posiciones que faltan en la última columna se rellenan con pad_with (predeterminado #N/A). Devuelve #VALUE! si el vector está vacío o wrap_count es 0 o menor.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=WRAPCOLS({1,2,3,4,5},2)` | {1,3,5;2,4,#N/A} | Dos valores por columna |
| `=WRAPCOLS({1,2,3,4,5},3,0)` | {1,4;2,5;3,0} | Relleno con 0 |

## Notas

- También se acepta una matriz bidimensional como entrada, que se aplana por filas.
- Función relacionada: WRAPROWS
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `WRAPCOLS` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.WRAPCOLS`.

# PERMUTATIONA

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Devuelve el número de permutaciones con repetición (number ^ number_chosen).

## Sintaxis

```
=PERMUTATIONA(number, number_chosen)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| number | Obligatorio | número total de elementos |
| number_chosen | Obligatorio | elementos en cada permutación |

## Devuelve

Devuelve el número de permutaciones con repetición, number^number_chosen (número escalar). Si un argumento no es numérico o es negativo, devuelve un error #NUM!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=PERMUTATIONA(3,2)` | 9 | 3^2 |
| `=PERMUTATIONA(2,5)` | 32 | 2^5 |

## Notas

- Las partes fraccionarias se descartan.
- Todos los argumentos aceptan matrices y se calculan elemento a elemento, derramando una matriz de la misma forma: los escalares se repiten, un vector columna × un vector fila se expande como producto externo, los elementos con tamaños no coincidentes dan #N/A y un elemento de error devuelve ese error.
- Función relacionada: COMBINA
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `PERMUTATIONA` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.PERMUTATIONA`.

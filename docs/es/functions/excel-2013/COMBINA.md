# COMBINA

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Devuelve el número de combinaciones con repetición.

## Sintaxis

```
=COMBINA(number, number_chosen)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| number | Obligatorio | número total de elementos |
| number_chosen | Obligatorio | elementos en cada combinación |

## Devuelve

Devuelve el número de combinaciones con repetición (número escalar). Si un argumento no es numérico o es negativo, devuelve un error #NUM!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=COMBINA(4,3)` | 20 | Combinaciones de 3 entre 4 con repetición |
| `=COMBINA(10,3)` | 220 | Combinaciones de 3 entre 10 con repetición |

## Notas

- Las partes fraccionarias se descartan; se calcula como C(n+k-1, k).
- Todos los argumentos aceptan matrices y se calculan elemento a elemento, derramando una matriz de la misma forma: los escalares se repiten, un vector columna × un vector fila se expande como producto externo, los elementos con tamaños no coincidentes dan #N/A y un elemento de error devuelve ese error.
- Función relacionada: PERMUTATIONA
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `COMBINA` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.COMBINA`.

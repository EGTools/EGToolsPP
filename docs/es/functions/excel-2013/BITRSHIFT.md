# BITRSHIFT

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Devuelve un número desplazado a la derecha el número de bits indicado.

## Sintaxis

```
=BITRSHIFT(number, shift_amount)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| number | Obligatorio | el número a desplazar (>= 0) |
| shift_amount | Obligatorio | bits a desplazar a la derecha (negativo: a la izquierda) |

## Devuelve

Devuelve el resultado del desplazamiento a la derecha como número (escalar, con redondeo hacia abajo). Devuelve #NUM! si number no es un entero en el rango [0, 2^48-1], si shift_amount no es un entero o su valor absoluto supera 53, o si el resultado supera 2^53.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=BITRSHIFT(13,2)` | 3 | Desplaza 2 bits a la derecha |
| `=BITRSHIFT(3,-2)` | 12 | Un valor negativo desplaza a la izquierda |

## Notas

- Si shift_amount es negativo, se comporta como un desplazamiento a la izquierda.
- Todos los argumentos aceptan matrices y se calculan elemento a elemento, derramando una matriz de la misma forma: los escalares se repiten, un vector columna × un vector fila se expande como producto externo, los elementos con tamaños no coincidentes dan #N/A y un elemento de error devuelve ese error.
- Función relacionada: BITLSHIFT
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `BITRSHIFT` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.BITRSHIFT`.

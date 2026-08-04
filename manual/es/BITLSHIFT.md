# BITLSHIFT

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Devuelve un número desplazado a la izquierda el número de bits indicado.

## Sintaxis

```
=BITLSHIFT(number, shift_amount)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| number | Obligatorio | el número a desplazar (>= 0) |
| shift_amount | Obligatorio | bits a desplazar a la izquierda (negativo: a la derecha) |

## Devuelve

Devuelve el resultado del desplazamiento a la izquierda como número (escalar). Devuelve #NUM! si number no es un entero en el rango [0, 2^48-1], si shift_amount no es un entero o su valor absoluto supera 53, o si el resultado supera 2^53.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=BITLSHIFT(4,2)` | 16 | Desplaza 2 bits a la izquierda |
| `=BITLSHIFT(4,-1)` | 2 | Un valor negativo desplaza a la derecha |

## Notas

- Si shift_amount es negativo, se comporta como un desplazamiento a la derecha.
- Función relacionada: BITRSHIFT
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `BITLSHIFT` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.BITLSHIFT`.

# BITAND

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Devuelve un AND bit a bit de dos números.

## Sintaxis

```
=BITAND(number1, number2)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| number1 | Obligatorio | primer número (>= 0) |
| number2 | Obligatorio | segundo número (>= 0) |

## Devuelve

Devuelve el AND bit a bit de dos números como número (escalar). Devuelve #NUM! si un argumento no es un número o valor lógico, es negativo o fraccionario, o supera 2^48-1 (281.474.976.710.655).

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=BITAND(13,25)` | 9 | 1101 AND 11001 = 1001 |
| `=BITAND(-1,2)` | #NUM! | Un operando negativo da error |

## Notas

- Los argumentos no numéricos también devuelven #NUM!, no #VALUE!.
- Funciones relacionadas: BITOR, BITXOR
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `BITAND` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.BITAND`.

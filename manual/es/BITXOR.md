# BITXOR

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Devuelve un XOR bit a bit de dos números.

## Sintaxis

```
=BITXOR(number1, number2)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| number1 | Obligatorio | primer número (>= 0) |
| number2 | Obligatorio | segundo número (>= 0) |

## Devuelve

Devuelve el XOR bit a bit de dos números como número (escalar). Devuelve #NUM! si un argumento no es un número o valor lógico, es negativo o fraccionario, o supera 2^48-1 (281.474.976.710.655).

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=BITXOR(5,3)` | 6 | 101 XOR 011 = 110 |
| `=BITXOR(5,"a")` | #NUM! | Un operando no numérico da error |

## Notas

- Los argumentos no numéricos también devuelven #NUM!, no #VALUE!.
- Funciones relacionadas: BITAND, BITOR
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `BITXOR` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.BITXOR`.

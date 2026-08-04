# BITOR

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Devuelve un OR bit a bit de dos números.

## Sintaxis

```
=BITOR(number1, number2)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| number1 | Obligatorio | primer número (>= 0) |
| number2 | Obligatorio | segundo número (>= 0) |

## Devuelve

Devuelve el OR bit a bit de dos números como número (escalar). Devuelve #NUM! si un argumento no es un número o valor lógico, es negativo o fraccionario, o supera 2^48-1 (281.474.976.710.655).

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=BITOR(23,10)` | 31 | 10111 OR 01010 = 11111 |
| `=BITOR(1.5,2)` | #NUM! | Un operando fraccionario da error |

## Notas

- Los argumentos no numéricos también devuelven #NUM!, no #VALUE!.
- Funciones relacionadas: BITAND, BITXOR
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `BITOR` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.BITOR`.

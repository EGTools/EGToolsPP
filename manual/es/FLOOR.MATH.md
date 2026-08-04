# FLOOR.MATH

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Redondea un número hacia abajo al entero o múltiplo más cercano.

## Sintaxis

```
=FLOOR.MATH(number, [significance], [mode])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| number | Obligatorio | el número a redondear |
| significance | Opcional | múltiplo (opcional, predeterminado 1) |
| mode | Opcional | si no es cero, redondea negativos hacia cero (opcional) |

## Devuelve

Devuelve el número redondeado hacia abajo al múltiplo de significance (escalar). Si number no es numérico, devuelve un error #VALUE!; si significance es 0, devuelve 0.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=FLOOR.MATH(6.7)` | 6 | Redondea hacia abajo al múltiplo predeterminado 1 |
| `=FLOOR.MATH(-5.5,2)` | -6 | Los negativos se redondean alejándose de cero |
| `=FLOOR.MATH(-5.5,2,1)` | -4 | mode≠0: hacia cero |

## Notas

- El signo de significance se ignora; se usa su valor absoluto.
- Función relacionada: CEILING.MATH
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `FLOOR.MATH` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.FLOOR.MATH`.

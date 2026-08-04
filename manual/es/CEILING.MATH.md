# CEILING.MATH

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Redondea un número hacia arriba al entero o múltiplo más cercano.

## Sintaxis

```
=CEILING.MATH(number, [significance], [mode])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| number | Obligatorio | el número a redondear |
| significance | Opcional | múltiplo (opcional, predeterminado 1) |
| mode | Opcional | si no es cero, redondea negativos alejándose de cero (opcional) |

## Devuelve

Devuelve el número redondeado hacia arriba al múltiplo de significance (escalar). Si number no es numérico, devuelve un error #VALUE!; si significance es 0, devuelve 0.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=CEILING.MATH(6.3)` | 7 | Redondea hacia arriba al múltiplo predeterminado 1 |
| `=CEILING.MATH(-5.5,2)` | -4 | Los negativos se redondean hacia cero |
| `=CEILING.MATH(-5.5,2,1)` | -6 | mode≠0: alejándose de cero |

## Notas

- El signo de significance se ignora; se usa su valor absoluto.
- Función relacionada: FLOOR.MATH
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `CEILING.MATH` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.CEILING.MATH`.

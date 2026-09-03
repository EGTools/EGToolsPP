# ARABIC

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Convierte un número romano en un número arábigo.

## Sintaxis

```
=ARABIC(text)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| text | Obligatorio | texto de número romano |

## Devuelve

Devuelve el valor numérico convertido del número romano (escalar). Cualquier carácter distinto de los símbolos romanos (I, V, X, L, C, D, M) y de un '-' inicial devuelve un error #VALUE!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=ARABIC("LVII")` | 57 | Número romano a número |
| `=ARABIC("MCMXII")` | 1912 | Convierte un número romano grande |
| `=ARABIC("-X")` | -10 | Admite la notación negativa |

## Notas

- Los espacios se ignoran y no se distinguen mayúsculas y minúsculas.
- El texto vacío devuelve 0.
- Si el texto es una matriz, la función se aplica elemento a elemento y se derrama una matriz de la misma forma (un elemento de error devuelve ese error).
- Funciones relacionadas: BASE, DECIMAL
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `ARABIC` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.ARABIC`.

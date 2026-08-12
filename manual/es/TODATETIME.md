# TODATETIME

**Categoría**: función exclusiva de EGTools

Convierte texto de fecha/hora con unidades coreanas o hanja en un valor de fecha-hora.

## Sintaxis

```
=TODATETIME(texto)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto | Obligatorio | texto con fecha y/u hora |

## Devuelve

Devuelve un valor de serie de fecha y hora (número escalar): un entero si solo hay fecha y una fracción menor que 1 si solo hay hora. Si no se encuentra fecha ni hora, devuelve texto vacío, y si ocurre un error durante el análisis, devuelve #VALUE!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=TODATETIME("2026년 1월 15일")` | 46037 | Fecha en coreano a valor de serie |
| `=TODATETIME("오후 2시 30분")` | 0.604166667 | Hora de la tarde a valor de hora |

## Notas

- También reconoce los caracteres chinos (年月日時分秒) y los marcadores 오전/오후, AM/PM y 上午/下午.
- Un texto con solo mes y día se interpreta con el año actual, y los años de 2 dígitos mayores que 31 se ajustan a los años 2000.
- Aplique un formato de número de fecha/hora a la celda de resultado.
- Si el texto es una matriz, la función se aplica elemento a elemento y se derrama una matriz de la misma forma (un elemento de error devuelve ese error).
- Compatibilidad: Excel 2010+. Se registra siempre como `TODATETIME` en todas las versiones de Excel.

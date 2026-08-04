# LET

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2021

Asigna nombres a valores y los usa en un cálculo final.

## Sintaxis

```
=LET(name1, value1, calculation, ...)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| name1 | Obligatorio | nombre a definir |
| value1 | Obligatorio | valor de este nombre |
| calculation | Obligatorio | el cálculo que usa los nombres anteriores |
| … | Opcional | Los argumentos anteriores se pueden repetir |

## Devuelve

Devuelve el resultado de evaluar el cálculo final con cada nombre sustituido por su valor; si el resultado es una matriz, se derrama. Devuelve #VALUE! si la estructura de los argumentos es incorrecta o no se puede leer la fórmula de la celda que llama; si la evaluación falla, devuelve la cadena de la fórmula sustituida en lugar de un error.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=LET(x,3,x*2)` | 6 | Define un nombre y luego calcula |
| `=LET(x,2,y,x+1,x*y)` | 6 | Se puede hacer referencia a nombres anteriores |

## Notas

- A diferencia de la función nativa, no vincula nombres: es una función de tipo hoja de macros que sustituye los nombres por sus valores en el texto de la fórmula de la celda que llama y la evalúa con Application.Evaluate.
- Los nombres "R" y "C" están reservados por el analizador R1C1 de Excel y no se pueden usar.
- Si la evaluación falla, devuelve tal cual la cadena de la fórmula sustituida para facilitar la depuración.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `LET` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.LET`. En algunas versiones antiguas se registra como `xLET` por un conflicto de palabra reservada.

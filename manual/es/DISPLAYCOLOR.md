# DISPLAYCOLOR

**Categoría**: función exclusiva de EGTools

Devuelve el color de relleno o de fuente mostrado de cada celda de un rango.

## Sintaxis

```
=DISPLAYCOLOR(rango_color, [fuente])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| rango_color | Obligatorio | el rango de celdas a inspeccionar |
| fuente | Opcional | TRUE = color de fuente (predeterminado FALSE = relleno) |

## Devuelve

Devuelve el valor del color de fondo que realmente se muestra en cada celda (con color_fuente=TRUE, el color de fuente) como una matriz numérica del mismo tamaño que el rango, y se derrama. Devuelve #VALUE! si el número de celdas supera 100.000 o si falla la conexión COM o la lectura del color.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=DISPLAYCOLOR(A1:B2)` |  | Número de color mostrado de cada celda (el resultado depende del formato de las celdas) |

## Notas

- Se basa en DisplayFormat, por lo que devuelve el color mostrado con el formato condicional aplicado.
- Cambiar solo el formato no provoca un recálculo, por lo que se debe recalcular con F9.
- Es una función de tipo macro, por lo que se excluye del recálculo multiproceso.
- Compatibilidad: Excel 2010+. Se registra siempre como `DISPLAYCOLOR` en todas las versiones de Excel.

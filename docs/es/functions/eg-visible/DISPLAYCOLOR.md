# DISPLAYCOLOR

**Categoría**: función exclusiva de EGTools

Devuelve el color de relleno o de fuente de cada celda de un rango.

## Sintaxis

```
=DISPLAYCOLOR(rango_color, [fuente], [condicional])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| rango_color | Obligatorio | el rango de celdas a inspeccionar |
| fuente | Opcional | TRUE = color de fuente (predeterminado FALSE = relleno) |
| condicional | Opcional | TRUE = usar el color mostrado con formato condicional (exacto solo con la hoja activa); predeterminado FALSE = color del formato de la celda |

## Devuelve

Devuelve el valor del color de fondo de cada celda (con fuente=TRUE, el color de fuente) como una matriz numérica del mismo tamaño que el rango, y se derrama. De forma predeterminada es el color asignado en el formato de la celda; con condicional=TRUE es el color mostrado con el formato condicional aplicado. Devuelve #VALUE! si el número de celdas supera 100.000 o si falla la conexión COM o la lectura del color.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=DISPLAYCOLOR(A1:B2)` |  | Número del color de fondo de cada celda (el resultado depende del formato de las celdas) |
| `=DISPLAYCOLOR(A1:B2,,TRUE)` |  | Número del color mostrado con formato condicional (el resultado depende del formato de las celdas) |

## Notas

- De forma predeterminada (condicional omitido/FALSE) solo se usa el color asignado directamente en el formato de la celda (Interior.Color / Font.Color); los colores aplicados por formato condicional se ignoran. El resultado es estable sea cual sea la hoja activa, aunque se edite otra hoja o se ejecute F9 o un recálculo completo, por lo que se recomienda el valor predeterminado salvo que realmente necesite los colores del formato condicional.
- Con condicional=TRUE se usa el color realmente mostrado con el formato condicional aplicado (DisplayFormat). Los colores del formato condicional solo se leen correctamente cuando el recálculo se produce con la hoja de la fórmula activa.
- Precaución: con condicional=TRUE, si el recálculo se produce mientras está activa otra hoja (u otro libro) —al editar esa hoja, con F9 o con un recálculo completo— no se leen bien los colores del formato condicional y se obtienen valores incorrectos. Como estas funciones son volátiles, basta con editar otra hoja para que ocurra, así que es frecuente en la práctica.
- Precaución: el valor incorrecto se mantiene aunque vuelva a la hoja de la fórmula; pulse F9 con la hoja de la fórmula activa para corregirlo.
- Precaución: si necesita agregar según los colores del formato condicional, se recomienda, siempre que sea posible, agregar usando la propia condición del formato condicional (p. ej. `=COUNTIF(A1:A10,">5")`).
- Al ser una función de tipo macro que recibe referencias de celda, Excel la trata como volátil y se recalcula cada vez que cambia un valor en cualquier parte del libro. En cambio, cambiar solo el color de una celda o una regla de formato condicional no provoca un recálculo, por lo que se necesita F9; también se excluye del recálculo multiproceso.
- Las fórmulas existentes que usan hasta el segundo argumento (fuente) siguen siendo compatibles, pero las que dependían de colores de formato condicional deben indicar TRUE en el tercer argumento para obtener el mismo resultado que antes.
- Compatibilidad: Excel 2010+. Se registra siempre como `DISPLAYCOLOR` en todas las versiones de Excel.

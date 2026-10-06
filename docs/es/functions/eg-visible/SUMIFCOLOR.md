# SUMIFCOLOR

**Categoría**: función exclusiva de EGTools

Suma las celdas cuyo color coincide con una celda de referencia.

## Sintaxis

```
=SUMIFCOLOR(rango_busqueda, celda_color, [fuente], [condicional])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| rango_busqueda | Obligatorio | el rango a sumar |
| celda_color | Obligatorio | celda con el color de referencia (se usa la primera) |
| fuente | Opcional | TRUE = comparar color de fuente (predeterminado FALSE) |
| condicional | Opcional | TRUE = usar el color mostrado con formato condicional (exacto solo con la hoja activa); predeterminado FALSE = color del formato de la celda |

## Devuelve

Devuelve la suma de los valores numéricos de las celdas con el mismo color que la celda de referencia (de forma predeterminada el color del formato de la celda; con condicional=TRUE, el color mostrado). Devuelve #VALUE! si falla la conexión COM, si no se puede leer el color de referencia o si se examinan más de 100.000 celdas; si una celda coincidente contiene un valor de error, devuelve ese error tal cual.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=SUMIFCOLOR(A1:A10,C1)` |  | Suma de las celdas con el mismo color de fondo que C1 (el resultado depende del formato de las celdas) |
| `=SUMIFCOLOR(A1:A10,C1,,TRUE)` |  | Suma según los colores del formato condicional (el resultado depende del formato de las celdas) |

## Notas

- De forma predeterminada (condicional omitido/FALSE) solo se usa el color asignado directamente en el formato de la celda (Interior.Color / Font.Color); los colores aplicados por formato condicional se ignoran. El resultado es estable sea cual sea la hoja activa, aunque se edite otra hoja o se ejecute F9 o un recálculo completo, por lo que se recomienda el valor predeterminado salvo que realmente necesite los colores del formato condicional.
- Con condicional=TRUE se usa el color realmente mostrado con el formato condicional aplicado (DisplayFormat). Los colores del formato condicional solo se leen correctamente cuando el recálculo se produce con la hoja de la fórmula activa.
- Precaución: con condicional=TRUE, si el recálculo se produce mientras está activa otra hoja (u otro libro) —al editar esa hoja, con F9 o con un recálculo completo— no se leen bien los colores del formato condicional y se obtienen valores incorrectos. Como estas funciones son volátiles, basta con editar otra hoja para que ocurra, así que es frecuente en la práctica.
- Precaución: el valor incorrecto se mantiene aunque vuelva a la hoja de la fórmula; pulse F9 con la hoja de la fórmula activa para corregirlo.
- Precaución: si necesita agregar según los colores del formato condicional, se recomienda, siempre que sea posible, agregar usando la propia condición del formato condicional (p. ej. `=SUMIF(A1:A10,">5")`).
- Al ser una función de tipo macro que recibe referencias de celda, Excel la trata como volátil y se recalcula cada vez que cambia un valor en cualquier parte del libro. En cambio, cambiar solo el color de una celda o una regla de formato condicional no provoca un recálculo, por lo que se necesita F9; también se excluye del recálculo multiproceso.
- Las fórmulas existentes que usan hasta el tercer argumento (fuente) siguen siendo compatibles, pero las que dependían de colores de formato condicional deben indicar TRUE en el cuarto argumento para obtener el mismo resultado que antes.
- Un área combinada se trata como una sola celda (solo se comprueba su primera celda, donde está el valor).
- Compatibilidad: Excel 2010+. Se registra siempre como `SUMIFCOLOR` en todas las versiones de Excel.

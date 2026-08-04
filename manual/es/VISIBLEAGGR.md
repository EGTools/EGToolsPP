# VISIBLEAGGR

**Categoría**: función exclusiva de EGTools

Lista o agrega solo las celdas visibles de un rango.

## Sintaxis

```
=VISIBLEAGGR(rango, [funcion], [opcion])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| rango | Obligatorio | el rango de celdas |
| funcion | Opcional | SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S·P/VAR.S·P/CONCAT/TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC); omitido = listar celdas visibles |
| opcion | Opcional | delimitador de TEXTJOIN, k de LARGE/SMALL, p de PERCENTILE (0-1), QUARTILE 0-4 |

## Devuelve

Si se omite la función, derrama solo las celdas visibles como una matriz bidimensional; si se especifica una función, devuelve un valor agregado escalar. Si no hay celdas visibles devuelve #N/A; si una celda visible contiene un error, lo devuelve tal cual salvo con COUNT/COUNTA; un nombre de función no admitido devuelve #VALUE!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=VISIBLEAGGR(A1:B10)` |  | Enumerar solo las celdas visibles (depende del estado de ocultación) |
| `=VISIBLEAGGR(A1:A10,"SUM")` |  | Suma de las celdas visibles (depende del estado de ocultación) |
| `=VISIBLEAGGR(A1:A10,"LARGE",2)` |  | Segundo valor más grande entre los valores visibles |

## Notas

- Agregadores admitidos (21): SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/TEXTJOIN/LARGE/SMALL/PERCENTILE(.INC/.EXC)/QUARTILE(.INC/.EXC). Cualquier otro devuelve #VALUE!.
- Argumento de opción: TEXTJOIN=separador, LARGE/SMALL=k, familia PERCENTILE=p (0~1), familia QUARTILE=0~4.
- Cambiar la ocultación de filas/columnas no provoca un recálculo automático, por lo que se necesita F9. Al ser de tipo macro, se excluye del recálculo multiproceso.
- Compatibilidad: Excel 2010+. Se registra siempre como `VISIBLEAGGR` en todas las versiones de Excel.

# UNPIVOT

**Categoría**: función exclusiva de EGTools

Convierte una tabla cruzada (dinamizada) en una tabla de datos normalizada.

## Sintaxis

```
=UNPIVOT(matriz, [col_fijas], [filas_fijas], [orden], [nombres_col], [con_conjunto])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| matriz | Obligatorio | tabla origen (con encabezados) |
| col_fijas | Opcional | columnas de encabezado de fila a la izquierda (por defecto 1) |
| filas_fijas | Opcional | filas de encabezado de columna arriba (por defecto 1) |
| orden | Opcional | ordenación del resultado: número de columna, negativo descendente |
| nombres_col | Opcional | títulos para las columnas del resultado |
| con_conjunto | Opcional | VERDADERO si las columnas de valores se repiten como conjuntos |

## Devuelve

Devuelve la tabla de datos normalizada, incluida la fila de títulos, como una matriz derramada de dos dimensiones. Si la entrada no es una matriz; si las columnas o filas fijas son negativas, están fuera de rango o ambas son 0; o si el orden de clasificación es 0 o queda fuera de las columnas del resultado, devuelve #VALUE!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=UNPIVOT({"Item","Jan","Feb";"A",10,20;"B",30,40},1,1,,{"Item","Month","Qty"})` | {"Item","Month","Qty";"A","Jan",10;"A","Feb",20;"B","Jan",30;"B","Feb",40} | Anula la dinamización de una tabla cruzada |
| `=UNPIVOT({"Item","Jan","Feb";"A",10,20;"B",30,40},1,1,-3,{"Item","Month","Qty"})` | {"Item","Month","Qty";"B","Feb",40;"B","Jan",30;"A","Feb",20;"A","Jan",10} | Ordena por la columna 3 en orden descendente |

## Notas

- Funciona solo con valores; no se admite el relleno de celdas combinadas.
- Los elementos (celdas) cuyos valores están todos vacíos se excluyen del resultado.
- Si se omiten los títulos de tabla, los títulos predeterminados dependen del idioma de la interfaz (coreano/inglés).
- Compatibilidad: Excel 2010+. Se registra siempre como `UNPIVOT` en todas las versiones de Excel.

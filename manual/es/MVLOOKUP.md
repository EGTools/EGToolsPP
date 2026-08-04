# MVLOOKUP

**Categoría**: función exclusiva de EGTools

Ejecuta VLOOKUP para muchas claves a la vez, devolviendo una o más columnas.

## Sintaxis

```
=MVLOOKUP(valor_buscado, matriz_busqueda, columna_devuelta, [modo_coincidencia], [si_no_encontrado])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| valor_buscado | Obligatorio | columna de claves (o valor único) |
| matriz_busqueda | Obligatorio | tabla de datos; claves en la primera columna |
| columna_devuelta | Obligatorio | número(s) de columna a devolver, p. ej. 2 o {2,3} |
| modo_coincidencia | Opcional | FALSO exacta (por defecto), VERDADERO aproximada |
| si_no_encontrado | Opcional | valor sin coincidencia (por defecto "") |

## Devuelve

Devuelve una matriz derramada de tamaño filas de valores buscados × columnas de retorno; las filas sin coincidencia reciben if_not_found (predeterminado ""). Si falta un argumento obligatorio, la matriz de búsqueda no es una matriz, el valor buscado tiene más de una columna o un número de columna de retorno no es válido, devuelve #VALUE!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=MVLOOKUP({"a";"c"},{"a",1;"b",2},2)` | {1;""} | Coincidencia exacta de varias claves |
| `=MVLOOKUP("b",{"a",1,"x";"b",2,"y"},{2,3})` | {2,"y"} | Devuelve varias columnas a la vez |
| `=MVLOOKUP(15,{10,"low";20,"high"},2,TRUE)` | low | Coincidencia aproximada (el mayor valor menor o igual) |

## Notas

- La coincidencia exacta se basa en tablas hash; si hay claves duplicadas, se devuelve la primera fila.
- La coincidencia aproximada ordena internamente y usa búsqueda binaria, por lo que no es necesario ordenar los datos previamente.
- La comparación de claves no distingue mayúsculas y minúsculas.
- Compatibilidad: Excel 2010+. Se registra siempre como `MVLOOKUP` en todas las versiones de Excel.

# ADVANCEDFILTER

**Categoría**: función exclusiva de EGTools

Filtra una tabla con una cuadrícula de criterios estilo filtro avanzado (filas=O, columnas=Y).

## Sintaxis

```
=ADVANCEDFILTER(matriz_datos, matriz_filtro, [unicos])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| matriz_datos | Obligatorio | datos con fila de encabezado |
| matriz_filtro | Obligatorio | criterios con encabezado; admite =, <>, >, <, >=, <=, *, ? |
| unicos | Opcional | VERDADERO elimina filas duplicadas |

## Devuelve

Devuelve las filas filtradas, incluida la fila de encabezado, como una matriz derramada de dos dimensiones. Si los datos o el filtro no son matrices o tienen menos de 2 filas, o si un encabezado del filtro no existe entre los encabezados de los datos, devuelve #VALUE!; si ninguna fila de datos coincide, devuelve #N/A.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=ADVANCEDFILTER({"Name","Qty";"A",10;"B",5;"A",7},{"Qty";">6"})` | {"Name","Qty";"A",10;"A",7} | Extrae solo las filas con Qty>6 |
| `=ADVANCEDFILTER({"Name","Qty";"Apple",10;"Banana",5;"Avocado",7},{"Name";"A*"})` | {"Name","Qty";"Apple",10;"Avocado",7} | Filtro con el carácter comodín A* |

## Notas

- Un criterio de texto sin operador funciona como coincidencia de "comienza por", igual que el Filtro avanzado de Excel.
- Los caracteres comodín * y ? se admiten en los criterios = y <>.
- El resultado siempre incluye la fila de encabezado de los datos.
- Compatibilidad: Excel 2010+. Se registra siempre como `ADVANCEDFILTER` en todas las versiones de Excel.

# SUMBYBOM

**Categoría**: función exclusiva de EGTools

Expande un nivel de BOM y suma cantidades requeridas por artículo.

## Sintaxis

```
=SUMBYBOM(matriz1, ...)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| matriz1 | Obligatorio | columnas en orden: padre, hijo, uso, nombre, columna(s) de cantidad |
| … | Opcional | Los argumentos anteriores se pueden repetir |

## Devuelve

Devuelve una matriz derramada compuesta por una columna de nombres de artículo más una columna de suma por cada columna de cantidad, ordenada por nombre de artículo en orden ascendente. Si hay menos de 5 columnas, las columnas tienen longitudes distintas o el consumo o la cantidad no es numérico, devuelve #VALUE!; si hay cantidades negativas devuelve #NUM!; si no hay elementos que agregar devuelve #N/A.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=SUMBYBOM({"X","P",2,"X",1;"X","Q",3,"P",5})` | {"P",7;"Q",3} | Expande X un nivel y suma |

## Notas

- La lista de materiales (BOM) se expande solo un nivel (no es una expansión multinivel).
- Los nombres de artículo se suman sin distinguir mayúsculas y minúsculas; el resultado se ordena por nombre en orden ascendente.
- Los artículos que no son padres se excluyen del resultado si su cantidad es 0.
- Compatibilidad: Excel 2010+. Se registra siempre como `SUMBYBOM` en todas las versiones de Excel.

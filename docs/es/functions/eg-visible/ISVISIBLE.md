# ISVISIBLE

**Categoría**: función exclusiva de EGTools

Indica si cada celda de un rango es visible (fila y columna no ocultas).

## Sintaxis

```
=ISVISIBLE(rango)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| rango | Obligatorio | el rango de celdas a comprobar |

## Devuelve

Devuelve una matriz del mismo tamaño que el rango con TRUE si cada celda es visible y FALSE si su fila o columna está oculta, y se derrama. Devuelve #VALUE! si el número de celdas supera 100.000 o si falla la conexión COM.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=ISVISIBLE(A1:A5)` |  | Visibilidad de cada celda (el resultado depende del estado de ocultación) |

## Notas

- Cambiar solo la ocultación de filas/columnas no provoca un recálculo, por lo que se debe recalcular con F9.
- Es una función de tipo macro, por lo que se excluye del recálculo multiproceso.
- Compatibilidad: Excel 2010+. Se registra siempre como `ISVISIBLE` en todas las versiones de Excel.

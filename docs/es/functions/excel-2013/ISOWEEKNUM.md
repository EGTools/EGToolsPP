# ISOWEEKNUM

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Devuelve el número de semana ISO 8601 del año para una fecha.

## Sintaxis

```
=ISOWEEKNUM(fecha)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| fecha | Obligatorio | la fecha |

## Devuelve

Devuelve el número de semana ISO 8601 (1 a 53) de la fecha como número. Si el argumento es una matriz, se calcula elemento por elemento y se devuelve una matriz del mismo tamaño, que se derrama en las versiones de Excel con matrices dinámicas. Devuelve #VALUE! si la fecha no es numérica o es negativa, y #NUM! si el valor de serie no se puede convertir en fecha.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=ISOWEEKNUM(DATE(2026,1,1))` | 1 | 2026-01-01 es la semana 1 |
| `=ISOWEEKNUM(DATE(2023,1,1))` | 52 | Última semana del año anterior |

## Notas

- Con una matriz de entrada, calcula elemento por elemento y devuelve una matriz del mismo tamaño.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `ISOWEEKNUM` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.ISOWEEKNUM`.

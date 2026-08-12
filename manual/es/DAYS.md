# DAYS

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Devuelve el número de días entre dos fechas (fin − inicio).

## Sintaxis

```
=DAYS(fecha_final, fecha_inicial)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| fecha_final | Obligatorio | fecha final |
| fecha_inicial | Obligatorio | fecha inicial |

## Devuelve

Devuelve el número de días que resulta de restar la fecha inicial a la fecha final, como número (escalar). La parte de hora de cada fecha se descarta en el cálculo. Devuelve #VALUE! si se omite un argumento.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=DAYS(DATE(2026,3,1),DATE(2026,1,15))` | 45 | Días entre dos fechas |
| `=DAYS(DATE(2026,1,1),DATE(2026,1,31))` | -30 | Negativo si la fecha final es anterior |

## Notas

- El texto con formato de fecha (como "2026-01-15") no se interpreta y devuelve #VALUE!; las celdas vacías se tratan como 0.
- Ambos argumentos de fecha aceptan matrices y se calculan elemento a elemento, derramando una matriz de la misma forma: los escalares se repiten, un vector columna × un vector fila se expande como producto externo, los elementos con tamaños no coincidentes dan #N/A y un elemento de error devuelve ese error.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `DAYS` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.DAYS`.

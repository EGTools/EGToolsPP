# SHEET

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Devuelve el número de hoja de una referencia (o la hoja actual).

## Sintaxis

```
=SHEET([valor])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| valor | Opcional | referencia (opcional; pred.: hoja actual) |

## Devuelve

Devuelve como número la posición (a partir de 1) de la hoja a la que pertenece la referencia (o, si se omite, la celda donde se introdujo la fórmula). Devuelve #N/A si no se encuentra la hoja y #VALUE! si falla la consulta de la lista de hojas.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=SHEET()` |  | Posición de la hoja actual (depende del libro) |

## Notas

- Solo acepta referencias de celda como argumento (la función nativa también admite el nombre de la hoja como texto).
- Se registra como función de tipo hoja de macros.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `SHEET` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.SHEET`.

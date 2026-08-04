# MAXIFS

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2016/2019

Devuelve el máximo de las celdas que cumplen todos los criterios.

## Sintaxis

```
=MAXIFS(rango_máx, rango_criterios1, criterios1, ...)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| rango_máx | Obligatorio | celdas a maximizar |
| rango_criterios1 | Obligatorio | rango a evaluar |
| criterios1 | Obligatorio | criterio, p. ej. ">5" |
| … | Opcional | Los argumentos anteriores se pueden repetir |

## Devuelve

Devuelve como escalar numérico el valor máximo entre los valores numéricos de las celdas que cumplen todos los criterios. Devuelve #VALUE! si el tamaño de un rango de criterios difiere del rango de destino, y 0 si ningún valor cumple los criterios.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=MAXIFS({1;5;3},{1;2;3},">1")` | 5 | Máximo que cumple el criterio |
| `=MAXIFS({10;20},{1;2},">5")` | 0 | 0 si no hay coincidencias |

## Notas

- Los criterios admiten números, operadores de comparación (p. ej. ">5") y caracteres comodín (*, ?); la comparación de texto no distingue mayúsculas y minúsculas.
- Función relacionada: MINIFS.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `MAXIFS` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.MAXIFS`.

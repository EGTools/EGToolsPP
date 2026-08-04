# MINIFS

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2016/2019

Devuelve el mínimo de las celdas que cumplen todos los criterios.

## Sintaxis

```
=MINIFS(rango_mín, rango_criterios1, criterios1, ...)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| rango_mín | Obligatorio | celdas a minimizar |
| rango_criterios1 | Obligatorio | rango a evaluar |
| criterios1 | Obligatorio | criterio, p. ej. "<5" |
| … | Opcional | Los argumentos anteriores se pueden repetir |

## Devuelve

Devuelve como escalar numérico el valor mínimo entre los valores numéricos de las celdas que cumplen todos los criterios. Devuelve #VALUE! si el tamaño de un rango de criterios difiere del rango de destino, y 0 si ningún valor cumple los criterios.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=MINIFS({10;20;30},{"a";"b";"a"},"a")` | 10 | Mínimo que cumple el criterio |
| `=MINIFS({7;8;9},{"x";"y";"z"},"?")` | 7 | Criterio con caracteres comodín |

## Notas

- Los criterios admiten números, operadores de comparación (p. ej. "<5") y caracteres comodín (*, ?); la comparación de texto no distingue mayúsculas y minúsculas.
- Función relacionada: MAXIFS.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `MINIFS` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.MINIFS`.

# CONCAT

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2016/2019

Concatena texto de valores y rangos (sin delimitador).

## Sintaxis

```
=CONCAT(texto1, ...)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto1 | Obligatorio | texto o rango |
| … | Opcional | Los argumentos anteriores se pueden repetir |

## Devuelve

Devuelve un único valor de texto (escalar) que concatena todos los argumentos (incluidos rangos y matrices) en orden. Si algún argumento contiene un valor de error, se devuelve ese error.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=CONCAT("A",1,"B")` | A1B | Concatena los valores tal cual |
| `=CONCAT({1,2;3,4})` | 1234 | Matriz concatenada por filas |

## Notas

- Concatena sin separador; si se necesita un separador, use TEXTJOIN.
- Admite hasta 255 argumentos.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `CONCAT` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.CONCAT`.

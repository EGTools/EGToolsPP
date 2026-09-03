# XOR

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Devuelve el O exclusivo lógico (VERDADERO si un número impar de argumentos es VERDADERO).

## Sintaxis

```
=XOR(logico1, ...)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| logico1 | Obligatorio | valor o rango a evaluar |
| … | Opcional | Los argumentos anteriores se pueden repetir |

## Devuelve

Devuelve TRUE como escalar lógico si el número de valores TRUE es impar, y FALSE si es par. Si hay errores entre los argumentos, devuelve el primer error tal cual; si no hay ningún valor evaluable como lógico, devuelve #VALUE!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=XOR(TRUE,FALSE)` | TRUE | Número impar de TRUE |
| `=XOR(TRUE,TRUE)` | FALSE | Número par de TRUE |
| `=XOR(1,0,1,"text")` | FALSE | El texto se ignora |

## Notas

- El texto y los valores vacíos se ignoran (la función nativa devuelve #VALUE! con texto no convertible).
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `XOR` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.XOR`.

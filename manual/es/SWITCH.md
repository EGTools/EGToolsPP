# SWITCH

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2016/2019

Compara una expresión con valores y devuelve la primera coincidencia (o el predeterminado).

## Sintaxis

```
=SWITCH(expresión, valor1, resultado1, [predeterminado], ...)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| expresión | Obligatorio | valor a comparar |
| valor1 | Obligatorio | valor a coincidir |
| resultado1 | Obligatorio | resultado si coincide el valor |
| predeterminado | Opcional | predeterminado / valor2 … |
| … | Opcional | Los argumentos anteriores se pueden repetir |

## Devuelve

Devuelve tal cual el resultado del primer valor que coincide con la expresión (o el valor predeterminado final). Devuelve #N/A si no hay ningún valor coincidente ni valor predeterminado.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=SWITCH(2,1,"one",2,"two","other")` | two | Resultado del valor coincidente |
| `=SWITCH(9,1,"one","other")` | other | Devuelve el valor predeterminado |

## Notas

- Si se deja vacío un argumento intermedio, se ignoran los pares valor/resultado posteriores.
- Función relacionada: IFS.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `SWITCH` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.SWITCH`.

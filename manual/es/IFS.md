# IFS

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2016/2019

Comprueba condiciones en orden y devuelve el valor de la primera VERDADERA.

## Sintaxis

```
=IFS(condición1, valor1, ...)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| condición1 | Obligatorio | condición a comprobar |
| valor1 | Obligatorio | resultado si condición1 es VERDADERO |
| … | Opcional | Los argumentos anteriores se pueden repetir |

## Devuelve

Devuelve tal cual el valor de la primera condición verdadera (escalar o matriz). Devuelve #N/A si ninguna condición es verdadera; si una condición es un valor de error, devuelve ese error tal cual.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=IFS(1>2,"a",2>1,"b")` | b | Valor de la primera condición TRUE |
| `=IFS(1>2,"a")` | #N/A | Ninguna condición verdadera |

## Notas

- Las condiciones de texto o vacías se tratan como FALSE (la función nativa devuelve #VALUE!).
- Función relacionada: SWITCH.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `IFS` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.IFS`.

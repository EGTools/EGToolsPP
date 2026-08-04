# UNICODE

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Devuelve el punto de código Unicode del primer carácter del texto.

## Sintaxis

```
=UNICODE(texto)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto | Obligatorio | texto cuyo primer carácter se examina |

## Devuelve

Devuelve el punto de código (número escalar) del primer carácter; si se especifica una matriz de textos, se evalúa elemento por elemento y el resultado se derrama. Devuelve #VALUE! si el texto está vacío.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=UNICODE("A")` | 65 | letra latina |
| `=UNICODE("가")` | 44032 | sílaba hangul |
| `=UNICODE("😀")` | 128512 | carácter fuera del BMP |

## Notas

- Reconoce los pares suplentes, por lo que los caracteres fuera del BMP también devuelven el punto de código completo.
- Función relacionada: UNICHAR
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `UNICODE` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.UNICODE`.

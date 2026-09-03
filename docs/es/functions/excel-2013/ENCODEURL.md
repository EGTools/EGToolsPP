# ENCODEURL

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Devuelve una cadena codificada para URL (codificación porcentual).

## Sintaxis

```
=ENCODEURL(texto)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto | Obligatorio | texto a codificar para URL |

## Devuelve

Devuelve el texto codificado como URL en forma de escalar. Si se introduce una matriz, se aplica elemento por elemento y se derrama una matriz del mismo tamaño; la entrada se convierte a texto, por lo que no se producen valores de error específicos.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=ENCODEURL("a b&c")` | a%20b%26c | Convierte los espacios y el carácter & en %XX |
| `=ENCODEURL("한글")` | %ED%95%9C%EA%B8%80 | Codificación UTF-8 de texto coreano |

## Notas

- Solo se conservan los caracteres no reservados de RFC 3986 (letras, dígitos y - _ . ~); el resto se codifica como %XX por cada byte UTF-8.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `ENCODEURL` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.ENCODEURL`.

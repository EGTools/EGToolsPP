# DECODEURL

**Categoría**: función exclusiva de EGTools

Decodifica una cadena URL con codificación porcentual.

## Sintaxis

```
=DECODEURL(texto)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto | Obligatorio | el texto codificado como URL |

## Devuelve

Devuelve la cadena resultante de revertir la codificación porcentual (UTF-8) al texto original. Si se introduce una matriz, se procesa elemento por elemento y se derrama como una matriz del mismo tamaño.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=DECODEURL("a%20b%26c")` | a b&c | Decodificar codificación porcentual |
| `=DECODEURL("%ED%95%9C%EA%B8%80")` | 한글 | Decodificar texto coreano en UTF-8 |

## Notas

- Las secuencias % no válidas se dejan tal cual, sin convertir.
- El signo + no se convierte en espacio (no se admite la codificación de formularios).
- Es la función inversa de ENCODEURL.
- Compatibilidad: Excel 2010+. Se registra siempre como `DECODEURL` en todas las versiones de Excel.

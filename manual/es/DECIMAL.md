# DECIMAL

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Convierte el texto de un número en una base dada a un número decimal.

## Sintaxis

```
=DECIMAL(text, radix)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| text | Obligatorio | el texto a convertir |
| radix | Obligatorio | base, de 2 a 36 |

## Devuelve

Devuelve el número decimal convertido a partir del texto en la base indicada (escalar). Si radix está fuera del rango 2 a 36 o si text contiene un carácter no válido en esa base, devuelve un error #NUM!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=DECIMAL("FF",16)` | 255 | Hexadecimal a decimal |
| `=DECIMAL("111",2)` | 7 | Binario a decimal |

## Notas

- No distingue mayúsculas y minúsculas.
- Todos los argumentos aceptan matrices y se calculan elemento a elemento, derramando una matriz de la misma forma: los escalares se repiten, un vector columna × un vector fila se expande como producto externo, los elementos con tamaños no coincidentes dan #N/A y un elemento de error devuelve ese error.
- Función relacionada: BASE
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `DECIMAL` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.DECIMAL`.

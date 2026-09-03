# BASE

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Convierte un número en texto en la base (radix) indicada.

## Sintaxis

```
=BASE(number, radix, [min_length])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| number | Obligatorio | el número a convertir (>= 0) |
| radix | Obligatorio | base, de 2 a 36 |
| min_length | Opcional | longitud mínima del resultado (opcional) |

## Devuelve

Devuelve el texto que representa el número en la base indicada (escalar). Si number es negativo o no es numérico, o si radix está fuera del rango 2 a 36, devuelve un error #NUM!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=BASE(255,16)` | FF | 255 en hexadecimal |
| `=BASE(15,2,8)` | 00001111 | Relleno con ceros hasta la longitud mínima |

## Notas

- La parte fraccionaria de number se descarta (redondeo hacia abajo).
- Si el resultado es más corto que min_length, se rellena con ceros por la izquierda.
- Todos los argumentos aceptan matrices y se calculan elemento a elemento, derramando una matriz de la misma forma: los escalares se repiten, un vector columna × un vector fila se expande como producto externo, los elementos con tamaños no coincidentes dan #N/A y un elemento de error devuelve ese error.
- Función relacionada: DECIMAL
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `BASE` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.BASE`.

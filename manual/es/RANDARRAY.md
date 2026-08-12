# RANDARRAY

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2021

Devuelve una matriz de números aleatorios.

## Sintaxis

```
=RANDARRAY([filas], [columnas], [mín], [máx], [entero])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| filas | Opcional | número de filas (pred. 1) |
| columnas | Opcional | número de columnas (pred. 1) |
| mín | Opcional | mínimo (pred. 0) |
| máx | Opcional | máximo (pred. 1) |
| entero | Opcional | VERDADERO para enteros |

## Devuelve

Devuelve una matriz de números aleatorios de tamaño rows×cols; se derrama. Devuelve #VALUE! si rows o cols es 0 o menor, o si max es menor que min.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=RANDARRAY(2,3,1,10,TRUE)` |  | Aleatorio; el resultado cambia cada vez |

## Notas

- Con integer=TRUE genera enteros de min a max (ambos incluidos); en caso contrario, números reales mayores o iguales que min y menores que max.
- Una matriz en Todos los argumentos devuelve una matriz con la forma de ese argumento, donde cada elemento es el primer valor del resultado de ese elemento (reducción al primer valor, igual que la función nativa).
- Función relacionada: SEQUENCE.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `RANDARRAY` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.RANDARRAY`.

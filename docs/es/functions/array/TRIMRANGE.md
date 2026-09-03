# TRIMRANGE

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Microsoft 365

Recorta filas/columnas vacías de los bordes de una matriz.

## Sintaxis

```
=TRIMRANGE(matriz, [recortar_filas], [recortar_columnas])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| matriz | Obligatorio | la matriz o los valores del rango |
| recortar_filas | Opcional | 0 ninguna, 1 iniciales, 2 finales, 3 ambas (predeterminado) |
| recortar_columnas | Opcional | 0 ninguna, 1 iniciales, 2 finales, 3 ambas (predeterminado) |

## Devuelve

Devuelve la matriz sin las filas/columnas vacías de los bordes y se derrama. Devuelve #VALUE! si la matriz está vacía, si un modo de recorte está fuera del rango 0~3 o si todos los valores están vacíos.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=TRIMRANGE({"","";"a",1;"",""})` | {"a",1} | Eliminar filas y columnas vacías de los bordes |
| `=TRIMRANGE({"";1;""},1)` | {1;""} | Eliminar solo las filas vacías iniciales |

## Notas

- La función TRIMRANGE nativa recorta una referencia; esta implementación se basa en valores, por lo que también se puede usar con literales de matriz y resultados de fórmulas.
- Las cadenas vacías ("") también se consideran valores vacíos.
- Una matriz en el modo de recorte devuelve una matriz con la forma de ese argumento, donde cada elemento es el primer valor del resultado de ese elemento (reducción al primer valor, igual que la función nativa).
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `TRIMRANGE` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.TRIMRANGE`.

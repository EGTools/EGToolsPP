# SEQUENCE

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2021

Genera una secuencia de números como matriz.

## Sintaxis

```
=SEQUENCE(filas, [columnas], [inicio], [paso])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| filas | Obligatorio | número de filas |
| columnas | Opcional | número de columnas (pred. 1) |
| inicio | Opcional | primer valor (pred. 1) |
| paso | Opcional | incremento (pred. 1) |

## Devuelve

Devuelve una matriz de números consecutivos de tamaño rows×cols (incrementados por filas); se derrama. Devuelve #VALUE! si rows o cols es 0 o menor.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=SEQUENCE(3)` | {1;2;3} | Secuencia básica de 3 filas |
| `=SEQUENCE(2,3,0,10)` | {0,10,20;30,40,50} | Valor inicial e incremento personalizados |

## Notas

- Si se omite rows, se toma como 1 (en la función nativa es un argumento obligatorio).
- Una matriz en Todos los argumentos devuelve una matriz con la forma de ese argumento, donde cada elemento es el primer valor del resultado de ese elemento (reducción al primer valor, igual que la función nativa).
- Función relacionada: RANDARRAY.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `SEQUENCE` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.SEQUENCE`.

# UNIQUE

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2021

Devuelve las filas (o columnas) únicas de una matriz.

## Sintaxis

```
=UNIQUE(matriz, [por_columna], [exactamente_una_vez])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| matriz | Obligatorio | datos |
| por_columna | Opcional | VERDADERO para columnas únicas |
| exactamente_una_vez | Opcional | VERDADERO para elementos que aparecen una vez |

## Devuelve

Devuelve una matriz con las filas (o columnas) únicas, sin duplicados; se derrama. Devuelve #N/A si no queda ningún elemento y #VALUE! en caso de error interno.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=UNIQUE({1;2;2;3})` | {1;2;3} | Eliminación de duplicados |
| `=UNIQUE({1;2;2;3},,TRUE)` | {1;3} | Valores que aparecen una sola vez |

## Notas

- Si el resultado está vacío, devuelve #N/A en lugar de #CALC!.
- Funciones relacionadas: FILTER, SORT.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `UNIQUE` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.UNIQUE`.

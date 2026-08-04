# XMATCH

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2021

Devuelve la posición de un valor dentro de una matriz.

## Sintaxis

```
=XMATCH(valor_buscado, matriz_buscada, [modo_coincidencia], [modo_búsqueda])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| valor_buscado | Obligatorio | valor a buscar |
| matriz_buscada | Obligatorio | matriz donde buscar |
| modo_coincidencia | Opcional | 0 exacto (pred.), -1 menor, 1 mayor |
| modo_búsqueda | Opcional | 1 inicio→fin (pred.), -1 fin→inicio |

## Devuelve

Devuelve la posición encontrada (a partir de 1) como escalar numérico. Devuelve #N/A si la matriz está vacía o no hay coincidencia, y #VALUE! en caso de error interno.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=XMATCH(3,{1;2;3;4})` | 3 | Posición de coincidencia exacta |
| `=XMATCH(2.5,{1;2;3},1)` | 3 | Posición del valor mayor o igual |
| `=XMATCH(5,{1;2;3})` | #N/A | #N/A si no hay coincidencia |

## Notas

- No se admiten match_mode 2 (caracteres comodín) ni search_mode 2/-2 (búsqueda binaria).
- Función relacionada: XLOOKUP.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `XMATCH` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.XMATCH`.

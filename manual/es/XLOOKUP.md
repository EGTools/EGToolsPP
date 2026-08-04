# XLOOKUP

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2021

Busca una coincidencia en un rango o matriz y devuelve el elemento correspondiente de un segundo rango o matriz.

## Sintaxis

```
=XLOOKUP(valor_buscado, matriz_buscada, matriz_devuelta, [si_no_encontrado], [modo_coincidencia], [modo_busqueda])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| valor_buscado | Obligatorio | el valor a buscar |
| matriz_buscada | Obligatorio | el rango/matriz donde buscar |
| matriz_devuelta | Obligatorio | el rango/matriz a devolver |
| si_no_encontrado | Opcional | valor si no hay coincidencia (opcional) |
| modo_coincidencia | Opcional | 0 exacto (pred.), -1 menor, 1 mayor |
| modo_busqueda | Opcional | 1 de inicio a fin (pred.), -1 de fin a inicio |

## Devuelve

Devuelve como matriz la fila (búsqueda vertical) o la columna (búsqueda horizontal) coincidente del rango de retorno; en Excel con matrices dinámicas se derrama. Devuelve #N/A si se omite un rango, la matriz está vacía o no hay coincidencia y no se indica un valor alternativo; #REF! si la posición coincidente queda fuera del rango de retorno; #VALUE! en caso de error interno.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=XLOOKUP("b",{"a";"b";"c"},{10;20;30})` | 20 | Búsqueda con coincidencia exacta |
| `=XLOOKUP(2,{1;2;3},{10,11;20,21;30,31})` | {20,21} | La fila coincidente completa se derrama |
| `=XLOOKUP(9,{1;2;3},{10;20;30},"none")` | none | Valor alternativo si no hay coincidencia |

## Notas

- No se admiten match_mode 2 (caracteres comodín) ni search_mode 2/-2 (búsqueda binaria).
- Si lookup_array tiene una sola columna se realiza una búsqueda vertical; en caso contrario se busca horizontalmente en la primera fila.
- Funciones relacionadas: XMATCH, FILTER.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `XLOOKUP` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.XLOOKUP`.

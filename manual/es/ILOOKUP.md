# ILOOKUP

**Categoría**: función exclusiva de EGTools

Busca una coincidencia y copia la imagen anclada en esa celda a la celda que llama.

## Sintaxis

```
=ILOOKUP(valor_buscado, matriz_busqueda, rango_imagenes, [si_no_encontrado], [modo_coincidencia], [modo_busqueda])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| valor_buscado | Obligatorio | el valor a buscar |
| matriz_busqueda | Obligatorio | la matriz o rango donde buscar |
| rango_imagenes | Obligatorio | rango cuyas celdas anclan las imágenes |
| si_no_encontrado | Opcional | valor si no hay coincidencia (predeterminado #N/A) |
| modo_coincidencia | Opcional | 0 exacta (predeterminado), -1 menor, 1 mayor, 2 comodines, 3 regex |
| modo_busqueda | Opcional | 1 desde el inicio (predeterminado), -1 desde el final |

## Devuelve

Busca una coincidencia y, si tiene éxito, devuelve una cadena vacía (""), y la imagen colocada en la celda coincidente se inserta por separado ajustada al tamaño de la celda que llama. Devuelve #REF! si el rango de búsqueda y el rango de imágenes tienen tamaños distintos, el valor de si_no_encontrado (predeterminado #N/A) si no hay coincidencia o no hay imagen, y #VALUE! si la expresión regular no es válida o si falla COM o el portapapeles.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=ILOOKUP("apple",A1:A5,B1:B5)` |  | Copiar la imagen colocada en la fila coincidente (inserción de imagen, depende del entorno) |

## Notas

- Copia a través del portapapeles la imagen (msoPicture/msoLinkedPicture) cuya TopLeftCell es la celda coincidente del rango de imágenes.
- modo_coincidencia: 0 exacta (predeterminado), -1 valor menor, 1 valor mayor, 2 caracteres comodín, 3 expresión regular; modo_busqueda: 1 desde el principio (predeterminado), -1 desde el final.
- Mover imágenes o cambiar formatos por sí solo no provoca un recálculo (se necesita F9); al ser de tipo macro, también se excluye del recálculo multiproceso.
- Compatibilidad: Excel 2010+. Se registra siempre como `ILOOKUP` en todas las versiones de Excel.

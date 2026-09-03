# BOMTREE

**Categoría**: función exclusiva de EGTools

Construye un árbol BOM a partir de listas padre/hijo/cantidad.

## Sintaxis

```
=BOMTREE(padres, hijos, valores, [encabezado], [wbs])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| padres | Obligatorio | columna de artículos padre |
| hijos | Obligatorio | columna de artículos hijo (misma longitud) |
| valores | Obligatorio | columna de cantidades (misma longitud) |
| encabezado | Opcional | VERDADERO añade fila de encabezado |
| wbs | Opcional | VERDADERO añade columna WBS |

## Devuelve

Devuelve el árbol de la lista de materiales (BOM) expandido como una matriz derramada de dos dimensiones con columnas de nivel, artículo padre, artículo hijo y valor (5 columnas si se selecciona WBS). Si las tres listas están vacías o tienen longitudes distintas devuelve #VALUE!; si un mismo par padre-hijo aparece con valores diferentes, devuelve el texto "Error: ...".

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=BOMTREE({"X";"X";"Y"},{"Y";"Z";"W"},{1;2;3})` | {0,"Top Level","X","";1,"X","Y",1;2,"Y","W",3;1,"X","Z",2} | Expansión del árbol BOM |
| `=BOMTREE({"X";"X";"Y"},{"Y";"Z";"W"},{1;2;3},TRUE)` | {"level","Parent","Child","Value";0,"Top Level","X","";1,"X","Y",1;2,"Y","W",3;1,"X","Z",2} | Incluye la fila de encabezado |

## Notas

- Los números WBS se rellenan hasta 5 dígitos con espacios de ancho cero (U+200B) para mantener el orden correcto.
- La profundidad de expansión está limitada a 64 niveles (protección contra referencias circulares).
- La lectura de las listas se detiene en la primera fila cuya celda de artículo padre está vacía.
- Compatibilidad: Excel 2010+. Se registra siempre como `BOMTREE` en todas las versiones de Excel.

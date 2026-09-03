# CPK

**Categoría**: función exclusiva de EGTools

Índice de capacidad de proceso Cpk a partir de mediciones y límites de especificación.

## Sintaxis

```
=CPK(datos, [usl], [lsl])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| datos | Obligatorio | rango o matriz de mediciones |
| usl | Opcional | límite superior (opcional si hay LSL) |
| lsl | Opcional | límite inferior (opcional si hay USL) |

## Devuelve

Devuelve el índice de capacidad de proceso Cpk = min((USL-μ)/3σ, (μ-LSL)/3σ) como número escalar. Si falta data o faltan tanto USL como LSL, devuelve #VALUE!; con menos de 2 valores numéricos devuelve #NUM!; si la desviación estándar es 0, devuelve #DIV/0!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=CPK({1;2;3;4;5},5,0)` | 0.421637021 | Cpk para los límites de especificación 0 a 5 |

## Notas

- Utiliza la desviación estándar muestral (n-1).
- Si solo se indica un límite, se calcula únicamente con ese lado.
- USL·LSL aceptan matrices y se calculan elemento a elemento, derramando una matriz de la misma forma: los escalares se repiten, un vector columna × un vector fila se expande como producto externo, los elementos con tamaños no coincidentes dan #N/A y un elemento de error devuelve ese error.
- Función relacionada: CP
- Compatibilidad: Excel 2010+. Se registra siempre como `CPK` en todas las versiones de Excel.

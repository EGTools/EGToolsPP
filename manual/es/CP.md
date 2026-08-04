# CP

**Categoría**: función exclusiva de EGTools

Índice de capacidad de proceso Cp a partir de mediciones y límites de especificación.

## Sintaxis

```
=CP(datos, [usl], [lsl])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| datos | Obligatorio | rango o matriz de mediciones |
| usl | Opcional | límite superior (opcional si hay LSL) |
| lsl | Opcional | límite inferior (opcional si hay USL) |

## Devuelve

Devuelve el índice de capacidad de proceso Cp = (USL-LSL)/(6σ) como número escalar. Si falta data o faltan tanto USL como LSL, devuelve #VALUE!; con menos de 2 valores numéricos devuelve #NUM!; si la desviación estándar es 0, devuelve #DIV/0!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=CP({1;2;3;4;5},6,0)` | 0.632455532 | Cp para los límites de especificación 0 a 6 |

## Notas

- Utiliza la desviación estándar muestral (n-1).
- Si solo se indica un límite, el lado faltante se considera igual a la media.
- Función relacionada: CPK
- Compatibilidad: Excel 2010+. Se registra siempre como `CP` en todas las versiones de Excel.

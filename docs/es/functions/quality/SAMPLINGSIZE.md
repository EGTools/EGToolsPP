# SAMPLINGSIZE

**Categoría**: función exclusiva de EGTools

Devuelve el tamaño de muestra (n) ISO 2859-1.

## Sintaxis

```
=SAMPLINGSIZE(tam_lote, aql, [nivel], [rigor], [plan], [ronda])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| tam_lote | Obligatorio | cantidad del lote |
| aql | Obligatorio | límite de calidad aceptable, p. ej. 1.0 |
| nivel | Opcional | G1/G2/G3/S1–S4 (por defecto G2) |
| rigor | Opcional | 0 normal (por defecto), 1 estricta, -1 reducida |
| plan | Opcional | 1 simple, 2 doble, 3–5 múltiple; +10 fuerza sin cambiar |
| ronda | Opcional | ronda de muestreo |

## Devuelve

Devuelve el tamaño de muestra n según ISO 2859-1 como número escalar. Si el AQL está fuera de 0~1000 o el nivel de inspección, el plan o la ronda no son válidos, devuelve #VALUE!; si la cantidad del lote es inferior a 2 o el AQL es inferior a 0.01, devuelve #NUM!; si no se encuentra el valor correspondiente en la tabla, devuelve #N/A.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=SAMPLINGSIZE(1000,1)` | 80 | G2, inspección normal, muestreo simple, AQL 1.0 |
| `=SAMPLINGSIZE(1000,1,"G2",0,2,1)` | 50 | Tamaño de la primera muestra en muestreo doble |
| `=SAMPLINGSIZE(30,0.065)` | 30 | Limitado por la cantidad del lote |

## Notas

- Si el n calculado supera la cantidad del lote, se devuelve la cantidad del lote; si la cantidad del lote es 1, se devuelve 1.
- Sumar 10 al plan fuerza ese plan e ignora las reglas de cambio de la tabla (*, ++).
- Todos los argumentos aceptan matrices, por lo que una lista de lotes o una tabla de AQL puede calcularse de una vez: elemento a elemento con derrame de la misma forma; los escalares se repiten, un vector columna × un vector fila se expande como producto externo y los elementos con tamaños no coincidentes dan #N/A.
- Funciones relacionadas: SAMPLINGLABEL, SAMPLINGAC, SAMPLINGRE
- Compatibilidad: Excel 2010+. Se registra siempre como `SAMPLINGSIZE` en todas las versiones de Excel.

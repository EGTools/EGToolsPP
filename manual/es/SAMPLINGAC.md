# SAMPLINGAC

**Categoría**: función exclusiva de EGTools

Devuelve el número de aceptación (Ac) ISO 2859-1.

## Sintaxis

```
=SAMPLINGAC(tam_lote, aql, [nivel], [rigor], [plan], [ronda])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| tam_lote | Obligatorio | cantidad del lote |
| aql | Obligatorio | límite de calidad aceptable |
| nivel | Opcional | por defecto G2 |
| rigor | Opcional | 0 normal, 1 estricta, -1 reducida |
| plan | Opcional | 1 simple, 2 doble, 3–5 múltiple |
| ronda | Opcional | ronda de muestreo |

## Devuelve

Devuelve el número de aceptación Ac según ISO 2859-1 como número escalar. Si el AQL está fuera de 0~1000 o el nivel de inspección, el plan o la ronda no son válidos, devuelve #VALUE!; si la cantidad del lote es inferior a 2 o el AQL es inferior a 0.01, devuelve #NUM!; si no se encuentra el valor correspondiente en la tabla, devuelve #N/A.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=SAMPLINGAC(1000,1)` | 2 | G2, inspección normal, muestreo simple, AQL 1.0 |
| `=SAMPLINGAC(1000,1,"G2",1)` | 1 | Inspección rigurosa |
| `=SAMPLINGAC(500,2.5)` | 3 | Lote de 500, AQL 2.5 |

## Notas

- Si la cantidad del lote es 1, devuelve 0.
- Funciones relacionadas: SAMPLINGSIZE, SAMPLINGRE, SAMPLINGLABEL
- Compatibilidad: Excel 2010+. Se registra siempre como `SAMPLINGAC` en todas las versiones de Excel.

# SAMPLINGRE

**Categoría**: función exclusiva de EGTools

Devuelve el número de rechazo (Re) ISO 2859-1.

## Sintaxis

```
=SAMPLINGRE(tam_lote, aql, [nivel], [rigor], [plan], [ronda])
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

Devuelve el número de rechazo Re según ISO 2859-1 como número escalar. Si el AQL está fuera de 0~1000 o el nivel de inspección, el plan o la ronda no son válidos, devuelve #VALUE!; si la cantidad del lote es inferior a 2 o el AQL es inferior a 0.01, devuelve #NUM!; si no se encuentra el valor correspondiente en la tabla, devuelve #N/A.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=SAMPLINGRE(1000,1)` | 3 | G2, inspección normal, muestreo simple, AQL 1.0 |
| `=SAMPLINGRE(1000,1,"G2",0,2,1)` | 3 | Re de la primera ronda en muestreo doble |

## Notas

- En el muestreo simple, Re = Ac + 1.
- Si la cantidad del lote es 1, devuelve 1.
- Funciones relacionadas: SAMPLINGSIZE, SAMPLINGAC, SAMPLINGLABEL
- Compatibilidad: Excel 2010+. Se registra siempre como `SAMPLINGRE` en todas las versiones de Excel.

# SAMPLINGLABEL

**Categoría**: función exclusiva de EGTools

Devuelve la letra de código de muestra ISO 2859-1 según el tamaño de lote y nivel de inspección.

## Sintaxis

```
=SAMPLINGLABEL(tam_lote, [nivel])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| tam_lote | Obligatorio | cantidad del lote (2 o más) |
| nivel | Opcional | nivel de inspección G1/G2/G3 o S1–S4 (por defecto G2) |

## Devuelve

Devuelve la letra código de muestra según ISO 2859-1 como texto escalar de un carácter. Si la cantidad del lote es inferior a 2, devuelve #NUM!; si el nivel de inspección no es S1~S4 ni G1~G3, devuelve #VALUE!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=SAMPLINGLABEL(1000)` | J | G2 predeterminado, lote de 1000 |
| `=SAMPLINGLABEL(1000,"S3")` | E | Nivel de inspección especial S-3 |

## Notas

- Los guiones y espacios en el nivel de inspección se ignoran ("G-2" = "G2").
- Calcula con las tablas ISO 2859-1:1999 integradas.
- Todos los argumentos aceptan matrices, por lo que una lista de lotes o una tabla de AQL puede calcularse de una vez: elemento a elemento con derrame de la misma forma; los escalares se repiten, un vector columna × un vector fila se expande como producto externo y los elementos con tamaños no coincidentes dan #N/A.
- Funciones relacionadas: SAMPLINGSIZE, SAMPLINGAC, SAMPLINGRE
- Compatibilidad: Excel 2010+. Se registra siempre como `SAMPLINGLABEL` en todas las versiones de Excel.

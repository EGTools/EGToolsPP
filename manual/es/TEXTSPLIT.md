# TEXTSPLIT

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2024 / Microsoft 365

Divide texto en una matriz usando delimitadores de columna y fila.

## Sintaxis

```
=TEXTSPLIT(texto, delim_columna, [delim_fila], [ignorar_vacíos], [modo_coincidencia], [rellenar_con])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto | Obligatorio | texto a dividir |
| delim_columna | Obligatorio | delimitador entre columnas |
| delim_fila | Opcional | delimitador entre filas (opcional) |
| ignorar_vacíos | Opcional | VERDADERO para omitir vacíos |
| modo_coincidencia | Opcional | 0 distingue mayúsc. (pred.), 1 no |
| rellenar_con | Opcional | valor para filas irregulares (pred. #N/A) |

## Devuelve

Devuelve una matriz con el texto dividido, que se derrama en las versiones de Excel que admiten matrices dinámicas. Las filas de distinta longitud se rellenan con el valor de relleno (predeterminado #N/A) y, si el resultado de la división queda vacío, devuelve #VALUE!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=TEXTSPLIT("a,b,c",",")` | {a,b,c} | división en columnas |
| `=TEXTSPLIT("a,b;c",",",";")` | {a,b;c,#N/A} | división en filas y columnas, con relleno de los huecos |
| `=TEXTSPLIT("1,,2",",",,TRUE)` | {1,2} | omite los valores vacíos |

## Notas

- El argumento de separador solo admite un texto único (no se admite la forma de matriz con varios separadores que permite la función nativa).
- Con match_mode=1 se ignoran mayúsculas y minúsculas.
- Funciones relacionadas: TEXTBEFORE, TEXTAFTER, TEXTJOIN
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `TEXTSPLIT` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.TEXTSPLIT`.

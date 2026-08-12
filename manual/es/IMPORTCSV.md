# IMPORTCSV

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Microsoft 365 (novedad de 2026)

Importa un archivo CSV (ruta local o URL) como matriz — coma y UTF-8.

## Sintaxis

```
=IMPORTCSV(ruta, [omitir_filas], [tomar_filas], [configuracion_regional])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| ruta | Obligatorio | ruta de archivo local o URL |
| omitir_filas | Opcional | filas a omitir; negativo desde el final |
| tomar_filas | Opcional | filas a tomar; negativo desde el final |
| configuracion_regional | Opcional | regional para números, p. ej. "de-DE" |

## Devuelve

Devuelve el contenido del CSV como una matriz derramada de dos dimensiones; el texto numérico y de fecha (formato yyyy-M-d, etc.) se convierte automáticamente en valores. Si falta la ruta o falla la lectura o descarga del archivo devuelve #VALUE!; si no queda ninguna fila de contenido devuelve #N/A.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=IMPORTCSV("C:\data\sales.csv",1)` |  | Depende del contenido del archivo |

## Notas

- El separador es la coma y la codificación es UTF-8, ambos fijos (si hay BOM, el BOM tiene prioridad). Si necesita otro separador u otra codificación, use IMPORTTEXT.
- Se reconocen los campos entre comillas dobles (incluido el escape "").
- Si se indican valores negativos en skip_rows/take_rows, se cuentan desde el final.
- Una matriz en la ruta y take_rows devuelve #VALUE!; para varios elementos, copie la fórmula fila por fila.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `IMPORTCSV` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.IMPORTCSV`.

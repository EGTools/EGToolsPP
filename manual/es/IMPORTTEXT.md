# IMPORTTEXT

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 365 (novedad de 2026)

Importa datos de un archivo de texto (ruta local o URL) como matriz.

## Sintaxis

```
=IMPORTTEXT(ruta, [delimitador], [omitir_filas], [tomar_filas], [codificacion], [configuracion_regional])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| ruta | Obligatorio | ruta de archivo local o URL |
| delimitador | Opcional | texto delimitador (por defecto tab), o números ascendentes como {1,3} para ancho fijo |
| omitir_filas | Opcional | filas a omitir; negativo omite desde el final |
| tomar_filas | Opcional | filas a tomar; negativo toma desde el final |
| codificacion | Opcional | por defecto "utf-8" — windows-1252, euc-kr… |
| configuracion_regional | Opcional | regional para números, p. ej. "de-DE" |

## Devuelve

Devuelve un archivo de texto (ruta local o URL) dividido por separadores o por anchos fijos como una matriz derramada de dos dimensiones; el texto numérico y de fecha se convierte automáticamente. Si falta la ruta, falla la lectura o descarga, o la especificación de separador o de ancho fijo no es válida, devuelve #VALUE!; si no queda ninguna fila tras aplicar skip/take, devuelve #N/A.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=IMPORTTEXT("C:\data\log.txt",";",1)` |  | Depende del contenido del archivo |

## Notas

- Si se omite delimiter, el valor predeterminado es la tabulación. Las cadenas se interpretan como separadores; una matriz numérica ({1,11,21}) se interpreta como posiciones iniciales de columnas de ancho fijo (base 1, en orden ascendente).
- encoding admite nombres de juego de caracteres como "euc-kr" o números de página de códigos; si hay BOM, el BOM tiene prioridad.
- Si se indican valores negativos en skip_rows/take_rows, se cuentan desde el final.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `IMPORTTEXT` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.IMPORTTEXT`.

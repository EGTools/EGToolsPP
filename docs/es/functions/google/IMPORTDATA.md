# IMPORTDATA

**Categoría**: función de compatibilidad con Google Sheets

Importa datos CSV/TSV desde una URL o una ruta de archivo local.

## Sintaxis

```
=IMPORTDATA(url_o_ruta, [delimitador], [configuracion_regional], [juego_caracteres])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| url_o_ruta | Obligatorio | URL web o ruta de archivo local con los datos |
| delimitador | Opcional | delimitador de campos (por defecto: autodetectar tab/coma) |
| configuracion_regional | Opcional | configuración regional para números, p. ej. "de-DE" |
| juego_caracteres | Opcional | codificación, por defecto "utf-8" — euc-kr o número de página de códigos |

## Devuelve

Devuelve el contenido CSV/TSV de una URL o de un archivo local como una matriz derramada de dos dimensiones; el texto numérico y de fecha se convierte automáticamente en valores. Si falta la ruta, falla la descarga o la lectura, o el separador indicado no es válido, devuelve #VALUE!; si el contenido está vacío devuelve #N/A.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=IMPORTDATA("https://example.com/data.csv")` |  | Depende de los datos externos |

## Notas

- Si se omite delimiter, se detecta automáticamente: tabulación si la primera línea contiene una tabulación; de lo contrario, coma.
- Las llamadas de versiones anteriores que pasan un nombre de juego de caracteres como segundo argumento (IMPORTDATA(url,"euc-kr")) también se admiten por compatibilidad.
- En lugar del comportamiento de Google, que conserva solo las líneas no vacías, se mantienen las líneas vacías interiores y solo se eliminan las líneas vacías del final.
- Una matriz en url devuelve #VALUE!; para varios elementos, copie la fórmula fila por fila.
- Compatibilidad: Excel 2010+. Se registra siempre como `IMPORTDATA` en todas las versiones de Excel.

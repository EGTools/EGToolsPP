# IMPORTRANGE

**Categoría**: función de compatibilidad con Google Sheets

Importa datos de una hoja de cálculo de Google Sheets compartida.

## Sintaxis

```
=IMPORTRANGE(url_hoja, [direccion_rango])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| url_hoja | Obligatorio | URL de Google Sheets (debe estar compartida) |
| direccion_rango | Opcional | rango a importar, p. ej. "A1:C10" (por defecto: toda la hoja) |

## Devuelve

Devuelve los valores de una hoja de cálculo compartida de Google Sheets como una matriz derramada de dos dimensiones; el texto numérico y de fecha se convierte automáticamente. Si la URL está vacía o no tiene el formato /d/<ID> devuelve #VALUE!; si falla la descarga, el documento no está compartido (respuesta HTML), la hoja está vacía o el rango queda fuera de los datos, devuelve #N/A.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1AbCdEf.../edit#gid=0","A1:C10")` |  | Depende del contenido de la hoja |

## Notas

- Solo se pueden importar hojas de Google Sheets con el uso compartido mediante enlace activado — los documentos no compartidos devuelven una página de inicio de sesión (HTML), lo que produce #N/A.
- La hoja se selecciona mediante el gid de la URL (gid=0 de forma predeterminada). La parte del nombre de hoja de range_address se ignora.
- Se usa el punto de conexión de exportación TSV, por lo que solo se importan los valores, sin conservar el formato.
- Compatibilidad: Excel 2010+. Se registra siempre como `IMPORTRANGE` en todas las versiones de Excel.

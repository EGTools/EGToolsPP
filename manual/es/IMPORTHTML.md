# IMPORTHTML

**Categoría**: función de compatibilidad con Google Sheets

Extrae una tabla o lista de una página web.

## Sintaxis

```
=IMPORTHTML(url, [buscar], [indice], [configuracion_regional])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| url | Obligatorio | URL de la página web |
| buscar | Opcional | "table" (por defecto) o "list" (ul/ol) |
| indice | Opcional | qué aparición extraer, desde 1 (por defecto 1) |
| configuracion_regional | Opcional | regional para números, p. ej. "de-DE" |

## Devuelve

Devuelve una tabla (matriz de dos dimensiones) o una lista (matriz de N×1) de una página web como matriz derramada; el texto numérico se convierte automáticamente. Si falta la URL, search no es "table"/"list" o index es menor que 1, devuelve #VALUE!; si falla la descarga o no existe una tabla o lista con ese index, devuelve #N/A.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=IMPORTHTML("https://en.wikipedia.org/wiki/Microsoft_Excel","table",1)` |  | Depende del contenido de la página |

## Notas

- La página se analiza con MSHTML tras eliminar los bloques <script>, por lo que no se puede importar el contenido dinámico generado por scripts.
- Los rowspan/colspan de las tablas se normalizan en una cuadrícula; el valor se coloca solo en la celda superior izquierda.
- Usa COM, por lo que se excluye del recálculo multiproceso y requiere conexión a Internet.
- Compatibilidad: Excel 2010+. Se registra siempre como `IMPORTHTML` en todas las versiones de Excel.

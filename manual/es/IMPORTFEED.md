# IMPORTFEED

**Categoría**: función de compatibilidad con Google Sheets

Importa un feed RSS o ATOM.

## Sintaxis

```
=IMPORTFEED(url, [consulta], [encabezados], [num_elementos])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| url | Obligatorio | URL del feed |
| consulta | Opcional | "items" (por defecto), "feed" o un campo: title/summary/description/author/url/created |
| encabezados | Opcional | VERDADERO para incluir fila de encabezado |
| num_elementos | Opcional | número de elementos a devolver (por defecto: todos) |

## Devuelve

Devuelve los elementos del feed (o la información del feed o un solo campo) como una matriz derramada de dos dimensiones. Si falta la URL o el campo de query no es compatible devuelve #VALUE!; si falla la descarga, falla el análisis del XML o no hay elementos, devuelve #N/A.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=IMPORTFEED("https://blog.example.com/rss","items",TRUE,5)` |  | Depende del contenido del feed |

## Notas

- RSS 2.0 y Atom se detectan automáticamente. El valor de query "items" (predeterminado) devuelve 5 columnas: Title/Summary/URL/Author/Created.
- Las fechas RFC822/ISO8601 del campo created se convierten en valores de serie de Excel (los desfases de zona horaria se ignoran).
- Usa COM (MSXML), por lo que se excluye del recálculo multiproceso y requiere conexión a Internet.
- Compatibilidad: Excel 2010+. Se registra siempre como `IMPORTFEED` en todas las versiones de Excel.

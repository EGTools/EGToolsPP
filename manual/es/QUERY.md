# QUERY

**Categoría**: función de compatibilidad con Google Sheets

Ejecuta una consulta SQL sobre los datos (nombre de tabla: EGTOOLS).

## Sintaxis

```
=QUERY(datos, consulta, [encabezados])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| datos | Obligatorio | rango o matriz a consultar |
| consulta | Obligatorio | SQL, p. ej. "SELECT F1, SUM(F2) FROM EGTOOLS GROUP BY F1" |
| encabezados | Opcional | número de filas de encabezado: N>=1 las N filas superiores son nombres, 0 ninguno (F1, F2, …), -1/omitido autodetectar |

## Devuelve

Devuelve el resultado de la consulta SQL como una matriz derramada de dos dimensiones (solo filas de valores, sin fila de encabezado). Si faltan data/query o la consulta está vacía devuelve #VALUE!; si no hay filas de resultado devuelve #N/A; los errores del proveedor ADO/ACE y los fallos al escribir el CSV temporal se devuelven como texto "ERROR: …".

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=QUERY(A1:C10,"SELECT F1, SUM(F3) FROM EGTOOLS GROUP BY F1",0)` |  | Depende de los datos |

## Notas

- Usa SQL basado en el controlador de texto ACE + ADO, no Google Query Language; el nombre de la tabla es EGTOOLS (las apariciones de EGTOOLS dentro de la consulta se sustituyen automáticamente por EGTOOLS#csv).
- Los nombres de columna proceden del texto de las filas de encabezado indicadas con headers; si no las hay, son F1, F2, … (headers: N≥1 = las N filas superiores son nombres de campo, 0 = ninguna, omitido/-1 = detección automática).
- Los datos se escriben en %TEMP%\EGTOOLS.csv y se consultan allí, por lo que se requiere el proveedor Microsoft ACE OLEDB (12.0/16.0); debido al uso de un archivo temporal compartido, se excluye del recálculo multiproceso.
- Compatibilidad: Excel 2010+. Se registra siempre como `QUERY` en todas las versiones de Excel.

# SEARCHADDRESS

**Categoría**: función exclusiva de EGTools

Busca direcciones coreanas (juso.go.kr).

## Sintaxis

```
=SEARCHADDRESS(texto, [info], [fila], [api_key])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto | Obligatorio | la palabra clave de dirección a buscar |
| info | Opcional | columna(s) de salida 1-27, p. ej. {1,2}: 1 código postal, 2 dirección vial (predeterminado), 3 jibun, 4 inglés |
| fila | Opcional | fila de resultado; 0/omitido = todas |
| api_key | Opcional | clave de juso.go.kr; se guarda en el registro tras la primera entrada |

## Devuelve

Devuelve las direcciones encontradas como una matriz de texto compuesta por las columnas de información seleccionadas, y se derrama (con fila=0 devuelve todas, hasta 100 resultados). Devuelve #VALUE! si el término de búsqueda está vacío o si una columna de información está fuera de 1~27, y #N/A si no hay resultados; la falta de clave y los errores de red o del servicio se devuelven como texto "ERROR: …".

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=SEARCHADDRESS("세종대로 110")` |  | Búsqueda de dirección de calle (requiere red y clave API) |
| `=SEARCHADDRESS("세종대로 110",{1,2},1)` |  | Código postal + dirección de calle de la fila 1 (requiere red y clave API) |

## Notas

- La clave API no se distribuye con el complemento. Obténgala en el [servicio de direcciones de Corea](https://www.juso.go.kr) e introdúzcala una vez como último argumento; se guarda en el Registro de Windows y puede omitirse después.
- Si el servicio rechaza la clave (caducada/no registrada), la clave guardada se elimina automáticamente y se devuelve una guía para obtener una nueva.
- Columnas de información: 1 código postal, 2 dirección de calle (predeterminado), 3 dirección de parcela (jibun), 4 dirección en inglés, etc. — 27 campos en total.
- Una matriz en el texto de búsqueda o api_key devuelve #VALUE! (una matriz en api_key se rechaza antes de guardarse, lo que evita almacenar una clave incorrecta); para varios elementos, copie la fórmula fila por fila.
- Compatibilidad: Excel 2010+. Se registra siempre como `SEARCHADDRESS` en todas las versiones de Excel.

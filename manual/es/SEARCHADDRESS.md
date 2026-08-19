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
- La lista completa de columnas de información (1~27) se muestra en la tabla siguiente. El valor predeterminado es 2 (dirección de calle); puede indicar varias columnas con una matriz como {1,2}.

| N.º | Campo | Descripción |
|---|---|---|
| 1 | zipNo | Código postal |
| 2 | roadAddr | Dirección de calle completa (predeterminado) |
| 3 | jibunAddr | Dirección jibun (por parcela) |
| 4 | engAddr | Dirección de calle en inglés |
| 5 | roadAddrPart1 | Dirección de calle sin elementos de referencia |
| 6 | roadAddrPart2 | Elementos de referencia de la dirección de calle |
| 7 | admCd | Código de división administrativa |
| 8 | rnMgtSn | Código del nombre de la calle |
| 9 | bdMgtSn | Número de gestión del edificio |
| 10 | detBdNmList | Lista de nombres detallados del edificio |
| 11 | bdNm | Nombre del edificio |
| 12 | bdKdcd | Vivienda colectiva (1: colectiva, 0: no colectiva) |
| 13 | siNm | Nombre de provincia/ciudad metropolitana (si/do) |
| 14 | sggNm | Nombre de ciudad/condado/distrito (si/gun/gu) |
| 15 | emdNm | Nombre de eup/myeon/dong |
| 16 | liNm | Nombre de ri legal (aldea) |
| 17 | rn | Nombre de la calle |
| 18 | udrtYn | Subterráneo (0: superficie, 1: subterráneo) |
| 19 | buldMnnm | Número principal del edificio |
| 20 | buldSlno | Número secundario del edificio |
| 21 | mtYn | Parcela de montaña (0: solar, 1: montaña) |
| 22 | lnbrMnnm | Número principal de parcela (beonji) |
| 23 | lnbrSlno | Número secundario de parcela (ho) |
| 24 | emdNo | Número de serie de eup/myeon/dong |
| 25 | hstryYn | Historial de cambios |
| 26 | relJibun | Parcelas relacionadas |
| 27 | hemdNm | Centro comunitario competente |

- Una matriz en el texto de búsqueda o api_key devuelve #VALUE! (una matriz en api_key se rechaza antes de guardarse, lo que evita almacenar una clave incorrecta); para varios elementos, copie la fórmula fila por fila.
- Compatibilidad: Excel 2010+. Se registra siempre como `SEARCHADDRESS` en todas las versiones de Excel.

# GEOSEARCH

**Categoría**: función exclusiva de EGTools

Busca una dirección o lugar y devuelve coordenadas GPS (VWorld).

## Sintaxis

```
=GEOSEARCH(palabra, [tipo], [epsg], [api_key])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| palabra | Obligatorio | dirección o lugar |
| tipo | Opcional | 1 dirección (predeterminado), 2 lugar |
| epsg | Opcional | código EPSG (predeterminado 4326 = WGS84) |
| api_key | Opcional | clave de vworld.kr; se guarda en el registro tras la primera entrada |

## Devuelve

Devuelve una matriz con las coordenadas x,y de los resultados de la búsqueda y se derrama (si hay varios resultados, es de n×3 e incluye una columna de dirección). Devuelve #VALUE! si el término de búsqueda está vacío o si tipo_busqueda/EPSG no son válidos, y #N/A si no hay resultados; la falta de clave y los errores de red o del servicio se devuelven como texto "ERROR: …".

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=GEOSEARCH("서울특별시청",2)` |  | Coordenadas de la búsqueda de lugares (requiere red y clave API) |

## Notas

- La clave API no se distribuye con el complemento. Obténgala en la plataforma abierta de información espacial VWorld (https://www.vworld.kr) e introdúzcala una vez como último argumento; se guarda en el Registro de Windows y, si la clave es rechazada, se elimina automáticamente y se devuelve una guía para obtener una nueva.
- La búsqueda de direcciones (tipo_busqueda=1) reintenta automáticamente en el orden dirección de calle → dirección de parcela (jibun).
- Una matriz en el texto de búsqueda o api_key devuelve #VALUE! (una matriz en api_key se rechaza antes de guardarse, lo que evita almacenar una clave incorrecta); para varios elementos, copie la fórmula fila por fila.
- Compatibilidad: Excel 2010+. Se registra siempre como `GEOSEARCH` en todas las versiones de Excel.

# GEOADDRESS

**Categoría**: función exclusiva de EGTools

Devuelve la dirección en una coordenada (geocodificación inversa VWorld).

## Sintaxis

```
=GEOADDRESS(x, y, [tipo_retorno], [epsg], [api_key])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| x | Obligatorio | longitud o x |
| y | Obligatorio | latitud o y |
| tipo_retorno | Opcional | 1 jibun (predeterminado), 2 vial, 3 ambas |
| epsg | Opcional | código EPSG (predeterminado 4326 = WGS84) |
| api_key | Opcional | clave de vworld.kr; se guarda en el registro tras la primera entrada |

## Devuelve

Devuelve el texto de la dirección correspondiente a las coordenadas (con tipo_direccion=3 se derrama en 2 filas: dirección de parcela y dirección de calle). Devuelve #VALUE! si tipo_direccion está fuera de 1~3, si el EPSG no es válido o si las coordenadas están fuera de rango en EPSG 4326, y #N/A si no hay dirección; la falta de clave y los errores de red o del servicio se devuelven como texto "ERROR: …".

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=GEOADDRESS(126.9779,37.5663)` |  | Dirección de parcela (jibun) de las coordenadas (requiere red y clave API) |

## Notas

- La clave API no se distribuye con el complemento. Obténgala en la [plataforma abierta de información espacial VWorld](https://www.vworld.kr) e introdúzcala una vez como último argumento; se guarda en el Registro de Windows y, si la clave es rechazada, se elimina automáticamente y se devuelve una guía para obtener una nueva.
- Una matriz en las coordenadas x·y o api_key devuelve #VALUE! (una matriz en api_key se rechaza antes de guardarse, lo que evita almacenar una clave incorrecta); para varios elementos, copie la fórmula fila por fila.
- Compatibilidad: Excel 2010+. Se registra siempre como `GEOADDRESS` en todas las versiones de Excel.

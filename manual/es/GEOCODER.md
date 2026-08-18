# GEOCODER

**Categoría**: función exclusiva de EGTools

Devuelve la coordenada GPS de una dirección (VWorld).

## Sintaxis

```
=GEOCODER(palabra, [epsg], [api_key])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| palabra | Obligatorio | la dirección |
| epsg | Opcional | código EPSG (predeterminado 4326 = WGS84) |
| api_key | Opcional | clave de vworld.kr; se guarda en el registro tras la primera entrada |

## Devuelve

Devuelve una matriz de 1×3 con {x, y, dirección depurada} y se derrama. Devuelve #VALUE! si la dirección está vacía o si el EPSG no es válido, y #N/A si no se encuentran las coordenadas; la falta de clave y los errores de red o del servicio se devuelven como texto "ERROR: …".

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=GEOCODER("세종대로 110")` |  | Coordenadas GPS de la dirección (requiere red y clave API) |

## Notas

- La clave API no se distribuye con el complemento. Obténgala en la plataforma abierta de información espacial VWorld (https://www.vworld.kr ) e introdúzcala una vez como último argumento; se guarda en el Registro de Windows y, si la clave es rechazada, se elimina automáticamente y se devuelve una guía para obtener una nueva.
- Primero intenta el geocodificador de direcciones de calle y, si falla, reintenta con la dirección de parcela (jibun).
- Una matriz en la dirección o api_key devuelve #VALUE! (una matriz en api_key se rechaza antes de guardarse, lo que evita almacenar una clave incorrecta); para varios elementos, copie la fórmula fila por fila.
- Compatibilidad: Excel 2010+. Se registra siempre como `GEOCODER` en todas las versiones de Excel.

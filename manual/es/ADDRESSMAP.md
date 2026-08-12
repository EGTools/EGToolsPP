# ADDRESSMAP

**Categoría**: función exclusiva de EGTools

Inserta una imagen del mapa catastral de una dirección en la celda que llama (VWorld WMS).

## Sintaxis

```
=ADDRESSMAP(direccion, [escala], [api_key])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| direccion | Obligatorio | dirección vial o jibun |
| escala | Opcional | nivel de zoom 1-10 (predeterminado 5) |
| api_key | Opcional | clave de vworld.kr; se guarda en el registro tras la primera entrada |

## Devuelve

Devuelve una cadena vacía ("") si tiene éxito, y la imagen del mapa catastral (VWorld WMS) de la zona de la dirección se inserta por separado ajustada al tamaño de la celda que llama. Devuelve #VALUE! si la dirección está vacía o si el nivel de zoom está fuera de 1~10, #N/A si no se encuentra la dirección y #REF! si no se puede determinar la celda que llama; la falta de clave y los errores de red o del servicio se devuelven como texto "ERROR: …".

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=ADDRESSMAP("세종대로 110",7)` |  | Insertar imagen del mapa catastral (inserción de imagen, requiere red y clave API) |

## Notas

- La clave API no se distribuye con el complemento. Obténgala en la plataforma abierta de información espacial VWorld (https://www.vworld.kr) e introdúzcala una vez como último argumento; se guarda en el Registro de Windows y, si la clave es rechazada, se elimina automáticamente y se devuelve una guía para obtener una nueva.
- Convierte la dirección en coordenadas (calle → parcela) y obtiene un mapa que incluye la capa del catastro continuo.
- Una matriz en la dirección o api_key devuelve #VALUE! (una matriz en api_key se rechaza antes de guardarse, lo que evita almacenar una clave incorrecta); para varios elementos, copie la fórmula fila por fila.
- Compatibilidad: Excel 2010+. Se registra siempre como `ADDRESSMAP` en todas las versiones de Excel.

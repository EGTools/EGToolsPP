# BRNSTATUS

**Categoría**: función exclusiva de EGTools

Consulta el estado de un número de registro mercantil coreano (NTS, data.go.kr).

## Sintaxis

```
=BRNSTATUS(numeros, [api_key])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| numeros | Obligatorio | número(s) de registro, guiones permitidos |
| api_key | Opcional | clave de data.go.kr; se guarda en el registro tras la primera entrada |

## Devuelve

Devuelve el texto del tipo de tributación (tax_type) de cada número de registro de empresa (Corea) con la misma forma que la entrada (si la entrada es una matriz, se derrama). Devuelve #VALUE! si la entrada está vacía; la falta de clave y los errores de red o del servicio se devuelven como texto "ERROR: …", y los números sin resultado se muestran como el texto "#N/A".

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=BRNSTATUS("123-45-67890")` |  | Consulta del estado de registro de empresa (requiere red y clave API) |

## Notas

- La clave API no se distribuye con el complemento. Obténgala en el Portal de Datos Públicos de Corea, servicio de consulta del estado del registro de empresas de la Agencia Tributaria Nacional (https://www.data.go.kr), e introdúzcala una vez como último argumento, o regístrela en la cinta [EGTools] → [Administrar claves API]; se guarda en la cuenta del usuario y puede omitirse después.
- La clave de autenticación de data.go.kr es única por cuenta, por lo que se comparte con KOREANHOLIDAYS y con los calendarios y agendas de la cinta. Una clave registrada por separado para el registro de empresas en una versión anterior se traslada automáticamente a la ranura compartida la primera vez que se usa (el usuario no necesita registrarla de nuevo).
- Si el servicio rechaza la clave (HTTP 401/403, etc.), la clave guardada se elimina automáticamente y se devuelve una guía para obtener una nueva.
- Los números de registro de empresa almacenados como **celdas numéricas** también se consultan tal cual: los números se convierten en cadenas de enteros sin notación exponencial, y los valores de 9 dígitos que perdieron su cero inicial se restauran a 10 dígitos. Los caracteres no numéricos, como guiones y espacios, se ignoran, y las celdas vacías devuelven cadenas vacías.
- Las solicitudes se agrupan en lotes de 100 por POST, por lo que es adecuada para consultas masivas.
- La consulta por lotes de una matriz de números de registro (100 por solicitud) se mantiene como antes. Una matriz en api_key devuelve #VALUE! antes de guardar nada (evita almacenar una clave incorrecta).
- Compatibilidad: Excel 2010+. Se registra siempre como `BRNSTATUS` en todas las versiones de Excel.

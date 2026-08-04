# RESTAPI

**Categoría**: función exclusiva de EGTools

Llama a una API REST y devuelve el texto de respuesta.

## Sintaxis

```
=RESTAPI(url_servicio, [metodo], [encabezados], [cuerpo])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| url_servicio | Obligatorio | URL de la API (con parámetros GET) |
| metodo | Opcional | "GET" (por defecto), "POST", "PUT", "PATCH", "DELETE"… |
| encabezados | Opcional | rango de dos columnas nombre/valor |
| cuerpo | Opcional | texto del cuerpo |

## Devuelve

Devuelve el cuerpo de la respuesta como un único valor escalar de texto (interpretado como UTF-8). Si la dirección falta o está vacía, o la solicitud HTTP falla, devuelve #VALUE!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=RESTAPI("https://api.example.com/users","GET")` |  | El resultado depende de la respuesta del servidor |

## Notas

- Usa WinHTTP y requiere conexión de red. El tiempo límite de recepción de la respuesta es de 30 segundos.
- El cuerpo de la respuesta se devuelve tal cual, independientemente del código de estado HTTP.
- Está registrada como no macro y segura para subprocesos, por lo que puede recalcularse en paralelo.
- Compatibilidad: Excel 2010+. Se registra siempre como `RESTAPI` en todas las versiones de Excel.

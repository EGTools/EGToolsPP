# WEBSERVICE

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Devuelve la respuesta de un servicio web en la URL (HTTP GET).

## Sintaxis

```
=WEBSERVICE(url)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| url | Obligatorio | la URL a solicitar |

## Devuelve

Devuelve el cuerpo de la respuesta de una solicitud HTTP(S) GET como texto escalar interpretado como UTF-8. Si la URL tiene un formato incorrecto o falla la conexión o la solicitud, devuelve #VALUE! (tiempos de espera: 5 s de conexión, 15 s de recepción de la respuesta).

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=WEBSERVICE("https://api.frankfurter.app/latest")` |  | Depende de la respuesta externa |

## Notas

- Funciona mediante una solicitud GET de WinHTTP y requiere conexión a Internet.
- El cuerpo de la respuesta siempre se interpreta como UTF-8.
- Está registrada como segura para subprocesos, por lo que se ejecuta en paralelo durante el recálculo multiproceso.
- Una matriz en url devuelve #VALUE!; para varios elementos, copie la fórmula fila por fila.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `WEBSERVICE` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.WEBSERVICE`.

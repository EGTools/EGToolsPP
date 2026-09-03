# GTRS

**Categoría**: función de compatibilidad con Google Sheets

Alias corto de GOOGLETRANSLATE — traduce texto con Google.

## Sintaxis

```
=GTRS(texto, [idioma_origen], [idioma_destino])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto | Obligatorio | texto (o rango) a traducir |
| idioma_origen | Opcional | código del idioma de origen, p. ej. "en" (por defecto "auto") |
| idioma_destino | Opcional | código del idioma del resultado (por defecto: idioma de la interfaz) |

## Devuelve

Al igual que GOOGLETRANSLATE, devuelve el texto traducido como un valor escalar. Si falla la solicitud HTTP devuelve #VALUE!; los errores del servicio, los códigos de idioma no válidos y los textos de más de 5000 caracteres devuelven un texto "ERROR:…"; un texto vacío devuelve una cadena vacía.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=GTRS("안녕하세요","ko","en")` |  | Depende del resultado de la traducción |

## Notas

- Alias corto de GOOGLETRANSLATE que comparte la misma implementación. Consulte la documentación de GOOGLETRANSLATE para conocer el comportamiento detallado.
- Compatibilidad: Excel 2010+. Se registra siempre como `GTRS` en todas las versiones de Excel.

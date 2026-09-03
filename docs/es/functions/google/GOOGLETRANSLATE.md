# GOOGLETRANSLATE

**Categoría**: función de compatibilidad con Google Sheets

Traduce texto con el servicio de Google Traductor.

## Sintaxis

```
=GOOGLETRANSLATE(texto, [idioma_origen], [idioma_destino])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto | Obligatorio | texto (o rango) a traducir |
| idioma_origen | Opcional | código del idioma de origen, p. ej. "en" (por defecto "auto") |
| idioma_destino | Opcional | código del idioma del resultado (por defecto: idioma de la interfaz) |

## Devuelve

Devuelve el texto traducido como escalar. Si la propia solicitud HTTP falla, devuelve #VALUE!; los errores del servicio, los cambios de formato de la página, los códigos de idioma no válidos y los textos de más de 5000 caracteres se devuelven como texto "ERROR:…"; y una entrada de texto vacía devuelve una cadena vacía.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=GOOGLETRANSLATE("Hello","en","ko")` |  | Depende del resultado de la traducción |

## Notas

- Extrae los datos de la página no oficial translate.google.com/m — puede dejar de funcionar si Google cambia el formato; requiere conexión a Internet.
- El valor predeterminado de source_language es auto y el de target_language es el idioma de la interfaz del complemento.
- Las entradas de rango/matriz se unen con saltos de línea y se traducen de una sola vez; si se superan los 5000 caracteres, devuelve "ERROR:Exceed max length".
- Compatibilidad: Excel 2010+. Se registra siempre como `GOOGLETRANSLATE` en todas las versiones de Excel.

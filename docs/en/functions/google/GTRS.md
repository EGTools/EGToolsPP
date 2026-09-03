# GTRS

**Category**: Google Sheets compatibility function

Short alias of GOOGLETRANSLATE — translates text via Google.

## Syntax

```
=GTRS(text, [source_language], [target_language])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | text (or range) to translate |
| source_language | Optional | source language code, e.g. "en" (default "auto") |
| target_language | Optional | target language code (default: UI language) |

## Returns

Identical to GOOGLETRANSLATE: returns the translated text as a scalar. An HTTP request failure gives #VALUE!; service errors, invalid language codes and texts over 5000 characters return "ERROR:…" text values; empty input returns an empty string.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=GTRS("안녕하세요","ko","en")` |  | result depends on the translation service |

## Notes

- Short alias of GOOGLETRANSLATE sharing the same implementation; see GOOGLETRANSLATE for details.
- Supported: Excel 2010+. Always registered as `GTRS` on every Excel version.

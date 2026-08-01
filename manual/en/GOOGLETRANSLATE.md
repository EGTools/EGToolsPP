# GOOGLETRANSLATE

**Category**: Google Sheets compatibility function

Translates text using the Google Translate service.

## Syntax

```
=GOOGLETRANSLATE(text, [source_language], [target_language])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | text (or range) to translate |
| source_language | Optional | source language code, e.g. "en" (default "auto") |
| target_language | Optional | target language code (default: UI language) |

## Returns

Returns the translated text as a scalar. If the HTTP request itself fails the result is #VALUE!; service errors, page-format changes, invalid language codes and texts over 5000 characters come back as "ERROR:…" text values; empty input returns an empty string.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=GOOGLETRANSLATE("Hello","en","ko")` |  | result depends on the translation service |

## Notes

- Scrapes the unofficial translate.google.com/m page — it may stop working if Google changes the format; requires an internet connection.
- source_language defaults to auto; target_language defaults to the add-in UI language.
- Range/array input is joined with newlines and translated as one text; over 5000 characters returns "ERROR:Exceed max length".
- Supported: Excel 2010+. Always registered as `GOOGLETRANSLATE` on every Excel version.

# FILTERXML

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Returns data from XML content using the specified XPath.

## Syntax

```
=FILTERXML(xml, xpath)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| xml | Required | XML text |
| xpath | Required | XPath query |

## Returns

Returns the matched node text as a scalar when the XPath matches one node, or a spilled N-by-1 vertical array of text when it matches several. Empty XML/XPath or an XML parse / XPath failure gives #VALUE!; no matching node gives #N/A.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=FILTERXML("<r><a>1</a><a>2</a></r>","//a")` | {"1";"2"} | all <a> node texts |
| `=FILTERXML(WEBSERVICE("https://example.com/rss.xml"),"//item/title")` |  | result depends on the fetched XML |

## Notes

- Parsed with MSXML6; the function itself does no networking — combine with WEBSERVICE for remote XML.
- Results are node text values and are always returned as text.
- Uses COM, so it is excluded from multi-threaded recalculation (not thread-safe).
- Supported: Excel 2010+. Registered as `FILTERXML` (drop-in) on hosts without the native function, and as `EG.FILTERXML` on modern Excel that has it.

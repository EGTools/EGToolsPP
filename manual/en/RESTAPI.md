# RESTAPI

**Category**: EGTools-only function

Calls a REST API and returns the response text.

## Syntax

```
=RESTAPI(service_url, [method], [request_headers], [request_body])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| service_url | Required | API URL (GET parameters included) |
| method | Optional | "GET" (default), "POST", "PUT", "PATCH", "DELETE"… |
| request_headers | Optional | two-column range of header name/value pairs |
| request_body | Optional | request body text |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=RESTAPI(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `RESTAPI` on every Excel version.

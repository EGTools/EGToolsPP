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

Returns the response body as a single text scalar (decoded as UTF-8). Returns #VALUE! when the URL is missing or empty, or the HTTP request fails.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=RESTAPI("https://api.example.com/users","GET")` |  | Result depends on the server response |

## Notes

- Uses WinHTTP and requires a network connection; the receive timeout is 30 seconds.
- The response body is returned as-is, regardless of the HTTP status code.
- Registered non-macro and thread-safe, so it can recalculate in parallel.
- Supported: Excel 2010+. Always registered as `RESTAPI` on every Excel version.

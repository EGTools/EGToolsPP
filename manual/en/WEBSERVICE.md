# WEBSERVICE

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Returns the response from a web service at the URL (HTTP GET).

## Syntax

```
=WEBSERVICE(url)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| url | Required | the URL to request |

## Returns

Returns the HTTP(S) GET response body as a scalar text value decoded as UTF-8. A malformed URL or a connection/request failure gives #VALUE! (timeouts: 5 s connect, 15 s receive).

## Examples

| Formula | Result | Description |
|---|---|---|
| `=WEBSERVICE("https://api.frankfurter.app/latest")` |  | result depends on the web service |

## Notes

- Implemented with a WinHTTP GET request; requires an internet connection.
- The response body is always interpreted as UTF-8.
- Registered thread-safe, so it runs in parallel under multi-threaded recalculation.
- Supported: Excel 2010+. Registered as `WEBSERVICE` (drop-in) on hosts without the native function, and as `EG.WEBSERVICE` on modern Excel that has it.

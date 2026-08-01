# SEARCHADDRESS

**Category**: EGTools-only function

Searches Korean addresses (juso.go.kr).

## Syntax

```
=SEARCHADDRESS(text, [info], [row], [api_key])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | the address keyword to search |
| info | Optional | output column(s) 1-27, e.g. {1,2}: 1 zip, 2 road addr (default), 3 jibun, 4 english |
| row | Optional | result row to return; 0/omitted = all rows |
| api_key | Optional | juso.go.kr key; entered once, then stored in the registry |

## Returns

Returns the matching addresses as a spilled text array with the selected info columns (row=0 returns all rows, up to 100 results). Returns #VALUE! for an empty keyword or an info column outside 1-27 and #N/A when nothing is found; a missing key, network, or service failure returns an "ERROR: ..." text.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SEARCHADDRESS("세종대로 110")` |  | Road-address search (needs network and an API key) |
| `=SEARCHADDRESS("세종대로 110",{1,2},1)` |  | Zip + road address of row 1 (needs network and an API key) |

## Notes

- No API key ships with the add-in. Get one at business.juso.go.kr and pass it once as the last argument; it is stored under HKCU\Software\EGTools++\ApiKeys and can be omitted afterwards.
- When the service rejects the key (expired/unregistered), the stored key is deleted automatically and a reissue guide is returned.
- info columns: 1 zip, 2 road address (default), 3 jibun, 4 English — 27 fields in total.
- Supported: Excel 2010+. Always registered as `SEARCHADDRESS` on every Excel version.

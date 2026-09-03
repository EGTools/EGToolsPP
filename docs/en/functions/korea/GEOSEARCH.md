# GEOSEARCH

**Category**: EGTools-only function

Searches an address or place and returns GPS coordinates (VWorld).

## Syntax

```
=GEOSEARCH(keyword, [search_type], [epsg], [api_key])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| keyword | Required | address or place keyword |
| search_type | Optional | 1 address (default), 2 place |
| epsg | Optional | coordinate system EPSG code (default 4326 = WGS84) |
| api_key | Optional | vworld.kr key; entered once, then stored in the registry |

## Returns

Returns the x,y coordinates of the results as a spilled array (n-by-3 including an address column when there are multiple hits). Returns #VALUE! for an empty keyword or invalid search_type/EPSG and #N/A when nothing is found; a missing key, network, or service failure returns an "ERROR: ..." text.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=GEOSEARCH("서울특별시청",2)` |  | Place-search coordinates (needs network and an API key) |

## Notes

- No API key ships with the add-in. Get one at https://www.vworld.kr and pass it once as the last argument; stored in the registry, and a rejected key is deleted automatically with a reissue guide.
- Address search (search_type 1) retries road then parcel (jibun) automatically.
- An array in the search text or api_key returns #VALUE! (an array api_key is rejected before saving, preventing a bad key from being stored) — for multiple items, copy the formula down row by row.
- Supported: Excel 2010+. Always registered as `GEOSEARCH` on every Excel version.

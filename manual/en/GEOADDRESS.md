# GEOADDRESS

**Category**: EGTools-only function

Returns the address at a coordinate (VWorld reverse geocoding).

## Syntax

```
=GEOADDRESS(x, y, [return_type], [epsg], [api_key])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| x | Required | longitude or x |
| y | Required | latitude or y |
| return_type | Optional | 1 jibun (default), 2 road, 3 both |
| epsg | Optional | coordinate system EPSG code (default 4326 = WGS84) |
| api_key | Optional | vworld.kr key; entered once, then stored in the registry |

## Returns

Returns the address text at the coordinate (return_type=3 spills jibun and road addresses as two rows). Returns #VALUE! when return_type is outside 1-3, EPSG is invalid, or the coordinate is out of range for EPSG 4326, and #N/A when no address exists; a missing key, network, or service failure returns an "ERROR: ..." text.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=GEOADDRESS(126.9779,37.5663)` |  | Jibun address at the coordinate (needs network and an API key) |

## Notes

- No API key ships with the add-in. Get one at https://www.vworld.kr and pass it once as the last argument; stored in the registry, and a rejected key is deleted automatically with a reissue guide.
- Supported: Excel 2010+. Always registered as `GEOADDRESS` on every Excel version.

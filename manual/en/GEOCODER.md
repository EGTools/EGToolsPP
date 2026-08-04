# GEOCODER

**Category**: EGTools-only function

Returns the GPS coordinate of an address (VWorld; road then jibun).

## Syntax

```
=GEOCODER(keyword, [epsg], [api_key])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| keyword | Required | the address |
| epsg | Optional | coordinate system EPSG code (default 4326 = WGS84) |
| api_key | Optional | vworld.kr key; entered once, then stored in the registry |

## Returns

Returns a spilled 1-by-3 array of {x, y, refined address}. Returns #VALUE! for an empty address or invalid EPSG and #N/A when no coordinate is found; a missing key, network, or service failure returns an "ERROR: ..." text.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=GEOCODER("세종대로 110")` |  | GPS coordinate of an address (needs network and an API key) |

## Notes

- No API key ships with the add-in. Get one at https://www.vworld.kr and pass it once as the last argument; stored in the registry, and a rejected key is deleted automatically with a reissue guide.
- Tries the road-address geocoder first, then falls back to parcel (jibun).
- Supported: Excel 2010+. Always registered as `GEOCODER` on every Excel version.

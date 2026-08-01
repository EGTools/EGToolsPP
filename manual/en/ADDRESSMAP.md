# ADDRESSMAP

**Category**: EGTools-only function

Inserts a cadastral map picture for an address into the calling cell (VWorld WMS).

## Syntax

```
=ADDRESSMAP(address, [scale], [api_key])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| address | Required | road or jibun address |
| scale | Optional | zoom level 1-10 (default 5), larger = closer |
| api_key | Optional | vworld.kr key; entered once, then stored in the registry |

## Returns

Returns an empty string ("") on success; a cadastral map picture (VWorld WMS) around the address is inserted separately, sized to the calling cell. Returns #VALUE! for an empty address or a scale outside 1-10, #N/A when the address cannot be geocoded, and #REF! when the calling cell cannot be determined; a missing key, network, or service failure returns an "ERROR: ..." text.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=ADDRESSMAP("세종대로 110",7)` |  | Insert a cadastral map picture (image insertion, network and API key needed) |

## Notes

- No API key ships with the add-in. Get one at vworld.kr and pass it once as the last argument; stored under HKCU\Software\EGTools++\ApiKeys, and a rejected key is deleted automatically with a reissue guide.
- Geocodes the address (road then parcel) and fetches a WMS map that includes the continuous cadastral layers.
- Supported: Excel 2010+. Always registered as `ADDRESSMAP` on every Excel version.

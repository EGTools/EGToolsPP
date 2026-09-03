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

- No API key ships with the add-in. Get one at https://www.juso.go.kr and pass it once as the last argument; it is stored in the registry and can be omitted afterwards.
- When the service rejects the key (expired/unregistered), the stored key is deleted automatically and a reissue guide is returned.
- The full list of info columns (1~27) is shown below. The default is 2 (road address); pass an array like {1,2} to select multiple columns.

| No. | Field | Description |
|---|---|---|
| 1 | zipNo | Postal code |
| 2 | roadAddr | Full road-name address (default) |
| 3 | jibunAddr | Jibun (lot-number) address |
| 4 | engAddr | Road-name address in English |
| 5 | roadAddrPart1 | Road-name address without reference items |
| 6 | roadAddrPart2 | Road-name address reference items |
| 7 | admCd | Administrative district code |
| 8 | rnMgtSn | Road-name code |
| 9 | bdMgtSn | Building management number |
| 10 | detBdNmList | Detailed building name list |
| 11 | bdNm | Building name |
| 12 | bdKdcd | Multi-unit housing flag (1: multi-unit, 0: not) |
| 13 | siNm | Province/metropolitan city (si/do) name |
| 14 | sggNm | City/county/district (si/gun/gu) name |
| 15 | emdNm | Town/township/neighborhood (eup/myeon/dong) name |
| 16 | liNm | Legal village (ri) name |
| 17 | rn | Road name |
| 18 | udrtYn | Underground flag (0: above ground, 1: underground) |
| 19 | buldMnnm | Building main number |
| 20 | buldSlno | Building sub-number |
| 21 | mtYn | Mountain-lot flag (0: land, 1: mountain) |
| 22 | lnbrMnnm | Lot main number (beonji) |
| 23 | lnbrSlno | Lot sub-number (ho) |
| 24 | emdNo | Eup/myeon/dong serial number |
| 25 | hstryYn | Change-history flag |
| 26 | relJibun | Related jibun (lot numbers) |
| 27 | hemdNm | Responsible community service center |

- An array in the search text or api_key returns #VALUE! (an array api_key is rejected before saving, preventing a bad key from being stored) — for multiple items, copy the formula down row by row.
- Supported: Excel 2010+. Always registered as `SEARCHADDRESS` on every Excel version.

# BRNSTATUS

**Category**: EGTools-only function

Checks Korean business-registration-number status (NTS via data.go.kr).

## Syntax

```
=BRNSTATUS(numbers, [api_key])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| numbers | Required | registration number(s), hyphens allowed |
| api_key | Optional | data.go.kr key; entered once, then stored in the registry |

## Returns

Returns the tax-type text (tax_type) for each business registration number in the same shape as the input (spills for an array). Returns #VALUE! for empty input; a missing key, network, or service failure returns an "ERROR: ..." text, and unresolved numbers show the text "#N/A".

## Examples

| Formula | Result | Description |
|---|---|---|
| `=BRNSTATUS("123-45-67890")` |  | Business-registration status (needs network and an API key) |

## Notes

- No API key ships with the add-in. Get one at https://www.data.go.kr (NTS business-registration status service) and pass it once as the last argument; stored under HKCU\Software\EGTools\ApiKeys and reused afterwards.
- When the service rejects the key (HTTP 401/403 etc.), the stored key is deleted automatically and a reissue guide is returned.
- Requests are batched 100 numbers per POST, so bulk lookups are efficient; hyphens are removed automatically and blank inputs return empty strings.
- Supported: Excel 2010+. Always registered as `BRNSTATUS` on every Excel version.

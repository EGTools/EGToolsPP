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

- No API key ships with the add-in. Get one at https://www.data.go.kr (NTS business-registration status service) and pass it once as the last argument, or register it under ribbon [EGTools] → [Manage API Keys]; it is stored in your user account and reused afterwards.
- The data.go.kr key is a single per-account key, so it is shared with KOREANHOLIDAYS and the ribbon Calendar commands. A key registered separately for business-registration lookups by an older version is moved to the shared slot automatically on first use — nothing to re-enter.
- When the service rejects the key (HTTP 401/403 etc.), the stored key is deleted automatically and a reissue guide is returned.
- Numbers stored as **numeric cells** are handled correctly: the number is converted to an integer string without exponent notation, and a 9-digit value whose leading zero was dropped is restored to 10 digits. Non-digit characters such as hyphens and spaces are ignored, and blank cells return empty strings.
- Requests are batched 100 numbers per POST, so bulk lookups are efficient.
- Supported: Excel 2010+. Always registered as `BRNSTATUS` on every Excel version.

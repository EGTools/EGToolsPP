# COMBINA

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Returns the number of combinations with repetitions for a given number of items.

## Syntax

```
=COMBINA(number, number_chosen)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| number | Required | total number of items |
| number_chosen | Required | number of items in each combination |

## Returns

Returns the number of combinations with repetition (scalar number). Non-numeric or negative arguments return a #NUM! error.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=COMBINA(4,3)` | 20 | Choose 3 of 4 with repetition |
| `=COMBINA(10,3)` | 220 | Choose 3 of 10 with repetition |

## Notes

- Fractional parts are truncated; computed as C(n+k-1, k).
- All arguments accept arrays and are computed element-wise, spilling an array of the same shape — scalars are broadcast, a column vector × a row vector expands to their outer product, size-mismatched elements yield #N/A, and an error element returns that error.
- Related function: PERMUTATIONA
- Supported: Excel 2010+. Registered as `COMBINA` (drop-in) on hosts without the native function, and as `EG.COMBINA` on modern Excel that has it.

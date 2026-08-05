# PERCENTOF

**Category**: MS Excel compatibility function · **Native since**: Microsoft 365

Returns the sum of a subset divided by the sum of all data.

## Syntax

```
=PERCENTOF(data_subset, data_all)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| data_subset | Required | the subset of data |
| data_all | Required | all of the data |

## Returns

Returns a scalar number: the sum of the subset divided by the sum of all data. Missing arguments give #VALUE!; if the total sum is 0 or contains no numbers the result is #DIV/0!.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=PERCENTOF({10;20},{10;20;30;40})` | 0.3 | share of subset in total |
| `=PERCENTOF(50,200)` | 0.25 | scalar ratio |

## Notes

- Non-numeric values are ignored when summing.
- Passing "PERCENTOF" as the function argument of GROUPBY/PIVOTBY computes per-group shares.
- Supported: Excel 2010+. Registered as `PERCENTOF` (drop-in) on hosts without the native function, and as `EG.PERCENTOF` on modern Excel that has it.

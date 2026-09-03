# PERMUTATIONA

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Returns the number of permutations with repetitions (number ^ number_chosen).

## Syntax

```
=PERMUTATIONA(number, number_chosen)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| number | Required | total number of items |
| number_chosen | Required | number of items in each permutation |

## Returns

Returns the number of permutations with repetition, number^number_chosen (scalar number). Non-numeric or negative arguments return a #NUM! error.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=PERMUTATIONA(3,2)` | 9 | 3^2 |
| `=PERMUTATIONA(2,5)` | 32 | 2^5 |

## Notes

- Fractional parts are truncated.
- All arguments accept arrays and are computed element-wise, spilling an array of the same shape — scalars are broadcast, a column vector × a row vector expands to their outer product, size-mismatched elements yield #N/A, and an error element returns that error.
- Related function: COMBINA
- Supported: Excel 2010+. Registered as `PERMUTATIONA` (drop-in) on hosts without the native function, and as `EG.PERMUTATIONA` on modern Excel that has it.

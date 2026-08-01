# EVAL

**Category**: EGTools-only function

Evaluates a formula given as text and returns the result.

## Syntax

```
=EVAL(formula)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| formula | Required | formula text, e.g. "SUM(A1:A3)*2" |

## Returns

Returns the evaluation result as-is — a single value for scalars, a spilled array for array results. Returns #VALUE! when the formula is empty or evaluation fails (including when COM is unavailable).

## Examples

| Formula | Result | Description |
|---|---|---|
| `=EVAL("SUM(A1:A3)*2")` |  | Result depends on sheet values |

## Notes

- Macro-type function; uses COM Application.Evaluate (not thread-safe).
- A leading "=" is added automatically if missing.
- Supported: Excel 2010+. Always registered as `EVAL` on every Excel version.

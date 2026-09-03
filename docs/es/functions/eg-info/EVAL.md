# EVAL

**Categoría**: función exclusiva de EGTools

Evalúa una fórmula dada como texto y devuelve el resultado.

## Sintaxis

```
=EVAL(formula)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| formula | Obligatorio | texto de fórmula, p. ej. "SUM(A1:A3)*2" |

## Devuelve

Devuelve el resultado del cálculo tal cual: un valor único si es escalar y una matriz derramada si es una matriz. Si la fórmula está vacía o el cálculo falla (incluida la falta de conexión COM), devuelve #VALUE!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=EVAL("SUM(A1:A3)*2")` |  | El resultado depende de los valores de la hoja |

## Notas

- Función de tipo macro; usa COM Application.Evaluate (no es segura para subprocesos).
- El signo "=" inicial de la fórmula se agrega automáticamente.
- Compatibilidad: Excel 2010+. Se registra siempre como `EVAL` en todas las versiones de Excel.

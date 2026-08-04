# EXRATE

**Categoría**: función exclusiva de EGTools

Consulta el tipo de cambio estándar en KRW para una fecha (hasta 10 días atrás).

## Sintaxis

```
=EXRATE([moneda], [fecha])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| moneda | Opcional | código de 3 letras, p. ej. USD (por defecto) |
| fecha | Opcional | fecha a consultar (por defecto hoy) |

## Devuelve

Devuelve el tipo de cambio de referencia en wones coreanos (KRW) como un valor escalar numérico. Si se indica una fecha futura devuelve #VALUE!; si el código de moneda no es compatible o no hay tipo de cambio publicado en el rango retroactivo de 10 días, devuelve #N/A.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=EXRATE("USD")` |  | El resultado depende de la fecha de consulta y de la red |

## Notas

- Consulta por HTTP el tipo de cambio de referencia de smbs.biz (Seoul Money Brokerage Services) — se requiere conexión de red.
- Si no hay tipo publicado para la fecha indicada, la consulta retrocede hasta 10 días (para cubrir los festivos).
- CNY se convierte automáticamente en CNH. Está registrada como no macro y segura para subprocesos, por lo que se recalcula en paralelo.
- Compatibilidad: Excel 2010+. Se registra siempre como `EXRATE` en todas las versiones de Excel.

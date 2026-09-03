# TOTALPAGES

**Categoría**: función exclusiva de EGTools

Número total de páginas con la configuración de impresión actual (recalcule antes de imprimir).

## Sintaxis

```
=TOTALPAGES([ignorar_inicio])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| ignorar_inicio | Opcional | VERDADERO numera desde 1 ignorando el número inicial |

## Devuelve

Devuelve el número total de páginas como un escalar numérico, reflejando el número de página inicial de la configuración de impresión (si ignorar_inicio es TRUE, cuenta desde 1). Devuelve #VALUE! si falla la consulta del número de páginas.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=TOTALPAGES()` |  | El resultado depende de la configuración de impresión |

## Notas

- Función de tipo macro que usa XLM GET.DOCUMENT (no es segura para subprocesos).
- Después de cambiar la configuración de impresión se debe recalcular para que el valor se actualice.
- Compatibilidad: Excel 2010+. Se registra siempre como `TOTALPAGES` en todas las versiones de Excel.

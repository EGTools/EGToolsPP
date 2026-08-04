# SHEETS

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Devuelve el número de hojas del libro (o de una referencia).

## Sintaxis

```
=SHEETS([referencia])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| referencia | Opcional | referencia (opcional) |

## Devuelve

Devuelve como número la cantidad de hojas del libro si se omite la referencia, y siempre 1 si se especifica una referencia. Devuelve #VALUE! si falla la consulta del número de hojas.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=SHEETS()` |  | Número de hojas del libro (depende del libro) |

## Notas

- Si se especifica una referencia, siempre devuelve 1 (no se admite contar las hojas de una referencia 3D).
- Se registra como función de tipo hoja de macros.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `SHEETS` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.SHEETS`.

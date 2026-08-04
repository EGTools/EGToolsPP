# ISFORMULA

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Devuelve VERDADERO si la celda referenciada contiene una fórmula.

## Sintaxis

```
=ISFORMULA(referencia)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| referencia | Obligatorio | celda a evaluar |

## Devuelve

Devuelve TRUE si la celda referenciada contiene una fórmula y FALSE en caso contrario. Devuelve #VALUE! si se omite la referencia o falla la consulta.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=ISFORMULA(A1)` |  | Depende del contenido de A1 |

## Notas

- Se registra como función de tipo hoja de macros.
- Función relacionada: FORMULATEXT.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `ISFORMULA` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.ISFORMULA`.

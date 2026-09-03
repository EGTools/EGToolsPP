# FORMULATEXT

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Devuelve la fórmula de la celda referenciada como texto.

## Sintaxis

```
=FORMULATEXT(referencia)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| referencia | Obligatorio | celda cuya fórmula se devuelve |

## Devuelve

Devuelve la fórmula de la celda referenciada como escalar de texto. Devuelve #N/A si se omite la referencia, la celda no contiene una fórmula o falla la consulta.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=FORMULATEXT(A1)` |  | Depende del contenido de A1 |

## Notas

- Se registra como función de tipo hoja de macros.
- Función relacionada: ISFORMULA.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `FORMULATEXT` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.FORMULATEXT`.

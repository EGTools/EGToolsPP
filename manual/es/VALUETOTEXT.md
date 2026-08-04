# VALUETOTEXT

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2024 / Microsoft 365

Devuelve la representación de texto de un valor.

## Sintaxis

```
=VALUETOTEXT(valor, [formato])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| valor | Obligatorio | el valor |
| formato | Opcional | 0 conciso (pred.), 1 estricto (texto entre comillas) |

## Devuelve

Devuelve la representación de texto del valor como un texto único (escalar). La implementación en sí no devuelve valores de error propios.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=VALUETOTEXT(1.5)` | 1.5 | número sin cambios |
| `=VALUETOTEXT("abc",1)` | "abc" | estricto: texto entre comillas |
| `=VALUETOTEXT(TRUE)` | TRUE | valor lógico |

## Notas

- En el formato estricto (1), las comillas dobles dentro del texto se escriben duplicadas ("").
- Función relacionada: ARRAYTOTEXT
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `VALUETOTEXT` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.VALUETOTEXT`.

# ARRAYTOTEXT

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2024 / Microsoft 365

Devuelve una representación de texto de una matriz.

## Sintaxis

```
=ARRAYTOTEXT(matriz, [formato])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| matriz | Obligatorio | la matriz |
| formato | Opcional | 0 conciso (pred.), 1 estricto |

## Devuelve

Devuelve un único valor de texto (escalar) que une toda la matriz. La implementación en sí no devuelve valores de error propios.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=ARRAYTOTEXT({1,2;3,4})` | 1, 2, 3, 4 | Unión en formato conciso |
| `=ARRAYTOTEXT({"a",1;"b",2},1)` | {"a",1;"b",2} | Formato estricto |

## Notas

- El formato conciso (0) une todos los elementos con ", " sin distinguir filas y columnas.
- El formato estricto (1) tiene la forma {a,b;c,d}; solo el texto va entre comillas dobles y las comillas internas se duplican ("").
- Función relacionada: VALUETOTEXT
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `ARRAYTOTEXT` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.ARRAYTOTEXT`.

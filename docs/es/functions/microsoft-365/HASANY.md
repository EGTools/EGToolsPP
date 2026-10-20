# HASANY

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Microsoft 365 (novedad de 2026)

Devuelve VERDADERO si alguno de los valores aparece en la lista o matriz.

## Sintaxis

```
=HASANY(matriz, valores)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| matriz | Obligatorio | texto de lista, rango o matriz donde buscar |
| valores | Obligatorio | valores a buscar (rango, matriz o texto de lista) |

## Devuelve

Devuelve TRUE/FALSE (escalar). Si el propio argumento matriz o valores es un error, devuelve ese error tal cual; si se omite valores, devuelve #VALUE!. Si valores está vacío, el resultado es FALSE.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=HASANY("a, b, c",{"x","c"})` | TRUE | TRUE si hay al menos uno |
| `=HASANY("1, 2",{"1"})` | FALSE | el texto "1" de la constante matricial no es el número 1 |

## Notas

- valores acepta un rango, una constante matricial o texto de lista; todos los elementos, incluidos los anidados, se aplanan en un único conjunto y cada uno se busca con las mismas reglas que HAS (elemento completo sin distinguir mayúsculas, igualdad numérica). Si se indica texto en valores, se interpreta como texto de lista: se divide por comas y las partes con forma numérica pasan a ser números ("z, B" indica varios valores; "1, 2" son los números 1 y 2). Para comparar un solo valor como texto, indíquelo como constante matricial {"1"}.
- matriz se interpreta exactamente igual que en HAS (texto de lista, rango o constante matricial; un texto único sin comas queda como texto; el texto no analizable es un solo elemento). Convierta primero las listas en celdas con la cinta [EGTools] › [Celdas de lista/matriz → texto].
- Funciones relacionadas: HAS, HASALL, FLATTEN
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `HASANY` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.HASANY`.

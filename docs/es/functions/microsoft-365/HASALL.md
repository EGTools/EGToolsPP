# HASALL

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Microsoft 365 (novedad de 2026)

Devuelve VERDADERO si todos los valores aparecen en la lista o matriz.

## Sintaxis

```
=HASALL(matriz, valores)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| matriz | Obligatorio | texto de lista, rango o matriz donde buscar |
| valores | Obligatorio | valores a buscar (rango, matriz o texto de lista) |

## Devuelve

Devuelve TRUE/FALSE (escalar). Si el propio argumento matriz o valores es un error, devuelve ese error tal cual; si se omite valores, devuelve #VALUE!. Si valores está vacío, el resultado es TRUE.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=HASALL("a, b, c",{"a","c"})` | TRUE | TRUE si están todos |
| `=HASALL("a, b",{"a","z"})` | FALSE | FALSE si falta alguno |

## Notas

- valores acepta un rango, una constante matricial o texto de lista; todos los elementos, incluidos los anidados, se aplanan en un único conjunto y cada uno se busca con las mismas reglas que HAS (elemento completo sin distinguir mayúsculas, igualdad numérica). Si se indica texto en valores, se interpreta como texto de lista: se divide por comas y las partes con forma numérica pasan a ser números ("z, B" indica varios valores; "1, 2" son los números 1 y 2). Para comparar un solo valor como texto, indíquelo como constante matricial {"1"}.
- Una celda vacía en el rango de valores cuenta como elemento de texto vacío, por lo que el resultado es FALSE salvo que la lista contenga un elemento vacío. matriz se interpreta exactamente igual que en HAS (texto de lista, rango o constante matricial; un texto único sin comas queda como texto; el texto no analizable es un solo elemento). Convierta primero las listas en celdas con la cinta [EGTools] › [Celdas de lista/matriz → texto].
- Funciones relacionadas: HAS, HASANY, FLATTEN
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `HASALL` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.HASALL`.

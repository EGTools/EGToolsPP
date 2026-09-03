# TEXTNUMSORT

**Categoría**: función exclusiva de EGTools

Ordena filas en orden natural (texto y números incrustados comparados por separado).

## Sintaxis

```
=TEXTNUMSORT(matriz, [col_indice], [orden], [ignorar_texto])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| matriz | Obligatorio | datos a ordenar |
| col_indice | Opcional | columna clave, desde 1 (por defecto 1) |
| orden | Opcional | 1 ascendente (por defecto), -1 descendente |
| ignorar_texto | Opcional | VERDADERO compara solo las partes numéricas |

## Devuelve

Devuelve una matriz ordenada del mismo tamaño que la entrada, que se derrama en las versiones de Excel que admiten matrices dinámicas. Un valor único que no sea una matriz se devuelve sin cambios. Devuelve #VALUE! cuando falta la matriz, la columna clave está fuera de rango o el orden de clasificación no es 1/-1.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=TEXTNUMSORT({"a10";"a2";"a1"})` | {a1;a2;a10} | orden natural ascendente |
| `=TEXTNUMSORT({"a1";"a10";"a2"},1,-1)` | {a10;a2;a1} | orden descendente |

## Notas

- Separa los tramos de texto y de números, y compara los números por su valor numérico (a2 < a10).
- Las comas dentro de los números se ignoran como separadores de miles y el '.' se trata como punto decimal.
- La ordenación es estable (stable sort), por lo que las filas con claves iguales conservan su orden.
- Compatibilidad: Excel 2010+. Se registra siempre como `TEXTNUMSORT` en todas las versiones de Excel.

# COMPARELIST

**Categoría**: función exclusiva de EGTools

Compara listas y devuelve el conjunto elegido: 0 unión, N solo-en-lista-N, -1 en-una-sola, -2 en-todas, -3 en-algunas.

## Sintaxis

```
=COMPARELIST(tipo_conjunto, por_fila, matriz1, ...)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| tipo_conjunto | Obligatorio | 0 unión, N solo en lista N, -1 exactamente una, -2 todas, -3 algunas |
| por_fila | Obligatorio | VERDADERO compara filas completas |
| matriz1 | Obligatorio | rango o matriz de lista |
| … | Opcional | Los argumentos anteriores se pueden repetir |

## Devuelve

Devuelve los elementos que pertenecen al conjunto seleccionado como una matriz derramada (varias columnas si by_row es TRUE; de lo contrario, una columna). Si no se proporciona ninguna lista o unique_type está fuera del rango -3..número de listas devuelve #VALUE!; si el resultado está vacío devuelve #N/A.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=COMPARELIST(0,FALSE,{"a";"b"},{"b";"c"})` | {"a";"b";"c"} | Unión (sin duplicados) |
| `=COMPARELIST(-2,FALSE,{"a";"b"},{"b";"c"})` | b | Elementos presentes en todas las listas |
| `=COMPARELIST(1,FALSE,{"a";"b"},{"b";"c"})` | a | Elementos exclusivos de la lista 1 |

## Notas

- La comparación no distingue mayúsculas y minúsculas; los números se comparan por su valor.
- El resultado conserva el orden de primera aparición de los elementos.
- Con by_row TRUE, cada fila completa se compara como una sola clave (las celdas vacías finales se ignoran).
- Compatibilidad: Excel 2010+. Se registra siempre como `COMPARELIST` en todas las versiones de Excel.

# SHEETLIST

**Categoría**: función exclusiva de EGTools

Devuelve los nombres de las hojas visibles del libro que llama.

## Sintaxis

```
=SHEETLIST([inicio], [fin], [solo_entre], [excluir_actual])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| inicio | Opcional | número/nombre de la primera hoja (por defecto: primera) |
| fin | Opcional | número/nombre de la última hoja (por defecto: última) |
| solo_entre | Opcional | VERDADERO excluye inicio y fin |
| excluir_actual | Opcional | VERDADERO excluye la hoja actual (por defecto) |

## Devuelve

Devuelve los nombres de las hojas visibles como una matriz derramada de una columna. Si no se puede determinar la celda de llamada o el libro, falla la conexión COM o la especificación de inicio/fin no es válida, devuelve #VALUE!; si no queda ninguna hoja tras el filtrado, devuelve #N/A.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=SHEETLIST()` |  | El resultado depende de la composición del libro |

## Notas

- Función de tipo macro; lee la información de las hojas mediante COM (no es segura para subprocesos).
- Las hojas ocultas siempre se excluyen y, de forma predeterminada, también se excluye la hoja desde la que se llama la función (la opción de excluir la hoja actual es TRUE de forma predeterminada).
- El inicio y el fin se pueden indicar mediante el número o el nombre de la hoja.
- Compatibilidad: Excel 2010+. Se registra siempre como `SHEETLIST` en todas las versiones de Excel.

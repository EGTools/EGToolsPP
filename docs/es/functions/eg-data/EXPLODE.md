# EXPLODE

**Categoría**: función exclusiva de EGTools

Divide una columna por delimitador(es) expandiendo en filas o columnas; las demás columnas se copian.

## Sintaxis

```
=EXPLODE(matriz, num_columna, [delimitador], [por_columna], [relleno], [ignorar_vacios])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| matriz | Obligatorio | tabla origen |
| num_columna | Obligatorio | columna a dividir, desde 1 |
| delimitador | Opcional | delimitador(es), por defecto espacio |
| por_columna | Opcional | VERDADERO expande en columnas (por defecto FALSO: en filas) |
| relleno | Opcional | valor de relleno con por_columna (por defecto "") |
| ignorar_vacios | Opcional | VERDADERO descarta fragmentos vacíos |

## Devuelve

Devuelve el resultado de la descomposición como una matriz derramada de dos dimensiones. Si el argumento de matriz no es una matriz devuelve #VALUE!; si el número de columna está fuera del rango 1..número de columnas devuelve #NUM!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=EXPLODE({"A","x,y";"B","z"},2,",")` | {"A","x";"A","y";"B","z"} | Expande la columna 2 en filas |
| `=EXPLODE({"A","x,y";"B","z"},2,",",TRUE)` | {"A","x","y";"B","z",""} | Expansión en columnas, con relleno de celdas vacías |

## Notas

- Se pueden indicar varios separadores; se aplican primero los más largos (predeterminado: espacio).
- En la expansión horizontal, las celdas que faltan se rellenan con el valor de relleno (predeterminado "").
- Con ignore_empty TRUE se eliminan los fragmentos vacíos (si todos están vacíos, se conserva una cadena vacía).
- Compatibilidad: Excel 2010+. Se registra siempre como `EXPLODE` en todas las versiones de Excel.

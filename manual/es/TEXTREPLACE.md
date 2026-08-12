# TEXTREPLACE

**Categoría**: función exclusiva de EGTools

Reemplaza cada sección ClaveInicio…ClaveFin del texto.

## Sintaxis

```
=TEXTREPLACE(texto, clave_inicio, clave_fin, [reemplazo], [incluir])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto | Obligatorio | texto fuente |
| clave_inicio | Obligatorio | marca de inicio |
| clave_fin | Obligatorio | marca de fin |
| reemplazo | Opcional | texto de reemplazo (por defecto vacío) |
| incluir | Opcional | VERDADERO conserva las marcas |

## Devuelve

Devuelve el texto con las sustituciones aplicadas como un texto único (escalar). Cuando el texto, la clave inicial o la clave final están vacíos, devuelve el texto original sin cambios en lugar de un error.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=TEXTREPLACE("a[x]b[y]c","[","]","*")` | a*b*c | sustituye todas las secciones por * |
| `=TEXTREPLACE("a<1>b","<",">","2",TRUE)` | a<2>b | sustitución conservando las claves |

## Notas

- La búsqueda de claves ignora mayúsculas y minúsculas.
- Si se encuentra una clave inicial sin clave final, solo se sustituye esa clave inicial y el resto del texto se conserva.
- Si el texto es una matriz, se procesa elemento a elemento y se derrama una matriz de la misma forma (un elemento de error devuelve ese error); una matriz en las claves inicial y final y los demás argumentos devuelve #VALUE!.
- Función relacionada: TEXTBETWEEN
- Compatibilidad: Excel 2010+. Se registra siempre como `TEXTREPLACE` en todas las versiones de Excel.

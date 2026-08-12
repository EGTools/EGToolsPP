# STREXT

**Categoría**: función exclusiva de EGTools

Extrae o elimina caracteres por tipo (dígitos, latinos, coreano, japonés, hanja) o por regex.

## Sintaxis

```
=STREXT(texto, opcion, [delimitador])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto | Obligatorio | texto fuente |
| opcion | Obligatorio | S=espacio A=latino N=dígitos K=coreano J=japonés H=hanja; "-" inicial elimina; otro texto es regex |
| delimitador | Opcional | unión entre fragmentos extraídos (por defecto ninguno) |

## Devuelve

Devuelve el resultado de la extracción (o eliminación) como un texto único (escalar). Devuelve #VALUE! cuando faltan el texto o la opción, la opción está vacía, no resulta ninguna clase de caracteres válida o la expresión regular no es válida.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=STREXT("abc123한글45","N")` | 12345 | extrae solo los dígitos |
| `=STREXT("abc123한글45","K")` | 한글 | extrae solo el texto coreano (hangul) |
| `=STREXT("a1b2","N","-")` | 1-2 | une con el separador |

## Notas

- La clase N también incluye el punto decimal '.'.
- Si la opción no es una combinación de las letras -SANKJH, toda la opción se interpreta como una expresión regular ECMAScript (std::wregex).
- Si la opción comienza con '-', los caracteres correspondientes se eliminan en lugar de extraerse.
- Si el texto es una matriz, se procesa elemento a elemento y se derrama una matriz de la misma forma (un elemento de error devuelve ese error); una matriz en las opciones y el separador devuelve #VALUE!.
- Compatibilidad: Excel 2010+. Se registra siempre como `STREXT` en todas las versiones de Excel.

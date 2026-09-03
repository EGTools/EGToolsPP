# JSONFILTER

**Categoría**: función exclusiva de EGTools

Analiza texto JSON en una tabla de dos columnas (ruta, valor).

## Sintaxis

```
=JSONFILTER(texto_json, [ruta_clave])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto_json | Obligatorio | texto JSON a analizar |
| ruta_clave | Opcional | ruta a filtrar, p. ej. "data.items[0].name" (por defecto: todo) |

## Devuelve

Devuelve una matriz derramada de dos columnas (ruta, valor): los números se mantienen como números, true/false se convierten en booleanos y null se convierte en una cadena vacía. Si el texto está vacío, falla el análisis del JSON o no se encuentra la ruta de clave, devuelve #VALUE!; si el resultado es un objeto o una matriz vacíos, devuelve #N/A.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=JSONFILTER("{""a"":1,""b"":[true,null]}")` | {"a",1;"b[0]",TRUE;"b[1]",""} | Aplana todo el documento JSON |
| `=JSONFILTER("{""a"":{""b"":5,""c"":""x""}}","a")` | {"b",5;"c","x"} | Muestra solo el contenido bajo la ruta de clave |

## Notas

- Usa un analizador ligero propio, sin bibliotecas externas.
- El texto completo debe ser JSON válido; no se admiten comentarios ni comas finales.
- La ruta de clave usa el formato "a.b[0].c"; los índices de matriz empiezan en 0.
- Compatibilidad: Excel 2010+. Se registra siempre como `JSONFILTER` en todas las versiones de Excel.

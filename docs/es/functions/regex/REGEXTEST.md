# REGEXTEST

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Microsoft 365

Comprueba si el texto coincide con una expresión regular.

## Sintaxis

```
=REGEXTEST(texto, patron, [distincion_mayusculas])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto | Obligatorio | texto a comprobar |
| patron | Obligatorio | patrón regex (ECMAScript) |
| distincion_mayusculas | Opcional | 0 distingue (por defecto), 1 ignora |

## Devuelve

Devuelve un valor lógico TRUE/FALSE (escalar) que indica si hay coincidencia. Devuelve #VALUE! si la expresión regular no es válida.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=REGEXTEST("abc123","\d+")` | TRUE | contiene dígitos |
| `=REGEXTEST("ABC","[a-z]+",1)` | TRUE | ignora mayúsculas y minúsculas |
| `=REGEXTEST("abc","^\d")` | FALSE | sin coincidencia |

## Notas

- La sintaxis de expresiones regulares es la de ECMAScript de std::wregex (puede diferir en algunos aspectos del PCRE2 de la versión nativa de 365).
- text·pattern aceptan matrices y se calculan elemento a elemento, derramando una matriz de la misma forma: los escalares se repiten, un vector columna × un vector fila se expande como producto externo, los elementos con tamaños no coincidentes dan #N/A y un elemento de error devuelve ese error.
- Funciones relacionadas: REGEXEXTRACT, REGEXREPLACE
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `REGEXTEST` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.REGEXTEST`.

# REGEXREPLACE

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Microsoft 365

Reemplaza el texto que coincide con una expresión regular ($1.. retro-referencias).

## Sintaxis

```
=REGEXREPLACE(texto, patron, reemplazo, [aparicion], [distincion_mayusculas])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto | Obligatorio | texto original |
| patron | Obligatorio | patrón regex |
| reemplazo | Obligatorio | texto de reemplazo ($1, $2…) |
| aparicion | Opcional | 0 todas (por defecto), N = N-ésima |
| distincion_mayusculas | Opcional | 0 distingue (por defecto), 1 ignora |

## Devuelve

Devuelve el texto sustituido como un texto único (escalar). Devuelve #VALUE! si la expresión regular no es válida.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=REGEXREPLACE("a1b2c3","\d","-")` | a-b-c- | sustituye todas las coincidencias |
| `=REGEXREPLACE("a1b2c3","\d","-",2)` | a1b-c3 | sustituye solo la 2.ª |
| `=REGEXREPLACE("john smith","(\w+) (\w+)","$2 $1")` | smith john | sustitución con retrorreferencias |

## Notas

- Se admiten las retrorreferencias $1, $2 en el texto de reemplazo.
- La ocurrencia 0 (predeterminada) sustituye todas las coincidencias; con N>0 se sustituye solo la N-ésima.
- La sintaxis de expresiones regulares es la de ECMAScript de std::wregex.
- text·pattern·replacement aceptan matrices y se calculan elemento a elemento, derramando una matriz de la misma forma: los escalares se repiten, un vector columna × un vector fila se expande como producto externo, los elementos con tamaños no coincidentes dan #N/A y un elemento de error devuelve ese error.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `REGEXREPLACE` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.REGEXREPLACE`.

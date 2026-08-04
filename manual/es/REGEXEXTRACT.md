# REGEXEXTRACT

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2024 / Microsoft 365

Extrae coincidencias de una expresión regular del texto.

## Sintaxis

```
=REGEXEXTRACT(texto, patron, [modo_retorno], [distincion_mayusculas])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto | Obligatorio | texto original |
| patron | Obligatorio | patrón regex |
| modo_retorno | Opcional | 0 primera, 1 todas, 2 grupos de captura |
| distincion_mayusculas | Opcional | 0 distingue (por defecto), 1 ignora |

## Devuelve

Con el modo de retorno 0 devuelve el texto de la primera coincidencia (escalar), con 1 todas las coincidencias (matriz vertical de una columna, se derrama) y con 2 los grupos de captura de la primera coincidencia (matriz horizontal de una fila, se derrama). Devuelve #N/A cuando no hay coincidencias (incluido el caso sin grupos en el modo 2) y #VALUE! si la expresión regular no es válida.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=REGEXEXTRACT("abc123def45","\d+")` | 123 | primera coincidencia |
| `=REGEXEXTRACT("abc123def45","\d+",1)` | {123;45} | todas las coincidencias (vertical) |
| `=REGEXEXTRACT("2026-08-01","(\d+)-(\d+)",2)` | {2026,08} | grupos de captura (horizontal) |

## Notas

- La sintaxis de expresiones regulares es la de ECMAScript de std::wregex (puede diferir en algunos aspectos del PCRE2 de la versión nativa de 365).
- Funciones relacionadas: REGEXTEST, REGEXREPLACE
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `REGEXEXTRACT` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.REGEXEXTRACT`.

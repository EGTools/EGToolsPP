# TEXTAFTER

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2024

Devuelve el texto posterior a un delimitador.

## Sintaxis

```
=TEXTAFTER(texto, delimitador, [num_instancia], [modo_coincidencia], [coincidir_final], [si_no_encontrado])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto | Obligatorio | texto original |
| delimitador | Obligatorio | delimitador a buscar |
| num_instancia | Opcional | qué aparición, negativo desde el final (por defecto 1) |
| modo_coincidencia | Opcional | 1 ignora mayúsculas (por defecto 0) |
| coincidir_final | Opcional | 1 trata el final del texto como delimitador (por defecto 0) |
| si_no_encontrado | Opcional | valor si no se encuentra (por defecto #N/A) |

## Devuelve

Devuelve el texto único (escalar) situado después del separador. Devuelve #VALUE! cuando el número de instancia es 0 y el valor de no encontrado (predeterminado #N/A) cuando no se encuentra el separador.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=TEXTAFTER("red-blue-green","-")` | blue-green | después del primer separador |
| `=TEXTAFTER("a,b,c",",",-1)` | c | después del último separador |
| `=TEXTAFTER("abc","x",1,0,0,"none")` | none | devuelve el valor de no encontrado |

## Notas

- Con match_end=1, el final del texto (instancia positiva) o su inicio (instancia negativa) se considera un separador virtual.
- Con match_mode=1 se ignoran mayúsculas y minúsculas.
- Si texto es una matriz, se aplica elemento a elemento y se derrama una matriz de la misma forma (un elemento de error devuelve ese error); un separador en matriz devuelve #VALUE!.
- Funciones relacionadas: TEXTBEFORE, TEXTSPLIT
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `TEXTAFTER` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.TEXTAFTER`.

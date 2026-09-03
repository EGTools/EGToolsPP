# NUMBERVALUE

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Convierte texto en número de forma independiente de la configuración regional.

## Sintaxis

```
=NUMBERVALUE(texto, [separador_decimal], [separador_grupo])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto | Obligatorio | texto a convertir |
| separador_decimal | Opcional | separador decimal (pred. .) |
| separador_grupo | Opcional | separador de miles (pred. ,) |

## Devuelve

Devuelve un número (escalar); si se especifica una matriz como texto, se calcula elemento a elemento y se derrama en Excel con matrices dinámicas. Devuelve #VALUE! si, tras quitar los separadores, el texto no se interpreta como número; el texto vacío (solo espacios) devuelve 0.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=NUMBERVALUE("2.500,27",",",".")` | 2500.27 | Conversión de notación europea |
| `=NUMBERVALUE("9%%")` | 0.0009 | Cada % divide entre 100 |
| `=NUMBERVALUE("abc")` | #VALUE! | No es un número |

## Notas

- Todos los espacios en blanco se ignoran independientemente de su posición, y los separadores de miles se eliminan estén donde estén.
- Los signos % finales dividen el resultado entre 100 una vez por cada signo.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `NUMBERVALUE` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.NUMBERVALUE`.

# HSTACK

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2024

Apila matrices horizontalmente en una sola.

## Sintaxis

```
=HSTACK(matriz1, ...)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| matriz1 | Obligatorio | matriz a apilar |
| … | Opcional | Los argumentos anteriores se pueden repetir |

## Devuelve

Devuelve una matriz con las matrices pasadas como argumentos unidas horizontalmente; en las versiones de Excel que admiten matrices dinámicas el resultado se derrama. Si las alturas difieren, las posiciones vacías de la matriz más corta se rellenan con #N/A. Devuelve #VALUE! si no se proporciona ningún argumento de matriz válido.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=HSTACK({1;2},{3,4;5,6})` | {1,3,4;2,5,6} | une dos matrices horizontalmente |
| `=HSTACK({1,2},{3;4;5})` | {1,2,3;#N/A,#N/A,4;#N/A,#N/A,5} | si las alturas difieren, #N/A |

## Notas

- Los argumentos omitidos se ignoran.
- Función relacionada: VSTACK
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `HSTACK` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.HSTACK`.

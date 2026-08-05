# VSTACK

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2024

Apila matrices verticalmente en una sola.

## Sintaxis

```
=VSTACK(matriz1, ...)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| matriz1 | Obligatorio | matriz a apilar |
| … | Opcional | Los argumentos anteriores se pueden repetir |

## Devuelve

Devuelve una matriz con las matrices dadas como argumentos apiladas verticalmente; en las versiones de Excel que admiten matrices dinámicas, el resultado se derrama. Si los anchos difieren, las posiciones vacías de las matrices más estrechas se rellenan con #N/A. Devuelve #VALUE! si no se proporciona ningún argumento de matriz válido.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=VSTACK({1,2},{3,4;5,6})` | {1,2;3,4;5,6} | Apila dos matrices verticalmente |
| `=VSTACK({1,2,3},{4,5})` | {1,2,3;4,5,#N/A} | Anchos distintos: se rellena con #N/A |

## Notas

- Los argumentos omitidos se ignoran.
- Función relacionada: HSTACK
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `VSTACK` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.VSTACK`.

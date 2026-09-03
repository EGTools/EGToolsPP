# RRI

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Devuelve una tasa de interés equivalente para el crecimiento de una inversión.

## Sintaxis

```
=RRI(nper, pv, fv)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| nper | Obligatorio | número de períodos |
| pv | Obligatorio | valor actual |
| fv | Obligatorio | valor futuro |

## Devuelve

Devuelve la tasa de interés equivalente por período correspondiente al crecimiento de una inversión (número escalar). Si un argumento no es numérico, si nper no es mayor que 0 o si pv es 0, devuelve un error #NUM!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=RRI(96,10000,11000)` | 0.000993307 | Tasa para pasar de 10000 a 11000 en 96 períodos |

## Notas

- Todos los argumentos aceptan matrices y se calculan elemento a elemento, derramando una matriz de la misma forma: los escalares se repiten, un vector columna × un vector fila se expande como producto externo, los elementos con tamaños no coincidentes dan #N/A y un elemento de error devuelve ese error.
- Función relacionada: PDURATION
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `RRI` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.RRI`.

# PERCENTOF

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Microsoft 365

Devuelve la suma de un subconjunto dividida por la suma de todos los datos.

## Sintaxis

```
=PERCENTOF(data_subset, data_all)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| data_subset | Obligatorio | subconjunto de datos |
| data_all | Obligatorio | todos los datos |

## Devuelve

Devuelve como número escalar la proporción resultante de dividir la suma del subconjunto de datos entre la suma de todos los datos. Si faltan argumentos, devuelve #VALUE!; si la suma total es 0 o el conjunto total no contiene números, devuelve #DIV/0!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=PERCENTOF({10;20},{10;20;30;40})` | 0.3 | Proporción de la suma parcial sobre la suma total |
| `=PERCENTOF(50,200)` | 0.25 | Proporción entre valores escalares |

## Notas

- Los valores no numéricos se ignoran al calcular las sumas.
- Si se especifica "PERCENTOF" en el argumento function de GROUPBY/PIVOTBY, se obtiene la proporción de cada grupo.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `PERCENTOF` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.PERCENTOF`.

# PDURATION

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Devuelve el número de períodos necesarios para que una inversión alcance un valor.

## Sintaxis

```
=PDURATION(rate, pv, fv)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| rate | Obligatorio | tasa de interés por período |
| pv | Obligatorio | valor actual |
| fv | Obligatorio | valor futuro |

## Devuelve

Devuelve el número de períodos necesarios para alcanzar el valor objetivo (número escalar). Si rate, pv o fv no es numérico o no es mayor que 0, devuelve un error #NUM!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=PDURATION(0.025,2000,2200)` | 3.859866163 | Períodos para pasar de 2000 a 2200 al 2,5 % |

## Notas

- Función relacionada: RRI
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `PDURATION` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.PDURATION`.

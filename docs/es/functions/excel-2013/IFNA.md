# IFNA

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Devuelve valor_si_nd si el valor es el error #N/A; de lo contrario, el valor.

## Sintaxis

```
=IFNA(valor, valor_si_nd)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| valor | Obligatorio | el valor o expresión |
| valor_si_nd | Obligatorio | valor a devolver si es #N/A |

## Devuelve

Devuelve el valor alternativo si el valor es el error #N/A; en caso contrario, devuelve el valor original. Si se proporciona una matriz, se aplica elemento a elemento y devuelve una matriz que se derrama.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=IFNA(NA(),"none")` | none | Sustituye #N/A |
| `=IFNA(1/0,0)` | #DIV/0! | Los errores distintos de #N/A se conservan |
| `=IFNA(5,0)` | 5 | El valor normal se conserva |

## Notas

- Tanto el valor como el valor alternativo aceptan matrices y se difunden entre sí, sustituyendo solo #N/A elemento a elemento (los escalares se repiten, los elementos de tamaño no coincidente dan #N/A; la matriz resultante se derrama).
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `IFNA` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.IFNA`.

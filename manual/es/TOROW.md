# TOROW

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2024

Devuelve la matriz como una sola fila.

## Sintaxis

```
=TOROW(matriz, [ignorar], [explorar_por_columna])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| matriz | Obligatorio | la matriz |
| ignorar | Opcional | 0 nada,1 vacíos,2 errores,3 ambos |
| explorar_por_columna | Opcional | VERDADERO para explorar por columna |

## Devuelve

Devuelve una matriz con todos los valores de la matriz dispuestos en una sola fila (1×N); en las versiones de Excel que admiten matrices dinámicas, el resultado se derrama. La opción de ignorar (1=valores vacíos, 2=errores, 3=ambos) permite excluir valores. Devuelve #VALUE! si la matriz está vacía o no queda ningún valor.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=TOROW({1,2;3,4})` | {1,2,3,4} | Aplanado por filas a una sola fila |
| `=TOROW({1,2;3,4},0,TRUE)` | {1,3,2,4} | Recorrido por columnas |
| `=TOROW({1,"",2},1)` | {1,2} | Ignora los valores vacíos |

## Notas

- Las cadenas vacías ("") también se consideran valores vacíos y se excluyen con las opciones de ignorar 1 y 3.
- Si se excluyen todos los valores, devuelve #VALUE! en lugar del #CALC! de la función nativa.
- Función relacionada: TOCOL
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `TOROW` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.TOROW`.

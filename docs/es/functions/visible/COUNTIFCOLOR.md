# COUNTIFCOLOR

**Categoría**: función exclusiva de EGTools

Cuenta las celdas cuyo color mostrado coincide con una celda de referencia.

## Sintaxis

```
=COUNTIFCOLOR(rango_busqueda, celda_color, [fuente])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| rango_busqueda | Obligatorio | el rango a comprobar |
| celda_color | Obligatorio | celda con el color de referencia (se usa la primera) |
| fuente | Opcional | TRUE = comparar color de fuente (predeterminado FALSE) |

## Devuelve

Devuelve como número la cantidad de celdas cuyo color mostrado coincide con el de la celda de referencia. Devuelve #VALUE! si falla la conexión COM, si no se puede leer el color de referencia o si se examinan más de 100.000 celdas, y devuelve 0 si el rango examinado no se superpone con el área usada.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=COUNTIFCOLOR(A1:A10,C1)` |  | Número de celdas con el mismo color de fondo que C1 (el resultado depende del formato de las celdas) |

## Notas

- Con color_fuente=TRUE solo se cuentan las celdas no vacías.
- Se basa en DisplayFormat, por lo que compara con el color mostrado con el formato condicional aplicado.
- Cambiar solo el formato no provoca un recálculo (se necesita F9); al ser de tipo macro, también se excluye del recálculo multiproceso.
- Compatibilidad: Excel 2010+. Se registra siempre como `COUNTIFCOLOR` en todas las versiones de Excel.

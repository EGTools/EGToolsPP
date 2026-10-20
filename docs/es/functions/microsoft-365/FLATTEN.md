# FLATTEN

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Microsoft 365 (novedad de 2026)

Elimina niveles de anidamiento de un texto de lista o matriz y derrama los elementos en la cuadrícula.

## Sintaxis

```
=FLATTEN(matriz, [valor_relleno], [niveles])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| matriz | Obligatorio | texto de lista, rango o matriz a aplanar |
| valor_relleno | Opcional | valor para filas irregulares (pred. #N/A) |
| niveles | Opcional | niveles de anidamiento a eliminar (pred. todos) |

## Devuelve

Devuelve la matriz 2-D sin anidamiento, que se derrama en las versiones de Excel con matrices dinámicas. Si el propio argumento matriz es un error, devuelve ese error tal cual; una matriz en valor_relleno, un niveles no numérico o negativo, o un resultado vacío devuelven #VALUE!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=FLATTEN("a, b, c")` | {a;b;c} | lista → una columna |
| `=FLATTEN("{""a"",""b"";""c""}")` | {a,b;c,#N/A} | texto de matriz 2-D, huecos rellenados con #N/A |
| `=FLATTEN({"a, b";"c"})` | {a;b;c} | una columna de celdas de lista se apila verticalmente |

## Notas

- Un nivel = en cada fila de la matriz exterior, los elementos (escalar = 1×1, matriz = su propia forma) se unen en horizontal y las filas se apilan, igualando el ancho con valor_relleno (predeterminado #N/A). Si niveles se omite o es 0, se repite hasta que no quede anidamiento; un N positivo elimina N niveles y los elementos anidados restantes se muestran como texto de lista ({…}). Ejemplo: =FLATTEN("{""a"",{""b"",""c""}}",,1) → {a,{"b","c"}} — solo se elimina el primer anidamiento de esa celda.
- Una lista de una columna da una columna vertical, un rango de columna con varias celdas de lista se apila verticalmente y una sola celda con texto de matriz 2-D conserva esa forma 2-D. Las filas irregulares dentro del propio texto de matriz se rellenan siempre con #N/A, independientemente de valor_relleno; los elementos de error de la lista (#N/A, etc.) se conservan como datos. Un texto único sin comas ni llaves se convierte en un texto 1×1.
- Las listas en celdas de la beta de Microsoft 365 solo se ven como #UNKNOWN! (#VALUE! en Excel 2016) en Excel antiguo, por lo que conviene convertirlas primero con la cinta [EGTools] › [Celdas de lista/matriz → texto]. Funciones relacionadas: HAS, HASANY, HASALL
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `FLATTEN` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.FLATTEN`.

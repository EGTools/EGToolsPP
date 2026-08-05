# Combinar/Separar celdas

**Ubicación**: cinta, pestaña `EGTools` → grupo `Combinar/Separar celdas`

Grupo de funciones para trabajar con celdas combinadas. Se compone de 2 botones divididos —
**Combinar** (acción predeterminada = Combinar sin formato) y **Separar** (acción predeterminada = Separar a filas sin formato).

- **9 menús de combinación** (en el orden de la cinta): Combinar valores iguales · Combinar conservando valores · Combinar iguales conservando ·
  Combinar sin formato / Combinar con formato · Combinar filas sin formato / Combinar filas con formato ·
  Combinar columnas sin formato / Combinar columnas con formato
- **5 menús de separación**: Separar y rellenar · Separar a filas sin formato / Separar a filas con formato ·
  Separar a columnas sin formato / Separar a columnas con formato

## Combinar valores iguales

Convierte en celdas combinadas reales los tramos de la selección donde el mismo valor se repite en vertical.
Se usa para agrupar visualmente las columnas de una lista ordenada en las que se repite la misma categoría.

### Procedimiento

1. Seleccione el rango a combinar (se recomienda una lista ordenada) y pulse el botón.
2. En cada columna se buscan y combinan los **tramos donde el mismo valor se
   repite en vertical durante 2 o más celdas**. Los tramos de celdas vacías no se combinan.
3. Si selecciona varias columnas, la combinación es **jerárquica** — los tramos
   de la columna derecha se cortan en los límites de grupo de la(s) columna(s)
   izquierda(s). Por ejemplo, si la columna A es el departamento y la B el
   equipo, la combinación de equipos también se corta donde cambia el departamento.

### Notas

- La comparación de valores se basa en el valor mostrado (en las fórmulas, en el resultado calculado).
- Como solo se combinan tramos con el mismo valor, la operación avanza sin
  ventana de confirmación y ningún valor se pierde (por la regla de combinación
  de Excel solo queda el valor superior izquierdo, pero todos los valores del tramo son iguales).
- Si selecciona varias áreas (Ctrl+clic), cada área se procesa por separado.

## Combinar conservando valores

Combina la(s) selección(es) **solo en apariencia**. El valor de cada celda se
conserva, de modo que funciones como `SUM` o `VLOOKUP`, los filtros y la
ordenación **siguen funcionando**, mientras que en pantalla solo se ve el valor
de la celda superior izquierda, como en una celda combinada.

1. Seleccione el rango que desea mostrar combinado y pulse el menú (puede
   seleccionar varias áreas — cada una se procesa por separado).
2. Internamente, la hoja actual se **copia a un libro temporal**, en la copia se
   realiza la combinación real y luego **solo su formato** se pega en la misma
   posición del original. Los valores y las fórmulas no cambian en absoluto, y
   el libro temporal se cierra automáticamente sin guardarse.

- Como la combinación es solo aparente, los valores de las celdas distintas de
  la superior izquierda quedan ocultos en pantalla pero no se eliminan. Si una
  fórmula los referencia, se leen con normalidad.

## Combinar iguales conservando

Busca los tramos con las mismas reglas que [Combinar valores iguales](#combinar-valores-iguales)
(tramos verticales de 2 o más celdas con el mismo valor; varias columnas de forma jerárquica),
pero en lugar de combinar de verdad, combina **solo en apariencia**, como
[Combinar conservando valores](#combinar-conservando-valores). Como el valor de
cada celda se conserva, los filtros y las agregaciones siguen operando sobre los
valores después de combinar.

- Las celdas ya combinadas dentro de la selección se separan primero (en la
  copia) y se rellenan con sus valores antes de calcular los tramos. Los valores
  del original no cambian.

## Combinar sin formato / Combinar con formato

Convierte toda la selección en **una sola celda combinada** uniendo también su contenido.

1. Seleccione el rango y pulse el menú. Si hay 2 o más columnas, se pide el
   **carácter separador de columnas** (predeterminado `,`; si cancela, la operación termina).
2. Las celdas de una misma fila se unen con el separador y **las filas se unen
   con saltos de línea**. Las celdas vacías se omiten. Las celdas con fórmulas
   se unen como el texto de su resultado calculado.
3. El área queda combinada y contiene el texto unido. La celda resultante pasa a
   tener **formato de texto (@)**.
   - **Sin formato**: restablece el estilo de la celda al estándar y activa el ajuste automático de texto.
   - **Con formato**: memoriza 11 atributos de formato de fuente de **cada
     carácter** del original (fuente, tamaño, color, negrita, cursiva,
     subrayado, tachado, subíndice/superíndice, etc.) y los vuelve a aplicar
     carácter por carácter al texto combinado.

## Combinar filas sin formato / Combinar filas con formato

Procesa **cada fila** de la selección con el método de combinación de contenido
anterior, creando una celda combinada por fila. Si hay 2 o más columnas, el
carácter separador se pide **una sola vez** y se reutiliza en todas las filas.
El progreso se muestra en la barra de estado.

## Combinar columnas sin formato / Combinar columnas con formato

Convierte **cada columna** de la selección en una celda combinada. Los valores
consecutivos en vertical se unen con **saltos de línea**, por lo que no se pide
carácter separador.

## Separar y rellenar

Separa las celdas combinadas de la selección y vuelve a rellenar en las celdas
separadas el valor que contenían. Se usa para que los filtros, la ordenación y
las referencias de fórmulas sigan operando sobre los valores tras deshacer la combinación.

### Procedimiento

1. Seleccione el rango que desea rellenar y pulse el botón. Si la selección no
   es un rango de celdas, la operación termina silenciosamente sin hacer nada.
2. Para cada celda combinada de la selección, se deshace la combinación y el
   valor que contenía se vuelve a rellenar en las celdas separadas. Las celdas
   combinadas que estaban vacías se dejan vacías.
3. Si la combinación abarcaba varias **columnas**, el valor se rellena **solo en
   la primera columna de la izquierda** (en las combinaciones de varias filas se
   rellena en todas las filas de la primera columna).

### Precaución

- **Las fórmulas se convierten en valores.** Si la celda combinada contenía una
  fórmula, se rellena el resultado calculado y la fórmula desaparece. Si
  necesita conservar las fórmulas, haga una copia de seguridad antes de ejecutar.

## Separar a filas sin formato / Separar a filas con formato

Divide el texto de **una celda combinada** por los **saltos de línea** y coloca
cada fragmento en una fila, de arriba abajo. Es la operación inversa a la
combinación de contenido.

1. Seleccione la celda combinada (o una sola celda) y pulse el menú.
2. El texto se divide por los saltos de línea y se coloca en cada fila tras
   deshacer la combinación. Si hay más fragmentos que filas combinadas, se
   **insertan filas** por la diferencia.
3. Cada celda de destino pasa a formato de texto (@) y la altura de la fila se
   ajusta automáticamente (AutoFit). La variante **con formato** conserva el
   formato de los caracteres del original en cada fragmento.

- Si no hay saltos de línea y solo resulta 1 fragmento, se muestra un aviso y la operación termina.

## Separar a columnas sin formato / Separar a columnas con formato

Divide el texto de una celda combinada por un **carácter separador** y coloca
cada fragmento en una columna, de izquierda a derecha.

1. Seleccione la celda combinada y pulse el menú; se pide el carácter separador
   (predeterminado `,`; si cancela o lo deja vacío, la operación termina).
2. Si hay más fragmentos que columnas combinadas, se **insertan columnas**, y
   cada celda de destino recibe formato de texto (@) + ajuste automático del
   ancho de columna. La variante **con formato** conserva el formato de los
   caracteres del original en cada fragmento.

## Notas comunes

### Casos en los que no se puede ejecutar

| Situación | Funciones afectadas | Mensaje |
|---|---|---|
| La selección no es un rango de celdas | Todas | "Seleccione primero un rango de celdas." |
| Se solapa con un área de tabla (ListObject) | Combinar valores iguales · Combinar iguales conservando · Combinar sin/con formato · Combinar filas/columnas | "No se pueden procesar áreas de tabla." — dentro de una tabla no se admite la combinación de celdas |
| Se seleccionó una sola fila | Combinar valores iguales · Combinar iguales conservando | "No se puede procesar una sola fila." |
| Selección de varias áreas (Ctrl+clic) | Combinar sin/con formato | "Solo se puede procesar un área a la vez." |
| Más de 100 filas | Combinar sin/con formato · Combinar columnas | "No se pueden procesar rangos de más de 100 filas." |
| Contiene celdas ya combinadas | Combinar sin/con formato (todo el rango ya es una celda combinada) · Combinar filas/columnas | "El rango contiene celdas ya combinadas. Sepárelas primero." |
| No es una única celda combinada | Separar a filas/columnas | "Seleccione una sola celda combinada para ejecutar esta acción." |
| Nada que dividir | Separar a filas | "La celda no tiene saltos de línea (Alt+Entrar) para dividir." |
| Nada que dividir | Separar a columnas | "La celda no tiene delimitador para dividir." |

### Precaución

- **Las variantes "con formato" leen y vuelven a aplicar el formato carácter por
  carácter, por lo que con textos largos tardan más.** El progreso se muestra en
  la barra de estado. Si no necesita el formato de los caracteres, la variante
  "sin formato" es mucho más rápida.
- Las celdas resultantes de combinar/separar contenido tienen formato de texto,
  por lo que no se reconocen automáticamente como números o fechas. Las celdas
  que contenían fórmulas pasan a ser el texto de su resultado calculado.

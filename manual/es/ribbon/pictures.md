# Imagen

**Ubicación**: pestaña `EGTools` de la cinta → grupo `Imagen`

Grupo de funciones para insertar, alinear y extraer imágenes ajustadas a las celdas.
Se compone de 2 botones divididos:
**Imagen** (acción predeterminada = Insertar imagen; menú: Insertar imagen · Insertar por nombre · Insertar en formulario) y
**Ajustar** (acción predeterminada = Ajustar selección; menú: Ajustar selección · Ajustar todo · Exportar todo).

## Insertar imagen

Permite elegir un archivo de imagen e insertarlo ajustado al tamaño de la selección actual.

1. Elija un archivo de imagen en el cuadro de diálogo de selección de archivos
   (jpg·jpeg·png·gif·bmp·tif·tiff·svg·emf·wmf, etc.; cancelar finaliza la operación).
2. La imagen se inserta en la posición y el tamaño de la selección con un **margen de 0.2 pt**
   (se guarda incrustada en el documento, no como vínculo a un archivo).
3. El **nombre de la imagen insertada se establece con el nombre del archivo** y el
   **bloqueo de la relación de aspecto se desactiva**, de modo que la imagen pueda
   deformarse libremente al ajustarla a celdas más adelante.

## Insertar por nombre

Permite elegir una carpeta, busca archivos de imagen **con el mismo nombre que el valor de la celda**
y los inserta en esa celda. Se usa para pegar fotos en bloque en tablas donde ya hay
nombres introducidos, como listados de personas o de productos.

1. Elija la carpeta que contiene las imágenes en la ventana de selección de carpeta (cancelar finaliza).
2. Se comprueban todos los valores de celda del área usada de la hoja activa (desde A1)
   y se buscan en la carpeta archivos de imagen cuyo nombre coincida con el valor.
   - La extensión puede incluirse u omitirse (el valor de celda `foto1` coincide tanto con
     el archivo `foto1.jpg` como con `foto1`). No se distinguen mayúsculas y minúsculas.
   - Solo se consideran archivos con extensión de imagen: jpg·jpeg·png·gif·bmp·tif·tiff·svg·emf·wmf, etc.
3. La imagen encontrada se inserta en esa celda (si es una celda combinada, con el tamaño
   de toda el área combinada) con un **margen de 0.3 pt** (incrustada en el documento).
   El nombre de la imagen se establece con el nombre del archivo sin extensión y el
   bloqueo de la relación de aspecto se desactiva.
4. Al terminar se muestra el resultado (carpeta procesada / ninguna imagen añadida).

Si el mismo valor aparece en varias celdas, la imagen se inserta en cada una de ellas.

## Insertar en formulario

Clona un formulario e inserta en orden las imágenes de una carpeta en las posiciones
del marcador de posición `{{}}`. Se usa para rellenar formularios de informe fotográfico.

1. **Especificar el rango del formulario** — indique el rango del formulario en la ventana
   de entrada de rango (valor predeterminado = selección actual; cancelar finaliza).
   El formulario debe contener al menos 1 celda que contenga solo `{{}}` (posición de la foto);
   si no existe, se muestra un aviso y finaliza.
2. **Introducir el método de clonado** — `1` = duplicar la hoja del formulario completa /
   `2` = repetir el rango del formulario hacia abajo / `3` = repetir hacia la derecha.
3. **Seleccionar la carpeta** — los archivos de imagen de la carpeta se usan ordenados
   **por nombre** (si no hay imágenes, se muestra un aviso y finaliza).
4. Se clona el formulario hasta agotar las imágenes, insertando en cada celda `{{}}`
   (o en el área combinada si la celda está combinada) la imagen con un **margen de 0.3 pt**.
   **El valor de la celda se registra con el nombre del archivo** y el nombre de la imagen
   también es el nombre del archivo. Los `{{}}` que queden tras agotarse las imágenes
   se borran. El progreso se muestra en la barra de estado.

- El método 1 duplica la hoja, por lo que los nombres de hoja crecen como
  `NombreOriginal_1`, `_2`, …
- Los métodos 2/3 copian filas/columnas completas, por lo que se conservan el alto de fila
  y el ancho de columna (el contenido existente en la posición de clonado se borra).

## Ajustar selección

Ajusta las imágenes seleccionadas a la celda donde se encuentra la esquina superior
izquierda de cada imagen.

1. Seleccione la(s) imagen(es) y pulse el botón; aparece la ventana de **entrada de margen**
   (0–10, predeterminado 0.3). Introducir un valor no numérico o cancelar finaliza.
2. Cada imagen se ajusta al tamaño de su celda superior izquierda, con el margen hacia dentro.
   Si la celda superior izquierda es un **área combinada**, se ajusta al tamaño de toda
   el área combinada.
3. Las imágenes **giradas 90/270 grados** intercambian ancho y alto para que encajen
   en la celda manteniendo su rotación.

Si no hay ninguna imagen seleccionada, se muestra el aviso "No hay ninguna imagen seleccionada."
Solo se procesan objetos de imagen (fotos); las formas, gráficos, etc. no se tocan.

## Ajustar todo

Ajusta de una vez **todas las imágenes** de la hoja activa, de la misma forma que
Ajustar selección.

1. Si acepta la ventana de confirmación (continuar o no), aparece la ventana de
   **entrada de margen** (0–5, predeterminado 0.3).
2. Todas las imágenes de la hoja se ajustan a su celda superior izquierda
   (o al área combinada, si la hay). El progreso se muestra en la **barra de estado**
   de Excel con el nombre de cada imagen.
3. Al terminar se informa del resultado (ajuste completado / no hay imágenes).

## Exportar todo

Guarda **todas las imágenes** de la hoja activa en la carpeta indicada
**con la calidad original intacta**. Se usa para recuperar como archivos las fotos
pegadas en una hoja.

1. Si la hoja no tiene imágenes, se muestra un aviso y finaliza. Seleccione la carpeta
   de destino (cancelar finaliza).
2. El nombre de archivo es **el valor de la celda donde está la esquina superior izquierda
   de la imagen**. Si la celda está vacía, se usa el formato `NombreDeImagen_RfilaCcolumna`.
   Los caracteres no válidos para nombres de archivo se sustituyen por `_` y, si hay
   nombres duplicados, se añade `_1`, `_2`, … La extensión conserva el formato original
   guardado en el documento (jpg, png, etc.).
3. Funcionamiento interno: se guarda una **copia del libro en una carpeta temporal**
   y de ella se extraen tal cual los datos de imagen originales del documento.
   Como no se vuelve a dibujar con la calidad de la pantalla, **se conserva exactamente
   la calidad con la que la imagen se insertó en el documento**, y el documento original
   no se modifica en absoluto (los archivos temporales se eliminan automáticamente).
4. Al terminar se informa del número de imágenes extraídas y de la carpeta de destino.

- El documento debe estar en **formato xlsx/xlsm**. En otros formatos (.xls, etc.)
  puede mostrarse el aviso "Falló la exportación de imágenes."
- El objetivo son los objetos de imagen (fotos). Las formas y gráficos no se guardan.

## Notas

- Con un margen de 0, la imagen llena la celda por completo; cuanto mayor es el valor,
  mayor es la separación con el borde de la celda.
- La referencia de ajuste es la **celda sobre la que cae la esquina superior izquierda
  de la imagen**. Si no se ajusta a la celda deseada, mueva la imagen sobre esa celda
  y ejecute de nuevo.

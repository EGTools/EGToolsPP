# Formularios

**Ubicación**: pestaña `EGTools++` de la cinta → grupo `Formularios`

Grupo de funciones que rellena de forma repetida un formulario (área con formato)
dibujado en una hoja con los datos de una lista.
Se compone de un único botón dividido (acción predeterminada = Crear etiquetas).

| Menú | Descripción |
|---|---|
| [Crear etiquetas](#crear-etiquetas) | Repite el formulario en una sola hoja, una vez por fila de datos, sustituyendo `{{nombredecolumna}}` por los valores |
| [Combinar correspondencia](mailmerge.md) | Clona la hoja del formulario por cada fila de datos en hojas/archivos individuales (xlsx/PDF), impresión individual y envío de correo (con archivo adjunto) |

## Crear etiquetas

Clona el formulario tantas veces como filas tenga la lista de datos, sustituyendo
los marcadores de posición `{{nombredecolumna}}` por los valores de cada fila, para
crear una hoja de etiquetas imprimible (identificadores, etiquetas de dirección, etc.).

### Procedimiento

1. **Designar el área del formulario** — Indique el área del formulario en el
   cuadro de entrada de rango (valor predeterminado = la región contigua de la
   selección actual; cancelar termina). Si designa una sola celda sin valor,
   se termina.
2. **Designar la lista de datos** — Indique el área de la lista de datos.
   La **primera fila contiene los nombres de columna** y los datos empiezan en la
   segunda fila.
3. **Introducir el número de etiquetas en horizontal** — Introduzca cuántas
   etiquetas se colocan en horizontal por línea (1–10, predeterminado 1).
4. En una hoja nueva (nombre de la hoja del formulario + sufijo `_1`) el formulario
   se copia una vez por cada fila de datos y cada `{{nombredecolumna}}` se sustituye
   por el valor de esa fila. Al terminar se muestra un aviso.

### Reglas de sustitución

| Elemento | Regla |
|---|---|
| Marcadores de posición | Solo se sustituyen las celdas que contienen `{{nombredecolumna}}`. Una celda puede contener varios marcadores y combinarlos con otro texto (p. ej. `Dirección: {{Dirección}}`) |
| Sin marcadores | Si el formulario no contiene ningún marcador `{{ }}`, simplemente se clona el formulario vacío tantas veces como filas de datos haya |
| Nombres de columna vacíos | Las columnas cuya primera fila está vacía se pueden referenciar con los nombres `Empty1`, `Empty2`, … |
| Nombres de columna inexistentes | Si se referencia un nombre de columna que no está en la lista, se muestra una advertencia y la operación se aborta |
| Formato | El formato del formulario (bordes, colores, etc.) se copia también, y **el ancho de columna y el alto de fila se clonan idénticos** a los del formulario |

### Notas

- Las celdas sustituidas se escriben como **valores de texto** (no como fórmulas).
- Para crear una hoja/archivo individual por cada fila, use
  [Combinar correspondencia](mailmerge.md)
  (Crear etiquetas = repetición en una sola hoja; Combinar correspondencia =
  hoja/archivo individual por fila).

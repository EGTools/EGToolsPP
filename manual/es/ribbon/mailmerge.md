# Combinar correspondencia

**Ubicación**: pestaña `EGTools++` de la cinta → grupo `Formularios` → menú del botón dividido `Etiquetas` → **Combinar correspondencia**

**Clona la hoja del formulario completa por cada fila** de la lista de datos,
sustituye los marcadores de posición `{{nombredecampo}}` por los valores de esa fila
y, según la opción elegida, guarda archivos individuales (xlsx/PDF), imprime cada
copia y/o realiza el **envío de correo** (con el archivo de esa fila adjunto). Los
datos del servidor SMTP usados para el envío se registran en
[Configuración SMTP](smtp-settings.md) (la contraseña no se guarda).

## Procedimiento

1. **Introducir la opción** — Introduzca la tarea a ejecutar como un número (0–7,
   predeterminado 1). Es una **combinación de bits** de 1 (guardar archivos),
   2 (imprimir) y 4 (enviar correo).
   `0` = solo clonar las hojas de formulario (las hojas clonadas se conservan) /
   `1` = guardar archivos / `2` = imprimir / `3` = guardar + imprimir /
   `4` = enviar correo / `5` = guardar + enviar / `6` = imprimir + enviar /
   `7` = guardar + imprimir + enviar.
2. **Configuración de impresión** — Si la opción incluye imprimir (2, 3, 6 o 7),
   primero se muestra el cuadro de diálogo de configuración de impresora. Si lo
   cancela, la operación se aborta.
3. **Designar la hoja del formulario** — Seleccione cualquier celda de la hoja del
   formulario en el cuadro de entrada de rango. Se usa como formulario la **hoja
   completa** a la que pertenece el área seleccionada.
4. **Designar el área de datos** — Seleccione una celda de la lista de datos. La
   selección se amplía automáticamente a su región contigua (CurrentRegion); la
   **primera fila contiene los nombres de campo** y los datos empiezan en la
   segunda fila.
5. **Preparación del envío de correo** (opciones con envío, 4–7) — Se comprueban la
   columna de destinatarios y la hoja de plantilla de correo; después, si la
   [Configuración SMTP](smtp-settings.md) está vacía, el cuadro de diálogo de
   configuración se abre automáticamente. A continuación se pide la **contraseña de
   la cuenta SMTP** (no se guarda; si cancela, la operación se aborta).
6. Por cada fila de datos, la hoja del formulario se clona al final del libro, cada
   `{{nombredecampo}}` se sustituye y, según la opción, se imprime, se guarda y se
   envía. Al terminar se muestra "Elementos completados: N" — y, si se eligió el
   envío, también "Correos enviados: N".

## Reglas

| Elemento | Regla |
|---|---|
| Marcadores de posición | `{{nombredecampo}}` se hace corresponder con los nombres de campo de la fila de encabezados de la lista **ignorando mayúsculas/minúsculas y espacios al inicio o al final**. Una celda puede contener varios marcadores o combinarlos con otro texto. Si hay un nombre de campo que no está en la lista, se muestra un error indicando ese nombre y la operación se aborta |
| Nombres de hoja/archivo | El **valor de la primera columna** da nombre a la hoja clonada y al archivo guardado. Las filas con la primera columna vacía se omiten, y los caracteres prohibidos en el nombre se sustituyen por `_`. Si ya existe una hoja con el mismo nombre, se añade un sufijo `_1`, `_2`, … |
| Ubicación de guardado | Los archivos se guardan en la subcarpeta `Output\` de la carpeta donde está el libro. El **libro debe estar guardado**; si no, se muestra un aviso y la operación se aborta (también se aplica al envío de correo, que crea archivos adjuntos) |
| Guardado en PDF | Si el **encabezado de la primera columna de la lista es `PDF`** (sin distinguir mayúsculas), los archivos se guardan en PDF en lugar de xlsx |
| Contraseña de apertura | Si la lista tiene una columna `PassWord` (o en coreano `암호`), el valor de esa fila se establece como **contraseña de apertura** del xlsx (si está vacío, sin contraseña; no se aplica a los PDF) |
| Limpieza de hojas clonadas | Con cualquier opción distinta de 0, las hojas clonadas se eliminan tras imprimir, guardar o enviar. Con la opción 0 las hojas permanecen en el libro |

## Envío de correo (opciones 4, 5, 6 y 7)

Si la opción elegida incluye el envío, se aplican además estas reglas.

| Elemento | Regla |
|---|---|
| Columna de destinatarios | La lista de datos debe contener **obligatoriamente** una columna con el encabezado `eMail` (o en coreano `이메일`, sin distinguir mayúsculas). Si no existe, se muestra un aviso y la operación se aborta. Las filas con la celda de destinatario **vacía solo omiten el envío**; la clonación, el guardado y la impresión se realizan igualmente |
| Hoja de plantilla de correo | El libro debe contener **obligatoriamente** una hoja llamada `Email` (o en coreano `이메일`). Introduzca **columna 1 = clave, columna 2 = valor**. Claves: `Subject`/`제목` (**obligatoria**), `Body`/`본문`, `CC`/`참조`, `BCC`/`숨은참조` |
| Sustitución en asunto/cuerpo | La sustitución `{{nombredecampo}}` también se aplica al asunto y al cuerpo, por lo que puede enviar **un correo personalizado por cada fila**. Los nombres de campo que no estén en la lista se dejan tal cual |
| Archivo adjunto | El archivo guardado para esa fila (xlsx o PDF, con contraseña de apertura si procede) se **adjunta** al correo. Si no se eligió guardar archivos (bit 1), el archivo creado para el adjunto se elimina tras el envío |
| Remitente/servidor | Se usan el nombre del remitente, el correo, el servidor, el puerto y el SSL de [Configuración SMTP](smtp-settings.md). Si la configuración está vacía, el cuadro de diálogo de configuración se abre automáticamente; si aun así queda incompleta, la operación se aborta |
| Contraseña | La contraseña de la cuenta SMTP se **pide en cada ejecución** y no se guarda en ningún sitio |
| Gestión de fallos | Si el envío de una fila falla, se **continúa con la fila siguiente**. El resumen final muestra las filas fallidas como una lista "nombre - descripción del error" |

## Notas

- Las celdas sustituidas se escriben como **valores de texto** (no como fórmulas).
- El guardado copia la hoja clonada a un libro nuevo de una sola hoja, lo guarda y
  lo cierra, por lo que cada archivo guardado contiene solo esa hoja.
- El motor de envío **solo admite SSL implícito (normalmente el puerto 465)**;
  STARTTLS en el puerto 587 no es compatible. Consulte los detalles de esta
  limitación en [Configuración SMTP](smtp-settings.md).
- Diferencia con [Crear etiquetas](forms.md#crear-etiquetas): las etiquetas
  **repiten el formulario en una sola hoja**, mientras que combinar correspondencia
  crea una **hoja/archivo individual por cada fila**.

# Menú de la cinta

Al cargar EGTools++ se añade la pestaña **EGTools++** a la cinta.
Los botones de funciones están organizados en **grupos** según su naturaleza, y a medida que se añadan funciones en el futuro, la cinta se ampliará por grupos.

## Grupo: Celdas visibles

Gestiona la copia y el pegado con filtros u ocultaciones activos (2 botones divididos: Copiar = copiar solo las celdas visibles / Pegar = pegar solo valores). Detalles: [Celdas visibles](visible-cells.md)

| Menú | Descripción |
|---|---|
| [Copiar celdas visibles](visible-cells.md#copiar-celdas-visibles) | Recuerda como origen de copia solo las celdas visibles de la selección y las copia al portapapeles |
| [Copiar todo](visible-cells.md#copiar-todo) | Recuerda como origen de copia toda la selección (incluidas las celdas ocultas) |
| [Pegar valores](visible-cells.md#pegar-valores--pegar-fórmulas--pegar-todo) | Pega el origen de copia recordado como valores, solo en las celdas visibles del destino |
| [Pegar fórmulas](visible-cells.md#pegar-valores--pegar-fórmulas--pegar-todo) | Pega como fórmulas, solo en las celdas visibles |
| [Pegar todo](visible-cells.md#pegar-valores--pegar-fórmulas--pegar-todo) | Pega valores, fórmulas y formato en bloque, solo en las celdas visibles |

## Grupo: Combinar/Separar celdas

Funciones para trabajar con celdas combinadas (2 botones divididos: Combinar = combinar contenido sin formato / Separar = separar a filas sin formato). Detalles: [Combinar/Separar celdas](merge-split.md)

| Menú | Descripción |
|---|---|
| [Combinar valores iguales](merge-split.md#combinar-valores-iguales) | Combina de verdad los tramos verticales consecutivos con el mismo valor (con varias columnas, de forma jerárquica) |
| [Combinar conservando valores](merge-split.md#combinar-conservando-valores) | Combina solo en apariencia, dejando los valores intactos — las funciones y los filtros siguen funcionando |
| [Combinar iguales conservando](merge-split.md#combinar-iguales-conservando) | Combina solo en apariencia cada tramo consecutivo con el mismo valor |
| [Combinar sin formato / Combinar con formato](merge-split.md#combinar-sin-formato--combinar-con-formato) | Combina el rango en una sola celda uniendo el contenido con separadores y saltos de línea (la variante con formato conserva el formato de los caracteres) |
| [Combinar filas sin formato / Combinar filas con formato](merge-split.md#combinar-filas-sin-formato--combinar-filas-con-formato) | Convierte cada fila en una celda combinada uniendo su contenido |
| [Combinar columnas sin formato / Combinar columnas con formato](merge-split.md#combinar-columnas-sin-formato--combinar-columnas-con-formato) | Convierte cada columna en una celda combinada uniendo su contenido con saltos de línea |
| [Separar y rellenar](merge-split.md#separar-y-rellenar) | Separa las celdas combinadas y vuelve a rellenar el valor que contenían (las fórmulas se convierten en valores) |
| [Separar a filas sin formato / Separar a filas con formato](merge-split.md#separar-a-filas-sin-formato--separar-a-filas-con-formato) | Divide el texto de una celda combinada por saltos de línea y lo reparte fila por fila (inserta filas si faltan) |
| [Separar a columnas sin formato / Separar a columnas con formato](merge-split.md#separar-a-columnas-sin-formato--separar-a-columnas-con-formato) | Divide el texto de una celda combinada por un separador y lo reparte columna por columna (inserta columnas si faltan) |

## Grupo: Imagen

Inserta, ajusta y extrae imágenes adaptadas a las celdas (2 botones divididos: Imagen = Insertar imagen / Ajustar = Ajustar selección).
Detalles: [Imagen](pictures.md)

| Menú | Descripción |
|---|---|
| [Insertar imagen](pictures.md#insertar-imagen) | Elige un archivo de imagen y lo inserta ajustado al tamaño de la selección |
| [Insertar por nombre](pictures.md#insertar-por-nombre) | Busca en una carpeta la imagen con el mismo nombre que el valor de la celda y la inserta en esa celda |
| [Insertar en formulario](pictures.md#insertar-en-formulario) | Clona un formulario e inserta en orden las imágenes de una carpeta en las posiciones `{{}}` |
| [Ajustar selección](pictures.md#ajustar-selección) | Ajusta las imágenes seleccionadas al tamaño de la celda bajo su esquina superior izquierda |
| [Ajustar todo](pictures.md#ajustar-todo) | Ajusta de una vez todas las imágenes de la hoja al tamaño de sus celdas |
| [Exportar todo](pictures.md#exportar-todo) | Guarda todas las imágenes de la hoja en una carpeta indicada, con su calidad original |

## Grupo: Calendario

Genera hojas de calendario y agenda (2 botones divididos: Calendario = Mensual / Agenda = Diaria B).
Detalles: [Calendario](calendar.md)

| Menú | Descripción |
|---|---|
| [Mensual](calendar.md#mensual) | Hoja de calendario mensual (con número de semana, festivos y calendario lunar) |
| [Anual A](calendar.md#anual-a--anual-b) | Calendario anual en formato vertical (3×4) |
| [Anual B](calendar.md#anual-a--anual-b) | Calendario anual en formato horizontal (4×3) |
| [Semanal](calendar.md#semanal) | Agenda semanal (domingo a sábado, eje horario 08:00~21:00) |
| [Diaria A](calendar.md#diaria-a--diaria-b) | Agenda diaria de 2 secciones (con arrastre de tareas pendientes) |
| [Diaria B](calendar.md#diaria-a--diaria-b) | Agenda diaria de 3 secciones (con arrastre de tareas pendientes) |

## Grupo: Formularios

Rellena un formulario con una lista de datos y lo genera de forma repetida (botón dividido, acción predeterminada = Crear etiquetas).
Detalles: [Formularios](forms.md)

| Menú | Descripción |
|---|---|
| [Crear etiquetas](forms.md#crear-etiquetas) | Clona el formulario por cada fila de datos sustituyendo `{{nombre de columna}}` por su valor |
| [Combinar correspondencia](mailmerge.md) | Clona la hoja de formulario por cada fila de datos para generar hojas/archivos individuales (xlsx/PDF), imprimir individualmente o enviar correos (con archivo adjunto) |

## Grupo: Herramientas

Colección de utilidades de uso frecuente (botón dividido, acción predeterminada = Limpiar errores). Detalles: [Herramientas](utilities.md)

| Menú | Descripción |
|---|---|
| [Eliminar estilos](utilities.md#eliminar-estilos) | Elimina en bloque los estilos de celda personalizados (con opción de conservar los integrados o en uso) |
| [Eliminar nombres](utilities.md#eliminar-nombres) | Elimina los nombres definidos con referencias rotas u ocultos |
| [Limpiar errores](utilities.md#limpiar-errores) | Envuelve las fórmulas con valores de error en `IFERROR` para que no se muestren |
| [Limpiar texto vacío](utilities.md#limpiar-texto-vacío) | Convierte las celdas con cadenas de longitud 0 en celdas realmente vacías |
| [Ordenar notas](utilities.md#ordenar-notas) | Mueve todas las notas de la hoja junto al lado derecho de su celda + tamaño automático |
| [Cambiar color de bordes](utilities.md#cambiar-color-de-bordes) | Cambia en bloque el color de todos los bordes de la selección (conservando el tipo de línea) |
| [Trazar forma](utilities.md#trazar-forma) | Crea una forma libre que sigue el contorno de las celdas con color de fondo |
| [Fijar valores UDF](utilities.md#fijar-valores-udf) | Convierte las fórmulas con funciones EGTools en valores y conserva la fórmula original como nota (para distribución) |
| [Aplicar CheckBox](utilities.md#aplicar-checkbox) | Aplica un formato de celda que muestra las entradas 0/1 como ✅/⬜ |
| [Recalcular todo](utilities.md#recalcular-todo) | Fuerza el recálculo completo de todos los libros abiertos |

## Grupo: EGTools

Botón dividido que reúne la conversión en bloque de nombres de función en las fórmulas, la configuración SMTP, la administración de claves API y la información del complemento (acción predeterminada = Acerca de EGTools++).

| Menú | Descripción |
|---|---|
| [Aplicar compatibilidad](apply-compat.md) | Convierte en bloque los tokens de funciones modernas/nativas (`_xlfn.*`) a los nombres de función compatibles de EGTools |
| [Restaurar funciones nativas](restore-native.md) | Restaura en bloque los nombres de función de EGTools a las funciones nativas integradas |
| [Configuración SMTP](smtp-settings.md) | Registra el servidor SMTP para el envío de correo de Combinar correspondencia (la contraseña no se guarda) |
| [Administrar claves API](api-keys.md) | Registra/elimina claves de API públicas (búsqueda de direcciones · data.go.kr · vworld — la clave de data.go.kr se comparte entre festivos, calendarios y registro de empresas) |
| Acerca de EGTools++ | Muestra la versión, la página de distribución y las bibliotecas base; si hay una nueva versión en GitHub Release, avisa y ofrece el enlace a la página de distribución |

## Menú contextual de celda

Además de la cinta, también se añaden elementos al menú contextual de las celdas. Detalles: [Menú contextual de celda](context-menu.md)

| Elemento | Descripción |
|---|---|
| [Selector de fecha](context-menu.md#selector-de-fecha) | Muestra un calendario junto a la celda e introduce en ella la fecha en la que se haga clic |
| [Abrir URL](context-menu.md#abrir-url) | Abre en el navegador predeterminado la URL http/https contenida en la celda |

## CheckBox en la pestaña Insertar

También se añade el botón [Aplicar CheckBox](utilities.md#aplicar-checkbox) a la pestaña `Insertar` de Excel.
No se muestra en Microsoft 365, que ya dispone de casillas de verificación nativas.

<!-- Las funciones que se añadan en el futuro se incorporarán aquí como nuevas secciones de grupo -->

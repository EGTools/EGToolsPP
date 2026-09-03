# Calendario

**Ubicación**: pestaña `EGTools` de la cinta → grupo `Calendario`

Grupo de funciones que crea hojas de calendario y de agenda. Se compone de 2 botones
divididos: **Calendario** (acción predeterminada = Mensual; menú: Mensual · Anual A · Anual B) y
**Agenda** (acción predeterminada = Diaria B; menú: Semanal · Diaria A · Diaria B).

Aspectos comunes:

- La ventana de entrada de fecha/año-mes se rellena de forma predeterminada con hoy
  (el mes actual) y admite `-`, `/` y `.` como separadores. Si se cancela o se introduce
  un valor que no puede interpretarse, finaliza silenciosamente.
- La nueva hoja se añade **después** de la hoja activa y se aplican a la vez la
  configuración de impresión (orientación del papel + ajuste a 1 página), la ocultación
  de las líneas de cuadrícula y la inmovilización de paneles.

## Mensual

Al introducir el año y el mes (`YYYY-MM`) crea una hoja de calendario mensual con una
cuadrícula de 6 semanas (42 días).

- En la primera columna de cada semana se muestra el **número de semana** con el formato
  `W##` (semana del año calculada con base en el miércoles).
- Los **festivos** muestran su nombre en la casilla junto a la fecha y todo el bloque
  del día se resalta con fondo amarillo claro + fuente roja.
- Los días **1, 11 y 21 del calendario lunar** llevan la anotación lunar `(MM/dd)`.
- Las fechas de los meses anterior y siguiente se muestran en gris como `M/D`
  (en enero se refleja también el festivo de Navidad del año anterior y en diciembre
  el de Año Nuevo del año siguiente), y las semanas que no pertenecen al mes se borran.
- Domingo en rojo, sábado en azul. Se configura la impresión horizontal en 1 página.

## Anual A / Anual B

Al introducir el año (2000–2050) crea una hoja de calendario anual con los 12 meses
en una sola página.

- **Anual A**: formato vertical de 3 columnas × 4 filas (impresión vertical) /
  **Anual B**: formato horizontal de 4 columnas × 3 filas (impresión horizontal).
  Solo cambia la disposición; el contenido es el mismo.
- Las fechas de **festivos** se resaltan con fondo amarillo + fuente roja, y el nombre
  del festivo se añade como **nota de celda**.
- Los títulos de mes llevan fondo verde; domingo en rojo y sábado en azul.

## Semanal

Al introducir una fecha (`YYYY-MM-DD`) crea una hoja de agenda semanal de la semana
(de domingo a sábado) a la que pertenece esa fecha.

- La fila de título muestra el año y el mes, las fechas de inicio y fin de la semana
  y el número de semana.
- En la columna de cada día se disponen la fecha, el calendario lunar (`MM/dd`),
  el nombre del festivo (resaltado) y la casilla de contenido de la agenda; debajo
  continúan las casillas de agenda horaria con un eje de **08:00 a 21:00 en intervalos
  de 30 minutos**.
- Domingo en rojo, sábado en azul; impresión horizontal en 1 página.

## Diaria A / Diaria B

Al introducir una fecha (`YYYY-MM-DD`) crea una hoja de agenda de un día.
El nombre de la hoja es `MMdd` (si está duplicado, `MMdd_1`, …).

- Se compone de **tareas principales** (columna de casillas de verificación + contenido
  de la tarea), **agenda horaria** (de las 08 a las 21 h) y **notas de trabajo**.
  Diaria A es de 2 secciones (tareas y agenda horaria + columna de notas) y Diaria B
  es un formato horizontal de 3 secciones (tareas | horario | notas).
- La columna de casillas de verificación lleva una validación 0/1 + el formato de
  visualización ✅/⬜, de modo que el estado de finalización puede totalizarse como
  números (mismo formato que [Aplicar CheckBox](utilities.md#aplicar-checkbox)).
- **Traslado automático de tareas pendientes**: de la hoja de agenda diaria anterior
  del libro (la de fecha pasada más reciente), las tareas sin marcar (vacías/0) se
  traen como tareas principales de la nueva agenda. Las hojas de agenda diaria se
  identifican mediante una nota oculta en la celda A1 (`EGCalendar:fecha`).
- El título muestra la fecha, el día de la semana y el calendario lunar; si es festivo,
  el nombre aparece bajo el título.
- Se configura la impresión horizontal en 1 página.

## Acerca de los festivos

- La base de los festivos es el **cálculo integrado**: incluye los festivos de calendario
  solar y lunar, el Día del Trabajo y las **reglas de festivos sustitutivos** (según el
  año de entrada en vigor). Comparte el mismo cálculo que la función
  [KOREANHOLIDAYS](../functions/korea/KOREANHOLIDAYS.md).
- A esto se **suma de forma predeterminada la API de festivos**. Con la clave de
  data.go.kr guardada ([EGTools → Administrar claves API](api-keys.md)) se consulta la
  información de días especiales y se reflejan en el calendario elementos que el cálculo
  integrado no incluye, como **festivos temporales y días electorales**. La misma regla
  se aplica a **todas las hojas**: Mensual, Anual A/B, Semanal y Diaria A/B.
- Si no hay conexión a Internet, no hay clave registrada, o la clave está caducada o
  fue rechazada, la hoja se crea **solo con el cálculo integrado** y, al terminar el
  comando, se muestra **una sola vez** el siguiente aviso (junto con el motivo del fallo).

  > La API de festivos no estaba disponible; se usó el cálculo integrado.
  > Los festivos temporales y días electorales no aparecen.

  La creación de la hoja no se interrumpe. Si recibió el aviso, registre o renueve la
  clave y ejecute de nuevo para que los festivos temporales queden reflejados.
- Los resultados de la consulta se almacenan en caché por año, de modo que crear varios
  calendarios del mismo año solo genera una consulta (los fallos no se almacenan en
  caché, así que al corregir la clave el cambio se aplica de inmediato).
- La anotación lunar se calcula con la tabla lunar integrada (igual que la función
  [TOLUNAR](../functions/korea/TOLUNAR.md)) y no se muestra en las fechas fuera del rango admitido.

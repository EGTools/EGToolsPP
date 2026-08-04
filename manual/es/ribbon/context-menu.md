# Menú contextual de las celdas

**Ubicación**: clic derecho sobre un área de celdas → menú contextual

Al cargar EGTools++ se añaden elementos al menú contextual de las celdas.
Aparecen en la misma posición tanto en celdas normales como en rangos de tabla
(lista) y en la vista Diseño de página.

## Selector de fecha

Muestra un calendario junto a la celda sobre la que hizo clic derecho para elegir
una fecha e introducirla. En lugar de escribir la fecha a mano, puede insertarla
con dos clics.

1. Haga clic derecho en la celda y pulse **Selector de fecha**; el calendario
   aparece a la derecha de la celda activa (si sobrepasa el borde de la pantalla,
   se recoloca a la izquierda). Si el valor actual de la celda es una fecha,
   empieza en **el mes de esa fecha seleccionada**; si no, empieza en hoy.
2. Use `◀◀`/`◀`/`▶`/`▶▶` en la parte superior para moverse por años/meses.
   **Al hacer clic en una fecha, esa fecha se introduce en la celda** y el
   calendario se cierra (también se puede hacer clic en los días del mes
   anterior/siguiente, mostrados en gris).
3. **Al hacer clic en la línea inferior "Hoy: yyyy-mm-dd" se va a la fecha de hoy**.
   Si cancela con `Esc` o cerrando la ventana (X), la celda no cambia.

- Los domingos y **los festivos se muestran en rojo**, los sábados en azul y el día
  de hoy con un borde. Al situar el ratón sobre una casilla de festivo, el nombre
  del festivo se muestra en la parte inferior.
- Los festivos se basan en el cálculo integrado; si hay una clave de data.go.kr
  registrada, también se reflejan los **festivos temporales** — consulte
  [Calendario/Agenda — Acerca de los festivos](calendar.md#acerca-de-los-festivos).

## Abrir URL

Busca en el contenido de la celda seleccionada una URL que empiece por `http://` o
`https://` y la abre en el navegador predeterminado.

- Aunque la URL esté mezclada con otros caracteres dentro del texto de la celda,
  se localiza y se abre la primera URL.
- Si hay varias celdas seleccionadas, se actúa sobre la **primera celda (superior
  izquierda)**.
- Si la celda no contiene ninguna URL, no se hace nada.

Permite abrir con dos clics una URL de texto normal sin formato de hipervínculo,
sin necesidad de entrar en el modo de edición de la celda ni de crear un
hipervínculo.

---

En el futuro se podrán añadir progresivamente más elementos al menú contextual.

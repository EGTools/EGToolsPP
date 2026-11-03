# Grupo Alineación de la pestaña Inicio (Centrar en la selección)

**Ubicación**: pestaña `Inicio` de Excel → grupo `Alineación` (sustituido por EGTools++) → `Centrar en la selección`, bajo el botón de combinar

Al cargar EGTools++, el grupo **Alineación** predeterminado de la pestaña `Inicio` de Excel se oculta y en
su lugar se reconstruye un grupo **Alineación** con la misma disposición. Un complemento no puede insertar
controles dentro de un grupo predeterminado de Office, por lo que se recrea el grupo completo; la única
diferencia es un botón adicional, **Centrar en la selección**, situado bajo el botón de combinar. Este botón
cambia solo la alineación horizontal a "Centrar en la selección" sin combinar celdas, de modo que un título
puede quedar centrado sobre varias columnas mientras la ordenación, los filtros y las referencias de las
fórmulas siguen funcionando.

La sustitución se realiza **solo cuando** el desplegable de combinar de Excel aún no tiene el elemento
nativo "Centrar en la selección" (idMso `CenterAcrossSelection`). En un Excel que ya lo tenga, el grupo
`Alineación` predeterminado se deja intacto y el botón de EGTools++ no aparece: use el elemento propio de
Excel.

## Disposición del grupo

| Posición | Control | Origen |
|---|---|---|
| Izquierda, fila 1 | Alinear arriba · Alinear en el medio · Alinear abajo · Orientación | Integrado en Excel |
| Izquierda, fila 2 | Alinear a la izquierda · Centrar · Alinear a la derecha · Disminuir sangría · Aumentar sangría | Integrado en Excel |
| Derecha, fila 1 | Ajustar texto | Integrado en Excel |
| Derecha, fila 2 | Combinar y centrar (desplegable) | Integrado en Excel |
| Derecha, fila 3 | **Centrar en la selección** | Añadido por EGTools++ |
| Iniciador de cuadro de diálogo | Formato de celdas: Alineación | Integrado en Excel |

Los controles integrados son los propios controles del grupo predeterminado oculto, colocados de nuevo,
así que su comportamiento, etiquetas y atajos de teclado son los de Excel.

## Procedimiento — Centrar en la selección

1. Seleccione, dentro de una fila, el rango de columnas sobre el que extender el contenido, con la celda
   que lo contiene en el extremo izquierdo (por ejemplo `A1:F1`). Si selecciona varias filas, cada una se
   trata por separado.
2. Pulse **Centrar en la selección** en la pestaña `Inicio` → grupo `Alineación`.
3. La alineación horizontal de la selección se establece en **Centrar en la selección**
   (`xlHAlignCenterAcrossSelection`): en cada fila el contenido se muestra centrado sobre el tramo de
   celdas vacías situado a su derecha. No se combina ninguna celda.
4. Al pulsar de nuevo sobre el mismo rango, la alineación horizontal vuelve a **General**
   (`xlHAlignGeneral`): el botón funciona como conmutador.

## Reglas

| Situación | Comportamiento |
|---|---|
| Toda la selección ya está en "Centrar en la selección" | Restablece la alineación **General** |
| Mixta (solo una parte la tiene) o no aplicada | Aplica **Centrar en la selección** a toda la selección |
| Hay seleccionado un objeto que no es una celda (forma, gráfico, imagen, …) | No hace nada |
| Dirección vertical | Solo cambia la alineación horizontal; el contenido no se centra verticalmente (así está definido "Centrar en la selección" en Excel). Use Alinear arriba/en el medio/abajo para la posición vertical |
| Deshacer | Es un comando del complemento, por lo que no se registra en Deshacer de Excel (`Ctrl+Z`). Para revertirlo, pulse el botón de nuevo o cambie la alineación horizontal en Formato de celdas |

## Notas

- Microsoft está desplegando gradualmente un elemento "Centrar en la selección" dentro del desplegable de
  combinar, empezando por el canal Insider de Microsoft 365 (What's New de septiembre de 2026). EGTools++
  ofrece la función en la pestaña `Inicio` para que también esté disponible en Excel 2010–2024 y en las
  compilaciones de Microsoft 365 que aún no la han recibido. En un Excel que ya tenga el elemento nativo,
  EGTools++ lo detecta al cargarse y no sustituye el grupo `Alineación`, de modo que ambos nunca aparecen
  juntos: use el elemento propio de Excel. La comprobación se hace una sola vez al cargar EGTools++, así
  que si un cambio de canal o compilación añade o quita el elemento nativo, el cambio se aplica tras
  reiniciar Excel.
- El mismo efecto se obtiene desde el cuadro de diálogo `Formato de celdas` (`Ctrl+1`) → pestaña
  `Alineación` → Horizontal `Centrar en la selección`. Este botón es un atajo de un clic para aplicarlo o
  quitarlo.
- Para que varias celdas con valores solo parezcan combinadas, consulte
  [Combinar conservando valores](merge-split.md#combinar-conservando-valores); ese comando aplica el
  formato de combinación, mientras que este botón solo cambia la alineación horizontal.
- Al descargar EGTools++, el grupo sustituto desaparece y vuelve a mostrarse el grupo `Alineación`
  predeterminado de Excel.

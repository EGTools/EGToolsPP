# Herramientas

**Ubicación**: pestaña `EGTools` de la cinta → grupo `Herramientas`

Botón dividido que reúne funciones auxiliares de uso frecuente. Al pulsar el botón
directamente se ejecuta la acción predeterminada, **Recalcular todo**; al pulsar la
flecha se despliega un menú con 10 elementos.

## Eliminar estilos

Elimina de una vez los **estilos de celda personalizados** acumulados en el libro
activo. Se usa cuando los estilos, multiplicados como una bola de nieve al copiar desde
otros documentos, hacen pesado el documento. En la ventana de confirmación se elige
el alcance de la eliminación.

| Opción | Acción |
|---|---|
| Sí | Conserva solo los estilos integrados (Built-In) y elimina todos los personalizados |
| No | Conserva los estilos integrados + los **realmente en uso** y elimina el resto |
| Cancelar | Cancela la operación |

- La comprobación de "en uso" recorre celda a celda el área usada de todas las hojas,
  por lo que **puede tardar en documentos grandes** (el progreso se muestra en la barra
  de estado).
- Los estilos bloqueados se desbloquean antes de eliminarlos; al terminar se informa
  del número de eliminados/fallidos.

## Eliminar nombres

Elimina, de entre los nombres definidos del libro activo, los que cumplen lo siguiente.

| Objetivo de eliminación | Condición |
|---|---|
| Nombres con referencias rotas | El destino de la referencia (RefersTo) contiene `#NAME?` / `#REF!` / `#N/A` |
| Restos de macros | El destino de la referencia contiene `BlankMacro1` |
| Nombres ocultos | Nombres no visibles en el Administrador de nombres (Visible=False) |

Los nombres internos de Excel (formato `_xl*.*`, p. ej. los de área de impresión o
filtros) nunca se tocan. Si hay muchos nombres, el progreso se muestra en la barra de
estado y al terminar se informa del número de eliminados (si no hay objetivos, se avisa
de que no hay ninguno).

## Limpiar errores

En el área usada de la hoja activa, envuelve las **fórmulas que están produciendo
valores de error** con `=IFERROR(fórmula_original,"")` para ocultar la indicación de
error. Las fórmulas de matriz CSE (`{=…}`) se envuelven como matriz completa.
Las celdas sin fórmula (constantes de error introducidas directamente) no se tocan;
al terminar se informa del número de correcciones.

- La causa del error no se corrige: simplemente se **muestra como cadena vacía**.
  Es adecuado para limpiar justo antes de imprimir o distribuir el documento.
- Si tras envolverla la fórmula sigue produciendo errores y se ejecuta repetidamente,
  los IFERROR pueden quedar anidados.

## Limpiar texto vacío

Convierte en celdas realmente vacías (borrando su contenido) las celdas que **solo
contienen una cadena de longitud 0 ("")**. Las celdas `""` que dejan los resultados de
fórmulas o los datos externos parecen vacías, pero no se tratan como tales en ISBLANK,
COUNTA ni en la detección de la última celda (Ctrl+End); esta función las limpia.

1. Seleccione un rango y ejecute. **Si solo hay una celda seleccionada, el objetivo es
   toda el área usada de la hoja.**
2. Ventana de confirmación "¿Excluir celdas con fórmulas?" — **Sí** = se conservan las
   celdas con fórmulas que devuelven `""` / **No** = también se borra el contenido de
   las celdas con fórmulas (la fórmula se elimina) / **Cancelar** = se interrumpe.
3. Al terminar se informa del número de celdas limpiadas. Las fórmulas de matriz CSE
   se borran como matriz completa.

## Ordenar notas

Mueve **todas las notas** de la hoja activa justo al lado derecho de su celda y ajusta
automáticamente su tamaño al contenido. Si la celda está combinada, la nota se mueve al
lado derecho del área combinada. Al terminar se informa del número de notas ordenadas.

## Cambiar color de bordes

Al elegir un color en el cuadro de diálogo de selección de color (cancelar finaliza),
cambia de una vez el color de todos los bordes de la selección
— superior/inferior/izquierdo/derecho + las 2 diagonales.
**El tipo de línea (continua, punteada, grosor) se mantiene tal cual**: solo cambia el
color, y no se crean bordes nuevos en las direcciones que no los tienen.

## Trazar forma

Crea una forma libre (Freeform) siguiendo el contorno de las **celdas con color de
fondo** dentro de la selección. Se usa para extraer como una sola forma mapas o zonas
dibujadas coloreando celdas.

1. Seleccione un rango que incluya celdas coloreadas y ejecute. Si no hay celdas
   coloreadas, se muestra el aviso "Seleccione un rango con celdas coloreadas."
   y finaliza.
2. **El fondo blanco (incluido sin relleno)** se considera fondo; las celdas con
   cualquier otro color se consideran área de la forma.
3. Se rastrea el contorno del primer **bloque conectado** de celdas coloreadas que se
   encuentre, siguiendo el ancho y alto reales de las celdas, se crea la forma y se le
   aplica un preajuste de estilo.

- Si hay varios bloques separados, solo el primero se convierte en forma. Para el
  resto, divida el rango y ejecute de nuevo.
- Si falla la creación de la forma, se muestra el aviso "Falló la conversión de la forma."

## Fijar valores UDF

Convierte en **valores calculados** las fórmulas que contienen funciones EGTools
(incluidas las variantes con prefijo `EG.`/`x`). Es una función pensada para la
distribución: al enviar el documento a un entorno sin el complemento, los valores
permanecen en lugar de que las fórmulas se rompan con `#NAME?`.

1. **Escaneo** — recorre en modo de solo lectura las fórmulas de todas las hojas y
   muestra en una ventana de confirmación el número de objetivos (si no hay ninguno,
   se avisa y finaliza).
2. **Comprobación de guardado + copia de seguridad** — si el libro no está guardado,
   se pregunta si desea guardarlo, y justo antes de aplicar se guarda automáticamente
   un archivo de copia de seguridad (SaveCopyAs). Si la copia de seguridad falla,
   no se continúa.
3. **Fijación de valores** — las fórmulas objetivo se convierten en valores.
   Las matrices CSE se convierten como matriz completa y las fórmulas derramadas,
   como todo el rango derramado. La fórmula original se conserva en la nota de la celda
   con el formato `EgUdfRemoved: <fórmula>` (si ya existe una nota, se añade al final).
4. Al terminar se informa del número de celdas fijadas y de la ruta del archivo de
   copia de seguridad.

- La lista de funciones objetivo se gestiona desde una única fuente junto con la lista
  de registro de funciones del complemento, por lo que siempre coincide.
- Mientras que [Aplicar compatibilidad](apply-compat.md) **convierte los nombres** de
  las funciones para mantener vivas las fórmulas, Fijar valores UDF **convierte las
  fórmulas en valores**.

## Aplicar CheckBox

Hace que la selección se muestre como casillas de verificación mediante entradas 0/1.

1. Aplica a la selección una **validación de datos que solo admite 0 o 1**.
2. Establece el formato de número `[=1]"✅";[=0]"⬜";""`
   — al introducir 1 se muestra ✅, con 0 se muestra ⬜ y en cualquier otro caso queda vacío.
3. Las celdas vacías se rellenan con 0 y se centra el contenido.

Los valores de celda siguen siendo números 0/1, por lo que pueden totalizarse
directamente con `SUM`, `COUNTIF`, etc.

## Recalcular todo

Fuerza el recálculo completo de **todos los libros abiertos** (`CalculateFull`).
Se usa para recalcular desde cero las celdas que el cálculo automático haya pasado por
alto, las funciones volátiles y las referencias externas.

> El registro y la eliminación de claves para las funciones de API públicas se han
> trasladado a [Administrar claves API, en el grupo EGTools](api-keys.md).

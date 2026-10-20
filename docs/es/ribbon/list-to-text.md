# Celdas de lista/matriz → texto

**Ubicación**: pestaña `EGTools` de la cinta → grupo `EGTools` → menú desplegable del botón dividido

Las **listas en celdas (Lists in Cells)** de la beta de Microsoft 365 (versión 2610, compilación
20520 o posterior) se guardan como valores enriquecidos, por lo que en versiones antiguas de Excel
(incluido el Microsoft 365 normal) solo se ven como `#UNKNOWN!` (`#VALUE!` en Excel 2016) y no es
posible leer sus elementos. Este comando lee directamente el richData del archivo guardado y
reescribe cada celda de lista con la **notación de texto** descrita abajo. El texto convertido se
puede evaluar con las funciones de compatibilidad [HAS](../functions/microsoft-365/HAS.md) ·
[HASANY](../functions/microsoft-365/HASANY.md) · [HASALL](../functions/microsoft-365/HASALL.md) ·
[FLATTEN](../functions/microsoft-365/FLATTEN.md) (en Microsoft 365 se registran con el prefijo `EG.`,
p. ej. `EG.HAS`, para convivir con las nativas).

## Notación de texto

| Forma de la lista | Notación | Ejemplo |
|---|---|---|
| Lista de una columna (sin anidamiento) | Separada por comas — igual que el formato conciso de `ARRAYTOTEXT`. Un elemento que contenga `,` `;` `{` `}` `"` o espacios al principio/final, o un texto que se leería como número/TRUE/error, se entrecomilla como `"…"` (las `"` internas se duplican a `""`) | `Mon, Tue, Wed` / `"1,000", "TRUE"` |
| Matriz 2-D o anidada | Forma de constante matricial — el formato estricto de `ARRAYTOTEXT` (todo el texto entrecomillado, sin espacios, `;` filas, `,` columnas) ampliado con llaves anidadas | `{"a",1;"b",2}` / `{"a",{"b","c"},"d"}` |
| Lista con un único elemento numérico/lógico | Se escribe como el **valor** numérico/lógico en sí, no como texto | `3`, `TRUE` |

- En el Excel beta, al anteponer `=` al texto en forma de constante matricial vuelve a ser directamente una fórmula de matriz en celda.
- Las funciones de la familia HAS leen ambas notaciones. Los elementos de una lista separada por comas o
  dentro de `{…}` se leen como número si tienen forma numérica, como lógico si son TRUE/FALSE y como error
  si son nombres de error como `#N/A`; un texto único sin comas ni llaves no se clasifica y queda como un
  solo elemento de texto. Si el texto no se puede analizar (p. ej. comillas sin cerrar), se trata como un
  único elemento sin producir error.

## Procedimiento

1. **Lectura del archivo** — Se descomprime una copia del libro actual (SaveCopyAs) en una carpeta
   temporal y se lee su richData (valores enriquecidos de lista/matriz). Las ediciones sin guardar
   también están en la copia, así que se reflejan igualmente.
2. **Comprobación de objetivos** — Si no hay celdas de lista, se avisa y se termina.
3. **Ventana de confirmación** — Informa de que se convertirán N celdas de lista con valores. Si hay M
   celdas de lista que son resultado de fórmula, pregunta si también deben fijarse a texto:
   **Sí** = todas / **No** = solo las celdas de valor / **Cancelar**. (Las celdas de lista con fórmula se
   guardan como `_xlfn._ARRAY(...)`, que en versiones antiguas es `#NAME?` de todos modos. Si elige No y
   no hay celdas de valor, se termina.)
4. **Comprobación de guardado + copia de seguridad** — Si el libro no está guardado, se pregunta si desea
   guardarlo (si rechaza o cancela, se aborta) y, justo antes de aplicar, se guarda automáticamente un
   archivo de copia de seguridad (en la misma carpeta que el libro,
   `nombreoriginal_egbak_<marca de tiempo>.<extensión>`, el mismo procedimiento que
   [Aplicar compatibilidad](apply-compat.md)). Si la copia de seguridad falla, no se continúa.
5. **Aplicación** — Cada celda de lista se reescribe como texto (o como valor) según las reglas de abajo.
6. **Ventana de resultado** — Muestra el número de celdas convertidas, el de celdas omitidas (si las hay)
   y la ruta del archivo de copia de seguridad.

## Reglas de aplicación

| Objetivo | Tratamiento |
|---|---|
| Celda de lista con valor | Solo se reescriben las celdas cuyo valor actual es un error (`#UNKNOWN!`/`#VALUE!`). Las celdas editadas después de guardar (que ya no son error) se omiten |
| Celda de lista con fórmula | Se convierte solo si eligió **Sí**. La fórmula original se conserva en la nota de la celda como `EgListRemoved: <fórmula>` (si ya hay una nota, se añade a continuación) |
| Área de matriz CSE | Primero se fija toda el área a valores y luego se escribe el texto en esa celda |
| Texto que Excel interpretaría automáticamente como fórmula/número/TRUE | Se cambia el formato de la celda a Texto (`@`) antes de escribir |
| Lista con un único elemento numérico/lógico | Se escribe como el valor numérico/lógico en sí, no como texto |
| Valores enriquecidos que no son `_array` (imágenes en celda, tipos de datos vinculados, etc.) | No son objetivo |

## Notas

- El comando también funciona en el propio Excel beta, por lo que sirve para convertir de antemano antes de compartir un libro.
- Tras la conversión, HAS / HASANY / HASALL / FLATTEN dan los mismos resultados. Si el libro se guardó en
  el Excel beta con fórmulas como `=HAS(...)`, [Aplicar compatibilidad](apply-compat.md) convierte los
  tokens `_xlfn.HAS` y `_xlfn._xlws.FLATTEN` a los nombres de función compatibles de EGTools.
- Orden recomendado: ① Celdas de lista/matriz → texto, ② [Aplicar compatibilidad](apply-compat.md) (o
  [Restaurar funciones nativas](restore-native.md) en un Microsoft 365 no beta). Respetar este orden es lo que
  hace que fórmulas como `=HAS(...)` devuelvan el resultado correcto en lugar de `#VALUE!`. Si respondió
  **No** y dejó sin tocar las celdas de lista con fórmula, Aplicar compatibilidad convierte sus fórmulas
  `_xlfn._ARRAY(...)` a `ARRAYTOTEXT(…,1)`.
- Para deshacer, abra el archivo de copia de seguridad. En las celdas de fórmula, la fórmula original
  también puede consultarse en la nota `EgListRemoved:`.

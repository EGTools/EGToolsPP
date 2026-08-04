# Celdas visibles

**Ubicación**: cinta, pestaña `EGTools++` → grupo `Celdas visibles`

Grupo de funciones para copiar y pegar en rangos con filtros o filas/columnas ocultas.
Se compone de 2 botones divididos — **Copiar** (acción predeterminada = Copiar celdas visibles) y
**Pegar** (acción predeterminada = Pegar valores).

## Copiar celdas visibles

Selecciona **solo las celdas visibles** del rango elegido, las recuerda como origen de copia y las copia al portapapeles.

1. Seleccione el rango que desea copiar y pulse el botón. Si la selección no es
   un rango de celdas, la operación termina silenciosamente sin hacer nada.
2. De la selección se toman **solo las celdas visibles**, excluyendo las ocultas
   por filtros u ocultación, y se recuerdan como origen de copia el nombre del
   libro y de la hoja junto con la dirección del área.
3. La misma área se copia también al portapapeles de Excel (con el borde en
   movimiento), lista para usarse de inmediato con el pegado normal.

## Copiar todo

**Solo recuerda** como origen de copia la selección **completa (incluidas las celdas ocultas)**.
No copia nada al portapapeles, por lo que no se aprecia ningún cambio en pantalla.

## Pegar valores / Pegar fórmulas / Pegar todo

Pega el origen de copia recordado, en orden, **solo en las celdas visibles** de la selección actual.
Los 3 menús se diferencian únicamente en lo que pegan — **Pegar valores** (resultados calculados,
acción predeterminada del botón dividido), **Pegar fórmulas** y **Pegar todo** (valores, fórmulas y formato en bloque).

### Procedimiento

1. Ejecute primero [Copiar celdas visibles](#copiar-celdas-visibles) o
   [Copiar todo](#copiar-todo) para registrar el origen de copia.
2. Seleccione el rango de destino y pulse el menú. En el destino se omiten las
   celdas ocultas por filtros u ocultación y se rellenan **solo las celdas
   visibles**, siguiendo el orden del origen.
   Si el destino es una sola celda, se pega desde esa celda hacia abajo y hacia la derecha.
3. Si el origen tiene varias celdas, internamente se crea un **libro temporal**
   para adaptar la forma a la posición de las filas/columnas ocultas del destino
   antes de pegar. El libro temporal se cierra automáticamente sin guardarse
   (sin riesgo de que quede sin limpiar).
4. Durante este proceso, el **autofiltro de la hoja de origen se desactiva
   momentáneamente y se restaura con sus mismas condiciones**. Si el origen es
   una sola celda, no hay pasos intermedios y su valor se rellena directamente
   en todas las celdas visibles del destino.

### Casos en los que no se puede ejecutar

| Situación | Mensaje |
|---|---|
| No hay origen de copia recordado | "Ejecute primero [Copiar celdas visibles]." |
| El libro/hoja de origen se cerró o cambió de nombre | "El libro/hoja de origen ya no existe. Copie de nuevo." |

## Notas

- El origen se recuerda por el nombre del libro y de la hoja más la dirección.
  Si cierra el libro de origen o cambia el nombre de la hoja, deberá copiar de nuevo.
- Si el área visible copiada con filtros activos queda muy fragmentada (cuando la
  dirección del área supera los 255 caracteres), la reinterpretación del origen
  puede fallar. En ese caso también se mostrará un aviso pidiendo copiar de nuevo.

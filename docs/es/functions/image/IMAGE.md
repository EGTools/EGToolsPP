# IMAGE

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2024

Inserta una imagen en la celda desde una URL o ruta de archivo.

## Sintaxis

```
=IMAGE(source, [alt_text], [sizing], [height], [width])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| source | Obligatorio | URL o ruta del archivo de imagen |
| alt_text | Opcional | texto alternativo (accesibilidad) (opcional) |
| sizing | Opcional | 0 ajustar (mantener proporción), 1 rellenar, 2 tamaño original, 3 personalizado (opcional) |
| height | Opcional | alto en píxeles cuando sizing es 3 (opcional) |
| width | Opcional | ancho en píxeles cuando sizing es 3 (opcional) |

## Devuelve

Siempre devuelve una cadena vacía ("") como escalar y, al terminar el cálculo, inserta una forma de imagen flotante ajustada al tamaño de la celda de llamada (o del área combinada). Si source está vacío, devuelve #VALUE!; si no se puede determinar la dirección de la celda de llamada, devuelve #REF!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=IMAGE("https://www.google.com/favicon.ico")` |  | Depende de la imagen externa |

## Notas

- A diferencia de la función nativa de 365, el resultado no es un valor dentro de la celda: inserta una forma de imagen flotante (EG.Image_hoja_celda) sobre la celda y, al recalcular, reemplaza la forma existente con el mismo nombre.
- Los argumentos alt_text, height y width se aceptan por compatibilidad pero se ignoran; con resize_mode 0 (predeterminado) se mantiene la relación de aspecto.
- En las versiones modernas de Excel que incluyen la función IMAGE nativa, esta función se registra como EG.IMAGE.
- Una matriz en source devuelve #VALUE! (una imagen por celda); para varios elementos, copie la fórmula fila por fila.
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `IMAGE` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.IMAGE`.

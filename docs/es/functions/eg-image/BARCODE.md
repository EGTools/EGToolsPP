# BARCODE

**Categoría**: función exclusiva de EGTools

Codifica texto como código de barras e inserta la imagen en la celda que llama (los tipos 2D mantienen la proporción y se centran si la celda es mayor).

## Sintaxis

```
=BARCODE(texto, [tipo], [opcion])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto | Obligatorio | el texto a codificar |
| tipo | Opcional | QRCODE (predeterminado), CODE128, DATAMATRIX, EAN13/8, ITF, PDF417, AZTEC, CODE39/93, UPC-A/E, CODABAR, GS1-128/DM/QR |
| opcion | Opcional | número = margen (2D) o TRUE = mostrar texto debajo (1D) |

## Devuelve

Devuelve una cadena vacía ("") si tiene éxito, y la imagen del código de barras se inserta por separado ajustada al tamaño de la celda que llama. Devuelve #VALUE! si el texto está vacío, si no se reconoce el tipo o si el texto no se puede codificar en ese formato, y #REF! si no se puede determinar la celda que llama.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=BARCODE("12345678","CODE128")` |  | Insertar imagen Code 128 (inserción de imagen, depende del entorno) |
| `=BARCODE("https://example.com")` |  | Insertar el código QR predeterminado (inserción de imagen, depende del entorno) |

## Notas

- Tipos admitidos: CODE128, CODE39/93, EAN13/8, UPC-A/E, ITF, CODABAR, PDF417, QRCODE (predeterminado), DATAMATRIX, AZTEC, GS1-128/GS1DATAMATRIX/GS1QRCODE (también se aceptan códigos numéricos).
- El argumento de opción se interpreta como margen en módulos si es numérico (formatos 2D) y, si es TRUE, como mostrar el texto bajo las barras (formatos 1D).
- Los formatos 2D mantienen su proporción normalizada (módulos cuadrados), se ajustan dentro de la celda que llama y se centran; los formatos 1D llenan la celda.
- Una matriz en el texto devuelve #VALUE! (una imagen por celda); para varios elementos, copie la fórmula fila por fila.
- Compatibilidad: Excel 2010+. Se registra siempre como `BARCODE` en todas las versiones de Excel.

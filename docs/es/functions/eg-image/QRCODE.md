# QRCODE

**Categoría**: función exclusiva de EGTools

Genera una imagen de código QR en la celda que llama.

## Sintaxis

```
=QRCODE(texto, [margen], [gs1])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto | Obligatorio | el texto a codificar |
| margen | Opcional | módulos de margen (predeterminado 0) |
| gs1 | Opcional | TRUE = tratar como pares GS1 (AI)valor |

## Devuelve

Devuelve una cadena vacía ("") si tiene éxito, y la imagen del código QR se inserta por separado ajustada al tamaño de la celda que llama. Devuelve #VALUE! si el texto está vacío o si falla la conversión GS1 o la codificación, y #REF! si no se puede determinar la celda que llama.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=QRCODE("https://example.com")` |  | Insertar imagen de código QR (inserción de imagen, depende del entorno) |

## Notas

- gs1=TRUE se aproxima con separadores GS (0x1D); no es una codificación FNC1 completa. Si se necesita GS1 completo, use CODE128 (GS1-128).
- La imagen mantiene su proporción cuadrada dentro de la celda que llama.
- Una matriz en el texto devuelve #VALUE! (una imagen por celda); para varios elementos, copie la fórmula fila por fila.
- Compatibilidad: Excel 2010+. Se registra siempre como `QRCODE` en todas las versiones de Excel.

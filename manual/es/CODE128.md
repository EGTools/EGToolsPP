# CODE128

**Categoría**: función exclusiva de EGTools

Genera una imagen de código de barras Code 128 en la celda que llama.

## Sintaxis

```
=CODE128(texto, [mostrar_texto], [gs1], [tamano_texto])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto | Obligatorio | el texto a codificar |
| mostrar_texto | Opcional | TRUE = dibujar el texto bajo las barras |
| gs1 | Opcional | TRUE = GS1-128 con FNC1 |
| tamano_texto | Opcional | tamaño relativo de la etiqueta |

## Devuelve

Devuelve una cadena vacía ("") si tiene éxito, y la imagen del código de barras Code 128 se inserta por separado ajustada al tamaño de la celda que llama. Devuelve #VALUE! si el texto está vacío o si falla la conversión GS1 o la codificación, y #REF! si no se puede determinar la celda que llama.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=CODE128("ABC-1234",TRUE)` |  | Insertar con el texto mostrado bajo las barras (inserción de imagen, depende del entorno) |
| `=CODE128("(01)04012345678901(10)LOT42",FALSE,TRUE)` |  | Insertar en formato GS1-128 (FNC1) (inserción de imagen, depende del entorno) |

## Notas

- Con gs1=TRUE, el texto en formato (AI)valor se codifica como GS1-128 totalmente conforme con FNC1 (identificador de simbología ]C1).
- La imagen llena por completo la celda que llama (formato 1D).
- tamaño_texto es un valor relativo que se refleja en la altura de la banda de texto inferior.
- Compatibilidad: Excel 2010+. Se registra siempre como `CODE128` en todas las versiones de Excel.

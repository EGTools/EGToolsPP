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
| gs1 | Opcional | TRUE = GS1-128 (FNC1): notación (AI)valor o [AI]valor, o flujo sin formato separado por GS |
| tamano_texto | Opcional | tamaño relativo de la etiqueta |

## Devuelve

Devuelve una cadena vacía ("") si tiene éxito, y la imagen del código de barras Code 128 se inserta por separado ajustada al tamaño de la celda que llama. Devuelve #VALUE! si el texto está vacío o si falla la conversión GS1 o la codificación, y #REF! si no se puede determinar la celda que llama.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=CODE128("ABC-1234",TRUE)` |  | Insertar con el texto mostrado bajo las barras (inserción de imagen, depende del entorno) |
| `=CODE128("(01)04012345678901(10)LOT42",FALSE,TRUE)` |  | Insertar en formato GS1-128 (FNC1) (inserción de imagen, depende del entorno) |

## Notas

- Con gs1=TRUE, el texto con AI entre paréntesis, como "(01)04012345678901(10)LOT42", se codifica como GS1-128 (identificador de simbología ]C1): FNC1 en primera posición y separadores tras los AI de longitud variable según la norma. La notación con corchetes, como "[01]04012345678901[10]LOT42", y un flujo sin procesar leído por un escáner como `]C10104012345678901<GS>10LOT42` (`<GS>` = separador 0x1D — también se reconoce si se escribe literalmente `<GS>`; el identificador inicial ]C1 se ignora si está presente) producen el mismo símbolo; las infracciones de formato, longitud o dígito de control de los AI, los AI desconocidos y los caracteres no ASCII devuelven #VALUE!.
- La imagen llena por completo la celda que llama (formato 1D).
- tamaño_texto es un valor relativo que se refleja en la altura de la banda de texto inferior.
- Una matriz en el texto devuelve #VALUE! (una imagen por celda); para varios elementos, copie la fórmula fila por fila.
- Compatibilidad: Excel 2010+. Se registra siempre como `CODE128` en todas las versiones de Excel.

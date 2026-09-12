# DATAMATRIX

**Categoría**: función exclusiva de EGTools

Genera una imagen DataMatrix en la celda que llama (mantiene la proporción y se centra si la celda es mayor).

## Sintaxis

```
=DATAMATRIX(texto, [margen], [gs1])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto | Obligatorio | el texto a codificar |
| margen | Opcional | módulos de margen (predeterminado 0) |
| gs1 | Opcional | TRUE = codificar como GS1 (FNC1): notación (AI)valor o [AI]valor, o flujo sin formato separado por GS |

## Devuelve

Devuelve una cadena vacía ("") si tiene éxito, y la imagen DataMatrix se inserta por separado ajustada al tamaño de la celda que llama. Devuelve #VALUE! si el texto está vacío o si falla la conversión GS1 o la codificación, y #REF! si no se puede determinar la celda que llama.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=DATAMATRIX("LOT-2026-08")` |  | Insertar imagen DataMatrix (inserción de imagen, depende del entorno) |

## Notas

- Con gs1=TRUE, el texto con AI entre paréntesis, como "(01)04012345678901(10)LOT42", se codifica como GS1 DataMatrix (identificador de simbología ]d2): FNC1 va en la primera posición y tras los AI de longitud variable se insertan los separadores según la norma. La notación con corchetes, como "[01]04012345678901[10]LOT42", y un flujo sin procesar leído por un escáner como `]d20104012345678901<GS>10LOT42` (`<GS>` = separador 0x1D — también se reconoce si se escribe literalmente `<GS>`; el identificador inicial ]d2 se ignora si está presente) producen el mismo símbolo; los AI desconocidos devuelven #VALUE!.
- Los datos GS1 se validan en formato, longitud y dígito de control de cada AI (p. ej., el dígito de control del GTIN en (01)); las infracciones devuelven #VALUE! y solo se admite ASCII imprimible (el coreano, etc. devuelve #VALUE!). El texto normal se codifica en UTF-8 y, para caracteres fuera de Latin-1, se inserta automáticamente un ECI (compatible con la mayoría de los lectores).
- La imagen mantiene la proporción normalizada del símbolo (módulos cuadrados) al ajustarse dentro de la celda que llama y se centra en el espacio sobrante; margin se expresa en módulos (predeterminado 0).
- Una matriz en el texto devuelve #VALUE! (una imagen por celda); para varios elementos, copie la fórmula fila por fila.
- Compatibilidad: Excel 2010+. Se registra siempre como `DATAMATRIX` en todas las versiones de Excel.

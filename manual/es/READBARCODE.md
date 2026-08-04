# READBARCODE

**Categoría**: función exclusiva de EGTools

Decodifica un código de barras desde un archivo de imagen o URL.

## Sintaxis

```
=READBARCODE(origen, [tipo_resultado])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| origen | Obligatorio | ruta de archivo local o URL http(s) |
| tipo_resultado | Opcional | 1 texto (predeterminado), 2 texto de código, 3 formato, 4 id de simbología, 5 hex; admite matriz {1,3} |

## Devuelve

Devuelve el contenido del código de barras leído de la imagen con la misma forma que el argumento tipo_resultado (si es una matriz como {1,3}, se derrama). Devuelve #VALUE! si el origen está vacío o si un valor de tipo_resultado no es válido, #NULL! si falla la descarga desde la URL y #N/A si el archivo no existe, si no se puede leer la imagen o si no se encuentra ningún código de barras.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=READBARCODE("C:\barcodes\qr.png")` |  | Texto del código de barras de la imagen (el resultado depende del contenido del archivo) |
| `=READBARCODE("https://example.com/code.png",{1,3})` |  | Devolver el texto y el formato juntos (depende de la red) |

## Notas

- Lee los formatos PNG/JPG/BMP/GIF compatibles con WIC, y admite tanto rutas de archivo locales como URL http(s).
- tipo_resultado: 1 texto (predeterminado), 2 texto de código (los caracteres de control se muestran como <GS>, etc.), 3 formato, 4 identificador de simbología, 5 HEX sin procesar.
- También intenta leer imágenes giradas o de baja calidad (TryHarder/TryRotate), pero lee solo un código de barras.
- Compatibilidad: Excel 2010+. Se registra siempre como `READBARCODE` en todas las versiones de Excel.

# TRIMENDS

**Categoría**: función exclusiva de EGTools

Elimina solo los espacios al inicio y final (conserva los interiores).

## Sintaxis

```
=TRIMENDS(texto, [opciones])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto | Obligatorio | texto o rango |
| opciones | Opcional | extras a eliminar: +1 caracteres de control, +2 nbsp, +4 Unicode invisible (por defecto 0) |

## Devuelve

Devuelve el texto con ambos extremos limpios (escalar); si se especifica una matriz o un rango, cada elemento se procesa por separado y el resultado se derrama. Devuelve #VALUE! si las opciones están fuera del rango 0–7, y las celdas vacías se convierten en cadenas vacías.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=TRIMENDS("  a  b  ")` | a  b | elimina solo los espacios de los extremos |
| `=TRIMENDS("a",8)` | #VALUE! | opción fuera de rango |

## Notas

- Incluso con la opción 0, los espacios, tabulaciones, CR y LF siempre se eliminan de ambos extremos.
- A diferencia de TRIM, los espacios del interior de la cadena se conservan.
- Las opciones se suman: +1 caracteres de control, +2 NBSP, +4 espacios Unicode invisibles (ZWSP, espacio de ancho completo, etc.).
- Compatibilidad: Excel 2010+. Se registra siempre como `TRIMENDS` en todas las versiones de Excel.

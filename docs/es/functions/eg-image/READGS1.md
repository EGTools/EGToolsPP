# READGS1

**Categoría**: función exclusiva de EGTools

Divide un texto GS1 en filas de (AI) y valor.

## Sintaxis

```
=READGS1(texto_gs1, [concatenar])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto_gs1 | Obligatorio | forma (AI)valor o cadena decodificada de un código |
| concatenar | Opcional | TRUE = unir en una sola cadena (AI)valor |

## Devuelve

Descompone el texto GS1 en una matriz de N×2 con (AI) y valor, y se derrama (con concatenar=TRUE devuelve una sola cadena en formato (AI)valor). Devuelve #VALUE! si la entrada está vacía, #NUM! si se encuentra un AI desconocido en el flujo sin procesar y #N/A si no hay elementos descompuestos.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=READGS1("(01)04012345678901(10)LOT42")` | {"(01)","04012345678901";"(10)","LOT42"} | Descomponer en filas de (AI) y valor |
| `=READGS1("(01)04012345678901(10)LOT42",TRUE)` | (01)04012345678901(10)LOT42 | Concatenar en una sola cadena |

## Notas

- Se puede introducir tanto el formato con paréntesis (AI)valor como el flujo sin procesar leído de un código de barras (con separadores GS).
- Usa las tablas de AI GS1 de longitud fija y longitud variable para descomponer el flujo.
- Compatibilidad: Excel 2010+. Se registra siempre como `READGS1` en todas las versiones de Excel.

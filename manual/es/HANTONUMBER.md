# HANTONUMBER

**Categoría**: función exclusiva de EGTools

Convierte notación monetaria coreana/hanja en número.

## Sintaxis

```
=HANTONUMBER(texto, [solo_numero])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto | Obligatorio | texto monetario coreano |
| solo_numero | Opcional | VERDADERO devuelve número (por defecto); FALSO texto formateado |

## Devuelve

De forma predeterminada (solo número=TRUE) devuelve un número (escalar); si la parte entera supera los 15 dígitos, el resultado se devuelve como texto para preservar la precisión. Con solo número=FALSE devuelve un texto con formato del tipo "금1,234원". Devuelve #VALUE! cuando el texto está vacío o contiene dos o más dígitos arábigos o puntos consecutivos.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=HANTONUMBER("일금오만원정")` | 50000 | importe en coreano a número |
| `=HANTONUMBER("삼억이천오백만")` | 325000000 | manejo de las unidades eok (억) y man (만) |
| `=HANTONUMBER("일금오만원정",FALSE)` | 일금50,000원정 | devuelve texto con formato |

## Notas

- También reconoce la notación en caracteres chinos (hanja) y sus variantes (一二三, 壹貳參, 拾佰仟萬, etc.).
- Los espacios y las comas de la entrada se ignoran; los prefijos 금/일금 y los sufijos 원/원정 se restauran tal cual cuando solo número=FALSE.
- Compatibilidad: Excel 2010+. Se registra siempre como `HANTONUMBER` en todas las versiones de Excel.

# TOLUNAR

**Categoría**: función exclusiva de EGTools

Convierte una fecha solar (gregoriana) al calendario lunar coreano (919–2050).

## Sintaxis

```
=TOLUNAR(fecha_solar)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| fecha_solar | Obligatorio | fecha, serial, texto "yyyy-mm-dd", número yyyymmdd o rango |

## Devuelve

Devuelve la fecha del calendario lunar como texto "yyyy-mm-dd" (con sufijo si es un mes intercalado) en forma de escalar, y se derrama como matriz del mismo tamaño si se introduce un rango. Si la fecha no se puede interpretar o está fuera del rango de la tabla lunar (años 918~2050), devuelve un error #VALUE!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=TOLUNAR(DATE(2025,1,29))` | 2025-01-01 | Año Nuevo Lunar de 2025 → 1/1 lunar |
| `=TOLUNAR(DATE(2025,10,6))` | 2025-08-15 | Chuseok de 2025 → 8/15 lunar |

## Notas

- Los meses intercalados llevan el sufijo "(윤)" (coreano) o "(LM)" (inglés) según el idioma configurado.
- Usa una tabla integrada (años 918~2050) generada con KoreanLunisolarCalendar de .NET.
- Función relacionada: TOSOLAR
- Compatibilidad: Excel 2010+. Se registra siempre como `TOLUNAR` en todas las versiones de Excel.

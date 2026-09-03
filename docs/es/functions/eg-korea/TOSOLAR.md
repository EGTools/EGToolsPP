# TOSOLAR

**Categoría**: función exclusiva de EGTools

Convierte una fecha lunar coreana a fecha solar (gregoriana) (919–2050).

## Sintaxis

```
=TOSOLAR(fecha_lunar)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| fecha_lunar | Obligatorio | texto como "2025-06-05" — mes intercalar con 윤/LM — o rango |

## Devuelve

Devuelve el valor de serie de la fecha del calendario solar (gregoriano) como número escalar, y se derrama como matriz del mismo tamaño si se introduce un rango. Si el texto no se puede interpretar, la fecha está fuera del rango de la tabla lunar (años 918~2050) o la fecha lunar no existe, devuelve un error #VALUE!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=TOSOLAR("2025-01-01")` | 45686 | Año Nuevo Lunar → 2025-01-29 |
| `=TOSOLAR("2025-08-15")` | 45936 | Chuseok lunar → 2025-10-06 |

## Notas

- Los meses intercalados se indican anteponiendo "윤" o "LM" (por ejemplo, "2025-윤06-01").
- Aplique un formato de número de fecha a la celda de resultado.
- Función relacionada: TOLUNAR
- Compatibilidad: Excel 2010+. Se registra siempre como `TOSOLAR` en todas las versiones de Excel.

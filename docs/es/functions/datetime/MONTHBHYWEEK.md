# MONTHBHYWEEK

**Categoría**: función exclusiva de EGTools

Mes al que pertenece la semana de una fecha según día inicial y propietario.

## Sintaxis

```
=MONTHBHYWEEK(fecha, [inicio_semana], [dia_propietario])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| fecha | Obligatorio | la fecha |
| inicio_semana | Opcional | 1=dom…7=sáb (por defecto 2=lun) |
| dia_propietario | Opcional | día que decide el mes (por defecto 5=jue) |

## Devuelve

Devuelve el mes (1~12, número escalar) al que pertenece la semana de la fecha. Si la fecha no se puede interpretar o el día de inicio o el día de referencia están fuera del rango 1~7, devuelve un error #VALUE!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=MONTHBHYWEEK(DATE(2026,2,1))` | 1 | La semana del domingo 1 de febrero pertenece a enero |
| `=MONTHBHYWEEK(DATE(2026,2,2))` | 2 | La semana del 2 de febrero pertenece a febrero |

## Notas

- Los números de día de la semana van de 1=domingo a 7=sábado; los valores predeterminados son inicio en lunes y jueves como día de referencia.
- El mes al que pertenece una semana se determina por el mes que contiene su día de referencia.
- Si la fecha es una matriz, la función se aplica elemento a elemento y se derrama una matriz de la misma forma (un elemento de error devuelve ese error).
- Función relacionada: WEEKNUMOFMONTH
- Compatibilidad: Excel 2010+. Se registra siempre como `MONTHBHYWEEK` en todas las versiones de Excel.

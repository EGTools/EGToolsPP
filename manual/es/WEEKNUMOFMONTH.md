# WEEKNUMOFMONTH

**Categoría**: función exclusiva de EGTools

Número de semana dentro del mes según día inicial y día propietario.

## Sintaxis

```
=WEEKNUMOFMONTH(fecha, [inicio_semana], [dia_propietario])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| fecha | Obligatorio | la fecha |
| inicio_semana | Opcional | 1=dom…7=sáb (por defecto 2=lun) |
| dia_propietario | Opcional | día que decide el mes (por defecto 5=jue) |

## Devuelve

Devuelve el número de semana dentro del mes (número escalar). Si la fecha no se puede interpretar o el día de inicio o el día de referencia están fuera del rango 1~7, devuelve un error #VALUE!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=WEEKNUMOFMONTH(DATE(2026,1,15))` | 3 | Tercera semana de enero |
| `=WEEKNUMOFMONTH(DATE(2026,2,1))` | 5 | El 1 de febrero pertenece a la semana 5 de enero |

## Notas

- Los números de día de la semana van de 1=domingo a 7=sábado; los valores predeterminados son inicio en lunes y jueves como día de referencia.
- El mes y el número de semana a los que pertenece una semana se calculan según el mes del día de referencia de esa semana.
- Función relacionada: MONTHBHYWEEK
- Compatibilidad: Excel 2010+. Se registra siempre como `WEEKNUMOFMONTH` en todas las versiones de Excel.

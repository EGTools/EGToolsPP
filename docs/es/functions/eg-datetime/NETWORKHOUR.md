# NETWORKHOUR

**Categoría**: función exclusiva de EGTools

Calcula las horas netas de trabajo restando los descansos.

## Sintaxis

```
=NETWORKHOUR(hora_inicio, hora_fin, [descansos])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| hora_inicio | Obligatorio | hora(s) de inicio |
| hora_fin | Obligatorio | hora(s) de fin, misma forma |
| descansos | Opcional | intervalos: 1.ª columna inicio, 2.ª fin |

## Devuelve

Se derrama como matriz con el tamaño resultante de difundir inicio y fin (escalar si son valores únicos); los valores son números en unidades de día (1 día = 1). Si los tamaños de inicio y fin no se pueden difundir (tamaños distintos que no sean 1) o la matriz de descansos tiene menos de 2 columnas, devuelve #VALUE!; las celdas de inicio/fin no numéricas producen texto vacío, y una celda cuyo fin con fecha es anterior a su inicio produce #VALUE!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=NETWORKHOUR(TIME(9,0,0),TIME(18,0,0))*24` | 9 | Horas de trabajo de 9:00 a 18:00 |
| `=NETWORKHOUR(TIME(9,0,0),TIME(18,0,0),HSTACK(TIME(12,0,0),TIME(13,0,0)))*24` | 8 | Excluye 1 hora de almuerzo |
| `=NETWORKHOUR(TIME(22,0,0),TIME(6,0,0))*24` | 8 | Turno que cruza la medianoche |

## Notas

- El resultado está en unidades de día; multiplique por 24 para obtener horas.
- Si solo se indican horas y el fin es menor que el inicio, se considera que cruza la medianoche.
- Los intervalos de descanso superpuestos se combinan automáticamente, y los descansos que cruzan la medianoche se dividen para el cálculo.
- Inicio y fin admiten difusión escalar↔matriz; p. ej., =NETWORKHOUR(A2:A100,$B$1) fija la hora de fin en una sola celda.
- Compatibilidad: Excel 2010+. Se registra siempre como `NETWORKHOUR` en todas las versiones de Excel.

# KOREANHOLIDAYS

**Categoría**: función exclusiva de EGTools

Devuelve los días festivos de Corea para un año (con reglas de festivo sustituto).

## Sintaxis

```
=KOREANHOLIDAYS([año], [con_nombre], [dia_trabajo], [api_key])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| año | Opcional | año (por defecto: actual) |
| con_nombre | Opcional | VERDADERO incluye nombres (por defecto) |
| dia_trabajo | Opcional | VERDADERO incluye el Día del Trabajo (por defecto) |
| api_key | Opcional | Clave de data.go.kr para festivos temporales. Se guarda al introducirla una vez y luego puede omitirse. Indique 0 para omitir la API y usar solo el cálculo integrado |

## Devuelve

Se derrama como matriz de 2 columnas donde cada fila es [valor de serie de fecha, nombre del festivo] (solo 1 columna de fechas cuando holiday_name=FALSE). Si el año está fuera del rango 919~2050, devuelve un error #VALUE!.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=ROWS(KOREANHOLIDAYS(2025,,,0))` | 18 | Número de filas de festivos de 2025 — el 0 final omite la API y usa solo el cálculo integrado |
| `=INDEX(KOREANHOLIDAYS(2025,FALSE,,0),1,1)` | 45658 | Valor de serie del primer festivo (1 de enero), solo fechas (sin nombres) |
| `=KOREANHOLIDAYS(2026)` |  | Comportamiento predeterminado — con la clave data.go.kr guardada se combinan también los festivos temporales (como el día de las elecciones locales simultáneas nacionales del 2026-06-03). El número de filas resultante depende de si hay una clave registrada y del estado de la red |

## Notas

- Las reglas de festivos sustitutos se aplican según su año de entrada en vigor (Día del Niño desde 2014; Día del Movimiento del 1 de Marzo, Día de la Liberación y Día de la Fundación Nacional desde 2022; Cumpleaños de Buda y Navidad desde 2023).
- Los nombres de los festivos siguen el idioma configurado del complemento (coreano/inglés); los nombres de festivos que coinciden en el mismo día se combinan con comas.
- Los festivos del calendario lunar (Año Nuevo Lunar, Cumpleaños de Buda, Chuseok) se calculan con la tabla lunar integrada (años 918~2050).
- El argumento api_key admite tres usos. **Si se omite** (comportamiento predeterminado), se consulta la API de festivos con la clave data.go.kr guardada y se combinan automáticamente los elementos que el cálculo integrado no incluye, como festivos temporales y días de elecciones; si no hay clave o la consulta falla, se devuelve solo el resultado del cálculo integrado (la fórmula no se rompe con un error). **Si se especifica una cadena de clave**, esa clave se guarda y se utiliza; si entonces la consulta falla, la celda muestra `ERROR: 공휴일 API - <motivo>` con la causa (sin clave / clave rechazada más el mensaje del servidor / fallo de conexión con el servidor / sin elementos en la respuesta). **Si se especifica 0**, no se llama a la API y se usa solo el cálculo integrado (sin acceso a la red).
- La clave de autenticación de data.go.kr (portal coreano de datos públicos) es una sola por cuenta y la comparten KOREANHOLIDAYS, los comandos de calendario/agenda de la cinta y BRNSTATUS. Puede pasarse una vez como último argumento o registrarse en la cinta [EGTools] → [Administrar claves API]; queda guardada en la cuenta del usuario y después puede omitirse. Una clave registrada en una versión anterior para consultas del número de registro de empresa (Corea) se traslada automáticamente a la ranura compartida la primera vez que se usa.
- Los resultados de la consulta se almacenan en caché por año, de modo que al recalcular cada año se consulta una sola vez. Los fallos no se almacenan en caché, por lo que corregir la clave surte efecto en el siguiente cálculo.
- Una matriz en api_key devuelve #VALUE! antes de guardar nada (evita almacenar una clave incorrecta).
- Compatibilidad: Excel 2010+. Se registra siempre como `KOREANHOLIDAYS` en todas las versiones de Excel.

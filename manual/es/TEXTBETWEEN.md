# TEXTBETWEEN

**Categoría**: función exclusiva de EGTools

Extrae el texto entre ClaveInicio y ClaveFin.

## Sintaxis

```
=TEXTBETWEEN(texto, texto_izq, texto_der, [delimitador_o_num], [incluir])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| texto | Obligatorio | texto fuente |
| texto_izq | Obligatorio | marca de inicio |
| texto_der | Obligatorio | marca de fin |
| delimitador_o_num | Opcional | unión de coincidencias (por defecto ",") o número N para solo la N-ésima |
| incluir | Opcional | VERDADERO incluye las marcas |

## Devuelve

Devuelve el texto extraído como un texto único (escalar). Devuelve #VALUE! cuando el texto, la clave inicial o la clave final están vacíos, o cuando el número de sección N está fuera del número de secciones; si no hay secciones coincidentes, devuelve una cadena vacía.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=TEXTBETWEEN("[a] and [b]","[","]")` | a,b | une todas las secciones con , |
| `=TEXTBETWEEN("[a] and [b]","[","]",2)` | b | solo la 2.ª sección |
| `=TEXTBETWEEN("<x><y>","<",">","/",TRUE)` | <x>/<y> | une incluyendo las claves |

## Notas

- La búsqueda de claves ignora mayúsculas y minúsculas.
- Si el 4.º argumento es un número, selecciona solo la N-ésima sección; si es texto, es el separador de unión de todas las secciones (predeterminado ",").
- Función relacionada: TEXTREPLACE
- Compatibilidad: Excel 2010+. Se registra siempre como `TEXTBETWEEN` en todas las versiones de Excel.

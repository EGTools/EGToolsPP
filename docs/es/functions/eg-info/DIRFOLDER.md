# DIRFOLDER

**Categoría**: función exclusiva de EGTools

Lista los archivos (y opcionalmente carpetas) de un directorio.

## Sintaxis

```
=DIRFOLDER(ruta, [filtro], [carpetas], [subcarpetas], [ocultos])
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| ruta | Obligatorio | ruta de la carpeta |
| filtro | Opcional | p. ej. "*.xlsx;*.csv" (por defecto *.*) |
| carpetas | Opcional | VERDADERO incluye nombres de carpeta |
| subcarpetas | Opcional | VERDADERO recorre subcarpetas |
| ocultos | Opcional | VERDADERO incluye elementos ocultos |

## Devuelve

Devuelve una matriz derramada con los nombres de los elementos; si se incluyen las subcarpetas, las columnas se escalonan según la profundidad de carpeta. Si la ruta falta o está vacía devuelve #VALUE!; si ningún elemento coincide devuelve #N/A.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=DIRFOLDER("C:\Data","*.xlsx;*.csv",FALSE,TRUE)` |  | El resultado depende del contenido de la carpeta |

## Notas

- Función no macro; lee la carpeta directamente mediante la API Win32.
- Se pueden indicar varios filtros separados por ";"; se admiten los caracteres comodín * y ?.
- Los elementos ocultos se excluyen de forma predeterminada; se pueden incluir estableciendo la opción de incluir ocultos en TRUE.
- Compatibilidad: Excel 2010+. Se registra siempre como `DIRFOLDER` en todas las versiones de Excel.

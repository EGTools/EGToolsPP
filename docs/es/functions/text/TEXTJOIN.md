# TEXTJOIN

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2016/2019

Une texto con un delimitador, opcionalmente omitiendo vacíos.

## Sintaxis

```
=TEXTJOIN(delimitador, ignorar_vacíos, texto1, ...)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| delimitador | Obligatorio | separador entre elementos |
| ignorar_vacíos | Obligatorio | VERDADERO para omitir celdas vacías |
| texto1 | Obligatorio | texto o rango |
| … | Opcional | Los argumentos anteriores se pueden repetir |

## Devuelve

Devuelve un texto único (escalar) unido con el separador. Si algún valor, el separador o el indicador de ignorar vacíos contiene un error, se devuelve ese error.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=TEXTJOIN("-",TRUE,"a","","b")` | a-b | omite los valores vacíos |
| `=TEXTJOIN("-",FALSE,"a","","b")` | a--b | conserva los valores vacíos |
| `=TEXTJOIN(",",TRUE,{1,2;3,4})` | 1,2,3,4 | concatena los elementos de la matriz |
| `=TEXTJOIN({"-";"+"},TRUE,"a","b","c")` | a-b+c | El separador en matriz rota |

## Notas

- Si se omite el argumento ignore_empty, se trata como TRUE.
- Se admiten hasta 255 argumentos de texto.
- Un separador en matriz se usa en rotación; p. ej., {"-";"+"} aplica -, +, -, … por turnos.
- Función relacionada: CONCAT
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `TEXTJOIN` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.TEXTJOIN`.

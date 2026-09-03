# UNICHAR

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Devuelve el carácter Unicode del número (punto de código) dado.

## Sintaxis

```
=UNICHAR(número)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| número | Obligatorio | punto de código Unicode |

## Devuelve

Devuelve el carácter (texto escalar) correspondiente al punto de código; si se especifica una matriz de números, se evalúa elemento por elemento y el resultado se derrama. Devuelve #VALUE! cuando el código está fuera del rango 1–1114111 o dentro del área de suplentes (55296–57343).

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=UNICHAR(44032)` | 가 | sílaba hangul |
| `=UNICHAR(128512)` | 😀 | carácter fuera del BMP |
| `=UNICHAR(0)` | #VALUE! | fuera de rango |

## Notas

- Los caracteres a partir de 0x10000 se devuelven como un par suplente (UTF-16).
- Función relacionada: UNICODE
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `UNICHAR` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.UNICHAR`.

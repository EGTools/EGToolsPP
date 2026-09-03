# FILTERXML

**Categoría**: función de compatibilidad con MS Excel · **Nativa desde**: Excel 2013

Devuelve datos de un contenido XML mediante la ruta XPath indicada.

## Sintaxis

```
=FILTERXML(xml, xpath)
```

## Argumentos

| Argumento | Obligatorio | Descripción |
|---|---|---|
| xml | Obligatorio | texto XML |
| xpath | Obligatorio | consulta XPath |

## Devuelve

Si el XPath coincide con un solo nodo, devuelve su texto como escalar; si coincide con varios, se derrama como matriz vertical de texto de N×1. Si el XML o el XPath están vacíos, o falla el análisis del XML o el procesamiento del XPath, devuelve #VALUE!; si no hay nodos coincidentes, devuelve #N/A.

## Ejemplos

| Fórmula | Resultado | Descripción |
|---|---|---|
| `=FILTERXML("<r><a>1</a><a>2</a></r>","//a")` | {"1";"2"} | Texto de todos los nodos a |
| `=FILTERXML(WEBSERVICE("https://example.com/rss.xml"),"//item/title")` |  | Depende del XML externo |

## Notas

- Se analiza con MSXML6 y la función en sí no accede a la red — para XML remoto, combínela con WEBSERVICE.
- El resultado es el texto de los nodos y siempre se devuelve como texto.
- Por usar COM, se excluye del recálculo multiproceso (no es segura para subprocesos).
- Una matriz en xml·xpath devuelve una matriz con la forma de ese argumento, donde cada elemento es el primer valor del resultado de ese elemento (reducción al primer valor, igual que la función nativa).
- Compatibilidad: Excel 2010+. En versiones antiguas sin la función nativa se registra como `FILTERXML` (sustitución directa); en Excel moderno que ya la incluye se registra como `EG.FILTERXML`.

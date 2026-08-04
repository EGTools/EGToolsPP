# Administrar claves API

**Ubicación**: pestaña `EGTools++` de la cinta → grupo `EGTools` → menú desplegable del botón dividido

Cuadro de diálogo para registrar/eliminar las claves API que usan las funciones de
API públicas y las funciones de la cinta.

La lista contiene **3** servicios.

| Servicio | Uso | Funciones relacionadas |
|---|---|---|
| juso | Búsqueda de direcciones (Corea) | [SEARCHADDRESS](../SEARCHADDRESS.md) |
| data.go.kr | Festivos y estado del registro de empresas | [KOREANHOLIDAYS](../KOREANHOLIDAYS.md) · [Calendario/Agenda](calendar.md#acerca-de-los-festivos) · [BRNSTATUS](../BRNSTATUS.md) |
| vworld | Mapas/coordenadas | [GEOSEARCH](../GEOSEARCH.md) · [GEOCODER](../GEOCODER.md) · [GEOADDRESS](../GEOADDRESS.md) · [ADDRESSMAP](../ADDRESSMAP.md) |

- Al seleccionar un servicio, el estado de registro de la clave se muestra **solo
  enmascarado (●●●●●●)**. La clave guardada nunca se vuelve a mostrar en claro.
- La clave **se guarda en el perfil del usuario (Registro de Windows)** y **no queda
  en el libro ni en las fórmulas**.
- Coexiste con el método tradicional de **pasar la clave directamente como último
  argumento** de la función — una clave pasada como argumento se guarda
  automáticamente en el mismo almacén y a partir de entonces se puede omitir.
- **Las claves no se distribuyen con EGTools++.** El usuario debe obtenerlas por su
  cuenta en cada servicio (juso.go.kr, data.go.kr, vworld.kr).

## data.go.kr usa una sola clave

El portal de datos públicos de Corea (data.go.kr) emite **una única clave de
autenticación por cuenta**, que se comparte entre todos los servicios solicitados.
Por eso EGTools++ tampoco la separa por servicio y la gestiona en una única entrada
`data.go.kr`.

- Esta única clave la **comparten** [KOREANHOLIDAYS](../KOREANHOLIDAYS.md)
  (información de días señalados), la incorporación de festivos temporales de
  [Calendario/Agenda](calendar.md#acerca-de-los-festivos) y
  [BRNSTATUS](../BRNSTATUS.md) (estado del registro de empresas de la agencia
  tributaria coreana).
- Una clave que en versiones anteriores se hubiera **registrado por separado para
  el registro de empresas** se **migra automáticamente** a la entrada unificada la
  primera vez que se usa — el usuario no necesita volver a registrarla ni hacer
  limpieza.
- Puede pegar cualquiera de las dos claves que muestra el portal, la de
  codificación (Encoding) o la de decodificación (Decoding). Si es necesario, se
  codifica internamente antes de usarla.
- Si una consulta falla y **se confirma que la clave fue rechazada, la clave
  guardada se elimina automáticamente** y se devuelve un aviso para obtener una
  nueva. Obtenga una clave nueva y regístrela de nuevo.

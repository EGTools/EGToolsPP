# Configuración SMTP

**Ubicación**: pestaña `EGTools` de la cinta → grupo `EGTools` → menú desplegable del botón dividido

Cuadro de diálogo para registrar los datos del servidor SMTP que usa el **envío de
correo** de [Combinar correspondencia](mailmerge.md) (opciones 4, 5, 6 y 7).

## Campos de configuración

| Campo | Descripción |
|---|---|
| Nombre del remitente | Nombre del remitente mostrado al destinatario (opcional — si se omite, solo se muestra la dirección de correo) |
| Correo del remitente | Dirección del remitente y, a la vez, **cuenta de inicio de sesión SMTP** |
| Servidor de correo saliente | Dirección del servidor SMTP (p. ej. `smtp.gmail.com`) |
| Puerto del servidor | Predeterminado **465** |
| SSL | Uso de SSL (activado de forma predeterminada) |

## Reglas

- La configuración se guarda en el perfil del usuario (Registro de Windows) y **no
  queda en el libro**.
- **La contraseña no se guarda nunca.** Se pide en cada envío de correo y se usa
  solo en memoria.
- Si en Combinar correspondencia elige una opción con envío y la configuración está
  vacía (falta la dirección de correo, el servidor o el puerto), este cuadro de
  diálogo **se abre automáticamente**. Si aun así no se completa, el envío se
  aborta.

## Limitación importante — solo SSL implícito

El motor de envío (CDO de Windows) **solo admite SSL implícito (implicit SSL)**.
Normalmente corresponde al **puerto 465**. **STARTTLS en el puerto 587 no es
compatible**, por lo que si introduce 587 el envío fallará.

- Gmail: servidor `smtp.gmail.com`, puerto `465`, SSL activado —
  en lugar de la contraseña de la cuenta, genere e introduzca una
  **contraseña de aplicación**.
- Otros servicios también deben usar un puerto de tipo "SSL (465)"; si la cuenta
  usa verificación en dos pasos, puede necesitar una contraseña de aplicación.

## Notas

- Los asuntos/cuerpos con caracteres no ASCII (coreano, etc.) se envían codificados
  en UTF-8.
- Si un envío falla, la descripción del error devuelta por el servidor se muestra
  tal cual en la lista de fallos del resumen final de
  [Combinar correspondencia](mailmerge.md).

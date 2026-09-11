# GLS – Emails de tracking desde Cierre Día (rutina automática LOCAL)

> Esto NO es una skill de invocación manual: es la definición de una rutina
> pensada para dispararse sola, de forma automática, mediante un programador
> de tareas local (p. ej. el Programador de tareas de Windows) que lance
> Claude Code con este prompt a la hora indicada. Vive en `rutinas/` en vez
> de en `.claude/skills/` precisamente para no aparecer como skill invocable
> a mano.

Eres una rutina **LOCAL** de lunes a jueves (15:30) y viernes (13:00) para
Lizarte S.A.U. El usuario (mgomez@lizarte.com) suele estar presente en su
equipo cuando esto se ejecuta, con Chrome abierto y una sesión ya logeada en
la extranet de GLS (https://extranet.gls-spain.es). Esta tarea depende de esa
sesión de Chrome local y de Outlook de escritorio — no es una tarea de nube,
no intentes usar conectores MS365/cloud para el correo.

## Objetivo

Por cada envío que esté pendiente en "Cierre Día" de GLS hoy, extraer sus
datos y enviar al cliente un email de seguimiento (tracking) en 3 idiomas,
siempre que el envío tenga un email de contacto válido.

## Paso 1 — Navegar a Cierre Día

1. Carga las herramientas de Claude in Chrome: ToolSearch
   "select:mcp__claude-in-chrome__tabs_context_mcp,mcp__claude-in-chrome__navigate,mcp__claude-in-chrome__computer,mcp__claude-in-chrome__read_page,mcp__claude-in-chrome__get_page_text,mcp__claude-in-chrome__find,mcp__claude-in-chrome__tabs_create_mcp,mcp__claude-in-chrome__browser_batch".
2. Llama a tabs_context_mcp con createIfEmpty:true, y navega (en esa pestaña)
   a https://extranet.gls-spain.es/Extranet/GrabaEnvios/Cierre.aspx . Al ser
   el mismo perfil de Chrome del usuario, debería cargar ya logeado. Si
   aparece una pantalla de login, DETENTE y termina la tarea explicando que
   no hay sesión activa (no intentes adivinar credenciales).
3. IMPORTANTE — NUNCA pulses el botón global "cerrar" de esta pantalla (eso
   cerraría el día para TODOS los envíos de forma irreversible). Esta tarea
   SOLO navega y lee datos, nunca hace clic en "cerrar".

## Paso 2 — Por cada envío en la sección "Envíos en Cierre"

Para cada fila/envío listado:

1. Anota su número de envío (ej. "839-115539686") y ábrelo en su pantalla de
   edición:
   https://extranet.gls-spain.es/Extranet/GrabaEnvios/GrabaEnvios.aspx?CodSerie=<parte numérica>
   (o pulsa el botón "editar" de esa fila).
2. Lee de esa pantalla de edición:
   - Campo **Kontakto / Contacto**: debe contener un email de cliente. Si el
     texto de ese campo NO es un email válido (regex tipo algo@dominio.tld),
     OMITE este envío por completo (no mandes correo, no es un error,
     simplemente pasa al siguiente).
   - Campo **País**.
   - Campo **Observaciones**: tómalo tal cual, literal — es el número de
     pedido del cliente (el formato varía por cliente, no lo interpretes ni
     lo reformatees). Si este campo está vacío o no contiene nada útil, NO es
     motivo para omitir el envío: simplemente omite la línea "Your order
     number / Ihre Bestellnummer / Votre numéro de commande" del email y
     sigue adelante con el resto (tracking, país, plazo). Omitir el envío
     completo SOLO aplica cuando falta el email válido en Kontakto/Contacto
     (ver punto anterior).
3. Vuelve (o abre en otra pestaña) la etiqueta del envío:
   https://extranet.gls-spain.es/Extranet/GrabaEnvios/etiPDF.aspx?codexp=GLSEXT<parte numérica>
   y lee el "Track-ID" de GLS (código alfanumérico tipo "Z6SZ1A0U", bajo el
   texto "Ihre GLS Track-ID").
4. Calcula el plazo de entrega estimado según el país:
   - Alemania, Austria, Holanda → "4 working days" / "4 Werktage" / "4 jours
     ouvrables"
   - Polonia, Croacia → "5 working days" / "5 Werktage" / "5 jours ouvrables"
   - Grecia → "6 working days" / "6 Werktage" / "6 jours ouvrables"
   - Cualquier otro país → "4 to 6 working days" / "4 bis 6 Werktage" / "4 à
     6 jours ouvrables"

## Paso 3 — Componer el email (si hay email válido)

Cuerpo del correo, SIEMPRE en este orden inglés → alemán → francés,
concatenados en un único email separados por una línea
"--------------------------------------------------". Sustituye
{tracking}, {order_number} y {days_en}/{days_de}/{days_fr} según el paso 2:

**INGLÉS:**
"Dear Customer,
Your shipment from Lizarte S.A.U. (central warehouse in Spain) has the
following GLS tracking number: {tracking}
Your order number: {order_number}
Estimated delivery time: {days_en}.
You can check the status of your shipment on the GLS website using this
number.
For any questions, please contact dach@lizarte.com.
Best regards,
Lizarte S.A.U."

**ALEMÁN:**
"Sehr geehrte Damen und Herren,
Ihre Sendung von Lizarte S.A.U. (Zentrallager in Spanien) hat folgende GLS
Tracking-Nummer: {tracking}
Ihre Bestellnummer: {order_number}
Voraussichtliche Lieferzeit: {days_de}.
Den Status Ihrer Sendung können Sie mit dieser Nummer auf der GLS-Website
verfolgen.
Bei Fragen wenden Sie sich bitte an dach@lizarte.com.
Mit freundlichen Grüßen
Lizarte S.A.U."

**FRANCÉS:**
"Cher client,
Votre expédition depuis Lizarte S.A.U. (entrepôt central en Espagne) porte
le numéro de suivi GLS suivant : {tracking}
Votre numéro de commande : {order_number}
Délai de livraison estimé : {days_fr}.
Vous pouvez suivre l'état de votre expédition sur le site de GLS avec ce
numéro.
Pour toute question, veuillez contacter dach@lizarte.com.
Cordialement,
Lizarte S.A.U."

Asunto: "Your shipment tracking / Ihre Sendungsverfolgung / Suivi de votre
expédition"

## Paso 4 — Enviar el correo (SIEMPRE vía PowerShell + Outlook COM, nunca vía conector de nube)

1. Comprueba primero qué cuentas hay en Outlook con:
   `$outlook = New-Object -ComObject Outlook.Application; $outlook.Session.Accounts | ForEach-Object { $_.SmtpAddress }`
2. Elige remitente con esta prioridad: dach@lizarte.com si existe esa cuenta
   en Outlook; si no, mgomez@lizarte.com si existe; si no, la cuenta del
   usuario (mgomez@lizarte.com). Para forzar el remitente cuando no es la
   cuenta por defecto, usa `$mail.SendUsingAccount = <cuenta encontrada en
   $outlook.Session.Accounts>`.
3. Crea el mail:
   `$mail = $outlook.CreateItem(0); $mail.To = "<email del Kontakto>"; $mail.CC = "dach@lizarte.com"; $mail.Subject = "..."; $mail.Body = "..."; $mail.Send()`
4. Escribe el script a un archivo .ps1 en el directorio scratchpad de la
   sesión y ejecútalo con `powershell -ExecutionPolicy Bypass -File`.

## Excepción SOLO para la ejecución del 2026-09-10

El usuario adelantó a mano el número de pedido de dos envíos concretos
porque hoy el campo Observaciones aún no lo tenía relleno. Comprueba la
fecha actual del sistema: SOLO si es 2026-09-10, aplica esto (en cualquier
otra fecha, ignora por completo esta sección y usa siempre el campo
Observaciones como en el Paso 2):

- Envío con destinatario "WM SE - AUGSBURG", ref. 7290920 → usa como
  {order_number} "108/199595" (en vez de leerlo de Observaciones).
- Envío con destinatario "WM SE - BERLIN" → usa como {order_number}
  "124/113300" (en vez de leerlo de Observaciones).

Para cualquier otro envío de hoy que no sea uno de estos dos, sigue la regla
normal del Paso 2 (leer Observaciones, o omitir la línea si está vacío).

## Paso 5 — Resumen final

Al terminar de procesar todos los envíos de "Cierre Día", escribe un resumen
breve: cuántos envíos había, a cuáles se les envió correo (envío, email,
tracking) y cuáles se omitieron por no tener email válido en Kontakto.

## Contexto de validación

Esta rutina fue probada manualmente el 2026-09-10 con un envío real
(tracking Z6SZ1A0U, destino Berlín/Alemania) enviando un correo de prueba a
mgomez@lizarte.com; funcionó correctamente. Hoy en Outlook solo estaba
configurada la cuenta mgomez@lizarte.com (no dach@), así que lo más
probable es que el envío real también salga desde mgomez@lizarte.com — es
normal, no es un fallo.

## Pendiente: disparo automático

Este documento define QUÉ hace la rutina. Para que se dispare SOLA (lunes a
jueves 15:30, viernes 13:00) todavía falta configurar el disparador en el
equipo local de mgomez@lizarte.com (por ejemplo, una tarea del Programador
de tareas de Windows que lance `claude` con este archivo como prompt). Esa
parte no se ha configurado aún — pendiente de decidir el mecanismo exacto y
si el envío de correos debe quedar totalmente desatendido o con una pausa de
confirmación antes de pulsar "Send".

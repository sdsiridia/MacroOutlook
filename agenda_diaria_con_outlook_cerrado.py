'''
### Script para extraer la agenda diaria de Outlook y enviarla por correo electrónico.

Si ves el error, `(-2147418111, 'La llamada fue rechazada por el destinatario.', None, None)`, 
es un código de error COM (Component Object Model) que indica que la **conexión con Outlook 
fue rechazada**. Este problema no es del script en sí, sino de la **configuración de seguridad 
de Outlook** que impide que programas externos (como tu script de Python) accedan 
y manipulen su contenido.

---

### Causa del Error

Outlook tiene medidas de seguridad incorporadas para proteger 
a los usuarios de programas maliciosos que intentan enviar correos 
electrónicos o acceder a datos de forma no autorizada. 
Cuando tu script intenta conectarse a través de 
`win32com.client.Dispatch("Outlook.Application")`, Outlook detecta 
que se trata de un acceso programático y, si no está configurado 
para permitirlo, lo rechaza. Este error es una señal de que Outlook 
ha bloqueado tu intento de automatización.

---

### Solución

Para resolver este problema, necesitas **ajustar la configuración 
del Centro de Confianza de Outlook** para permitir que las 
aplicaciones externas accedan a él. Sigue estos pasos:

1.  **Abre Outlook.**
2.  Ve a **Archivo** en la esquina superior izquierda.
3.  Selecciona **Opciones**.
4.  En la ventana de Opciones de Outlook, haz clic en **Centro de confianza**.
5.  Haz clic en **Configuración del Centro de confianza...**.
6.  En el menú de la izquierda, selecciona **Acceso programático**.

Aquí verás las opciones de seguridad para el acceso programático. 
La configuración por defecto suele ser **"Advertir siempre sobre 
actividades sospechosas"** o algo similar. 
Debes cambiarla para permitir el acceso. 

* **Opción recomendada:** Si estás en un entorno seguro y el script 
es de tu autoría, puedes seleccionar **"Permitir siempre 
el acceso programático a Mi equipo"**. Esto evitará que 
Outlook muestre la advertencia cada vez que ejecutes el script.
* **Opción alternativa (más segura):** 
Deja la configuración por defecto, pero ten en cuenta que Outlook 
te mostrará una ventana emergente cada vez que el script intente conectarse. 
Debes hacer clic en **"Permitir"** en esa ventana 
para que la ejecución del script continúe.

Después de cambiar esta configuración, 
**reinicia Outlook y vuelve a ejecutar tu script**. 
Debería funcionar sin el error de rechazo.'''
import subprocess
import time
from datetime import datetime
import win32com.client


def send_daily_outlook_agenda():
    """
    Se conecta a Outlook, extrae los eventos del día y los envía por correo.
    """
    try:
        # Intenta crear un objeto de la aplicación de Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
    except Exception as e:
        print(f"Outlook no se está ejecutando. Intentando abrirlo... ({e})")
        # Si falla, intenta iniciar Outlook
        try:
            subprocess.Popen("outlook")
            # Espera unos segundos para que Outlook se inicie completamente
            time.sleep(10)
            # Intenta conectarse de nuevo
            outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception as e_open:
            print(f"Error al intentar abrir Outlook: {e_open}")
            return  # Termina la ejecución del script si no se puede abrir Outlook

    try:
        namespace = outlook.GetNamespace("MAPI")

        # Accede a la carpeta del calendario principal
        # 9 es el índice de la carpeta de calendario
        calendar = namespace.GetDefaultFolder(9)

        # Obtiene todas las citas del calendario
        citas = calendar.Items
        citas.IncludeRecurrences = True
        citas.Sort("[Start]")

        # Define la fecha de hoy para la comparación
        today = datetime.now().date()

        # Crea el cuerpo del correo en HTML
        cuerpo_html = f"<h3>Agenda del día: {today.strftime('%d/%m/%Y')}</h3><br>"

        # Una bandera para verificar si se encontraron eventos
        events_found = False

        # Itera sobre todas las citas y las filtra manualmente por fecha
        for cita in citas:
            # Comprueba si la fecha de inicio de la cita es la de hoy
            if cita.Start.date() == today:
                # Excluye citas de todo el día que no tienen una hora de inicio y fin
                if not cita.AllDayEvent:
                    hora_inicio = cita.Start.strftime("%I:%M %p")
                    hora_fin = cita.End.strftime("%I:%M %p")
                    cuerpo_html += f"<p><strong>{hora_inicio} - \
                         {hora_fin}</strong>: {cita.Subject}</p>"
                else:
                    # Si es un evento de todo el día, solo muestra el asunto
                    cuerpo_html += f"<p><strong>Todo el día</strong>: {cita.Subject}</p>"
                events_found = True

        # Agrega un mensaje si no se encontraron eventos
        if not events_found:
            cuerpo_html += "<p>No hay eventos programados para hoy.</p>"

        # Crea el correo electrónico y lo envía
        mail = outlook.CreateItem(0)  # 0 es el índice para un correo
        mail.To = "tuemail@aqui.com"  # CAMBIA ESTO
        mail.Subject = f"Agenda Diaria: {today.strftime('%d/%m/%Y')}"
        mail.HTMLBody = cuerpo_html

        mail.Send()
        print("Correo de la agenda enviado con éxito.")

    except Exception as e:
        print(f"Error al ejecutar el script: {e}")


if __name__ == "__main__":
    send_daily_outlook_agenda()

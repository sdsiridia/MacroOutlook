'''Script para extraer la agenda diaria de Outlook y enviarla por correo electrónico.'''
from datetime import datetime  # , timedelta
import win32com.client


def send_daily_outlook_agenda():
    """
    Se conecta a Outlook, extrae los eventos del día y los envía por correo.
    """
    try:
        # Crea un objeto de la aplicación de Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
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
        mail.To = "buzonsugerencias@asociacionnuevavida.org"  # CAMBIA ESTO
        mail.Subject = f"Agenda Diaria: {today.strftime('%d/%m/%Y')}"
        mail.HTMLBody = cuerpo_html

        mail.Send()
        print("Correo de la agenda enviado con éxito.")

    except Exception as e:
        print(f"Error al ejecutar el script: {e}")


if __name__ == "__main__":
    send_daily_outlook_agenda()

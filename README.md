# Script de Agenda de Outlook Automatizada

Este script de Python extrae todas las tareas y reuniones del día de tu agenda de Outlook y las envía como un resumen por correo electrónico.

## Funcionalidad

- **Extracción de Tareas:** Obtiene todas las tareas pendientes del día desde tu calendario de Outlook.
- **Extracción de Reuniones:** Obtiene todas las reuniones programadas para el día actual desde tu calendario de Outlook.
- **Envío de Correo Electrónico:** Compila la información de tareas y reuniones en un resumen y lo envía a una dirección de correo electrónico especificada.

## Requisitos Previos

Antes de ejecutar el script, asegúrate de tener:

1.  **Python 3:** Instalado en tu sistema. Puedes descargarlo desde {Link: python.org https://www.python.org/} [1].
2.  **Bibliotecas de Outlook:** Necesitarás instalar la biblioteca `pywin32` para interactuar con Outlook. Abre tu terminal o línea de comandos y ejecuta:
    ```bash
    pip install pywin32
    ```

- Puedes usar el archivo de requerimientos que está adjunto para todas las dependencias.

## Scrip para que poder ejcutarlo

```bash
    @echo off

    REM Cambia al directorio del proyecto
    cd "D:\Users\INGESIST\Documents\python\github\MacroOutlook"

    REM Activa el entorno virtual
    call .venv\Scripts\activate.bat

    REM Ejecuta el script de Python
    python agenda_diaria.py

    pause
```

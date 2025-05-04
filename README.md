# Automatización de Eventos en Outlook desde Excel

## 📝 Introducción
Este documento describe el proceso de automatización mediante un script que genera eventos en Microsoft Outlook a partir de un archivo Excel. La finalidad es recibir un correo electrónico como recordatorio de cada acontecimiento detallado en la hoja de cálculo, evitando que pasemos por alto eventos importantes.

## 📂 Estructura del Archivo Excel
El archivo Excel debe contener las siguientes columnas con información relevante:

- **Asunto**: Título del evento.
- **Fecha**: Día en el que se debe agendar el recordatorio.
- **Hora**: Momento exacto en que se debe enviar el aviso.
- **Detalle**: Descripción adicional sobre el evento.
- **Destinatarios**: Correos electrónicos de las personas que recibirán el aviso.

Ejemplo de estructura del Excel:

| Asunto          | Fecha       | Hora  | Detalle                  | Destinatarios                  |
|---------------|------------|------|--------------------------|--------------------------------|
| Reunión equipo | 2025-05-10 | 10:00 | Evaluación de estrategias | ejemplo@empresa.com, otro@empresa.com |
| Entrega informe | 2025-05-15 | 15:30 | Envío del reporte mensual | gerente@empresa.com           |

## ⚙️ Funcionamiento del Script
El script recorrerá cada línea del Excel y generará un evento en Outlook con los siguientes pasos:

1. **Lectura del archivo Excel**: Se extraen los datos de cada fila.
2. **Creación del evento en Outlook**: Se utilizan los datos obtenidos para agendar el evento.
3. **Envío de correo de aviso**: Outlook enviará un correo a los destinatarios como recordatorio.

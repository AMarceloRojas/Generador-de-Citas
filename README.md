# Generador de Citas Masivas para Transportistas

Este proyecto automatiza la generación y envío de correos para gestionar el acceso de transportistas a diversas sedes logísticas, utilizando Google Sheets y Google Apps Script.

# ¿Cómo funciona?

1. *Carga de datos de ingreso:* 
   Se pegan los datos generales (fecha, pedido, dirección, correo, etc.) en la hoja principal.

2. *Carga de transportistas:*  
   En una segunda hoja, se insertan datos de los transportes como placa, chofer, licencia y auxiliares.

3. *Automatización disponible desde un menú personalizado:*
   - `Limpiar Datos Generales` y `Limpiar Todas las Hojas`: para reiniciar el archivo.
   - `Crear Carpetas de Transporte`: genera carpetas por placa para añadir archivos necesarios.
   - `Enviar Correos`: envía automáticamente correos masivos a los destinatarios con formato HTML y datos personalizados (placa, nombre del chofer, dirección, etc.).

# Ejemplo del correo generado

El sistema envía correos con un formato empresarial (como se muestra en la imagen adjunta) a cada transportista o grupo, adjuntando archivos y con todos los datos generados dinámicamente desde la base.

# Tecnologías utilizadas

- Google Sheets
- Google Apps Script
- GmailApp
- HTML para cuerpo del correo

# Casos de uso

- Gestión diaria de transportes para entregas logísticas
- Envío masivo con plantillas empresariales
- Reducción de errores humanos y tiempo en la preparación de correos

# Autor

*Anthtonny Marcelo Rojas* 

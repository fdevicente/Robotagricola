# Bitácora de Desarrollo - Bot Agrícola Santa Elisa

## Objetivo General
Automatizar la recepción y extracción de datos de facturas (imágenes y PDFs) enviadas por Telegram, utilizando Inteligencia Artificial local para leer los datos y consolidarlos automáticamente en un archivo Excel.

## Lo que hemos logrado en esta sesión:

1. **Configuración Base y Telegram:** 
   - Conexión exitosa del Bot de Telegram para recibir fotos y archivos PDF directamente en el chat.
   - Configuración de rutas automatizadas (`.env`) para descargar las imágenes en tu carpeta local de "Facturas Recibidas por Telegram".

2. **Integración de Motores de Inteligencia Artificial (IA Local):**
   - Integramos **Ollama** de forma local para no depender de internet y proteger la privacidad de los datos.
   - Empezamos usando `llava`, pero lo actualizamos al modelo avanzado **`llama3.2-vision`** (11B) que ya tenías descargado, el cual es muy superior leyendo tablas de facturas.

3. **Arquitectura "Híbrida" (Visión + OCR Tradicional):**
   - Notamos que a veces las IA visuales inventaban números ("alucinaban") o copiaban el ejemplo al pie de la letra.
   - Lo solucionamos creando un motor dual: Primero, un OCR tradicional (Tesseract) escanea la imagen y saca el texto crudo. Luego, nuestro bot le "sopla" ese texto a `llama3.2-vision` como pista. Esto garantiza que la IA lea y estruture exactamente lo que dice el papel.

4. **Corrección Matemática y de Lógica (Python):**
   - Le quitamos la responsabilidad matemática a la IA. Ahora el modelo solo saca la cantidad y el valor unitario. Mi código en Python multiplica los valores, les calcula el 19% de IVA y cuadra el total con precisión absoluta.
   - Si la factura no trae Fecha de Vencimiento, el bot se la calcula automáticamente sumando exactamente 1 mes a la fecha de emisión.
   - Logramos que si una factura trae múltiples ítems (como la factura de Mecánica), el bot la desglose inteligentemente en 2 o más filas independientes en el Excel.

5. **Manejo Seguro de Excel y Comando `/deshacer`:**
   - El bot ahora inserta pacíficamente las filas nuevas en el archivo `Facturas recibidas.xlsx`.
   - Implementamos el comando salvavidas **`/deshacer`** en Telegram. Si envías una factura y notas que los datos salieron mal, escribes el comando y el bot borra mágicamente las filas que acababa de escribir en el Excel y la foto descargada.

## Próximos Pasos (Objetivos pendientes):
- [ ] Retomar las pruebas enviando fotos y PDFs de facturas variadas al bot para poner a prueba el nuevo sistema híbrido (Llama 3.2 Vision + OCR).
- [ ] Comprobar que el comando `/deshacer` funcione fluido en el día a día sin que se queden procesos de fondo colgados.
- [ ] Continuar iterando el prompt o el formato si topamos con alguna factura demasiado exótica que quiebre el sistema actual.

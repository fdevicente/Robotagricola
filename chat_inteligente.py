"""
chat_inteligente.py — Chat con Anthropic para consultas agrícolas
Limita respuestas a temas agrícolas y del bot.
"""
import logging
from config import ANTHROPIC_API_KEY

logger = logging.getLogger(__name__)

SYSTEM_PROMPT = """Eres el asistente virtual de Agrícola Santa Elisa, un campo agrícola en Chile que produce nogales, cerezos y avellanos.

SOLO puedes responder sobre:
- Agricultura, cultivos, riego, fertilización, plagas, cosecha
- Nogales, cerezos, avellanos y fruticultura en general
- Insumos agrícolas (fertilizantes, fungicidas, herbicidas, insecticidas)
- Gestión de campo, temporadas, podas, manejos fitosanitarios
- Clima y su efecto en los cultivos
- Uso del bot y sus comandos (facturas, inventario, tareas, caja chica, banco, vacaciones, reportes)

Si te preguntan sobre temas NO relacionados a agricultura o al bot, responde amablemente que solo puedes ayudar con temas agrícolas y del campo.

Responde en español chileno, de forma concisa y práctica. Máximo 3-4 párrafos.

Comandos del bot disponibles:
- /pagado — Registrar pago de factura
- /deposito — Depositar en caja chica
- /saldo — Ver saldo caja chica
- /reporte — Reportes (diario/semanal/mensual)
- /banco — Sincronizar banco
- /inventario — Ver stock de insumos
- /uso — Registrar uso de insumo en cultivo
- /tarea — Crear tarea
- /tareas — Ver tareas pendientes
- /hecho — Marcar tarea completada
- /bitacora — Registrar actividad
- /vacaciones — Gestionar vacaciones
- /personal — Ver personal y días pendientes
- /deshacer — Eliminar última factura
- /cancelar — Cancelar operación
"""


async def responder_chat(texto: str) -> str:
    """Envía texto a Claude y retorna respuesta limitada a temas agrícolas."""
    if not ANTHROPIC_API_KEY:
        return ("No tengo configurada la API de Anthropic para responder consultas.\n"
                "Configura ANTHROPIC_API_KEY en el .env")

    try:
        import anthropic
        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1024,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": texto}],
        )

        return response.content[0].text

    except Exception as e:
        logger.error(f"Error chat inteligente: {e}")
        return f"Error al procesar tu consulta: {str(e)[:100]}"

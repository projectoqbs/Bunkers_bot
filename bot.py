import os
import json
import logging
import httpx
from telegram import Update, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from datetime import datetime

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "")
ANTHROPIC_KEY  = os.environ.get("ANTHROPIC_API_KEY", "")

SYSTEM_PROMPT = """Eres el Agente Bunkers QBS, asistente operativo de CI Quality Bunkers Supply S.A.S para gestión de suministro de combustible a buques en Colombia.

Tu trabajo es recopilar datos de operaciones de bunkering conversacionalmente para generar documentos DIMAR.

DATOS QUE NECESITAS:
- NOMBRE_BUQUE
- IMO
- BANDERA
- CANTIDAD_VLSO (MT de VLSFO 380, puede ser vacío)
- CANTIDAD_MGO (MT de MGO, puede ser vacío)
- PUERTO (SPRB / PALERMO / SPSM)
- AGENCIA
- ETA (fecha estimada)
- CIUDAD_OPERACION (defecto: MALAMBO)

REGLAS:
- Habla en español, tono profesional y directo
- Si el usuario da varios datos juntos, extráelos todos de una vez
- Solo pregunta por datos que realmente falten (IMO y cantidad son obligatorios)
- Cuando tengas todos los datos principales muestra EXACTAMENTE este formato:

RESUMEN:
🚢 Buque: [NOMBRE]
🏳️ Bandera: [BANDERA]
🔢 IMO: [IMO]
⛽ Combustible: [CANTIDAD] MT [TIPO]
⚓ Puerto: [PUERTO]
🏢 Agencia: [AGENCIA]
📅 ETA: [ETA]
📍 Ciudad: [CIUDAD]

¿Confirmas el registro? Responde SÍ para registrar o dime qué corregir.

CUANDO EL USUARIO CONFIRME con sí/confirmo/ok/correcto responde EXACTAMENTE:
REGISTRO_CONFIRMADO
JSON:{"buque":"X","imo":"X","bandera":"X","cantidad_vlso":"X","cantidad_mgo":"X","puerto":"X","agencia":"X","eta":"X","ciudad":"X"}

Comandos especiales que debes reconocer:
- "listar" o "lista" → muestra los buques registrados en la sesión
- "ayuda" o "help" → explica el flujo
- "nuevo" → inicia un nuevo registro limpiando contexto
"""

# Almacenamiento en memoria por usuario
user_sessions = {}
registros_sesion = []

def get_session(user_id):
    if user_id not in user_sessions:
        user_sessions[user_id] = []
    return user_sessions[user_id]

def clear_session(user_id):
    user_sessions[user_id] = []

async def call_claude(messages: list, system: str) -> str:
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": ANTHROPIC_KEY,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json={
                "model": "claude-sonnet-4-20250514",
                "max_tokens": 1000,
                "system": system,
                "messages": messages,
            }
        )
        data = resp.json()
        return data["content"][0]["text"]

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    clear_session(user_id)
    keyboard = [
        [KeyboardButton("🚢 Nuevo buque"), KeyboardButton("📋 Listar buques")],
        [KeyboardButton("❓ Ayuda"),        KeyboardButton("🔄 Limpiar sesión")],
    ]
    markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text(
        "👋 *Bienvenido al Agente Bunkers QBS*\n\n"
        "Soy tu asistente para registrar operaciones de suministro de combustible a buques.\n\n"
        "📝 Para comenzar escríbeme el nombre del buque o usa los botones del menú.",
        parse_mode="Markdown",
        reply_markup=markup
    )

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id  = update.effective_user.id
    text     = update.message.text.strip()
    session  = get_session(user_id)

    # Comandos de teclado
    if text in ["🔄 Limpiar sesión", "limpiar"]:
        clear_session(user_id)
        await update.message.reply_text("✅ Sesión limpiada. Escribe el nombre de un buque para iniciar.")
        return

    if text in ["📋 Listar buques", "listar", "lista"]:
        if not registros_sesion:
            await update.message.reply_text("📋 No hay buques registrados en esta sesión aún.")
        else:
            lista = "📋 *Buques registrados esta sesión:*\n\n"
            for i, r in enumerate(registros_sesion, 1):
                lista += f"{i}. 🚢 *{r.get('buque','?')}* — IMO {r.get('imo','?')} — {r.get('puerto','?')}\n"
            await update.message.reply_text(lista, parse_mode="Markdown")
        return

    if text in ["🚢 Nuevo buque", "nuevo"]:
        clear_session(user_id)
        await update.message.reply_text("🚢 Listo, cuéntame: ¿cuál es el nombre del buque y los datos que tienes?")
        return

    if text in ["❓ Ayuda", "ayuda", "help"]:
        await update.message.reply_text(
            "ℹ️ *Cómo funciona el Agente Bunkers:*\n\n"
            "1️⃣ Escríbeme el nombre del buque y los datos que tengas\n"
            "2️⃣ Te pregunto lo que falte (IMO, cantidad, puerto)\n"
            "3️⃣ Te muestro un resumen para confirmar\n"
            "4️⃣ Al confirmar, queda registrado\n"
            "5️⃣ Desde el Excel ejecutas la macro para generar los Word y enviar a DIMAR\n\n"
            "💡 Puedes dar todos los datos de una vez:\n"
            "_CTI QUEEN, IMO 9240079, Panamá, 650 MT VLSO, SPRB, NAVES_",
            parse_mode="Markdown"
        )
        return

    # Enviar a Claude con historial
    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")

    session.append({"role": "user", "content": text})

    try:
        response = await call_claude(session, SYSTEM_PROMPT)
        session.append({"role": "assistant", "content": response})

        # Detectar registro confirmado
        if "REGISTRO_CONFIRMADO" in response:
            try:
                json_start = response.index("JSON:") + 5
                json_str   = response[json_start:].strip()
                # Extraer solo el JSON
                end = json_str.index("}") + 1
                data = json.loads(json_str[:end])
                registros_sesion.append(data)
                clear_session(user_id)  # Limpiar para siguiente buque

                buque_txt = (
                    f"✅ *Registro guardado exitosamente*\n\n"
                    f"🚢 *{data.get('buque','?')}*\n"
                    f"🔢 IMO: {data.get('imo','?')}\n"
                    f"⚓ Puerto: {data.get('puerto','?')}\n\n"
                    f"📌 *Próximo paso:* Abre el Excel, verás este buque en la tabla. "
                    f"Pon *ENVIAR* en la columna ACCION y ejecuta la macro para generar "
                    f"los documentos Word y enviar a DIMAR.\n\n"
                    f"¿Hay otro buque que registrar?"
                )
                await update.message.reply_text(buque_txt, parse_mode="Markdown")
                return
            except Exception as e:
                logger.error(f"Error parsing JSON: {e}")

        # Respuesta normal
        await update.message.reply_text(response, parse_mode="Markdown")

    except Exception as e:
        logger.error(f"Error calling Claude: {e}")
        await update.message.reply_text(
            "⚠️ Error de conexión. Por favor intenta de nuevo en unos segundos."
        )

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    logger.info("Bot iniciado...")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()

import os
import json
import logging
import httpx
import gspread
from google.oauth2.service_account import Credentials
from telegram import Update, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from datetime import datetime
import io

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN  = os.environ.get("TELEGRAM_TOKEN", "")
GROQ_API_KEY    = os.environ.get("GROQ_API_KEY", "")
GOOGLE_CREDS    = os.environ.get("GOOGLE_CREDS", "")
SHEET_ID        = "1cKS-P5T9hO3Ayv78gAGC8i_3JhOmDxlothPfxiWToCY"

SYSTEM_PROMPT = """Eres el Agente Bunkers QBS, asistente operativo de CI Quality Bunkers Supply S.A.S para gestión de suministro de combustible a buques en Colombia.

Tu trabajo es recopilar datos de operaciones de bunkering conversacionalmente.

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

Comandos especiales:
- "listar" → muestra buques registrados en la sesión
- "ayuda" → explica el flujo
- "nuevo" → inicia nuevo registro
"""

user_sessions  = {}
registros_sesion = []

def get_sheet():
    creds_dict = json.loads(GOOGLE_CREDS)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds  = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    sh     = client.open_by_key(SHEET_ID)
    try:
        ws = sh.worksheet("BUQUES")
    except:
        ws = sh.add_worksheet(title="BUQUES", rows=1000, cols=15)
        ws.append_row([
            "FECHA REGISTRO","MN","AGENCIA","IMO","BANDERA",
            "MT VLSO","MT MGO","PUERTO","ETA","CIUDAD","ESTADO"
        ])
        # Formato encabezado
        ws.format("A1:K1", {
            "backgroundColor": {"red": 0.12, "green": 0.22, "blue": 0.39},
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "horizontalAlignment": "CENTER"
        })
    return ws

def get_session(uid):
    if uid not in user_sessions:
        user_sessions[uid] = []
    return user_sessions[uid]

def clear_session(uid):
    user_sessions[uid] = []

async def call_groq(history, system):
    messages = [{"role": "system", "content": system}]
    for m in history:
        messages.append({"role": m["role"], "content": m["content"]})
    payload = {
        "model": "llama-3.3-70b-versatile",
        "messages": messages,
        "max_tokens": 1000,
        "temperature": 0.7
    }
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.post(
            "https://api.groq.com/openai/v1/chat/completions",
            headers={"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"},
            json=payload
        )
        if resp.status_code == 429:
            raise Exception("RATE_LIMIT")
        if resp.status_code != 200:
            raise Exception(f"Groq error: {resp.status_code}")
        return resp.json()["choices"][0]["message"]["content"]

def registrar_en_sheet(data):
    try:
        ws = get_sheet()
        cantidad = ""
        if data.get("cantidad_vlso") and data.get("cantidad_vlso") != "X":
            cantidad_vlso = data["cantidad_vlso"]
        else:
            cantidad_vlso = ""
        if data.get("cantidad_mgo") and data.get("cantidad_mgo") != "X":
            cantidad_mgo = data["cantidad_mgo"]
        else:
            cantidad_mgo = ""

        fila = [
            datetime.now().strftime("%d/%m/%Y %H:%M"),
            data.get("buque", ""),
            data.get("agencia", ""),
            data.get("imo", ""),
            data.get("bandera", ""),
            cantidad_vlso,
            cantidad_mgo,
            data.get("puerto", ""),
            data.get("eta", ""),
            data.get("ciudad", "MALAMBO"),
            "PENDIENTE"
        ]
        ws.append_row(fila)
        
        # Obtener numero de fila recien agregada
        all_rows = ws.get_all_values()
        row_num = len(all_rows)
        
        # Colorear fila alternada
        color = {"red": 0.93, "green": 0.96, "blue": 1.0} if row_num % 2 == 0 else {"red": 1.0, "green": 1.0, "blue": 1.0}
        ws.format(f"A{row_num}:K{row_num}", {"backgroundColor": color})
        
        return True, row_num
    except Exception as e:
        logger.error(f"Error Google Sheets: {e}")
        return False, 0

def generar_imagen_registro(data):
    """Genera texto formateado del registro para enviar como mensaje"""
    cantidad_vlso = data.get("cantidad_vlso", "")
    cantidad_mgo  = data.get("cantidad_mgo", "")
    
    if cantidad_vlso and cantidad_vlso != "X" and cantidad_mgo and cantidad_mgo != "X":
        comb = f"{cantidad_vlso} MT VLSFO + {cantidad_mgo} MT MGO"
    elif cantidad_vlso and cantidad_vlso != "X":
        comb = f"{cantidad_vlso} MT VLSFO 380"
    elif cantidad_mgo and cantidad_mgo != "X":
        comb = f"{cantidad_mgo} MT MGO"
    else:
        comb = "A confirmar"

    texto = (
        f"📋 *REGISTRO EN GOOGLE SHEETS*\n"
        f"{'─'*30}\n"
        f"🚢 *Buque:* {data.get('buque','')}\n"
        f"🏳️ *Bandera:* {data.get('bandera','')}\n"
        f"🔢 *IMO:* {data.get('imo','')}\n"
        f"⛽ *Combustible:* {comb}\n"
        f"⚓ *Puerto:* {data.get('puerto','')}\n"
        f"🏢 *Agencia:* {data.get('agencia','')}\n"
        f"📅 *ETA:* {data.get('eta','')}\n"
        f"📍 *Ciudad:* {data.get('ciudad','MALAMBO')}\n"
        f"📌 *Estado:* PENDIENTE\n"
        f"🕐 *Fecha:* {datetime.now().strftime('%d/%m/%Y %H:%M')}\n"
        f"{'─'*30}\n"
        f"✅ *Guardado en Google Sheets exitosamente*"
    )
    return texto

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    clear_session(uid)
    keyboard = [
        [KeyboardButton("🚢 Nuevo buque"), KeyboardButton("📋 Listar buques")],
        [KeyboardButton("❓ Ayuda"),        KeyboardButton("🔄 Limpiar sesión")],
    ]
    await update.message.reply_text(
        "👋 *Bienvenido al Agente Bunkers QBS*\n\n"
        "Soy tu asistente para registrar operaciones de suministro de combustible.\n\n"
        "📝 Escríbeme el nombre del buque para comenzar.",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid     = update.effective_user.id
    text    = update.message.text.strip()
    session = get_session(uid)

    if text in ["🔄 Limpiar sesión", "limpiar"]:
        clear_session(uid)
        await update.message.reply_text("✅ Sesión limpiada.")
        return

    if text in ["📋 Listar buques", "listar", "lista"]:
        if not registros_sesion:
            await update.message.reply_text("📋 No hay buques registrados aún.")
        else:
            lista = "📋 *Buques registrados esta sesión:*\n\n"
            for i, r in enumerate(registros_sesion, 1):
                lista += f"{i}. 🚢 *{r.get('buque','?')}* — IMO {r.get('imo','?')} — {r.get('puerto','?')}\n"
            await update.message.reply_text(lista, parse_mode="Markdown")
        return

    if text in ["🚢 Nuevo buque", "nuevo"]:
        clear_session(uid)
        await update.message.reply_text("🚢 Cuéntame los datos del buque:")
        return

    if text in ["❓ Ayuda", "ayuda"]:
        await update.message.reply_text(
            "ℹ️ *Cómo funciona:*\n\n"
            "1️⃣ Escríbeme los datos del buque\n"
            "2️⃣ Te pregunto lo que falte\n"
            "3️⃣ Te muestro resumen para confirmar\n"
            "4️⃣ Al confirmar se guarda en Google Sheets\n"
            "5️⃣ Te muestro cómo quedó el registro\n\n"
            "💡 Ejemplo:\n"
            "_CTI QUEEN, IMO 9240079, Panamá, 650 MT VLSO, SPRB, NAVES_",
            parse_mode="Markdown"
        )
        return

    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")
    session.append({"role": "user", "content": text})

    try:
        response = await call_groq(session, SYSTEM_PROMPT)
        session.append({"role": "assistant", "content": response})

        if "REGISTRO_CONFIRMADO" in response:
            try:
                json_start = response.index("JSON:") + 5
                json_str   = response[json_start:].strip()
                data       = json.loads(json_str[:json_str.index("}")+1])
                
                # Guardar en Google Sheets
                await update.message.reply_text("⏳ Guardando en Google Sheets...")
                ok, row_num = registrar_en_sheet(data)
                
                if ok:
                    registros_sesion.append(data)
                    clear_session(uid)
                    
                    # Enviar resumen del registro
                    resumen = generar_imagen_registro(data)
                    await update.message.reply_text(resumen, parse_mode="Markdown")
                    
                    await update.message.reply_text(
                        f"🔗 Ver Google Sheets:\n"
                        f"https://docs.google.com/spreadsheets/d/{SHEET_ID}\n\n"
                        f"¿Hay otro buque que registrar?",
                        disable_web_page_preview=True
                    )
                else:
                    await update.message.reply_text(
                        "⚠️ No se pudo guardar en Google Sheets. "
                        "Verifique la configuración e intente de nuevo."
                    )
                return
            except Exception as e:
                logger.error(f"Error procesando registro: {e}")

        await update.message.reply_text(response, parse_mode="Markdown")

    except Exception as e:
        logger.error(f"Error: {e}")
        if "RATE_LIMIT" in str(e):
            await update.message.reply_text("⏳ Espera unos segundos e intenta de nuevo.")
        else:
            await update.message.reply_text("⚠️ Error de conexión. Intenta de nuevo.")

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    logger.info("Bot Bunkers QBS con Google Sheets iniciado...")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()

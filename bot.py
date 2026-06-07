import os
import json
import logging
import httpx
import gspread
from google.oauth2.service_account import Credentials
from telegram import Update, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from datetime import datetime

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "")
GROQ_API_KEY   = os.environ.get("GROQ_API_KEY", "")
GOOGLE_CREDS   = os.environ.get("GOOGLE_CREDS", "")
SHEET_ID       = "1cKS-P5T9hO3Ayv78gAGC8i_3JhOmDxlothPfxiWToCY"

# Estructura de columnas
HEADERS = [
    "FECHA REGISTRO", "MN", "ETA", "AGENCIA", "IMO", "BANDERA",
    "MT VLSO", "MT HSFO", "MT MGO", "PUERTO", "HORAS OP.", "CIUDAD", "ESTADO"
]

SYSTEM_PROMPT = """Eres el Agente Bunkers QBS, asistente operativo de CI Quality Bunkers Supply S.A.S para gestión de suministro de combustible a buques en Colombia.

Tu trabajo es recopilar datos de operaciones de bunkering conversacionalmente.

DATOS QUE NECESITAS:
- NOMBRE_BUQUE (obligatorio)
- IMO (obligatorio)
- BANDERA
- ETA (fecha estimada llegada)
- AGENCIA
- MT_VLSO (MT de VLSFO 380, puede ser vacío)
- MT_HSFO (MT de HSFO, puede ser vacío)
- MT_MGO (MT de MGO, puede ser vacío)
- PUERTO (SPRB / PALERMO / SPSM)
- HORAS_OP (horas estimadas de operación, puede ser vacío)
- CIUDAD_OPERACION (defecto: MALAMBO)

REGLAS:
- Habla en español, tono profesional y directo
- Si el usuario da varios datos juntos, extráelos todos de una vez
- Al menos uno de MT_VLSO, MT_HSFO o MT_MGO debe tener valor
- Cuando tengas todos los datos principales muestra EXACTAMENTE este formato:

RESUMEN:
🚢 Buque: [NOMBRE]
📅 ETA: [ETA]
🏳️ Bandera: [BANDERA]
🔢 IMO: [IMO]
🏢 Agencia: [AGENCIA]
⛽ MT VLSO: [MT_VLSO o "—"]
⛽ MT HSFO: [MT_HSFO o "—"]
⛽ MT MGO: [MT_MGO o "—"]
⚓ Puerto: [PUERTO]
⏱️ Horas op.: [HORAS_OP o "—"]
📍 Ciudad: [CIUDAD]

¿Confirmas el registro? Responde SÍ para registrar o dime qué corregir.

CUANDO EL USUARIO CONFIRME con sí/confirmo/ok/correcto responde EXACTAMENTE:
REGISTRO_CONFIRMADO
JSON:{"buque":"X","eta":"X","agencia":"X","imo":"X","bandera":"X","mt_vlso":"X","mt_hsfo":"X","mt_mgo":"X","puerto":"X","horas_op":"X","ciudad":"X"}

Comandos especiales:
- "listar" → muestra buques registrados en la sesión
- "ayuda" → explica el flujo
- "nuevo" → inicia nuevo registro
"""

user_sessions    = {}
registros_sesion = []

def get_session(uid):
    if uid not in user_sessions:
        user_sessions[uid] = []
    return user_sessions[uid]

def clear_session(uid):
    user_sessions[uid] = []

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
        # Verificar si los headers están actualizados
        current_headers = ws.row_values(1)
        if current_headers != HEADERS:
            ws.clear()
            ws.append_row(HEADERS)
            _format_header(ws)
    except:
        ws = sh.add_worksheet(title="BUQUES", rows=1000, cols=15)
        ws.append_row(HEADERS)
        _format_header(ws)
    return ws

def _format_header(ws):
    ws.format("A1:M1", {
        "backgroundColor": {"red": 0.12, "green": 0.22, "blue": 0.39},
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
        "horizontalAlignment": "CENTER"
    })

def registrar_en_sheet(data):
    try:
        ws = get_sheet()
        fila = [
            datetime.now().strftime("%d/%m/%Y %H:%M"),
            data.get("buque", ""),
            data.get("eta", ""),
            data.get("agencia", ""),
            data.get("imo", ""),
            data.get("bandera", ""),
            data.get("mt_vlso", "") if data.get("mt_vlso", "") not in ["X", "0", ""] else "",
            data.get("mt_hsfo", "") if data.get("mt_hsfo", "") not in ["X", "0", ""] else "",
            data.get("mt_mgo",  "") if data.get("mt_mgo",  "") not in ["X", "0", ""] else "",
            data.get("puerto", ""),
            data.get("horas_op", "") if data.get("horas_op", "") not in ["X", "0", ""] else "",
            data.get("ciudad", "MALAMBO"),
            "PENDIENTE"
        ]
        ws.append_row(fila)
        all_rows = ws.get_all_values()
        row_num  = len(all_rows)
        color = {"red": 0.93, "green": 0.96, "blue": 1.0} if row_num % 2 == 0 else {"red": 1.0, "green": 1.0, "blue": 1.0}
        ws.format(f"A{row_num}:M{row_num}", {"backgroundColor": color})
        return True, row_num
    except Exception as e:
        logger.error(f"Error Sheets: {e}")
        return False, 0

def editar_en_sheet(nombre_buque, campo, nuevo_valor):
    try:
        ws    = get_sheet()
        datos = ws.get_all_values()
        headers = datos[0]
        
        # Mapeo de nombres amigables a columnas
        campo_map = {
            "eta": "ETA", "agencia": "AGENCIA", "imo": "IMO",
            "bandera": "BANDERA", "vlso": "MT VLSO", "hsfo": "MT HSFO",
            "mgo": "MT MGO", "puerto": "PUERTO", "horas": "HORAS OP.",
            "ciudad": "CIUDAD", "estado": "ESTADO"
        }
        col_name = campo_map.get(campo.lower(), campo.upper())
        
        if col_name not in headers:
            return False, f"Campo '{col_name}' no encontrado"
        
        col_idx = headers.index(col_name) + 1
        
        # Buscar fila del buque (más reciente)
        row_idx = None
        for i in range(len(datos)-1, 0, -1):
            if datos[i][1].upper() == nombre_buque.upper():
                row_idx = i + 1
                break
        
        if not row_idx:
            return False, f"Buque '{nombre_buque}' no encontrado"
        
        ws.update_cell(row_idx, col_idx, nuevo_valor)
        return True, f"✅ Actualizado: {col_name} = {nuevo_valor}"
    except Exception as e:
        logger.error(f"Error editando: {e}")
        return False, str(e)

def generar_resumen_registro(data):
    def val(k):
        v = data.get(k, "")
        return v if v and v not in ["X", "0"] else "—"
    
    return (
        f"📋 *REGISTRO GUARDADO EN GOOGLE SHEETS*\n"
        f"{'─'*32}\n"
        f"🚢 *Buque:* {val('buque')}\n"
        f"📅 *ETA:* {val('eta')}\n"
        f"🏳️ *Bandera:* {val('bandera')}\n"
        f"🔢 *IMO:* {val('imo')}\n"
        f"🏢 *Agencia:* {val('agencia')}\n"
        f"⛽ *MT VLSO:* {val('mt_vlso')}\n"
        f"⛽ *MT HSFO:* {val('mt_hsfo')}\n"
        f"⛽ *MT MGO:* {val('mt_mgo')}\n"
        f"⚓ *Puerto:* {val('puerto')}\n"
        f"⏱️ *Horas op.:* {val('horas_op')}\n"
        f"📍 *Ciudad:* {val('ciudad')}\n"
        f"📌 *Estado:* PENDIENTE\n"
        f"🕐 *Fecha:* {datetime.now().strftime('%d/%m/%Y %H:%M')}\n"
        f"{'─'*32}\n"
        f"✅ *Guardado exitosamente*"
    )

async def call_groq(history, system):
    messages = [{"role": "system", "content": system}]
    for m in history:
        messages.append({"role": m["role"], "content": m["content"]})
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.post(
            "https://api.groq.com/openai/v1/chat/completions",
            headers={"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"},
            json={"model": "llama-3.3-70b-versatile", "messages": messages, "max_tokens": 1000, "temperature": 0.7}
        )
        if resp.status_code == 429: raise Exception("RATE_LIMIT")
        if resp.status_code != 200: raise Exception(f"Groq error: {resp.status_code}")
        return resp.json()["choices"][0]["message"]["content"]

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    clear_session(update.effective_user.id)
    keyboard = [
        [KeyboardButton("🚢 Nuevo buque"),   KeyboardButton("📋 Listar buques")],
        [KeyboardButton("✏️ Editar registro"), KeyboardButton("🔄 Limpiar sesión")],
        [KeyboardButton("❓ Ayuda")],
    ]
    await update.message.reply_text(
        "👋 *Bienvenido al Agente Bunkers QBS*\n\n"
        "Soy tu asistente para registrar operaciones de suministro de combustible.\n\n"
        "📝 Escríbeme los datos del buque para comenzar.",
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

    if text in ["📋 Listar buques", "listar"]:
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

    if text in ["✏️ Editar registro", "editar"]:
        await update.message.reply_text(
            "✏️ *Editar registro*\n\n"
            "Escríbeme en este formato:\n"
            "_editar [NOMBRE BUQUE] / [CAMPO] / [NUEVO VALOR]_\n\n"
            "Campos disponibles: eta, agencia, bandera, vlso, hsfo, mgo, puerto, horas, ciudad, estado\n\n"
            "Ejemplo:\n"
            "_editar CTI QUEEN / vlso / 700_",
            parse_mode="Markdown"
        )
        return

    if text in ["❓ Ayuda", "ayuda"]:
        await update.message.reply_text(
            "ℹ️ *Cómo funciona:*\n\n"
            "1️⃣ Escríbeme los datos del buque\n"
            "2️⃣ Te pregunto lo que falte\n"
            "3️⃣ Te muestro resumen para confirmar\n"
            "4️⃣ Al confirmar se guarda en Google Sheets\n"
            "5️⃣ Te muestro cómo quedó el registro\n\n"
            "✏️ Para editar: _editar CTI QUEEN / vlso / 700_\n\n"
            "💡 Ejemplo registro:\n"
            "_CTI QUEEN, IMO 9240079, Panamá, 650 MT VLSO, SPRB, NAVES, ETA 10/06_",
            parse_mode="Markdown"
        )
        return

    # Detectar comando editar
    if text.lower().startswith("editar ") and "/" in text:
        partes = text[7:].split("/")
        if len(partes) == 3:
            nombre   = partes[0].strip()
            campo    = partes[1].strip()
            nuevo_val = partes[2].strip()
            await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")
            ok, msg = editar_en_sheet(nombre, campo, nuevo_val)
            await update.message.reply_text(msg if ok else f"⚠️ {msg}")
            if ok:
                await update.message.reply_text(
                    f"🔗 Ver Google Sheets:\nhttps://docs.google.com/spreadsheets/d/{SHEET_ID}",
                    disable_web_page_preview=True
                )
        else:
            await update.message.reply_text("⚠️ Formato incorrecto. Usa: _editar NOMBRE / campo / valor_", parse_mode="Markdown")
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
                await update.message.reply_text("⏳ Guardando en Google Sheets...")
                ok, row_num = registrar_en_sheet(data)
                if ok:
                    registros_sesion.append(data)
                    clear_session(uid)
                    await update.message.reply_text(generar_resumen_registro(data), parse_mode="Markdown")
                    await update.message.reply_text(
                        f"🔗 Ver Google Sheets:\nhttps://docs.google.com/spreadsheets/d/{SHEET_ID}\n\n¿Hay otro buque?",
                        disable_web_page_preview=True
                    )
                else:
                    await update.message.reply_text("⚠️ No se pudo guardar. Intenta de nuevo.")
                return
            except Exception as e:
                logger.error(f"Error registro: {e}")

        # Limpiar caracteres markdown que pueden causar errores
        safe_response = response.replace("*", "").replace("_", "").replace("`", "").replace("[", "").replace("]", "")
        await update.message.reply_text(safe_response)

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
    logger.info("Bot Bunkers QBS iniciado...")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()

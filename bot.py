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

HEADERS = [
    "FECHA REGISTRO", "MN", "ETA", "AGENCIA", "IMO", "BANDERA",
    "MT VLSO", "MT HSFO", "MT MGO", "PUERTO", "HORAS OP.", "CIUDAD", "ESTADO"
]

CAMPO_MAP = {
    "1": ("ETA",       "eta"),
    "2": ("AGENCIA",   "agencia"),
    "3": ("IMO",       "imo"),
    "4": ("BANDERA",   "bandera"),
    "5": ("MT VLSO",   "mt_vlso"),
    "6": ("MT HSFO",   "mt_hsfo"),
    "7": ("MT MGO",    "mt_mgo"),
    "8": ("PUERTO",    "puerto"),
    "9": ("HORAS OP.", "horas_op"),
    "10":("CIUDAD",    "ciudad"),
    "11":("ESTADO",    "estado"),
}

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
Buque: [NOMBRE]
ETA: [ETA]
Bandera: [BANDERA]
IMO: [IMO]
Agencia: [AGENCIA]
MT VLSO: [MT_VLSO o ninguno]
MT HSFO: [MT_HSFO o ninguno]
MT MGO: [MT_MGO o ninguno]
Puerto: [PUERTO]
Horas op.: [HORAS_OP o ninguno]
Ciudad: [CIUDAD]

Confirmas el registro? Responde SI para registrar o dime que corregir.

CUANDO EL USUARIO CONFIRME con si/confirmo/ok/correcto responde EXACTAMENTE:
REGISTRO_CONFIRMADO
JSON:{"buque":"X","eta":"X","agencia":"X","imo":"X","bandera":"X","mt_vlso":"X","mt_hsfo":"X","mt_mgo":"X","puerto":"X","horas_op":"X","ciudad":"X"}

Comandos especiales:
- listar: muestra buques registrados en la sesion
- ayuda: explica el flujo
- nuevo: inicia nuevo registro
"""

user_sessions    = {}
registros_sesion = []

# Estados para edicion paso a paso
edit_states   = {}
delete_states = {}
# edit_states[uid]   = {"step": "buque"|"campo"|"valor", "buque": "...", "campo_num": "..."}
# delete_states[uid] = {"step": "buque"|"confirmar", "buque": "...", "fila": {...}}

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
        def v(k):
            val = data.get(k, "")
            return val if val and val not in ["X", "0", "ninguno", "none", "-"] else ""
        fila = [
            datetime.now().strftime("%d/%m/%Y %H:%M"),
            data.get("buque", ""),
            v("eta"), v("agencia"), v("imo"), v("bandera"),
            v("mt_vlso"), v("mt_hsfo"), v("mt_mgo"),
            v("puerto"), v("horas_op"),
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

def buscar_buque_en_sheet(nombre_buque):
    """Busca el buque mas reciente y devuelve (row_idx, fila_data) o (None, None)"""
    try:
        ws    = get_sheet()
        datos = ws.get_all_values()
        headers = datos[0]
        for i in range(len(datos)-1, 0, -1):
            if datos[i][1].upper() == nombre_buque.upper():
                fila_dict = dict(zip(headers, datos[i]))
                return i + 1, fila_dict
        return None, None
    except Exception as e:
        logger.error(f"Error buscando buque: {e}")
        return None, None

def editar_campo_en_sheet(nombre_buque, col_name, nuevo_valor):
    try:
        ws    = get_sheet()
        datos = ws.get_all_values()
        headers = datos[0]
        if col_name not in headers:
            return False, f"Campo '{col_name}' no encontrado"
        col_idx = headers.index(col_name) + 1
        row_idx = None
        for i in range(len(datos)-1, 0, -1):
            if datos[i][1].upper() == nombre_buque.upper():
                row_idx = i + 1
                break
        if not row_idx:
            return False, f"Buque '{nombre_buque}' no encontrado"
        ws.update_cell(row_idx, col_idx, nuevo_valor)
        return True, "ok"
    except Exception as e:
        logger.error(f"Error editando: {e}")
        return False, str(e)

def generar_resumen_registro(data):
    def v(k):
        val = data.get(k, "")
        return val if val and val not in ["X", "0", "ninguno", "none", "-"] else "—"
    return (
        f"REGISTRO GUARDADO EN GOOGLE SHEETS\n"
        f"{'─'*32}\n"
        f"Buque: {v('buque')}\n"
        f"ETA: {v('eta')}\n"
        f"Bandera: {v('bandera')}\n"
        f"IMO: {v('imo')}\n"
        f"Agencia: {v('agencia')}\n"
        f"MT VLSO: {v('mt_vlso')}\n"
        f"MT HSFO: {v('mt_hsfo')}\n"
        f"MT MGO: {v('mt_mgo')}\n"
        f"Puerto: {v('puerto')}\n"
        f"Horas op.: {v('horas_op')}\n"
        f"Ciudad: {v('ciudad')}\n"
        f"Estado: PENDIENTE\n"
        f"Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M')}\n"
        f"{'─'*32}\n"
        f"Guardado exitosamente"
    )

def eliminar_buque_en_sheet(nombre_buque):
    try:
        ws    = get_sheet()
        datos = ws.get_all_values()
        row_idx = None
        for i in range(len(datos)-1, 0, -1):
            if datos[i][1].upper() == nombre_buque.upper():
                row_idx = i + 1
                break
        if not row_idx:
            return False, f"Buque '{nombre_buque}' no encontrado"
        ws.delete_rows(row_idx)
        return True, "ok"
    except Exception as e:
        logger.error(f"Error eliminando: {e}")
        return False, str(e)

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

def menu_principal():
    keyboard = [
        [KeyboardButton("Nuevo buque"),     KeyboardButton("Listar buques")],
        [KeyboardButton("Editar registro"), KeyboardButton("Eliminar buque")],
        [KeyboardButton("Limpiar sesion"),  KeyboardButton("Ayuda")],
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

def menu_campos():
    keyboard = [
        [KeyboardButton("1. ETA"),        KeyboardButton("2. Agencia")],
        [KeyboardButton("3. IMO"),         KeyboardButton("4. Bandera")],
        [KeyboardButton("5. MT VLSO"),    KeyboardButton("6. MT HSFO")],
        [KeyboardButton("7. MT MGO"),      KeyboardButton("8. Puerto")],
        [KeyboardButton("9. Horas op."),  KeyboardButton("10. Ciudad")],
        [KeyboardButton("11. Estado"),    KeyboardButton("Cancelar edicion")],
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    clear_session(uid)
    edit_states.pop(uid, None)
    await update.message.reply_text(
        "Bienvenido al Agente Bunkers QBS\n\n"
        "Soy tu asistente para registrar operaciones de suministro de combustible.\n\n"
        "Escribeme los datos del buque para comenzar.",
        reply_markup=menu_principal()
    )

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid  = update.effective_user.id
    text = update.message.text.strip()

    # ── FLUJO DE ELIMINACION PASO A PASO ────────────────────────
    if uid in delete_states:
        estado = delete_states[uid]

        if text in ["Cancelar", "cancelar", "Cancelar edicion"]:
            delete_states.pop(uid, None)
            await update.message.reply_text("Operacion cancelada.", reply_markup=menu_principal())
            return

        # Paso 1: recibir nombre del buque
        if estado["step"] == "buque":
            row_idx, fila = buscar_buque_en_sheet(text)
            if not row_idx:
                await update.message.reply_text(
                    f"No encontre el buque '{text}' en Google Sheets.\n"
                    f"Verifique el nombre e intente de nuevo o escriba 'Cancelar'."
                )
                return
            delete_states[uid]["buque"] = text
            delete_states[uid]["fila"]  = fila
            delete_states[uid]["step"]  = "confirmar"

            resumen = (
                f"Buque encontrado:\n"
                f"{'─'*28}\n"
                f"MN:      {fila.get('MN','')}\n"
                f"ETA:     {fila.get('ETA','—')}\n"
                f"IMO:     {fila.get('IMO','—')}\n"
                f"Puerto:  {fila.get('PUERTO','—')}\n"
                f"MT VLSO: {fila.get('MT VLSO','—')}\n"
                f"MT HSFO: {fila.get('MT HSFO','—')}\n"
                f"MT MGO:  {fila.get('MT MGO','—')}\n"
                f"Estado:  {fila.get('ESTADO','—')}\n"
                f"{'─'*28}\n"
                f"ATENCION: Esta accion eliminara este registro permanentemente.\n"
                f"Confirma la eliminacion?"
            )
            keyboard = ReplyKeyboardMarkup(
                [[KeyboardButton("SI, ELIMINAR"), KeyboardButton("Cancelar")]],
                resize_keyboard=True
            )
            await update.message.reply_text(resumen, reply_markup=keyboard)
            return

        # Paso 2: confirmar eliminacion
        if estado["step"] == "confirmar":
            if text in ["SI, ELIMINAR", "si, eliminar", "SI"]:
                buque = delete_states[uid]["buque"]
                delete_states.pop(uid, None)
                ok, msg = eliminar_buque_en_sheet(buque)
                if ok:
                    await update.message.reply_text(
                        f"Registro eliminado exitosamente\n\n"
                        f"Buque: {buque}\n\n"
                        f"Ver Google Sheets:\nhttps://docs.google.com/spreadsheets/d/{SHEET_ID}",
                        reply_markup=menu_principal()
                    )
                else:
                    await update.message.reply_text(f"Error al eliminar: {msg}", reply_markup=menu_principal())
            else:
                delete_states.pop(uid, None)
                await update.message.reply_text("Eliminacion cancelada.", reply_markup=menu_principal())
            return


    if uid in edit_states:
        estado = edit_states[uid]

        if text == "Cancelar edicion":
            edit_states.pop(uid, None)
            await update.message.reply_text("Edicion cancelada.", reply_markup=menu_principal())
            return

        # Paso 1: recibir nombre del buque
        if estado["step"] == "buque":
            row_idx, fila = buscar_buque_en_sheet(text)
            if not row_idx:
                await update.message.reply_text(
                    f"No encontre el buque '{text}' en Google Sheets.\n"
                    f"Verifique el nombre exacto e intente de nuevo, o escriba 'Cancelar edicion'."
                )
                return
            edit_states[uid]["buque"]   = text
            edit_states[uid]["row_idx"] = row_idx
            edit_states[uid]["step"]    = "campo"

            # Mostrar datos actuales
            resumen = (
                f"Buque encontrado: {fila.get('MN','')}\n"
                f"{'─'*28}\n"
                f"1.  ETA:       {fila.get('ETA','—')}\n"
                f"2.  Agencia:   {fila.get('AGENCIA','—')}\n"
                f"3.  IMO:       {fila.get('IMO','—')}\n"
                f"4.  Bandera:   {fila.get('BANDERA','—')}\n"
                f"5.  MT VLSO:   {fila.get('MT VLSO','—')}\n"
                f"6.  MT HSFO:   {fila.get('MT HSFO','—')}\n"
                f"7.  MT MGO:    {fila.get('MT MGO','—')}\n"
                f"8.  Puerto:    {fila.get('PUERTO','—')}\n"
                f"9.  Horas op.: {fila.get('HORAS OP.','—')}\n"
                f"10. Ciudad:    {fila.get('CIUDAD','—')}\n"
                f"11. Estado:    {fila.get('ESTADO','—')}\n"
                f"{'─'*28}\n"
                f"Que campo desea editar? Toque el boton o escriba el numero."
            )
            await update.message.reply_text(resumen, reply_markup=menu_campos())
            return

        # Paso 2: recibir campo a editar
        if estado["step"] == "campo":
            # Aceptar "N. Nombre" o solo "N"
            num = text.split(".")[0].strip()
            if num not in CAMPO_MAP:
                await update.message.reply_text(
                    "Por favor seleccione un campo del menu o escriba el numero (1-11).",
                    reply_markup=menu_campos()
                )
                return
            col_name, _ = CAMPO_MAP[num]
            edit_states[uid]["campo_num"] = num
            edit_states[uid]["col_name"]  = col_name
            edit_states[uid]["step"]      = "valor"
            await update.message.reply_text(
                f"Campo seleccionado: {col_name}\n\nEscriba el nuevo valor:",
                reply_markup=ReplyKeyboardMarkup([["Cancelar edicion"]], resize_keyboard=True)
            )
            return

        # Paso 3: recibir nuevo valor
        if estado["step"] == "valor":
            buque    = edit_states[uid]["buque"]
            col_name = edit_states[uid]["col_name"]
            ok, msg  = editar_campo_en_sheet(buque, col_name, text)
            edit_states.pop(uid, None)
            if ok:
                await update.message.reply_text(
                    f"Actualizado exitosamente\n\n"
                    f"Buque: {buque}\n"
                    f"Campo: {col_name}\n"
                    f"Nuevo valor: {text}\n\n"
                    f"Ver Google Sheets:\nhttps://docs.google.com/spreadsheets/d/{SHEET_ID}",
                    reply_markup=menu_principal()
                )
            else:
                await update.message.reply_text(f"Error al actualizar: {msg}", reply_markup=menu_principal())
            return

    # ── COMANDOS NORMALES ────────────────────────────────────────
    session = get_session(uid)

    if text in ["Limpiar sesion", "limpiar"]:
        clear_session(uid)
        await update.message.reply_text("Sesion limpiada.", reply_markup=menu_principal())
        return

    if text in ["Listar buques", "listar"]:
        if not registros_sesion:
            await update.message.reply_text("No hay buques registrados en esta sesion.")
        else:
            lista = "Buques registrados esta sesion:\n\n"
            for i, r in enumerate(registros_sesion, 1):
                lista += f"{i}. {r.get('buque','?')} — IMO {r.get('imo','?')} — {r.get('puerto','?')}\n"
            await update.message.reply_text(lista)
        return

    if text in ["Nuevo buque", "nuevo"]:
        clear_session(uid)
        await update.message.reply_text("Cuentame los datos del buque:", reply_markup=menu_principal())
        return

    if text in ["Eliminar buque", "eliminar"]:
        clear_session(uid)
        delete_states[uid] = {"step": "buque"}
        await update.message.reply_text(
            "Eliminar buque\n\n"
            "Escriba el nombre exacto del buque que desea eliminar:",
            reply_markup=ReplyKeyboardMarkup([["Cancelar"]], resize_keyboard=True)
        )
        return

    if text in ["Editar registro", "editar"]:
        clear_session(uid)
        edit_states[uid] = {"step": "buque"}
        await update.message.reply_text(
            "Editar registro\n\n"
            "Escriba el nombre exacto del buque que desea editar:",
            reply_markup=ReplyKeyboardMarkup([["Cancelar edicion"]], resize_keyboard=True)
        )
        return

    if text in ["Ayuda", "ayuda"]:
        await update.message.reply_text(
            "Como funciona:\n\n"
            "1. Escribeme los datos del buque\n"
            "2. Te pregunto lo que falte\n"
            "3. Te muestro resumen para confirmar\n"
            "4. Al confirmar se guarda en Google Sheets\n"
            "5. Te muestro como quedo el registro\n\n"
            "Para editar: boton 'Editar registro'\n\n"
            "Ejemplo:\n"
            "CTI QUEEN, IMO 9240079, Panama, 650 MT VLSO, SPRB, NAVES, ETA 10/06, 96 horas",
            reply_markup=menu_principal()
        )
        return

    # ── CONVERSACION CON IA ──────────────────────────────────────
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
                await update.message.reply_text("Guardando en Google Sheets...")
                ok, row_num = registrar_en_sheet(data)
                if ok:
                    registros_sesion.append(data)
                    clear_session(uid)
                    await update.message.reply_text(generar_resumen_registro(data))
                    await update.message.reply_text(
                        f"Ver Google Sheets:\nhttps://docs.google.com/spreadsheets/d/{SHEET_ID}\n\n"
                        f"Hay otro buque que registrar?",
                        reply_markup=menu_principal()
                    )
                else:
                    await update.message.reply_text("No se pudo guardar. Intenta de nuevo.")
                return
            except Exception as e:
                logger.error(f"Error registro: {e}")

        # Limpiar markdown de la respuesta
        safe = response.replace("*","").replace("_","").replace("`","").replace("[","").replace("]","")
        await update.message.reply_text(safe)

    except Exception as e:
        logger.error(f"Error: {e}")
        if "RATE_LIMIT" in str(e):
            await update.message.reply_text("Espera unos segundos e intenta de nuevo.")
        else:
            await update.message.reply_text("Error de conexion. Intenta de nuevo.")

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    logger.info("Bot Bunkers QBS iniciado...")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()

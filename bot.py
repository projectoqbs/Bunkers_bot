import os
import json
import logging
import httpx
import gspread
from PIL import Image, ImageDraw, ImageFont
import io
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
    "FECHA REGISTRO", "MN", "ETA", "AGENCIA", "ETD",
    "MT VLSO", "MT HSFO", "MT MGO",
    "PUERTO", "HORAS OP.", "CONTRATO",
    "IMO", "BANDERA", "CIUDAD", "ESTADO"
]

CAMPO_MAP = {
    "1":  ("ETA",       "eta"),
    "2":  ("AGENCIA",   "agencia"),
    "3":  ("ETD",       "etd"),
    "4":  ("MT VLSO",   "mt_vlso"),
    "5":  ("MT HSFO",   "mt_hsfo"),
    "6":  ("MT MGO",    "mt_mgo"),
    "7":  ("PUERTO",    "puerto"),
    "8":  ("HORAS OP.", "horas_op"),
    "9":  ("CONTRATO",  "contrato"),
    "10": ("IMO",       "imo"),
    "11": ("BANDERA",   "bandera"),
    "12": ("CIUDAD",    "ciudad"),
    "13": ("ESTADO",    "estado"),
}

SYSTEM_PROMPT = """Eres el Agente Bunkers QBS, asistente operativo de CI Quality Bunkers Supply S.A.S para gestion de suministro de combustible a buques en Colombia.

Tu trabajo es recopilar datos de operaciones de bunkering conversacionalmente.

DATOS QUE NECESITAS:
- NOMBRE_BUQUE (unico obligatorio)
- IMO (opcional)
- BANDERA (opcional)
- ETA (opcional)
- AGENCIA (opcional)
- ETD (opcional)
- MT_VLSO (opcional)
- MT_HSFO (opcional)
- MT_MGO (opcional)
- PUERTO (opcional, SPRB / PALERMO / SPSM)
- HORAS_OP (opcional)
- CONTRATO (opcional)
- CIUDAD_OPERACION (defecto: MALAMBO)

REGLAS:
- Habla en espanol, tono profesional y directo
- Si el usuario da varios datos juntos, extraelos todos de una vez
- El UNICO campo obligatorio es el NOMBRE_BUQUE
- Si solo da el nombre del buque, procede con los demas vacios
- No insistas en pedir datos opcionales si el usuario no los proporciona
- Cuando tengas el nombre del buque muestra el resumen con lo que tenga y pregunta si confirma
- Cuando tengas todos los datos principales muestra EXACTAMENTE este formato:

RESUMEN:
Buque: [NOMBRE]
ETA: [ETA]
Agencia: [AGENCIA]
ETD: [ETD o ninguno]
MT VLSO: [MT_VLSO o ninguno]
MT HSFO: [MT_HSFO o ninguno]
MT MGO: [MT_MGO o ninguno]
Puerto: [PUERTO]
Horas op.: [HORAS_OP o ninguno]
Contrato: [CONTRATO o ninguno]
IMO: [IMO]
Bandera: [BANDERA]
Ciudad: [CIUDAD]

Confirmas el registro? Responde SI para registrar o dime que corregir.

CUANDO EL USUARIO CONFIRME con si/confirmo/ok/correcto responde EXACTAMENTE:
REGISTRO_CONFIRMADO
JSON:{"buque":"X","eta":"X","agencia":"X","etd":"X","mt_vlso":"X","mt_hsfo":"X","mt_mgo":"X","puerto":"X","horas_op":"X","contrato":"X","imo":"X","bandera":"X","ciudad":"X"}
"""

def get_sheet():
    creds_dict = json.loads(GOOGLE_CREDS)
    scopes = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
    creds  = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    sh     = client.open_by_key(SHEET_ID)
    try:
        ws = sh.worksheet("BUQUES")
        # Solo agregar headers si la hoja esta completamente vacia
        if not ws.row_values(1):
            ws.append_row(HEADERS)
            _fmt(ws)
    except:
        ws = sh.add_worksheet(title="BUQUES", rows=1000, cols=16)
        ws.append_row(HEADERS)
        _fmt(ws)
    return ws

def _fmt(ws):
    ws.format("A1:O1", {
        "backgroundColor": {"red":0.12,"green":0.22,"blue":0.39},
        "textFormat": {"bold":True,"foregroundColor":{"red":1,"green":1,"blue":1}},
        "horizontalAlignment": "CENTER"
    })

def _clean(data, key):
    v = data.get(key, "")
    return v if v and v not in ["X","0","ninguno","none","-","N/A",""] else ""

def registrar_en_sheet(data):
    try:
        ws = get_sheet()
        fila = [
            datetime.now().strftime("%d/%m/%Y %H:%M"),
            data.get("buque",""),
            _clean(data,"eta"), _clean(data,"agencia"), _clean(data,"etd"),
            _clean(data,"mt_vlso"), _clean(data,"mt_hsfo"), _clean(data,"mt_mgo"),
            _clean(data,"puerto"), _clean(data,"horas_op"), _clean(data,"contrato"),
            _clean(data,"imo"), _clean(data,"bandera"),
            data.get("ciudad","MALAMBO"), "PENDIENTE"
        ]
        ws.append_row(fila)
        n = len(ws.get_all_values())
        color = {"red":0.93,"green":0.96,"blue":1.0} if n%2==0 else {"red":1.0,"green":1.0,"blue":1.0}
        ws.format(f"A{n}:O{n}", {"backgroundColor": color})
        return True
    except Exception as e:
        logger.error(f"Error Sheets: {e}")
        return False

def buscar_buque(nombre):
    try:
        ws    = get_sheet()
        datos = ws.get_all_values()
        for i in range(len(datos)-1, 0, -1):
            if datos[i][1].upper() == nombre.upper():
                return i+1, dict(zip(datos[0], datos[i]))
        return None, None
    except Exception as e:
        logger.error(f"Error buscando: {e}")
        return None, None

def editar_campo(nombre, col_name, nuevo_valor):
    try:
        ws    = get_sheet()
        datos = ws.get_all_values()
        headers = datos[0]
        if col_name not in headers:
            return False, f"Campo '{col_name}' no encontrado"
        col_idx = headers.index(col_name) + 1
        for i in range(len(datos)-1, 0, -1):
            if datos[i][1].upper() == nombre.upper():
                ws.update_cell(i+1, col_idx, nuevo_valor)
                return True, "ok"
        return False, f"Buque '{nombre}' no encontrado"
    except Exception as e:
        return False, str(e)

def eliminar_buque(nombre):
    try:
        ws    = get_sheet()
        datos = ws.get_all_values()
        for i in range(len(datos)-1, 0, -1):
            if datos[i][1].upper() == nombre.upper():
                ws.delete_rows(i+1)
                return True, "ok"
        return False, f"Buque '{nombre}' no encontrado"
    except Exception as e:
        return False, str(e)

def generar_imagen_tabla():
    try:
        ws    = get_sheet()
        datos = ws.get_all_values()
        if len(datos) <= 1:
            return None, "No hay buques registrados."

        cols_mostrar = ["MN", "ETA", "AGENCIA", "ETD", "MT VLSO", "MT HSFO", "MT MGO", "PUERTO", "HORAS OP."]
        headers = datos[0]
        indices = [headers.index(c) for c in cols_mostrar if c in headers]

        filas = []
        for row in datos[1:]:
            if any(row):
                filas.append([row[i] if i < len(row) else "" for i in indices])

        if not filas:
            return None, "No hay buques registrados."

        # Escala 2x para mejor calidad
        SCALE      = 3
        FONT_SIZE  = 32 * SCALE
        PAD_X      = 22 * SCALE
        PAD_Y      = 16 * SCALE
        ROW_H      = FONT_SIZE + PAD_Y * 2
        MARGIN     = 32 * SCALE

        # Calcular anchos de columna basado en contenido
        col_widths = []
        for ci, col in enumerate(cols_mostrar):
            max_w = len(col)
            for fila in filas:
                if ci < len(fila):
                    max_w = max(max_w, len(str(fila[ci])))
            col_widths.append(max(max_w, 4))

        COL_WIDTHS = [w * (FONT_SIZE // 2 + 1) + PAD_X * 2 for w in col_widths]
        TABLE_W    = sum(COL_WIDTHS)
        TITLE_H    = FONT_SIZE + MARGIN

        img_w = TABLE_W + MARGIN * 2
        img_h = ROW_H * (len(filas) + 1) + TITLE_H + MARGIN * 2

        img  = Image.new("RGB", (img_w, img_h), color=(240, 243, 248))
        draw = ImageDraw.Draw(img)

        try:
            font       = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", FONT_SIZE)
            font_bold  = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", FONT_SIZE)
            font_title = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", FONT_SIZE + 8)
        except:
            font = font_bold = font_title = ImageFont.load_default()

        # Fondo título
        draw.rectangle([0, 0, img_w, TITLE_H + MARGIN // 2], fill=(31, 56, 100))
        titulo = "CI QUALITY BUNKERS SUPPLY S.A.S  —  BUQUES"
        draw.text((MARGIN, MARGIN // 2), titulo, fill=(255, 255, 255), font=font_title)

        ox = MARGIN
        oy = TITLE_H + MARGIN // 2

        # Encabezados
        x = ox
        for ci, col in enumerate(cols_mostrar):
            draw.rectangle([x, oy, x + COL_WIDTHS[ci], oy + ROW_H], fill=(46, 117, 182))
            # Centrar texto en encabezado
            tw = draw.textlength(col, font=font_bold)
            tx = x + (COL_WIDTHS[ci] - tw) // 2
            draw.text((tx, oy + PAD_Y), col, fill=(255, 255, 255), font=font_bold)
            x += COL_WIDTHS[ci]

        # Filas de datos
        for ri, fila in enumerate(filas):
            y = oy + ROW_H * (ri + 1)
            bg = (214, 228, 247) if ri % 2 == 0 else (255, 255, 255)
            x  = ox
            for ci in range(len(cols_mostrar)):
                draw.rectangle([x, y, x + COL_WIDTHS[ci], y + ROW_H], fill=bg)
                val = str(fila[ci]) if ci < len(fila) else ""
                draw.text((x + PAD_X, y + PAD_Y), val, fill=(25, 25, 25), font=font_bold)
                x += COL_WIDTHS[ci]

        # Bordes horizontales entre filas
        for ri in range(len(filas) + 2):
            y = oy + ROW_H * ri
            draw.line([(ox, y), (ox + TABLE_W, y)], fill=(180, 195, 215), width=1)

        # Bordes verticales
        x = ox
        for w in COL_WIDTHS:
            draw.line([(x, oy), (x, oy + ROW_H * (len(filas) + 1))], fill=(180, 195, 215), width=1)
            x += w
        draw.line([(x, oy), (x, oy + ROW_H * (len(filas) + 1))], fill=(180, 195, 215), width=1)

        # Borde exterior tabla
        draw.rectangle(
            [ox, oy, ox + TABLE_W, oy + ROW_H * (len(filas) + 1)],
            outline=(31, 56, 100), width=2
        )

        # Guardar con alta calidad
        buf = io.BytesIO()
        img.save(buf, format="PNG", optimize=False)
        buf.seek(0)
        return buf, None

    except Exception as e:
        logger.error(f"Error generando imagen: {e}")
        return None, str(e)

def resumen_registro(data):
    def v(k): return _clean(data,k) or "—"
    return (
        f"REGISTRO GUARDADO\n{'─'*30}\n"
        f"Buque:     {data.get('buque','')}\n"
        f"ETA:       {v('eta')}\n"
        f"Agencia:   {v('agencia')}\n"
        f"ETD:       {v('etd')}\n"
        f"MT VLSO:   {v('mt_vlso')}\n"
        f"MT HSFO:   {v('mt_hsfo')}\n"
        f"MT MGO:    {v('mt_mgo')}\n"
        f"Puerto:    {v('puerto')}\n"
        f"Horas op.: {v('horas_op')}\n"
        f"Contrato:  {v('contrato')}\n"
        f"IMO:       {v('imo')}\n"
        f"Bandera:   {v('bandera')}\n"
        f"Ciudad:    {data.get('ciudad','MALAMBO')}\n"
        f"Estado:    PENDIENTE\n"
        f"{'─'*30}\nGuardado exitosamente"
    )

async def call_groq(history, system):
    messages = [{"role":"system","content":system}] + [{"role":m["role"],"content":m["content"]} for m in history]
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.post(
            "https://api.groq.com/openai/v1/chat/completions",
            headers={"Authorization":f"Bearer {GROQ_API_KEY}","Content-Type":"application/json"},
            json={"model":"llama-3.3-70b-versatile","messages":messages,"max_tokens":1000,"temperature":0.7}
        )
        if resp.status_code == 429: raise Exception("RATE_LIMIT")
        if resp.status_code != 200: raise Exception(f"Groq {resp.status_code}")
        return resp.json()["choices"][0]["message"]["content"]

def menu():
    return ReplyKeyboardMarkup([
        [KeyboardButton("Nuevo buque"),     KeyboardButton("Listar buques")],
        [KeyboardButton("Editar registro"), KeyboardButton("Eliminar buque")],
        [KeyboardButton("Limpiar sesion"),  KeyboardButton("Ayuda")],
    ], resize_keyboard=True)

def menu_campos():
    return ReplyKeyboardMarkup([
        [KeyboardButton("1. ETA"),       KeyboardButton("2. Agencia")],
        [KeyboardButton("3. ETD"),       KeyboardButton("4. MT VLSO")],
        [KeyboardButton("5. MT HSFO"),   KeyboardButton("6. MT MGO")],
        [KeyboardButton("7. Puerto"),    KeyboardButton("8. Horas op.")],
        [KeyboardButton("9. Contrato"),  KeyboardButton("10. IMO")],
        [KeyboardButton("11. Bandera"),  KeyboardButton("12. Ciudad")],
        [KeyboardButton("13. Estado"),   KeyboardButton("Cancelar")],
    ], resize_keyboard=True)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(
        "Bienvenido al Agente Bunkers QBS\n\n"
        "Soy tu asistente para registrar operaciones de suministro de combustible.\n\n"
        "Escribeme los datos del buque para comenzar.",
        reply_markup=menu()
    )

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    ud   = context.user_data  # persiste en memoria por sesion de usuario
    modo = ud.get("modo")

    # ── CANCELAR GLOBAL ──────────────────────────────────────────
    if text == "Cancelar":
        ud.clear()
        await update.message.reply_text("Operacion cancelada.", reply_markup=menu())
        return

    # ── FLUJO ELIMINACION ────────────────────────────────────────
    if modo == "eliminar":
        paso = ud.get("paso")

        if paso == "buque":
            row_idx, fila = buscar_buque(text)
            if not row_idx:
                await update.message.reply_text(f"No encontre '{text}'. Verifique el nombre o escriba Cancelar.")
                return
            ud["buque"] = text
            ud["paso"]  = "confirmar"
            resumen = (
                f"Buque encontrado:\n{'─'*26}\n"
                f"MN:       {fila.get('MN','')}\n"
                f"ETA:      {fila.get('ETA','—')}\n"
                f"IMO:      {fila.get('IMO','—')}\n"
                f"Puerto:   {fila.get('PUERTO','—')}\n"
                f"MT VLSO:  {fila.get('MT VLSO','—')}\n"
                f"MT HSFO:  {fila.get('MT HSFO','—')}\n"
                f"MT MGO:   {fila.get('MT MGO','—')}\n"
                f"Estado:   {fila.get('ESTADO','—')}\n"
                f"{'─'*26}\n"
                f"Desea eliminar este registro?\n"
                f"Escriba CONFIRMAR o Cancelar"
            )
            await update.message.reply_text(resumen, reply_markup=ReplyKeyboardMarkup(
                [[KeyboardButton("CONFIRMAR"), KeyboardButton("Cancelar")]],
                resize_keyboard=True, one_time_keyboard=True
            ))
            return

        if paso == "confirmar":
            if text.upper() == "CONFIRMAR":
                buque = ud["buque"]
                ud.clear()
                ok, msg = eliminar_buque(buque)
                if ok:
                    await update.message.reply_text(
                        f"Registro eliminado: {buque}\n\n"
                        f"Ver Sheets: https://docs.google.com/spreadsheets/d/{SHEET_ID}",
                        reply_markup=menu()
                    )
                else:
                    await update.message.reply_text(f"Error: {msg}", reply_markup=menu())
            else:
                ud.clear()
                await update.message.reply_text("Eliminacion cancelada.", reply_markup=menu())
            return

    # ── FLUJO EDICION ────────────────────────────────────────────
    if modo == "editar":
        paso = ud.get("paso")

        if paso == "buque":
            row_idx, fila = buscar_buque(text)
            if not row_idx:
                await update.message.reply_text(f"No encontre '{text}'. Verifique el nombre o escriba Cancelar.")
                return
            ud["buque"] = text
            ud["paso"]  = "campo"
            resumen = (
                f"Buque: {fila.get('MN','')}\n{'─'*26}\n"
                f"1.  ETA:       {fila.get('ETA','—')}\n"
                f"2.  Agencia:   {fila.get('AGENCIA','—')}\n"
                f"3.  ETD:       {fila.get('ETD','—')}\n"
                f"4.  MT VLSO:   {fila.get('MT VLSO','—')}\n"
                f"5.  MT HSFO:   {fila.get('MT HSFO','—')}\n"
                f"6.  MT MGO:    {fila.get('MT MGO','—')}\n"
                f"7.  Puerto:    {fila.get('PUERTO','—')}\n"
                f"8.  Horas op.: {fila.get('HORAS OP.','—')}\n"
                f"9.  Contrato:  {fila.get('CONTRATO','—')}\n"
                f"10. IMO:       {fila.get('IMO','—')}\n"
                f"11. Bandera:   {fila.get('BANDERA','—')}\n"
                f"12. Ciudad:    {fila.get('CIUDAD','—')}\n"
                f"13. Estado:    {fila.get('ESTADO','—')}\n"
                f"{'─'*26}\nQue campo desea editar?"
            )
            await update.message.reply_text(resumen, reply_markup=menu_campos())
            return

        if paso == "campo":
            num = text.split(".")[0].strip()
            if num not in CAMPO_MAP:
                await update.message.reply_text("Seleccione un campo del menu (1-13) o escriba Cancelar.", reply_markup=menu_campos())
                return
            col_name, _ = CAMPO_MAP[num]
            ud["col_name"] = col_name
            ud["paso"]     = "valor"
            await update.message.reply_text(
                f"Campo: {col_name}\nEscriba el nuevo valor:",
                reply_markup=ReplyKeyboardMarkup([[KeyboardButton("Cancelar")]], resize_keyboard=True)
            )
            return

        if paso == "valor":
            buque    = ud["buque"]
            col_name = ud["col_name"]
            ud.clear()
            ok, msg = editar_campo(buque, col_name, text)
            if ok:
                await update.message.reply_text(
                    f"Actualizado\nBuque: {buque}\n{col_name}: {text}\n\n"
                    f"Ver Sheets: https://docs.google.com/spreadsheets/d/{SHEET_ID}",
                    reply_markup=menu()
                )
            else:
                await update.message.reply_text(f"Error: {msg}", reply_markup=menu())
            return

    # ── FLUJO REGISTRO (IA) ──────────────────────────────────────
    if modo == "registro":
        history = ud.get("history", [])

        if text.upper() in ["SI", "SÍ", "CONFIRMO", "OK", "CORRECTO"]:
            pass  # dejar que la IA procese

        history.append({"role":"user","content":text})
        await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")
        try:
            response = await call_groq(history, SYSTEM_PROMPT)
            history.append({"role":"assistant","content":response})
            ud["history"] = history

            if "REGISTRO_CONFIRMADO" in response:
                try:
                    j = response.index("JSON:") + 5
                    data = json.loads(response[j:response.index("}", j)+1])
                    await update.message.reply_text("Guardando en Google Sheets...")
                    ok = registrar_en_sheet(data)
                    if ok:
                        ud.clear()
                        await update.message.reply_text(resumen_registro(data))
                        await update.message.reply_text(
                            f"Ver Sheets: https://docs.google.com/spreadsheets/d/{SHEET_ID}\n\nHay otro buque?",
                            reply_markup=menu()
                        )
                    else:
                        await update.message.reply_text("No se pudo guardar. Intenta de nuevo.")
                    return
                except Exception as e:
                    logger.error(f"Error JSON: {e}")

            safe = response.replace("*","").replace("_","").replace("`","")
            await update.message.reply_text(safe)
        except Exception as e:
            logger.error(f"Error Groq: {e}")
            if "RATE_LIMIT" in str(e):
                await update.message.reply_text("Espera unos segundos e intenta de nuevo.")
            else:
                await update.message.reply_text("Error de conexion. Intenta de nuevo.")
        return

    # ── MENU PRINCIPAL ───────────────────────────────────────────
    if text in ["Nuevo buque", "nuevo"]:
        ud.clear()
        ud["modo"] = "registro"
        ud["history"] = []
        await update.message.reply_text("Cuentame los datos del buque:", reply_markup=menu())
        return

    if text in ["Eliminar buque", "eliminar"]:
        ud.clear()
        ud["modo"] = "eliminar"
        ud["paso"] = "buque"
        await update.message.reply_text(
            "Eliminar buque\n\nEscriba el nombre exacto del buque:",
            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("Cancelar")]], resize_keyboard=True)
        )
        return

    if text in ["Editar registro", "editar"]:
        ud.clear()
        ud["modo"] = "editar"
        ud["paso"] = "buque"
        await update.message.reply_text(
            "Editar registro\n\nEscriba el nombre exacto del buque:",
            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("Cancelar")]], resize_keyboard=True)
        )
        return

    if text in ["Listar buques", "listar"]:
        await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="upload_photo")
        buf, error = generar_imagen_tabla()
        if buf:
            await update.message.reply_photo(photo=buf)
        else:
            await update.message.reply_text(error or "No se pudo generar la imagen.")
        return

    if text in ["Limpiar sesion", "limpiar"]:
        ud.clear()
        await update.message.reply_text("Sesion limpiada.", reply_markup=menu())
        return

    if text in ["Ayuda", "ayuda"]:
        await update.message.reply_text(
            "Como funciona:\n\n"
            "1. Nuevo buque: registra una operacion\n"
            "2. Editar registro: modifica un campo\n"
            "3. Eliminar buque: borra un registro\n\n"
            "Ejemplo registro:\n"
            "CTI QUEEN, IMO 9240079, Panama, 650 MT VLSO, SPRB, NAVES, ETA 10/06",
            reply_markup=menu()
        )
        return

    # Si no hay modo activo, iniciar registro automaticamente
    ud.clear()
    ud["modo"] = "registro"
    ud["history"] = [{"role":"user","content":text}]
    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")
    try:
        response = await call_groq(ud["history"], SYSTEM_PROMPT)
        ud["history"].append({"role":"assistant","content":response})
        safe = response.replace("*","").replace("_","").replace("`","")
        await update.message.reply_text(safe)
    except Exception as e:
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

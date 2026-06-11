import os, json, logging, httpx, gspread, sys, io
from google.oauth2.service_account import Credentials
from PIL import Image, ImageDraw, ImageFont
from telegram import Update, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from datetime import datetime

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "")
GROQ_API_KEY   = os.environ.get("GROQ_API_KEY", "")
GOOGLE_CREDS   = os.environ.get("GOOGLE_CREDS", "")
SHEET_ID       = "1cKS-P5T9hO3Ayv78gAGC8i_3JhOmDxlothPfxiWToCY"
BASE_DIR       = os.path.dirname(os.path.abspath(__file__))

sys.path.insert(0, BASE_DIR)
try:
    from generar_docs import generar_documentos, generar_pdf_carta
    DOCS_DISPONIBLES = True
except Exception as e:
    logger.warning(f"generar_docs no disponible: {e}")
    DOCS_DISPONIBLES = False

HEADERS = [
    "FECHA REGISTRO","MN","ETA","AGENCIA","ETD",
    "MT VLSO","MT HSFO","MT MGO","PUERTO","HORAS OP.","CONTRATO",
    "IMO","BANDERA","CIUDAD","ESTADO"
]

CAMPO_MAP = {
    "1":  ("ETA","eta"), "2":  ("AGENCIA","agencia"), "3":  ("ETD","etd"),
    "4":  ("MT VLSO","mt_vlso"), "5":  ("MT HSFO","mt_hsfo"), "6":  ("MT MGO","mt_mgo"),
    "7":  ("PUERTO","puerto"), "8":  ("HORAS OP.","horas_op"), "9":  ("CONTRATO","contrato"),
    "10": ("IMO","imo"), "11": ("BANDERA","bandera"), "12": ("CIUDAD","ciudad"),
    "13": ("ESTADO","estado"), "14": ("MN","buque"),
}

SYSTEM_PROMPT = """Eres el Agente Bunkers QBS, asistente operativo de CI Quality Bunkers Supply S.A.S para gestion de suministro de combustible a buques en Colombia.

DATOS QUE NECESITAS:
- NOMBRE_BUQUE (unico obligatorio)
- IMO, BANDERA, ETA, AGENCIA, ETD (opcionales)
- MT_VLSO, MT_HSFO, MT_MGO (opcionales)
- PUERTO (SPRB / PALERMO / SPSM, opcional)
- HORAS_OP, CONTRATO (opcionales)
- CIUDAD_OPERACION (defecto: MALAMBO)

REGLAS:
- Habla en espanol, tono profesional y directo
- Extrae todos los datos que el usuario de de una vez
- El UNICO campo obligatorio es NOMBRE_BUQUE
- Si solo da el nombre, procede con los demas vacios
- Muestra resumen y pregunta si confirma

Cuando tengas el nombre del buque muestra EXACTAMENTE:
RESUMEN:
Buque: [NOMBRE]
ETA: [ETA o ninguno]
Agencia: [AGENCIA o ninguno]
ETD: [ETD o ninguno]
MT VLSO: [MT_VLSO o ninguno]
MT HSFO: [MT_HSFO o ninguno]
MT MGO: [MT_MGO o ninguno]
Puerto: [PUERTO o ninguno]
Horas op.: [HORAS_OP o ninguno]
Contrato: [CONTRATO o ninguno]
IMO: [IMO o ninguno]
Bandera: [BANDERA o ninguno]
Ciudad: [CIUDAD]

Confirmas el registro?

CUANDO CONFIRME responde EXACTAMENTE:
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
        "backgroundColor":{"red":0.12,"green":0.22,"blue":0.39},
        "textFormat":{"bold":True,"foregroundColor":{"red":1,"green":1,"blue":1}},
        "horizontalAlignment":"CENTER"
    })

def _clean(data, key):
    v = data.get(key,"")
    return v if v and v not in ["X","0","ninguno","none","-","N/A",""] else ""

def registrar_en_sheet(data):
    try:
        ws = get_sheet()
        fila = [
            datetime.now().strftime("%d/%m/%Y %H:%M"),
            data.get("buque",""),
            _clean(data,"eta"),_clean(data,"agencia"),_clean(data,"etd"),
            _clean(data,"mt_vlso"),_clean(data,"mt_hsfo"),_clean(data,"mt_mgo"),
            _clean(data,"puerto"),_clean(data,"horas_op"),_clean(data,"contrato"),
            _clean(data,"imo"),_clean(data,"bandera"),
            data.get("ciudad","MALAMBO"),"PENDIENTE"
        ]
        ws.append_row(fila)
        n = len(ws.get_all_values())
        color = {"red":0.93,"green":0.96,"blue":1.0} if n%2==0 else {"red":1.0,"green":1.0,"blue":1.0}
        ws.format(f"A{n}:O{n}",{"backgroundColor":color})
        return True
    except Exception as e:
        logger.error(f"Error Sheets: {e}")
        return False

def buscar_buque(nombre):
    try:
        ws    = get_sheet()
        datos = ws.get_all_values()
        for i in range(len(datos)-1,0,-1):
            if datos[i][1].upper()==nombre.upper():
                return i+1, dict(zip(datos[0],datos[i]))
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
        col_idx = headers.index(col_name)+1
        for i in range(len(datos)-1,0,-1):
            if datos[i][1].upper()==nombre.upper():
                ws.update_cell(i+1,col_idx,nuevo_valor)
                return True,"ok"
        return False, f"Buque '{nombre}' no encontrado"
    except Exception as e:
        return False, str(e)

def eliminar_buque(nombre):
    try:
        ws    = get_sheet()
        datos = ws.get_all_values()
        for i in range(len(datos)-1,0,-1):
            if datos[i][1].upper()==nombre.upper():
                ws.delete_rows(i+1)
                return True,"ok"
        return False, f"Buque '{nombre}' no encontrado"
    except Exception as e:
        return False, str(e)

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
        f"{'─'*30}"
    )

def generar_imagen_tabla():
    try:
        ws    = get_sheet()
        datos = ws.get_all_values()
        if len(datos)<=1: return None,None,"No hay buques registrados."
        headers    = datos[0]
        filas_data = [row for row in datos[1:] if any(row)]
        if not filas_data: return None,None,"No hay buques registrados."
        cols_mostrar = ["MN","ETA","AGENCIA","ETD","MT VLSO","MT HSFO","MT MGO","PUERTO","HORAS OP."]
        indices = [headers.index(c) for c in cols_mostrar if c in headers]
        filas   = [[row[i] if i<len(row) else "" for i in indices] for row in filas_data]
        FONT_SIZE=15; PAD_X=10; PAD_Y=6; ROW_H=FONT_SIZE+PAD_Y*2; TITLE_H=36
        try:
            font      = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",FONT_SIZE)
            font_bold = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",FONT_SIZE)
        except:
            font=font_bold=ImageFont.load_default()
        tmp=ImageDraw.Draw(Image.new("RGB",(1,1)))
        COL_WIDTHS=[]
        for ci,col in enumerate(cols_mostrar):
            max_w=tmp.textlength(col,font=font_bold)
            for fila in filas:
                val=str(fila[ci]) if ci<len(fila) else ""
                max_w=max(max_w,tmp.textlength(val,font=font))
            COL_WIDTHS.append(int(max_w)+PAD_X*2)
        TABLE_W=sum(COL_WIDTHS)
        img=Image.new("RGB",(TABLE_W,TITLE_H+ROW_H*(len(filas)+1)),color=(255,255,255))
        draw=ImageDraw.Draw(img)
        draw.rectangle([0,0,TABLE_W,TITLE_H],fill=(31,56,100))
        draw.text((10,(TITLE_H-FONT_SIZE)//2),"CI QUALITY BUNKERS SUPPLY S.A.S  —  BUQUES",fill=(255,255,255),font=font_bold)
        oy=TITLE_H
        x=0
        for ci,col in enumerate(cols_mostrar):
            draw.rectangle([x,oy,x+COL_WIDTHS[ci],oy+ROW_H],fill=(217,225,242),outline=(180,180,180))
            draw.text((x+PAD_X,oy+PAD_Y),col,fill=(0,0,0),font=font_bold)
            x+=COL_WIDTHS[ci]
        for ri,fila in enumerate(filas):
            y=oy+ROW_H*(ri+1); bg=(242,242,242) if ri%2==0 else (255,255,255); x=0
            for ci in range(len(cols_mostrar)):
                draw.rectangle([x,y,x+COL_WIDTHS[ci],y+ROW_H],fill=bg,outline=(200,200,200))
                val=str(fila[ci]) if ci<len(fila) else ""
                draw.text((x+PAD_X,y+PAD_Y),val,fill=(0,0,0),font=font)
                x+=COL_WIDTHS[ci]
        buf=io.BytesIO(); img.save(buf,format="PNG"); buf.seek(0)
        return buf,None,None
    except Exception as e:
        logger.error(f"Error imagen: {e}")
        return None,None,str(e)

def generar_imagen_estado():
    try:
        ws    = get_sheet()
        datos = ws.get_all_values()
        if len(datos)<=1: return None,"No hay buques registrados."
        headers    = datos[0]
        filas_data = [row for row in datos[1:] if any(row)]
        if not filas_data: return None,"No hay buques registrados."
        cols_mostrar = ["MN","CONTRATO","IMO","BANDERA","CIUDAD","ESTADO"]
        indices = [headers.index(c) for c in cols_mostrar if c in headers]
        filas   = [[row[i] if i<len(row) else "" for i in indices] for row in filas_data]
        FONT_SIZE=15; PAD_X=10; PAD_Y=6; ROW_H=FONT_SIZE+PAD_Y*2; TITLE_H=36
        try:
            font      = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",FONT_SIZE)
            font_bold = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",FONT_SIZE)
        except:
            font=font_bold=ImageFont.load_default()
        tmp=ImageDraw.Draw(Image.new("RGB",(1,1)))
        COL_WIDTHS=[]
        for ci,col in enumerate(cols_mostrar):
            max_w=tmp.textlength(col,font=font_bold)
            for fila in filas:
                val=str(fila[ci]) if ci<len(fila) else ""
                max_w=max(max_w,tmp.textlength(val,font=font))
            COL_WIDTHS.append(int(max_w)+PAD_X*2)
        TABLE_W=sum(COL_WIDTHS)
        img=Image.new("RGB",(TABLE_W,TITLE_H+ROW_H*(len(filas)+1)),color=(255,255,255))
        draw=ImageDraw.Draw(img)
        draw.rectangle([0,0,TABLE_W,TITLE_H],fill=(31,56,100))
        draw.text((10,(TITLE_H-FONT_SIZE)//2),"CI QUALITY BUNKERS  —  ESTADO SOLICITUDES",fill=(255,255,255),font=font_bold)
        oy=TITLE_H; x=0
        for ci,col in enumerate(cols_mostrar):
            draw.rectangle([x,oy,x+COL_WIDTHS[ci],oy+ROW_H],fill=(217,225,242),outline=(180,180,180))
            draw.text((x+PAD_X,oy+PAD_Y),col,fill=(0,0,0),font=font_bold)
            x+=COL_WIDTHS[ci]
        for ri,fila in enumerate(filas):
            y=oy+ROW_H*(ri+1); bg=(242,242,242) if ri%2==0 else (255,255,255); x=0
            for ci in range(len(cols_mostrar)):
                val=str(fila[ci]) if ci<len(fila) else ""
                # Color especial para estado
                cell_bg=bg
                if col_name_at(cols_mostrar,ci)=="ESTADO":
                    if val.upper()=="ENVIADO": cell_bg=(198,239,206)
                    elif val.upper()=="PENDIENTE": cell_bg=(255,235,156)
                draw.rectangle([x,y,x+COL_WIDTHS[ci],y+ROW_H],fill=cell_bg,outline=(200,200,200))
                draw.text((x+PAD_X,y+PAD_Y),val,fill=(0,0,0),font=font)
                x+=COL_WIDTHS[ci]
        buf=io.BytesIO(); img.save(buf,format="PNG"); buf.seek(0)
        return buf,None
    except Exception as e:
        logger.error(f"Error imagen estado: {e}")
        return None,str(e)

def col_name_at(cols,idx):
    return cols[idx] if idx<len(cols) else ""

async def call_groq(history,system):
    messages=[{"role":"system","content":system}]+[{"role":m["role"],"content":m["content"]} for m in history]
    async with httpx.AsyncClient(timeout=30) as client:
        resp=await client.post(
            "https://api.groq.com/openai/v1/chat/completions",
            headers={"Authorization":f"Bearer {GROQ_API_KEY}","Content-Type":"application/json"},
            json={"model":"llama-3.3-70b-versatile","messages":messages,"max_tokens":1000,"temperature":0.7}
        )
        if resp.status_code==429: raise Exception("RATE_LIMIT")
        if resp.status_code!=200: raise Exception(f"Groq {resp.status_code}")
        return resp.json()["choices"][0]["message"]["content"]

def menu():
    return ReplyKeyboardMarkup([
        [KeyboardButton("Nuevo buque"),      KeyboardButton("Listar buques")],
        [KeyboardButton("Editar registro"),  KeyboardButton("Eliminar buque")],
        [KeyboardButton("Generar documentos"),KeyboardButton("Ayuda")],
    ],resize_keyboard=True)

def menu_listar():
    return ReplyKeyboardMarkup([
        [KeyboardButton("BUQUES"),           KeyboardButton("ESTADO SOLICITUD")],
        [KeyboardButton("Cancelar")],
    ],resize_keyboard=True)

def menu_gendocs():
    return ReplyKeyboardMarkup([
        [KeyboardButton("DIMAR"),            KeyboardButton("CARTA CAPITANIA")],
        [KeyboardButton("Cancelar")],
    ],resize_keyboard=True)

def menu_campos():
    return ReplyKeyboardMarkup([
        [KeyboardButton("1. ETA"),      KeyboardButton("2. Agencia")],
        [KeyboardButton("3. ETD"),      KeyboardButton("4. MT VLSO")],
        [KeyboardButton("5. MT HSFO"),  KeyboardButton("6. MT MGO")],
        [KeyboardButton("7. Puerto"),   KeyboardButton("8. Horas op.")],
        [KeyboardButton("9. Contrato"), KeyboardButton("10. IMO")],
        [KeyboardButton("11. Bandera"), KeyboardButton("12. Ciudad")],
        [KeyboardButton("13. Estado"),  KeyboardButton("14. Nombre buque")],
        [KeyboardButton("Cancelar")],
    ],resize_keyboard=True)

async def start(update:Update,context:ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(
        "Bienvenido al Agente Bunkers QBS\n\nEscribeme los datos del buque para comenzar.",
        reply_markup=menu()
    )

async def handle_message(update:Update,context:ContextTypes.DEFAULT_TYPE):
    text=update.message.text.strip()
    ud=context.user_data
    modo=ud.get("modo")

    if text=="Cancelar":
        ud.clear()
        await update.message.reply_text("Operacion cancelada.",reply_markup=menu())
        return

    # ── ELIMINACION ──────────────────────────────────────────────
    if modo=="eliminar":
        paso=ud.get("paso")
        if paso=="buque":
            row_idx,fila=buscar_buque(text)
            if not row_idx:
                await update.message.reply_text(f"No encontre '{text}'. Verifique o escriba Cancelar.")
                return
            ud["buque"]=text; ud["paso"]="confirmar"
            resumen=(f"Buque encontrado:\n{'─'*26}\nMN: {fila.get('MN','')}\n"
                     f"ETA: {fila.get('ETA','—')}  IMO: {fila.get('IMO','—')}\n"
                     f"Puerto: {fila.get('PUERTO','—')}  Estado: {fila.get('ESTADO','—')}\n{'─'*26}\n"
                     f"Desea eliminar este registro? Escriba CONFIRMAR o Cancelar")
            await update.message.reply_text(resumen,reply_markup=ReplyKeyboardMarkup(
                [[KeyboardButton("CONFIRMAR"),KeyboardButton("Cancelar")]],resize_keyboard=True,one_time_keyboard=True))
            return
        if paso=="confirmar":
            if text.upper()=="CONFIRMAR":
                buque=ud["buque"]; ud.clear()
                ok,msg=eliminar_buque(buque)
                await update.message.reply_text(
                    f"Registro eliminado: {buque}" if ok else f"Error: {msg}",reply_markup=menu())
            else:
                ud.clear()
                await update.message.reply_text("Eliminacion cancelada.",reply_markup=menu())
            return

    # ── EDICION ──────────────────────────────────────────────────
    if modo=="editar":
        paso=ud.get("paso")
        if paso=="buque":
            row_idx,fila=buscar_buque(text)
            if not row_idx:
                await update.message.reply_text(f"No encontre '{text}'. Verifique o escriba Cancelar.")
                return
            ud["buque"]=text; ud["paso"]="campo"
            resumen=(f"Buque: {fila.get('MN','')}\n{'─'*26}\n"
                     f"1.  ETA:          {fila.get('ETA','—')}\n"
                     f"2.  Agencia:      {fila.get('AGENCIA','—')}\n"
                     f"3.  ETD:          {fila.get('ETD','—')}\n"
                     f"4.  MT VLSO:      {fila.get('MT VLSO','—')}\n"
                     f"5.  MT HSFO:      {fila.get('MT HSFO','—')}\n"
                     f"6.  MT MGO:       {fila.get('MT MGO','—')}\n"
                     f"7.  Puerto:       {fila.get('PUERTO','—')}\n"
                     f"8.  Horas op.:    {fila.get('HORAS OP.','—')}\n"
                     f"9.  Contrato:     {fila.get('CONTRATO','—')}\n"
                     f"10. IMO:          {fila.get('IMO','—')}\n"
                     f"11. Bandera:      {fila.get('BANDERA','—')}\n"
                     f"12. Ciudad:       {fila.get('CIUDAD','—')}\n"
                     f"13. Estado:       {fila.get('ESTADO','—')}\n"
                     f"14. Nombre buque: {fila.get('MN','—')}\n"
                     f"{'─'*26}\nQue campo desea editar?")
            await update.message.reply_text(resumen,reply_markup=menu_campos())
            return
        if paso=="campo":
            num=text.split(".")[0].strip()
            if num not in CAMPO_MAP:
                await update.message.reply_text("Seleccione un campo (1-14) o Cancelar.",reply_markup=menu_campos())
                return
            col_name,_=CAMPO_MAP[num]; ud["col_name"]=col_name; ud["paso"]="valor"
            await update.message.reply_text(f"Campo: {col_name}\nEscriba el nuevo valor:",
                reply_markup=ReplyKeyboardMarkup([[KeyboardButton("Cancelar")]],resize_keyboard=True))
            return
        if paso=="valor":
            buque=ud["buque"]; col_name=ud["col_name"]; ud.clear()
            ok,msg=editar_campo(buque,col_name,text)
            await update.message.reply_text(
                f"Actualizado\nBuque: {buque}\n{col_name}: {text}" if ok else f"Error: {msg}",
                reply_markup=menu())
            return

    # ── GENERAR DOCUMENTOS ───────────────────────────────────────
    if modo=="gendocs":
        paso=ud.get("paso")
        if paso=="buque":
            row_idx,fila=buscar_buque(text)
            if not row_idx:
                await update.message.reply_text(f"No encontre '{text}'. Verifique o escriba Cancelar.")
                return
            ud["buque_data"]=fila; ud["paso"]="confirmar"
            tipo_doc=ud.get("tipo_doc","DIMAR")
            resumen=(f"Generar {tipo_doc} para:\n{'─'*26}\n"
                     f"Buque:   {fila.get('MN','')}\n"
                     f"IMO:     {fila.get('IMO','—')}\n"
                     f"Bandera: {fila.get('BANDERA','—')}\n"
                     f"ETA:     {fila.get('ETA','—')}\n"
                     f"Puerto:  {fila.get('PUERTO','—')}\n"
                     f"MT VLSO: {fila.get('MT VLSO','—')}\n"
                     f"MT HSFO: {fila.get('MT HSFO','—')}\n"
                     f"MT MGO:  {fila.get('MT MGO','—')}\n"
                     f"Agencia: {fila.get('AGENCIA','—')}\n{'─'*26}\n"
                     f"Fecha suministro: ETA + 5 dias\nConfirma?")
            await update.message.reply_text(resumen,reply_markup=ReplyKeyboardMarkup(
                [[KeyboardButton("SI GENERAR"),KeyboardButton("Cancelar")]],resize_keyboard=True,one_time_keyboard=True))
            return
        if paso=="confirmar":
            if text.upper() in ["SI GENERAR","SI","SÍ","CONFIRMAR","OK"]:
                fila=ud["buque_data"]
                tipo_doc=ud.get("tipo_doc","DIMAR")
                ud.clear()
                await update.message.reply_text(f"Generando {tipo_doc}...")
                await context.bot.send_chat_action(chat_id=update.effective_chat.id,action="upload_document")
                try:
                    buque_data={
                        "buque":   fila.get("MN",""),
                        "imo":     fila.get("IMO",""),
                        "bandera": fila.get("BANDERA",""),
                        "eta":     fila.get("ETA",""),
                        "agencia": fila.get("AGENCIA",""),
                        "mt_vlso": fila.get("MT VLSO",""),
                        "mt_hsfo": fila.get("MT HSFO",""),
                        "mt_mgo":  fila.get("MT MGO",""),
                        "puerto":  fila.get("PUERTO",""),
                    }
                    output_dir=os.path.join(BASE_DIR,"docs_generados")
                    if tipo_doc=="DIMAR":
                        doc1,doc2=generar_documentos(buque_data,output_dir,BASE_DIR)
                        with open(doc1,"rb") as f:
                            await update.message.reply_document(document=f,filename=os.path.basename(doc1))
                        with open(doc2,"rb") as f:
                            await update.message.reply_document(document=f,filename=os.path.basename(doc2))
                        await update.message.reply_text(
                            f"Documentos DIMAR generados para {buque_data['buque']}",reply_markup=menu())
                    else:  # CARTA CAPITANIA
                        pdf_path=generar_pdf_carta(buque_data,output_dir,BASE_DIR)
                        with open(pdf_path,"rb") as f:
                            await update.message.reply_document(document=f,filename=os.path.basename(pdf_path))
                        await update.message.reply_text(
                            f"Carta Capitania generada para {buque_data['buque']}",reply_markup=menu())
                except Exception as e:
                    logger.error(f"Error generando docs: {e}")
                    await update.message.reply_text(f"Error: {e}",reply_markup=menu())
            else:
                ud.clear()
                await update.message.reply_text("Cancelado.",reply_markup=menu())
            return

    # ── MENU PRINCIPAL ───────────────────────────────────────────
    if text in ["Nuevo buque","nuevo"]:
        ud.clear(); ud["modo"]="registro"; ud["history"]=[]
        await update.message.reply_text("Cuentame los datos del buque:",reply_markup=menu())
        return
    if text in ["Eliminar buque","eliminar"]:
        ud.clear(); ud["modo"]="eliminar"; ud["paso"]="buque"
        await update.message.reply_text("Escriba el nombre exacto del buque a eliminar:",
            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("Cancelar")]],resize_keyboard=True))
        return
    if text in ["Editar registro","editar"]:
        ud.clear(); ud["modo"]="editar"; ud["paso"]="buque"
        await update.message.reply_text("Escriba el nombre exacto del buque a editar:",
            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("Cancelar")]],resize_keyboard=True))
        return
    if text in ["Generar documentos","generar","documentos"]:
        if not DOCS_DISPONIBLES:
            await update.message.reply_text("Modulo de documentos no disponible.")
            return
        await update.message.reply_text("Seleccione el tipo de documento:",reply_markup=menu_gendocs())
        return

    if text in ["DIMAR","CARTA CAPITANIA"] and modo!="gendocs":
        if not DOCS_DISPONIBLES:
            await update.message.reply_text("Modulo de documentos no disponible.")
            return
        ud.clear(); ud["modo"]="gendocs"; ud["paso"]="buque"; ud["tipo_doc"]=text
        await update.message.reply_text(
            f"Generar: {text}\n\nEscriba el nombre exacto del buque:",
            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("Cancelar")]],resize_keyboard=True))
        return
    if text in ["Listar buques","listar"]:
        await update.message.reply_text("Seleccione una opcion:",reply_markup=menu_listar())
        return

    if text=="BUQUES":
        await context.bot.send_chat_action(chat_id=update.effective_chat.id,action="upload_photo")
        buf,_,error=generar_imagen_tabla()
        if error: await update.message.reply_text(error,reply_markup=menu())
        else: await update.message.reply_photo(photo=buf,reply_markup=menu())
        return

    if text=="ESTADO SOLICITUD":
        await context.bot.send_chat_action(chat_id=update.effective_chat.id,action="upload_photo")
        buf,error=generar_imagen_estado()
        if error: await update.message.reply_text(error,reply_markup=menu())
        else: await update.message.reply_photo(photo=buf,reply_markup=menu())
        return
    if text in ["Ayuda","ayuda"]:
        await update.message.reply_text(
            "Como funciona:\n\n"
            "1. Nuevo buque: registra una operacion\n"
            "2. Listar buques: ver tabla en imagen\n"
            "3. Editar registro: modifica un campo\n"
            "4. Eliminar buque: borra un registro\n"
            "5. Generar documentos: crea los Word para DIMAR\n\n"
            "Ejemplo:\nCTI QUEEN, IMO 9240079, Panama, 650 MT VLSO, SPRB, NAVES, ETA 10/06",
            reply_markup=menu())
        return

    # ── CONVERSACION IA ──────────────────────────────────────────
    if modo!="registro":
        ud.clear(); ud["modo"]="registro"; ud["history"]=[]
    history=ud.get("history",[])
    history.append({"role":"user","content":text})
    await context.bot.send_chat_action(chat_id=update.effective_chat.id,action="typing")
    try:
        response=await call_groq(history,SYSTEM_PROMPT)
        history.append({"role":"assistant","content":response})
        ud["history"]=history
        if "REGISTRO_CONFIRMADO" in response:
            try:
                j=response.index("JSON:")+5
                data=json.loads(response[j:response.index("}",j)+1])
                await update.message.reply_text("Guardando en Google Sheets...")
                ok=registrar_en_sheet(data)
                if ok:
                    ud.clear()
                    await update.message.reply_text(resumen_registro(data))
                    await update.message.reply_text("Hay otro buque?",reply_markup=menu())
                else:
                    await update.message.reply_text("No se pudo guardar. Intenta de nuevo.")
                return
            except Exception as e:
                logger.error(f"Error JSON: {e}")
        safe=response.replace("*","").replace("_","").replace("`","")
        await update.message.reply_text(safe)
    except Exception as e:
        logger.error(f"Error: {e}")
        msg="Espera unos segundos e intenta de nuevo." if "RATE_LIMIT" in str(e) else "Error de conexion. Intenta de nuevo."
        await update.message.reply_text(msg)

def main():
    app=Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start",start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND,handle_message))
    logger.info("Bot Bunkers QBS iniciado...")
    app.run_polling(drop_pending_updates=True)

if __name__=="__main__":
    main()

import zipfile
import shutil
import os
import re
from datetime import datetime, timedelta

MESES_ES = {
    1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril", 5:"Mayo", 6:"Junio",
    7:"Julio", 8:"Agosto", 9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"
}

PUERTO_COMPLETO = {
    "SPRB":    "SOCIEDAD PORTUARIA REGIONAL DE BARRANQUILLA",
    "PALERMO": "POR PALERMO SOCIEDAD PORTUARIA",
    "SPSM":    "SOCIEDAD PORTUARIA DE SANTA MARTA",
}

def fecha_larga(dt):
    return f"{dt.day:02d} de {MESES_ES[dt.month]} del {dt.year}"

def fecha_corta(dt):
    return f"{dt.day:02d} de {MESES_ES[dt.month]}"

def parsear_eta(eta_str):
    """Convierte ETA string a datetime. Acepta: 06/06/2026, 06/06, 6-jun, 06-jun-2026"""
    hoy = datetime.now()
    eta_str = eta_str.strip()
    formatos = ["%d/%m/%Y", "%d/%m", "%d-%m-%Y", "%d-%b-%Y", "%d-%b"]
    for fmt in formatos:
        try:
            dt = datetime.strptime(eta_str, fmt)
            if dt.year == 1900:
                dt = dt.replace(year=hoy.year)
            return dt
        except:
            pass
    # Intentar "6-jun" o "06-jun"
    m = re.match(r"(\d{1,2})[-/]([a-zA-Z]+)", eta_str)
    if m:
        day = int(m.group(1))
        mon_str = m.group(2).lower()[:3]
        mon_map = {"ene":1,"feb":2,"mar":3,"abr":4,"may":5,"jun":6,
                   "jul":7,"ago":8,"sep":9,"oct":10,"nov":11,"dic":12}
        if mon_str in mon_map:
            return datetime(hoy.year, mon_map[mon_str], day)
    return hoy

def generar_documentos(buque_data, output_dir, base_dir):
    """
    buque_data: dict con buque, imo, bandera, eta, agencia, mt_vlso, mt_hsfo, mt_mgo, puerto
    output_dir: carpeta donde guardar los Word generados
    base_dir:   carpeta donde están las plantillas .docx
    Retorna: (path_doc1, path_doc2) o lanza excepción
    """
    os.makedirs(output_dir, exist_ok=True)

    buque   = buque_data.get("buque", "").upper()
    imo     = buque_data.get("imo", "")
    bandera = buque_data.get("bandera", "").upper()
    agencia = buque_data.get("agencia", "").upper()
    puerto  = buque_data.get("puerto", "").upper()
    eta_str = buque_data.get("eta", "")

    # Calcular combustible
    vlso = buque_data.get("mt_vlso", "")
    hsfo = buque_data.get("mt_hsfo", "")
    mgo  = buque_data.get("mt_mgo", "")
    partes = []
    if vlso: partes.append(f"{vlso} MT VLSFO 380")
    if hsfo: partes.append(f"{hsfo} MT HSFO")
    if mgo:  partes.append(f"{mgo} MT MGO")
    cantidad_total = " + ".join(partes) if partes else "A CONFIRMAR"
    # Para plantilla necesitamos cantidad y tipo separados
    if vlso and not hsfo and not mgo:
        cantidad_doc = vlso
        tipo_doc     = "VLSFO 380"
    elif mgo and not vlso and not hsfo:
        cantidad_doc = mgo
        tipo_doc     = "MGO"
    elif hsfo and not vlso and not mgo:
        cantidad_doc = hsfo
        tipo_doc     = "HSFO"
    else:
        cantidad_doc = cantidad_total
        tipo_doc     = "COMBUSTIBLE"

    # Calcular fechas
    hoy     = datetime.now()
    eta_dt  = parsear_eta(eta_str) if eta_str else hoy
    fin_dt  = eta_dt + timedelta(days=5)

    fecha_solicitud   = fecha_larga(hoy)
    fecha_inicio_num  = f"{eta_dt.day:02d}"
    rango_fechas      = f" de {MESES_ES[eta_dt.month]} al {fin_dt.day:02d} de {MESES_ES[fin_dt.month]} de {fin_dt.year}"
    rango_fechas_corto= f"{eta_dt.day:02d} al {fin_dt.day:02d} de {MESES_ES[eta_dt.month]} de {fin_dt.year}"
    lugar_suministro  = PUERTO_COMPLETO.get(puerto, puerto)

    # Diccionario de reemplazos
    reemplazos = {
        "{{FECHA_SOLICITUD}}":    fecha_solicitud,
        "{{NOMBRE_BUQUE}}":       buque,
        "{{IMO}}":                imo,
        "{{BANDERA}}":            bandera,
        "{{CANTIDAD}}":           cantidad_doc,
        "{{TIPO_COMB}}":          tipo_doc,
        "{{PUERTO}}":             puerto,
        "{{LUGAR_SUMINISTRO}}":   lugar_suministro,
        "{{AGENCIA}}":            agencia,
        "{{FECHA_INICIO}}":       fecha_inicio_num,
        "{{RANGO_FECHAS}}":       rango_fechas,
        "{{RANGO_FECHAS_CORTO}}": rango_fechas_corto,
    }

    buque_id = buque.replace(" ", "_")
    doc1_out = os.path.join(output_dir, f"SOLICITUD_MUELLE_{buque_id}.docx")
    doc2_out = os.path.join(output_dir, f"FORMATO_SOLICITUD_{buque_id}.docx")

    _reemplazar_en_docx(
        os.path.join(base_dir, "plantilla_solicitud.docx"),
        doc1_out, reemplazos
    )
    _reemplazar_en_docx(
        os.path.join(base_dir, "plantilla_formato.docx"),
        doc2_out, reemplazos
    )
    return doc1_out, doc2_out

def _reemplazar_en_docx(plantilla_path, output_path, reemplazos):
    shutil.copy2(plantilla_path, output_path)
    with zipfile.ZipFile(output_path, 'r') as zin:
        nombres   = zin.namelist()
        contenidos = {}
        for nombre in nombres:
            with zin.open(nombre) as f:
                contenidos[nombre] = f.read()

    for xml_file in ["word/document.xml", "word/header1.xml", "word/footer1.xml"]:
        if xml_file in contenidos:
            texto = contenidos[xml_file].decode("utf-8")
            for key, val in reemplazos.items():
                texto = texto.replace(key, val)
            contenidos[xml_file] = texto.encode("utf-8")

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for nombre, data in contenidos.items():
            zout.writestr(nombre, data)

if __name__ == "__main__":
    # Test
    test_data = {
        "buque": "CTI QUEEN", "imo": "9240079", "bandera": "PANAMA",
        "eta": "10/06/2026", "agencia": "NAVES", "mt_vlso": "650",
        "mt_hsfo": "", "mt_mgo": "", "puerto": "SPRB"
    }
    d1, d2 = generar_documentos(test_data, "/home/claude/test_docs", "/home/claude/bunkers_bot")
    print(f"Doc1: {d1}")
    print(f"Doc2: {d2}")

# ── GENERAR PDF CARTA CAPITANIA ──────────────────────────────────
def generar_pdf_carta(buque_data, output_dir, base_dir):
    """Genera la Carta Capitanía en PDF con el mismo contenido que plantilla_solicitud"""
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch, cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY

    os.makedirs(output_dir, exist_ok=True)

    buque   = buque_data.get("buque","").upper()
    imo     = buque_data.get("imo","")
    bandera = buque_data.get("bandera","").upper()
    agencia = buque_data.get("agencia","").upper()
    puerto  = buque_data.get("puerto","").upper()
    eta_str = buque_data.get("eta","")

    vlso = buque_data.get("mt_vlso","")
    hsfo = buque_data.get("mt_hsfo","")
    mgo  = buque_data.get("mt_mgo","")
    partes = []
    if vlso: partes.append(f"{vlso} MT VLSFO 380")
    if hsfo: partes.append(f"{hsfo} MT HSFO")
    if mgo:  partes.append(f"{mgo} MT MGO")
    cantidad_total = " + ".join(partes) if partes else "A CONFIRMAR"

    if vlso and not hsfo and not mgo:
        cantidad_doc = vlso; tipo_doc = "VLSFO 380"
    elif mgo and not vlso and not hsfo:
        cantidad_doc = mgo; tipo_doc = "MGO"
    elif hsfo and not vlso and not mgo:
        cantidad_doc = hsfo; tipo_doc = "HSFO"
    else:
        cantidad_doc = cantidad_total; tipo_doc = "COMBUSTIBLE"

    hoy    = datetime.now()
    eta_dt = parsear_eta(eta_str) if eta_str else hoy
    fin_dt = eta_dt + timedelta(days=5)

    fecha_solicitud    = fecha_larga(hoy)
    fecha_inicio_num   = f"{eta_dt.day:02d}"
    mes_nombre         = MESES_ES[eta_dt.month]
    fin_dia            = f"{fin_dt.day:02d}"
    fin_mes            = MESES_ES[fin_dt.month]
    anio               = str(fin_dt.year)
    lugar_suministro   = PUERTO_COMPLETO.get(puerto, puerto)

    buque_id = buque.replace(" ","_")
    pdf_path = os.path.join(output_dir, f"CARTA_CAPITANIA_{buque_id}.pdf")

    doc = SimpleDocTemplate(
        pdf_path, pagesize=letter,
        leftMargin=3*cm, rightMargin=3*cm,
        topMargin=2.5*cm, bottomMargin=2.5*cm
    )

    styles = getSampleStyleSheet()
    normal = ParagraphStyle("normal", fontName="Helvetica", fontSize=11,
                             leading=16, alignment=TA_JUSTIFY)
    bold   = ParagraphStyle("bold", fontName="Helvetica-Bold", fontSize=11,
                             leading=16, alignment=TA_JUSTIFY)
    center = ParagraphStyle("center", fontName="Helvetica", fontSize=11,
                             leading=16, alignment=TA_CENTER)
    footer_style = ParagraphStyle("footer", fontName="Helvetica", fontSize=9,
                                   leading=12, alignment=TA_CENTER, textColor=colors.grey)

    story = []

    # Fecha y ciudad
    story.append(Paragraph(f"<b>Malambo, {fecha_solicitud}</b>", normal))
    story.append(Spacer(1, 14))

    # Destinatario
    story.append(Paragraph("<b>Señores</b>", normal))
    story.append(Spacer(1, 4))
    story.append(Paragraph("<b>DIMAR - Capitanía del Puerto de Barranquilla.</b>", normal))
    story.append(Paragraph("<b>Atención Sr. CF Bernardo Silva Flórez.</b>", normal))
    story.append(Paragraph("<b>Barranquilla.</b>", normal))
    story.append(Spacer(1, 14))

    # REF
    story.append(Paragraph(
        f"<b>REF:&nbsp;&nbsp;&nbsp;&nbsp;SOLICITUD DE AUTORIZACIÓN PARA EL SUMINISTRO DE "
        f"{cantidad_doc} MT DE {tipo_doc}</b>", normal))
    story.append(Spacer(1, 14))

    # Cuerpo
    cuerpo = (
        f"C.I QUALITY BUNKERS SUPPLY S.A.S, identificada con NIT. No. 900.346.136-3, licencia "
        f"de explotación comercial No.224 y autorizada mediante la resolución del Ministerio de "
        f"Minas y Energía No. 00124440 del 25 de Julio de 2011 para ejercer la actividad de "
        f"distribuidor minorista a través de la Estación de Servicio Marítima QBS002 con números "
        f"de matrícula MC-05-0190-AN, en convoy con el remolcador WILLIAMS, con números de "
        f"matrícula MC-03-0359, que comprende el recibo, almacenamiento y distribución de diésel "
        f"marinos e IFO's, destinados exclusivamente para buques y al servicio de los terminales "
        f"marítimos de los Puertos de Santa Marta, Barranquilla, Cartagena y fluvial asentados a "
        f"lo largo del canal del Dique y el Río Magdalena en los términos señalados en la "
        f"Resolución No. 000311 de 2011 emitida por la Corporación Autónoma Regional del "
        f"Atlántico – C.R.A, solicita permiso previo para la venta de <b>{cantidad_doc} MT de "
        f"{tipo_doc}</b> a la motonave <b>{buque}</b> con bandera de <b>{bandera}</b> e "
        f"<b>IMO NUMBER {imo}</b> que serán suministradas anclados en "
        f"<b>{lugar_suministro}</b> entre los días <b>{fecha_inicio_num} de {mes_nombre} al "
        f"{fin_dia} de {fin_mes} de {anio}</b>."
    )
    story.append(Paragraph(cuerpo, normal))
    story.append(Spacer(1, 14))

    story.append(Paragraph(
        "La operación iniciará en Puerto PIMSA y para el traslado de la barcaza se utilizará "
        "el remolcador Williams con números de matrícula MC-03-0359, se reportará los movimientos "
        "a la Oficina de Tráfico Marítimo siguiendo el protocolo establecido.", normal))
    story.append(Spacer(1, 14))

    story.append(Paragraph(
        "Para la operación en muelle seguiremos el protocolo establecido en el sistema NGS "
        "(Norma de Gestión de Seguridad) que fue aprobado por la Capitanía en el año 2012.", normal))
    story.append(Spacer(1, 14))

    story.append(Paragraph(
        f"Solicitado en <b>MALAMBO</b> el día {hoy.day:02d} del mes de "
        f"{MESES_ES[hoy.month]} del <b>{hoy.year}.</b>", normal))
    story.append(Spacer(1, 40))

    # Firma
    story.append(Paragraph("<b>ANDRES BALDIRIS</b>", normal))
    story.append(Paragraph("<b>COORDINADOR DE OPERACIONES JR</b>", normal))
    story.append(Paragraph("<b>CI QUALITY BUNKERS SUPPLY S.A.S</b>", normal))
    story.append(Paragraph("<b>Celular No. 3115466535</b>", normal))
    story.append(Paragraph("<b>Correo qbsops@qualitybunkers.com</b>", normal))
    story.append(Spacer(1, 20))

    story.append(Paragraph(
        "Parque Industrial de Malambo S.A, Manzana 05, Lote 12, Oficinas de QBS, Malambo, Atlántico.",
        footer_style))
    story.append(Paragraph("TELEFONO: +57 5 376 9038 - www.qualitybunkers.com", footer_style))

    doc.build(story)
    return pdf_path

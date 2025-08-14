from flask import Flask, render_template, request, send_file, session
from uuid import uuid4
import os
import re
import unicodedata
from datetime import datetime
from io import BytesIO
import base64
from sys import platform

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image

import html

from docx2pdf import convert
from dotenv import load_dotenv

# En Windows con docx2pdf + Word
try:
    import pythoncom
    HAS_PYTHONCOM = True
except Exception:
    HAS_PYTHONCOM = False

# Email
import correo_util

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET", os.urandom(24))

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
TEMPLATE_DOCX = os.path.join(BASE_DIR, "Contrato_Plantilla.docx")

# Config empresa
RESPONSABLE_EMPRESA = "Alan Arndt, Dueño de la Empresa"
FIRMA_EMPRESA_PATH = os.path.join(STATIC_DIR, "firma_empresa.png")  # o logo_empresa_firma.png
EMAIL_EMPRESA = os.getenv("EMAIL_EMPRESA")  # definir en .env
CONTACTO_TELEFONO = os.getenv("CONTACTO_TELEFONO", "")

# Tamaños de firma (puede ajustar a gusto)
SIGNATURE_IMAGE_WIDTH_IN = 2.8   # ancho de la firma (antes 2.0)
SIGNATURE_LABEL_FONT_PT  = 12    # tamaño de rótulos "Firma del ..."

# -------------------- Helpers --------------------

def _now_tag() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def _slug(s: str) -> str:
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[^a-zA-Z0-9]+", "_", s).strip("_").lower()
    return s or f"cliente_{uuid4().hex[:6]}"

def _docx_to_html_paragraphs(path_docx: str) -> str:
    """
    Lee el .docx y lo convierte a párrafos HTML simples (<p>...</p>).
    (Solo párrafos; si necesitás tablas, lo ampliamos luego.)
    """
    doc = Document(path_docx)
    parts = []
    for p in doc.paragraphs:
        text = html.escape(p.text).strip()
        parts.append(f"<p>{text or '&nbsp;'}</p>")
    return "\n".join(parts)

def _b64_to_pil_image(b64data: str) -> Image.Image:
    """
    Espera 'data:image/png;base64,AAAA...' o solo el base64.
    """
    if "," in b64data:
        b64data = b64data.split(",", 1)[1]
    raw = base64.b64decode(b64data)
    return Image.open(BytesIO(raw)).convert("RGBA")

def _insert_text_placeholders(doc: Document, mapping: dict):
    """
    Reemplazo robusto: opera a nivel de párrafo/celda para evitar el problema
    de placeholders “partidos” en runs.
    """
    def replace_text(text: str) -> str:
        for k, v in mapping.items():
            text = text.replace(k, v)
        return text

    # Párrafos fuera de tablas
    for p in doc.paragraphs:
        new = replace_text(p.text)
        if new != p.text:
            p.text = new  # reasigna el texto del párrafo (python-docx crea un run único)

    # Celdas de tablas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    new = replace_text(p.text)
                    if new != p.text:
                        p.text = new

def _ensure_company_signature(path_png: str, texto: str = "Alan Arndt"):
    """Genera una firma básica de empresa si no existe el archivo."""
    if os.path.exists(path_png):
        return
    from PIL import Image, ImageDraw, ImageFont
    img = Image.new("RGBA", (600, 220), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("arial.ttf", 72)
    except Exception:
        font = ImageFont.load_default()
    bbox = draw.textbbox((0, 0), texto, font=font)
    w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
    draw.text(((600 - w) // 2, (220 - h) // 2), texto, fill=(0, 0, 0, 255), font=font)
    img.save(path_png, "PNG")

def _add_signatures_section(doc: Document, firma_cliente_path: str, firma_empresa_path: str):
    """
    Inserta ambas firmas en una tabla 2x2:
      Fila 0: imágenes    (cliente izq | empresa der)
      Fila 1: rótulos     ("Firma del Cliente" | "Firma de la Empresa")
    Reutiliza tabla existente si detecta esos rótulos.
    """
    target = None
    for tbl in doc.tables:
        try:
            if len(tbl.rows) >= 2 and len(tbl.columns) >= 2:
                if ("firma del cliente" in tbl.cell(1, 0).text.lower()
                        and "firma de la empresa" in tbl.cell(1, 1).text.lower()):
                    target = tbl
                    break
        except Exception:
            continue

    if target is None:
        target = doc.add_table(rows=2, cols=2)
        target.cell(1, 0).text = "Firma del Cliente"
        target.cell(1, 1).text = "Firma de la Empresa"

    while len(target.rows) < 2:
        target.add_row()

    # --- Ancho fijo y parejo de columnas (evita que Word las deforme) ---
    try:
        target.autofit = False
    except Exception:
        pass

    # Forzar layout "fixed" a nivel de tabla (algunos Word lo respetan mejor)
    try:
        tbl = target._tbl
        tblPr = tbl.tblPr
        tblLayout = OxmlElement('w:tblLayout')
        tblLayout.set(qn('w:type'), 'fixed')
        tblPr.append(tblLayout)
    except Exception:
        pass

    # Ancho preferido de columnas y de celdas (parejo)
    col_w = Inches(3.2)  # ajustá a gusto
    for i in range(2):
        try:
            target.columns[i].width = col_w
        except Exception:
            pass
        for cell in target.columns[i].cells:
            try:
                cell.width = col_w
            except Exception:
                pass

    # Fila 0: imágenes centradas (más grandes)
    for col, img_path in enumerate([firma_cliente_path, firma_empresa_path]):
        c = target.cell(0, col)
        c.text = ""  # limpiar celda
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if os.path.exists(img_path) and os.path.getsize(img_path) > 0:
            ancho = SIGNATURE_IMAGE_WIDTH_IN if 'SIGNATURE_IMAGE_WIDTH_IN' in globals() else 2.8
            p.add_run().add_picture(img_path, width=Inches(ancho))

    # Fila 1: rótulos centrados (negrita + tamaño)
    for col, texto in enumerate(["Firma del Cliente", "Firma de la Empresa"]):
        c = target.cell(1, col)
        c.text = ""
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(texto)
        run.bold = True
        try:
            run.font.size = Pt(SIGNATURE_LABEL_FONT_PT if 'SIGNATURE_LABEL_FONT_PT' in globals() else 12)
        except Exception:
            pass

def _export_to_pdf_safe(out_docx: str, out_pdf: str) -> bool:
    """
    Convierte DOCX -> PDF. Devuelve True si el PDF quedó bien.
    Requiere Windows + Word para docx2pdf. Si falla, devolvemos False.
    """
    try:
        if platform.startswith("win") and HAS_PYTHONCOM:
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass
        convert(out_docx, out_pdf)
        return os.path.exists(out_pdf) and os.path.getsize(out_pdf) > 0
    except Exception:
        return False
    finally:
        if platform.startswith("win") and HAS_PYTHONCOM:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

# -------------------- Rutas --------------------

@app.route("/", methods=["GET"])
def index():
    if not os.path.exists(TEMPLATE_DOCX):
        contrato_html = "<p><strong>No se encontró la plantilla del contrato.</strong></p>"
    else:
        contrato_html = _docx_to_html_paragraphs(TEMPLATE_DOCX)
    return render_template("formulario_contrato.html", contrato_html=contrato_html)

@app.route("/generar", methods=["POST"])
def generar():
    # 1) Validación básica
    nombre = request.form.get("nombre", "").strip()
    dni = request.form.get("dni", "").strip()
    direccion = request.form.get("direccion", "").strip()
    email = request.form.get("email", "").strip()
    ubicacion = (request.form.get("ubicacion_monitoreo") or request.form.get("ubicacion") or "").strip()
    firma_b64 = (request.form.get("firmaBase64") or request.form.get("firma") or "").strip()

    faltantes = [k for k, v in {
        "nombre": nombre, "dni": dni, "direccion": direccion,
        "email": email, "ubicacion": ubicacion, "firma": firma_b64
    }.items() if not v]
    if faltantes:
        return f"Faltan campos obligatorios: {', '.join(faltantes)}", 400

    # 2) Preparar rutas de salida
    slug = f"{_slug(nombre)}_{_now_tag()}_{uuid4().hex[:6]}"
    out_docx = os.path.join(STATIC_DIR, f"{slug}.docx")
    out_pdf = os.path.join(STATIC_DIR, f"{slug}.pdf")
    firma_cliente_path = os.path.join(STATIC_DIR, f"firma_{slug}.png")

    # 3) Guardar firma del cliente
    try:
        img = _b64_to_pil_image(firma_b64)
        img.save(firma_cliente_path, "PNG")
    except Exception:
        return "La imagen de la firma es inválida.", 400

    # 4) Crear DOCX desde plantilla y reemplazar placeholders
    if not os.path.exists(TEMPLATE_DOCX):
        return "No se encontró la plantilla del contrato.", 500

    doc = Document(TEMPLATE_DOCX)
    fecha_firma = datetime.now().strftime("%d/%m/%Y")

    # Soportar ambos estilos de tokens en la plantilla para evitar mismatches
    mapping = {
        # estilo con espacios y minúsculas
        "{{ nombre }}": nombre,
        "{{ dni }}": dni,
        "{{ direccion }}": direccion,
        "{{ email }}": email,
        "{{ ubicacion }}": ubicacion,
        "{{ ubicacion_monitoreo }}": ubicacion,
        "{{ responsable_empresa }}": RESPONSABLE_EMPRESA,
        "{{ fecha_firma }}": fecha_firma,

        # estilo “corporativo” en mayúsculas sin espacios
        "{{NOMBRE_ABONADO}}": nombre,
        "{{DNI_ABONADO}}": dni,
        "{{DIRECCION_ABONADO}}": direccion,
        "{{EMAIL_ABONADO}}": email,
        "{{UBICACION_MONITOREO}}": ubicacion,
        "{{RESPONSABLE_EMPRESA}}": RESPONSABLE_EMPRESA,
        "{{FECHA_FIRMA}}": fecha_firma,
    }

    _insert_text_placeholders(doc, mapping)

    # Firmas
    _ensure_company_signature(FIRMA_EMPRESA_PATH, "Alan Arndt")
    _add_signatures_section(doc, firma_cliente_path, FIRMA_EMPRESA_PATH)

    # Guardar DOCX
    try:
        doc.save(out_docx)
    except Exception as e:
        return f"No se pudo guardar el DOCX: {e}", 500

    # 5) Convertir a PDF con fallback
    pdf_ok = _export_to_pdf_safe(out_docx, out_pdf)
    if pdf_ok:
        session["archivo_pdf"] = out_pdf
    else:
        # No rompemos el flujo: dejamos DOCX listo para descarga/envío
        session["archivo_pdf"] = out_docx

    # 6) Enviar correos (cliente y empresa) sin romper flujo si falla SMTP
    try:
        adjunto = session["archivo_pdf"]
        asunto = "Contrato firmado - Seguridad Ituzaingó"
        cuerpo = (
            f"Estimado/a {nombre},\n\n"
            f"Adjuntamos el contrato firmado correspondiente al servicio de monitoreo en {ubicacion}.\n"
            f"Le recomendamos conservar el archivo para su referencia.\n\n"
            "Quedamos a disposición por cualquier consulta.\n\n"
            "Atentamente,\n"
            "Seguridad Ituzaingó\n"
            "Alan Arndt — Dueño de la Empresa\n"
            f"Tel.: {CONTACTO_TELEFONO or '-'}\n"
            f"Email: {EMAIL_EMPRESA or '-'}\n"
        )

        # Enviar al cliente
        if email:
            correo_util.enviar_email(email, asunto, cuerpo, adjunto)

        # Copia a la empresa
        if EMAIL_EMPRESA:
            correo_util.enviar_email(
                EMAIL_EMPRESA,
                "Nuevo contrato firmado - Seguridad Ituzaingó",
                f"El cliente {nombre} firmó un contrato.\n\n{cuerpo}",
                adjunto
            )

    except Exception as e:
        print(f"[WARN] Error enviando emails: {e}")

    # 7) Agradecimiento
    return render_template("agradecimiento.html", telefono=CONTACTO_TELEFONO)

@app.route("/descargar", methods=["GET"])
def descargar():
    archivo = session.get("archivo_pdf")
    if not archivo or not os.path.exists(archivo):
        return "Error: No se encontró el contrato para descargar.", 404
    return send_file(archivo, as_attachment=True, download_name=os.path.basename(archivo))

# -------------------- Main --------------------

if __name__ == "__main__":
    os.makedirs(STATIC_DIR, exist_ok=True)
    # En dev: debug=True; en prod: usar WSGI/Gunicorn
    app.run(debug=True)

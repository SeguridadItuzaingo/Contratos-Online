# correo_util.py
import os
import base64
import json
import logging
import mimetypes
import requests

# ==============================
# Config desde variables de entorno
# ==============================
BREVO_API_KEY   = os.environ.get("BREVO_API_KEY")
FROM_EMAIL      = os.environ.get("FROM_EMAIL", "contratos@seguridadituzaingo.com")
FROM_NAME       = os.environ.get("FROM_NAME", "Seguridad Ituzaingó")
CC_EMPRESA      = os.environ.get("CC_EMPRESA")          # ej: administracion@seguridadituzaingo.com
CONTACTO_TEL    = os.environ.get("CONTACTO_TELEFONO", "")  # ej: 3786-617492

BREVO_URL = "https://api.brevo.com/v3/smtp/email"


def _adjunto_to_b64(path):
    """Convierte un archivo a adjunto base64 para Brevo."""
    if not path:
        return None
    ctype, enc = mimetypes.guess_type(path)
    if ctype is None or enc is not None:
        ctype = "application/octet-stream"
    with open(path, "rb") as f:
        content_b64 = base64.b64encode(f.read()).decode("utf-8")
    return {"content": content_b64, "name": os.path.basename(path)}


def enviar_email(to, asunto, cuerpo_texto, adjunto_path=None, cc=None, reply_to=None):
    """
    Envía email vía Brevo API.
      - to: str o lista de str
      - asunto: str
      - cuerpo_texto: str (texto plano)
      - adjunto_path: ruta archivo (opcional)
      - cc: str o lista de str (opcional)
      - reply_to: str (opcional)
    Return: (ok: bool, info: str)
    """
    if not BREVO_API_KEY:
        return False, "BREVO_API_KEY no configurada"

    # Normalizar destinatarios
    to_list = to if isinstance(to, list) else [to]
    cc_list = []
    if cc:
        cc_list = cc if isinstance(cc, list) else [cc]
    # agrega siempre CC_EMPRESA si está configurado
    if CC_EMPRESA and CC_EMPRESA not in cc_list:
        cc_list.append(CC_EMPRESA)

    # --- construir HTML sin f-strings problemáticos ---
    texto_plano = cuerpo_texto or ""
    texto_html  = texto_plano.replace("\n", "<br>")
    html_body   = "<p>" + texto_html + "</p>"
    if CONTACTO_TEL:
        html_body += "<p><strong>Teléfono:</strong> " + CONTACTO_TEL + "</p>"

    payload = {
        "sender": {"email": FROM_EMAIL, "name": FROM_NAME},
        "to": [{"email": x} for x in to_list],
        "subject": asunto,
        "textContent": texto_plano,
        "htmlContent": html_body,
    }

    if reply_to:
        payload["replyTo"] = {"email": reply_to}
    if cc_list:
        payload["cc"] = [{"email": x} for x in cc_list if x]

    adj = _adjunto_to_b64(adjunto_path)
    if adj:
        payload["attachment"] = [adj]

    headers = {
        "accept": "application/json",
        "api-key": BREVO_API_KEY,
        "content-type": "application/json",
    }

    try:
        r = requests.post(BREVO_URL, headers=headers, data=json.dumps(payload), timeout=20)
        if r.status_code in (200, 201, 202):
            data = r.json()
            message_id = data.get("messageId") or (
                data.get("messageIds", [""])[0] if isinstance(data.get("messageIds"), list) else ""
            )
            logging.info("[Brevo] Enviado OK -> to=%s id=%s", to_list, message_id)
            return True, str(message_id)
        else:
            logging.error("[Brevo] Error %s: %s", r.status_code, r.text)
            return False, f"HTTP {r.status_code}: {r.text}"
    except Exception as e:
        logging.exception("[Brevo] Excepción enviando email")
        return False, str(e)

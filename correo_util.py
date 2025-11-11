# correo_util.py — Envío por API HTTP (SendGrid), sin SMTP
import os
import base64
import json
import requests

PROVIDER = os.getenv("EMAIL_PROVIDER", "sendgrid").lower()
FROM_EMAIL = os.getenv("FROM_EMAIL", "no-reply@seguridadituzaingo.com")
CC_EMPRESA = os.getenv("CC_EMPRESA")  # opcional

class EmailError(Exception):
    pass

def _enviar_sendgrid(to_email, subject, body_text, attachment_path):
    api_key = os.getenv("SENDGRID_API_KEY")
    if not api_key:
        raise EmailError("Falta SENDGRID_API_KEY en variables de entorno")

    data = {
        "personalizations": [{
            "to": [{"email": to_email}],
            "subject": subject
        }],
        "from": {"email": FROM_EMAIL},
        "content": [{"type": "text/plain", "value": body_text or ""}]
    }

    if CC_EMPRESA:
        data["personalizations"][0]["cc"] = [{"email": CC_EMPRESA}]

    if attachment_path:
        with open(attachment_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode("utf-8")
        data.setdefault("attachments", []).append({
            "content": b64,
            "type": "application/pdf",
            "filename": os.path.basename(attachment_path),
            "disposition": "attachment"
        })

    resp = requests.post(
        "https://api.sendgrid.com/v3/mail/send",
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        },
        data=json.dumps(data),
        timeout=15
    )
    if resp.status_code >= 300:
        raise EmailError(f"SendGrid error {resp.status_code}: {resp.text}")

def enviar_email(destinatario, asunto, cuerpo="Adjunto el contrato firmado.", adjunto_path=None):
    """Envía correo usando API HTTP (SendGrid). Evita SMTP bloqueado en Render."""
    if PROVIDER == "sendgrid":
        return _enviar_sendgrid(destinatario, asunto, cuerpo, adjunto_path)
    raise EmailError(f"Proveedor no soportado: {PROVIDER}")

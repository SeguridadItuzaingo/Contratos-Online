import os
import smtplib
import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

def enviar_email(destinatario, asunto, cuerpo="Adjunto el contrato firmado.", adjunto_path=None):
    remitente = os.getenv("EMAIL_REMITENTE")
    clave = os.getenv("EMAIL_CLAVE")
    if not remitente or not clave:
        raise RuntimeError("Falta EMAIL_REMITENTE o EMAIL_CLAVE en .env")

    # Mensaje base
    msg = MIMEMultipart()
    msg["From"] = remitente
    msg["To"] = destinatario
    msg["Subject"] = asunto
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    # Adjunto (usa el nombre real y extensión correcta: .pdf o .docx)
    if adjunto_path:
        fname = os.path.basename(adjunto_path)
        ctype, _ = mimetypes.guess_type(adjunto_path)
        if ctype is None:
            ctype = "application/octet-stream"
        with open(adjunto_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=fname)
        part.add_header("Content-Disposition", "attachment", filename=fname)
        msg.attach(part)

    # SMTP configurable por .env (Gmail por defecto)
    host = os.getenv("SMTP_HOST", "smtp.gmail.com")
    use_tls = os.getenv("SMTP_TLS", "false").lower() in ("1", "true", "yes", "on")
    # STARTTLS suele ser 587; SSL directo suele ser 465
    port = int(os.getenv("SMTP_PORT", "587" if use_tls else "465"))

    if use_tls:
        # STARTTLS
        with smtplib.SMTP(host, port, timeout=30) as s:
            s.starttls()
            s.login(remitente, clave)
            s.send_message(msg)
    else:
        # SSL directo
        with smtplib.SMTP_SSL(host, port, timeout=30) as s:
            s.login(remitente, clave)
            s.send_message(msg)



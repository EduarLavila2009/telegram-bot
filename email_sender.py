"""Envío de correos electrónicos con archivos adjuntos a través de SMTP."""

from __future__ import annotations

import logging
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from mimetypes import guess_type
from pathlib import Path

from . import config

logger = logging.getLogger(__name__)

def send_report_email(
    recipient_email: str,
    subject: str,
    body_text: str,
    attachments: list[Path],
) -> None:
    """
    Envía un correo electrónico formal con múltiples archivos adjuntos usando SMTP configurado en .env.
    """
    if not config.SMTP_SERVER or not config.SMTP_USER or not config.SMTP_PASSWORD:
        raise ValueError(
            "Faltan configurar las variables SMTP en el archivo .env:\n"
            "Debes completar SMTP_SERVER, SMTP_USER y SMTP_PASSWORD."
        )

    # 1. Crear el contenedor MIME multipart
    msg = MIMEMultipart()
    msg["From"] = config.SMTP_USER
    msg["To"] = recipient_email.strip()
    msg["Subject"] = subject

    # 2. Agregar cuerpo del mensaje en texto plano
    msg.attach(MIMEText(body_text, "plain", "utf-8"))

    # 3. Adjuntar archivos
    for path in attachments:
        if not path.exists():
            logger.warning("El archivo a adjuntar no existe en disco: %s", path)
            continue
            
        filename = path.name
        ctype, encoding = guess_type(str(path))
        if ctype is None or encoding is not None:
            # Default fallback si no se puede determinar
            ctype = "application/octet-stream"
            
        maintype, subtype = ctype.split("/", 1)
        try:
            with path.open("rb") as f:
                part = MIMEBase(maintype, subtype)
                part.set_payload(f.read())
                
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename={filename}",
            )
            msg.attach(part)
            logger.info("Archivo adjuntado correctamente: %s", filename)
        except Exception as e:
            logger.exception("Error adjuntando el archivo %s: %s", filename, e)
            raise RuntimeError(f"Error adjuntando archivo {filename}: {e!s}") from e

    # 4. Establecer conexión SMTP y enviar
    logger.info("Conectando al servidor SMTP: %s:%s...", config.SMTP_SERVER, config.SMTP_PORT)
    try:
        # Usar SMTP estándar con STARTTLS
        server = smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT, timeout=15)
        server.ehlo()
        
        # Activar cifrado TLS de forma segura si el servidor lo soporta
        if server.has_extn("STARTTLS"):
            server.starttls()
            server.ehlo()
            
        # Iniciar sesión
        logger.info("Iniciando sesión como %s...", config.SMTP_USER)
        server.login(config.SMTP_USER, config.SMTP_PASSWORD)
        
        # Enviar correo
        logger.info("Enviando correo a %s...", recipient_email)
        server.sendmail(config.SMTP_USER, [recipient_email], msg.as_string())
        logger.info("¡Correo enviado exitosamente a %s!", recipient_email)
        
        server.quit()
    except Exception as e:
        logger.exception("Fallo en la conexión SMTP o en el envío del correo")
        raise RuntimeError(f"Error de conexión SMTP / envío de correo: {e!s}") from e

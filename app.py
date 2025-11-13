from flask import Flask, request, jsonify, render_template
import imaplib
import smtplib
import email
import os
import json
from email.header import decode_header
from email.utils import parseaddr, formataddr
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv
import openai
import requests
import msal
import time
import re
from datetime import datetime, timedelta, timezone
from urllib.parse import quote
from werkzeug.middleware.proxy_fix import ProxyFix
from io import BytesIO
import threading


# Cargar variables de entorno (.env) si existe
load_dotenv()

app = Flask(__name__)
# Si se despliega detrás de un proxy (NGINX/Traefik) que termina HTTPS,
# ProxyFix ayuda a que Flask detecte correctamente scheme/host/puerto
# a partir de los encabezados X-Forwarded-*. Ajusta los valores si tu proxy
# añade múltiples saltos.
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_port=1, x_prefix=1)


# Categorias configurables (nombre de carpeta + prompt)
CATEGORIES_FILE = 'categories.json'
DEFAULT_CATEGORIES = [
    {"name": "Prioridad baja", "prompt": "Mensajes informativos o de seguimiento sin impacto inmediato."},
    {"name": "Prioridad media", "prompt": "Solicitudes que requieren accion en las proximas 24-48 horas o con cierto impacto."},
    {"name": "Prioridad alta", "prompt": "Incidentes o bloqueos con urgencia explicita o impacto inmediato en clientes o ingresos."},
]


# Config por defecto (solo se usa si no hay cuentas guardadas)
EMAIL = os.getenv("DEFAULT_EMAIL")
PASSWORD = os.getenv("DEFAULT_EMAIL_APP_PASSWORD")
# Permitir definir proveedor por defecto si se desea (gmail | outlook)
DEFAULT_PROVIDER = (os.getenv("DEFAULT_EMAIL_PROVIDER") or "gmail").lower()
IMAP_SERVERS = {
    "gmail": "imap.gmail.com",
    "outlook": "outlook.office365.com",
}
SMTP_SERVERS = {
    "gmail": ("smtp.gmail.com", 587),
    "outlook": ("smtp.office365.com", 587),
}

# Microsoft Graph OAuth (para cuentas Outlook con OAuth)
MS_CLIENT_ID = os.getenv("MS_CLIENT_ID")
MS_CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET")
MS_TENANT_ID = os.getenv("MS_TENANT_ID", "common")
MS_REDIRECT_URI = os.getenv("MS_REDIRECT_URI", "http://localhost:5000/microsoft/callback")
MS_SCOPES = ["Mail.Read", "Mail.ReadWrite", "Mail.Send"]
# Microsoft Graph Subscriptions (para notificaciones automáticas)
MS_GRAPH_NOTIFICATION_URL = os.getenv("MS_GRAPH_NOTIFICATION_URL", "http://localhost:5000/webhook/microsoft-graph")
MS_GRAPH_SUBSCRIPTION_DURATION_MINUTES = int(os.getenv("MS_GRAPH_SUBSCRIPTION_DURATION_MINUTES", "4200"))  # 3 días por defecto
MS_GRAPH_SUBSCRIPTION_CLIENT_STATE = os.getenv("MS_GRAPH_SUBSCRIPTION_CLIENT_STATE", "CapDataDefaultClientStateSecret")
MS_GRAPH_RENEW_BEFORE_EXPIRY_HOURS = int(os.getenv("MS_GRAPH_RENEW_BEFORE_EXPIRY_HOURS", "24"))  # Renovar 24h antes

def _ms_authority():
    return f"https://login.microsoftonline.com/{MS_TENANT_ID}"

def _ms_app(token_cache: msal.SerializableTokenCache | None = None):
    return msal.ConfidentialClientApplication(
        MS_CLIENT_ID,
        authority=_ms_authority(),
        client_credential=MS_CLIENT_SECRET,
        token_cache=token_cache,
    )


def _effective_redirect_uri() -> str:
    """Resolve the redirect_uri that will be sent to Microsoft OAuth."""
    env_value = os.getenv('MS_REDIRECT_URI')
    base = (request.url_root or 'http://localhost:5000/').rstrip('/')
    dynamic_uri = f"{base}/microsoft/callback"

    if not env_value:
        return dynamic_uri

    candidates = [item.strip() for item in env_value.split(',') if item.strip()]
    if not candidates:
        return dynamic_uri

    current_host_raw = (request.host or '').strip()
    current_host = current_host_raw.split(':')[0].lower()
    current_port = None
    if ':' in current_host_raw:
        try:
            current_port = int(current_host_raw.split(':')[1])
        except (ValueError, IndexError):
            current_port = None
    current_scheme = request.headers.get('X-Forwarded-Proto', request.scheme or 'http')

    try:
        from urllib.parse import urlparse
    except Exception:
        urlparse = None

    for raw in candidates:
        if not raw:
            continue
        if any(token in raw for token in ('{host}', '{host_with_port}', '{scheme}', '{base}', '{dynamic}')):
            try:
                return raw.format(
                    host=current_host,
                    host_with_port=current_host_raw,
                    scheme=current_scheme,
                    base=base,
                    dynamic=dynamic_uri,
                )
            except Exception:
                continue
        if raw.startswith('/'):
            return f"{current_scheme}://{current_host_raw}{raw}"
        if urlparse:
            try:
                parsed = urlparse(raw)
            except Exception:
                parsed = None
            if not parsed:
                continue
            candidate_host = (parsed.hostname or '').lower()
            candidate_port = parsed.port
            if candidate_host:
                if candidate_host == current_host:
                    if candidate_port and current_port and candidate_port != current_port:
                        continue
                    return raw
                if candidate_host == current_host_raw.lower():
                    return raw
            elif not parsed.scheme and parsed.path:
                return f"{current_scheme}://{current_host_raw}{parsed.path}"

    first = candidates[0]
    if urlparse:
        try:
            parsed = urlparse(first)
            candidate_host = (parsed.hostname or '').lower()
            if candidate_host and candidate_host not in {current_host, current_host_raw.lower()}:
                return dynamic_uri
        except Exception:
            pass
    if first.startswith('/'):
        return f"{current_scheme}://{current_host_raw}{first}"
    return first or dynamic_uri

# OpenAI
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
if not OPENAI_API_KEY:
    raise RuntimeError("Falta OPENAI_API_KEY en variables de entorno o .env")
openai.api_key = OPENAI_API_KEY


# Ficheros de datos
ACCOUNTS_FILE = "accounts.json"
PROCESSED_FILE = "processed.json"  # e-mails ya tratados


def load_json(path: str, default):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                content = f.read().strip()
                if not content:
                    return default
                return json.loads(content)
        except (json.JSONDecodeError, ValueError):
            # Si el archivo está corrupto o vacío, devolver el valor por defecto
            return default
    return default


def save_json(path: str, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def normalize_categories(raw):
    normalized = []
    seen = set()
    for item in (raw or []):
        if not isinstance(item, dict):
            continue
        name = str(item.get('name', '')).strip()
        prompt = str(item.get('prompt', '')).strip()
        if not name:
            continue
        key = name.lower()
        if key in seen:
            continue
        seen.add(key)
        auto_reply = bool(item.get('autoReply', False))
        reply_text = str(item.get('replyText', '')).strip()
        normalized.append({
            'name': name, 
            'prompt': prompt,
            'autoReply': auto_reply,
            'replyText': reply_text
        })
    return normalized


def load_categories():
    # Si el archivo no existe, usar categorías por defecto
    # Si existe pero está vacío, usar lista vacía
    if not os.path.exists(CATEGORIES_FILE):
        data = DEFAULT_CATEGORIES
    else:
        data = load_json(CATEGORIES_FILE, [])
    
    normalized = normalize_categories(data)
    # Solo usar categorías por defecto si el archivo no existe y no hay categorías normalizadas
    if not os.path.exists(CATEGORIES_FILE) and not normalized:
        result = [dict(cat) for cat in DEFAULT_CATEGORIES]
        save_json(CATEGORIES_FILE, result)
        return result
    
    return normalized if normalized else []


def save_categories(categories):
    save_json(CATEGORIES_FILE, categories)


def category_names():
    return [item.get('name') for item in CATEGORIES if item.get('name')]


def pick_default_category(names):
    if not names:
        return 'Prioridad media'
    preferred = ['prioridad media', 'media', 'default', 'general']
    for target in preferred:
        for name in names:
            if name and name.lower() == target:
                return name
    return names[0]


CATEGORIES = load_categories()


# Cuentas
ACCOUNTS = load_json(ACCOUNTS_FILE, [])
# Solo añadir cuenta por defecto si viene por variables de entorno
if EMAIL and PASSWORD:
    if not any(acct.get("email") == EMAIL for acct in ACCOUNTS):
        ACCOUNTS.append({"email": EMAIL, "password": PASSWORD, "provider": DEFAULT_PROVIDER})
        save_json(ACCOUNTS_FILE, ACCOUNTS)


# Emails ya procesados: { account_email: { email_id: label } }
PROCESSED_DATA = load_json(PROCESSED_FILE, {})


def is_already_processed(account_email: str, email_id: str) -> bool:
    # Recargar desde archivo para evitar condiciones de carrera
    global PROCESSED_DATA
    PROCESSED_DATA = load_json(PROCESSED_FILE, {})
    return email_id in PROCESSED_DATA.get(account_email, {})


def get_processed_label(account_email: str, email_id: str):
    # Recargar desde archivo para asegurar datos actualizados
    global PROCESSED_DATA
    PROCESSED_DATA = load_json(PROCESSED_FILE, {})
    return PROCESSED_DATA.get(account_email, {}).get(email_id)


def mark_processed(account_email: str, email_id: str, label: str):
    # Recargar desde archivo antes de escribir para evitar sobrescrituras
    global PROCESSED_DATA
    PROCESSED_DATA = load_json(PROCESSED_FILE, {})
    PROCESSED_DATA.setdefault(account_email, {})[email_id] = label
    # Limitar tamaño por cuenta
    if len(PROCESSED_DATA[account_email]) > 10_000:
        oldest_keys = list(PROCESSED_DATA[account_email])[:1000]
        for k in oldest_keys:
            PROCESSED_DATA[account_email].pop(k, None)
    save_json(PROCESSED_FILE, PROCESSED_DATA)


def get_one_week_ago_datetime():
    """Retorna la fecha/hora de hace 1 semana en formato UTC."""
    return datetime.now(timezone.utc) - timedelta(days=7)


def parse_email_date(date_str: str) -> datetime | None:
    """Parsea una fecha de email (ISO 8601 o RFC 2822) y retorna datetime en UTC."""
    if not date_str:
        return None
    try:
        # Intentar parsear como ISO 8601 (formato de Microsoft Graph)
        if 'T' in date_str or '+' in date_str or date_str.endswith('Z'):
            # Formato ISO 8601
            if date_str.endswith('Z'):
                date_str = date_str[:-1] + '+00:00'
            dt = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            return dt
        # Intentar parsear como RFC 2822
        from email.utils import parsedate_to_datetime
        return parsedate_to_datetime(date_str)
    except Exception as e:
        try:
            safe_date = str(date_str).encode('ascii', 'ignore').decode('ascii')[:50]
            safe_error = str(e).encode('ascii', 'ignore').decode('ascii')
            print(f"Error al parsear fecha '{safe_date}': {safe_error}")
        except Exception:
            print("Error al parsear fecha")
        return None


def format_email_date(dt: datetime) -> str:
    """Formatea una fecha datetime a string legible."""
    if not dt:
        return "Fecha no disponible"
    try:
        # Convertir a zona horaria local si tiene tzinfo
        if dt.tzinfo:
            dt = dt.astimezone()
        return dt.strftime("%d/%m/%Y %H:%M")
    except Exception:
        return "Fecha no disponible"


def is_within_last_week(email_date: datetime) -> bool:
    """Verifica si una fecha está dentro de la última semana."""
    if not email_date:
        return False
    one_week_ago = get_one_week_ago_datetime()
    now = datetime.now(timezone.utc)
    # Normalizar ambas fechas a UTC para comparar
    if email_date.tzinfo is None:
        email_date = email_date.replace(tzinfo=timezone.utc)
    else:
        email_date = email_date.astimezone(timezone.utc)
    return one_week_ago <= email_date <= now


def decode_mime_words(s: str) -> str:
    if not s:
        return ""
    decoded_fragments = decode_header(s)
    out = ""
    for fragment, enc in decoded_fragments:
        if isinstance(fragment, bytes):
            encoding = enc or "utf-8"
            try:
                out += fragment.decode(encoding, errors="replace")
            except Exception:
                out += fragment.decode("latin-1", errors="replace")
        else:
            out += fragment
    return out


def get_body(msg) -> str:
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                payload = part.get_payload(decode=True)
                if payload is None:
                    continue
                charset = part.get_content_charset() or "utf-8"
                try:
                    return payload.decode(charset, errors="replace")
                except Exception:
                    return payload.decode("latin-1", errors="replace")
        return ""
    else:
        payload = msg.get_payload(decode=True)
        if payload is None:
            return ""
        try:
            return payload.decode("utf-8", errors="replace")
        except Exception:
            return payload.decode("latin-1", errors="replace")


def strip_html(html: str) -> str:
    if not html:
        return ""
    text = re.sub(r"<script[\s\S]*?</script>", " ", html, flags=re.I)
    text = re.sub(r"<style[\s\S]*?</style>", " ", text, flags=re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def get_email_attachments(access_token: str, email_id: str) -> list:
    """Obtiene la lista de adjuntos de un email desde Microsoft Graph API."""
    try:
        import sys
        headers = {"Authorization": f"Bearer {access_token}"}
        url = f"https://graph.microsoft.com/v1.0/me/messages/{email_id}/attachments"
        print(f"[DEBUG] Buscando adjuntos en email ID: {email_id[:50]}...", file=sys.stderr, flush=True)
        r = requests.get(url, headers=headers)
        if r.status_code >= 300:
            try:
                safe_error = str(r.text[:200]).encode('ascii', 'ignore').decode('ascii')
                print(f"[DEBUG] Error al obtener lista de adjuntos: {r.status_code} - {safe_error}", file=sys.stderr, flush=True)
            except Exception:
                print(f"[DEBUG] Error al obtener lista de adjuntos: {r.status_code}", file=sys.stderr, flush=True)
            return []
        attachments = r.json().get("value", [])
        print(f"[DEBUG] Se encontraron {len(attachments)} adjunto(s) en el email", file=sys.stderr, flush=True)
        for i, att in enumerate(attachments):
            try:
                att_name = att.get("name", "Sin nombre")
                att_size = att.get("size", 0)
                att_type = att.get("contentType", "Sin tipo")
                safe_name = str(att_name).encode('ascii', 'ignore').decode('ascii')
                print(f"[DEBUG] Adjunto {i+1}: {safe_name} (tamaño: {att_size} bytes, tipo: {att_type})", file=sys.stderr, flush=True)
            except Exception:
                print(f"[DEBUG] Adjunto {i+1}: (información no disponible)", file=sys.stderr, flush=True)
        return attachments
    except Exception as e:
        try:
            import sys
            safe_error = str(e).encode('ascii', 'ignore').decode('ascii')
            print(f"[DEBUG] Error al obtener adjuntos: {safe_error}", file=sys.stderr, flush=True)
        except Exception:
            pass
        return []


def extract_attachment_content(access_token: str, email_id: str, attachment_id: str, attachment_name: str, content_type: str) -> str:
    """Extrae el contenido de texto de un adjunto."""
    try:
        import sys
        try:
            safe_name = str(attachment_name).encode('ascii', 'ignore').decode('ascii')
            print(f"[DEBUG] Extrayendo contenido de adjunto: {safe_name} (tipo: {content_type})", file=sys.stderr, flush=True)
        except Exception:
            print(f"[DEBUG] Extrayendo contenido de adjunto (tipo: {content_type})", file=sys.stderr, flush=True)
        
        headers = {"Authorization": f"Bearer {access_token}"}
        # URL correcta según Microsoft Graph API: /me/messages/{message-id}/attachments/{attachment-id}/$value
        url = f"https://graph.microsoft.com/v1.0/me/messages/{email_id}/attachments/{attachment_id}/$value"
        print(f"[DEBUG] Descargando adjunto desde: {url[:100]}...", file=sys.stderr, flush=True)
        r = requests.get(url, headers=headers)
        if r.status_code >= 300:
            try:
                safe_error = str(r.text[:200]).encode('ascii', 'ignore').decode('ascii')
                print(f"[DEBUG] Error al descargar adjunto: {r.status_code} - {safe_error}", file=sys.stderr, flush=True)
            except Exception:
                print(f"[DEBUG] Error al descargar adjunto: {r.status_code}", file=sys.stderr, flush=True)
            return ""
        
        content = r.content
        if not content:
            print(f"[DEBUG] Adjunto descargado pero está vacío", file=sys.stderr, flush=True)
            return ""
        
        print(f"[DEBUG] Adjunto descargado correctamente (tamaño: {len(content)} bytes)", file=sys.stderr, flush=True)
        
        # Procesar según el tipo de contenido
        content_type_lower = (content_type or "").lower()
        attachment_name_lower = (attachment_name or "").lower()
        
        print(f"[DEBUG] Tipo de contenido detectado: {content_type_lower}, nombre: {attachment_name_lower}", file=sys.stderr, flush=True)
        
        # Texto plano
        if content_type_lower.startswith("text/"):
            try:
                return content.decode("utf-8", errors="replace")
            except Exception:
                try:
                    return content.decode("latin-1", errors="replace")
                except Exception:
                    return ""
        
        # PDF
        if content_type_lower == "application/pdf" or attachment_name_lower.endswith(".pdf"):
            print(f"[DEBUG] Detectado como PDF, intentando extraer texto...", file=sys.stderr, flush=True)
            try:
                import sys
                import PyPDF2
                from io import BytesIO
                print(f"[DEBUG] PyPDF2 importado correctamente, intentando extraer texto de PDF...", file=sys.stderr, flush=True)
                pdf_file = BytesIO(content)
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                text = ""
                num_pages = len(pdf_reader.pages)
                print(f"[DEBUG] PDF tiene {num_pages} página(s)", file=sys.stderr, flush=True)
                for i, page in enumerate(pdf_reader.pages):
                    try:
                        page_text = page.extract_text()
                        if page_text:
                            text += page_text + "\n"
                            print(f"[DEBUG] Texto extraído de página {i+1}/{num_pages} (longitud: {len(page_text)} caracteres)", file=sys.stderr, flush=True)
                        else:
                            print(f"[DEBUG] Página {i+1}/{num_pages} no contiene texto extraíble", file=sys.stderr, flush=True)
                    except Exception as e:
                        try:
                            safe_error = str(e).encode('ascii', 'ignore').decode('ascii')
                            print(f"[DEBUG] Error al extraer texto de página {i+1}: {safe_error}", file=sys.stderr, flush=True)
                        except Exception:
                            print(f"[DEBUG] Error al extraer texto de página {i+1}", file=sys.stderr, flush=True)
                        continue
                result = text.strip()
                if result:
                    print(f"[DEBUG] Texto total extraído del PDF: {len(result)} caracteres", file=sys.stderr, flush=True)
                else:
                    print(f"[DEBUG] No se pudo extraer texto del PDF (puede ser un PDF escaneado o protegido)", file=sys.stderr, flush=True)
                return result
            except ImportError:
                # Si PyPDF2 no está instalado, intentar con pdfplumber
                try:
                    import sys
                    import pdfplumber
                    from io import BytesIO
                    print(f"[DEBUG] PyPDF2 no disponible, intentando con pdfplumber...", file=sys.stderr, flush=True)
                    with pdfplumber.open(BytesIO(content)) as pdf:
                        text = ""
                        num_pages = len(pdf.pages)
                        print(f"[DEBUG] PDF tiene {num_pages} página(s)", file=sys.stderr, flush=True)
                        for i, page in enumerate(pdf.pages):
                            try:
                                page_text = page.extract_text()
                                if page_text:
                                    text += page_text + "\n"
                                    print(f"[DEBUG] Texto extraído de página {i+1}/{num_pages} (longitud: {len(page_text)} caracteres)", file=sys.stderr, flush=True)
                                else:
                                    print(f"[DEBUG] Página {i+1}/{num_pages} no contiene texto extraíble", file=sys.stderr, flush=True)
                            except Exception as e:
                                try:
                                    safe_error = str(e).encode('ascii', 'ignore').decode('ascii')
                                    print(f"[DEBUG] Error al extraer texto de página {i+1}: {safe_error}", file=sys.stderr, flush=True)
                                except Exception:
                                    print(f"[DEBUG] Error al extraer texto de página {i+1}", file=sys.stderr, flush=True)
                                continue
                        result = text.strip()
                        if result:
                            print(f"[DEBUG] Texto total extraído del PDF: {len(result)} caracteres", file=sys.stderr, flush=True)
                        else:
                            print(f"[DEBUG] No se pudo extraer texto del PDF (puede ser un PDF escaneado o protegido)", file=sys.stderr, flush=True)
                        return result
                except ImportError:
                    import sys
                    print(f"[DEBUG] ERROR: No se encontró PyPDF2 ni pdfplumber. Instala una de estas librerías para procesar PDFs.", file=sys.stderr, flush=True)
                    print(f"[DEBUG] Ejecuta: pip install PyPDF2 o pip install pdfplumber", file=sys.stderr, flush=True)
                    return ""
            except Exception as e:
                import sys
                import traceback
                try:
                    safe_error = str(e).encode('ascii', 'ignore').decode('ascii')
                    print(f"[DEBUG] Error al procesar PDF: {safe_error}", file=sys.stderr, flush=True)
                    # Imprimir el traceback completo para debugging
                    try:
                        traceback_str = traceback.format_exc()
                        safe_traceback = traceback_str.encode('ascii', 'ignore').decode('ascii')
                        print(f"[DEBUG] Traceback completo: {safe_traceback}", file=sys.stderr, flush=True)
                    except Exception:
                        pass
                except Exception:
                    print(f"[DEBUG] Error al procesar PDF", file=sys.stderr, flush=True)
                return ""
        
        # Si llegamos aquí, el tipo de archivo no es soportado
        print(f"[DEBUG] Tipo de archivo no soportado para extracción de texto: {content_type_lower}", file=sys.stderr, flush=True)
        
        # Word (docx)
        if content_type_lower in ["application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                                   "application/msword"] or attachment_name_lower.endswith((".docx", ".doc")):
            try:
                from docx import Document
                from io import BytesIO
                doc_file = BytesIO(content)
                doc = Document(doc_file)
                text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
                return text.strip()
            except ImportError:
                return ""
            except Exception:
                return ""
        
        # Excel (xlsx)
        if content_type_lower in ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   "application/vnd.ms-excel"] or attachment_name_lower.endswith((".xlsx", ".xls")):
            try:
                import pandas as pd
                from io import BytesIO
                excel_file = BytesIO(content)
                df = pd.read_excel(excel_file, sheet_name=None)
                text = ""
                for sheet_name, sheet_df in df.items():
                    text += f"Hoja: {sheet_name}\n"
                    text += sheet_df.to_string() + "\n\n"
                return text.strip()
            except ImportError:
                return ""
            except Exception:
                return ""
        
        return ""
    except Exception as e:
        try:
            safe_error = str(e).encode('ascii', 'ignore').decode('ascii')
            print(f"Error al extraer contenido de adjunto: {safe_error}")
        except Exception:
            pass
        return ""


def get_attachments_content(access_token: str, email_id: str) -> str:
    """Obtiene el contenido de texto de todos los adjuntos de un email."""
    import sys
    attachments = get_email_attachments(access_token, email_id)
    if not attachments:
        print(f"[DEBUG] No se encontraron adjuntos en el email", file=sys.stderr, flush=True)
        return ""
    
    all_content = []
    for i, att in enumerate(attachments):
        att_id = att.get("id")
        att_name = att.get("name", "")
        att_content_type = att.get("contentType", "")
        att_size = att.get("size", 0)
        
        try:
            safe_name = str(att_name).encode('ascii', 'ignore').decode('ascii')
            print(f"[DEBUG] Procesando adjunto {i+1}/{len(attachments)}: {safe_name}", file=sys.stderr, flush=True)
        except Exception:
            print(f"[DEBUG] Procesando adjunto {i+1}/{len(attachments)}", file=sys.stderr, flush=True)
        
        # Limitar tamaño de adjuntos para evitar problemas de memoria (max 10MB)
        if att_size > 10 * 1024 * 1024:
            print(f"[DEBUG] Adjunto {i+1} demasiado grande ({att_size} bytes), omitiendo", file=sys.stderr, flush=True)
            continue
        
        # Pasar email_id a extract_attachment_content para construir la URL correcta
        content = extract_attachment_content(access_token, email_id, att_id, att_name, att_content_type)
        if content:
            try:
                safe_content_preview = content[:200].encode('ascii', 'ignore').decode('ascii')
                print(f"[DEBUG] Contenido extraído del adjunto {i+1} (primeros 200 chars): {safe_content_preview}...", file=sys.stderr, flush=True)
            except Exception:
                print(f"[DEBUG] Contenido extraído del adjunto {i+1} (longitud: {len(content)} caracteres)", file=sys.stderr, flush=True)
            all_content.append(f"Adjunto: {att_name}\n{content}")
        else:
            print(f"[DEBUG] No se pudo extraer contenido del adjunto {i+1}", file=sys.stderr, flush=True)
    
    result = "\n\n---\n\n".join(all_content)
    if result:
        print(f"[DEBUG] Contenido total de adjuntos extraído (longitud: {len(result)} caracteres)", file=sys.stderr, flush=True)
    else:
        print(f"[DEBUG] No se pudo extraer contenido de ningún adjunto", file=sys.stderr, flush=True)
    return result


def get_graph_token_for_account(account_email: str):
    acct = next((a for a in ACCOUNTS if a.get("email") == account_email and a.get("provider") == "outlook_oauth"), None)
    if not acct:
        return None
    cache = msal.SerializableTokenCache()
    cache_state = acct.get("ms_cache")
    if cache_state:
        try:
            cache.deserialize(cache_state)
        except Exception:
            pass
    app_msal = _ms_app(cache)
    ms_accounts = app_msal.get_accounts(username=account_email) or app_msal.get_accounts()
    ms_account = ms_accounts[0] if ms_accounts else None
    result = app_msal.acquire_token_silent(MS_SCOPES, account=ms_account)
    if not result or "access_token" not in result:
        return None
    try:
        new_state = cache.serialize()
        if new_state and new_state != cache_state:
            acct["ms_cache"] = new_state
            save_json(ACCOUNTS_FILE, ACCOUNTS)
    except Exception:
        pass
    return result.get("access_token")


def setup_ms_graph_subscription(account_email: str, access_token: str) -> bool:
    """
    Crea una suscripción a notificaciones de Microsoft Graph para recibir notificaciones
    cuando llegan nuevos emails a la cuenta.
    """
    if not account_email or not access_token:
        print(f"[SUBSCRIPTION] Error: Cuenta o token no válidos para {account_email}")
        return False
    
    if not MS_GRAPH_NOTIFICATION_URL:
        print("[SUBSCRIPTION] Error: MS_GRAPH_NOTIFICATION_URL no configurado en variables de entorno")
        return False
    
    graph_api_base_url = "https://graph.microsoft.com/v1.0"
    subscriptions_url = f"{graph_api_base_url}/subscriptions"
    
    expiration_duration_minutes = MS_GRAPH_SUBSCRIPTION_DURATION_MINUTES
    # Usar datetime.utcnow() como en el proyecto de referencia
    expiration_datetime = datetime.utcnow() + timedelta(minutes=expiration_duration_minutes)
    client_state_secret = MS_GRAPH_SUBSCRIPTION_CLIENT_STATE
    
    # Formatear fecha exactamente como en el proyecto de referencia: .isoformat() + "Z"
    expiration_iso = expiration_datetime.isoformat() + "Z"
    
    subscription_payload = {
        "changeType": "created",
        "notificationUrl": MS_GRAPH_NOTIFICATION_URL,
        "resource": "me/mailFolders/inbox/messages",
        "expirationDateTime": expiration_iso,
        "clientState": client_state_secret
    }
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    try:
        # Debug: imprimir payload antes de enviar
        try:
            safe_payload_debug = json.dumps(subscription_payload, ensure_ascii=False, indent=2).encode('ascii', 'ignore').decode('ascii')
            print(f"[SUBSCRIPTION] Creando suscripcion para {account_email}")
            print(f"[SUBSCRIPTION] Payload: {safe_payload_debug}")
            print(f"[SUBSCRIPTION] URL: {MS_GRAPH_NOTIFICATION_URL}")
        except Exception:
            pass
        
        response = requests.post(subscriptions_url, headers=headers, json=subscription_payload, timeout=15)
        response.raise_for_status()
        subscription_data = response.json()
        try:
            safe_email = str(account_email).encode('ascii', 'ignore').decode('ascii')
            safe_id = str(subscription_data.get('id', 'N/A')).encode('ascii', 'ignore').decode('ascii')
            print(f"[SUBSCRIPTION] Suscripcion creada para {safe_email}. ID: {safe_id}")
        except Exception:
            print(f"[SUBSCRIPTION] Suscripcion creada exitosamente")
        
        sub_id = subscription_data.get("id")
        sub_expires_str = subscription_data.get("expirationDateTime")
        
        if not sub_id or not sub_expires_str:
            print(f"[SUBSCRIPTION] Error: Respuesta de suscripción inválida para {account_email}")
            return False
        
        # Actualizar la cuenta con la información de la suscripción
        acct = next((a for a in ACCOUNTS if a.get("email") == account_email), None)
        if acct:
            acct["ms_subscription_id"] = sub_id
            try:
                if sub_expires_str.endswith('Z'):
                    parsed_expiry = datetime.fromisoformat(sub_expires_str[:-1] + '+00:00')
                else:
                    parsed_expiry = datetime.fromisoformat(sub_expires_str)
                if parsed_expiry.tzinfo is None:
                    parsed_expiry = parsed_expiry.replace(tzinfo=timezone.utc)
                else:
                    parsed_expiry = parsed_expiry.astimezone(timezone.utc)
                acct["ms_subscription_expires_at"] = parsed_expiry.isoformat()
            except ValueError:
                # Si falla el parseo, usar la fecha calculada
                acct["ms_subscription_expires_at"] = expiration_datetime.isoformat()
            save_json(ACCOUNTS_FILE, ACCOUNTS)
        
        return True
        
    except requests.exceptions.HTTPError as e_http:
        error_text = ""
        try:
            if e_http.response and e_http.response.text:
                error_text = e_http.response.text[:500]
                # Intentar parsear JSON para obtener más detalles
                try:
                    error_json = e_http.response.json()
                    error_text = json.dumps(error_json, ensure_ascii=False)[:500]
                except Exception:
                    pass
        except Exception:
            error_text = "Sin detalles"
        
        try:
            safe_email = str(account_email).encode('ascii', 'ignore').decode('ascii')
            safe_error = str(error_text).encode('ascii', 'ignore').decode('ascii')
            print(f"[SUBSCRIPTION] Error HTTP creando suscripcion para {safe_email}: {e_http.response.status_code} - {safe_error}")
            # También imprimir el payload para debugging
            try:
                safe_payload = json.dumps(subscription_payload, ensure_ascii=False).encode('ascii', 'ignore').decode('ascii')
                print(f"[SUBSCRIPTION] Payload enviado: {safe_payload}")
            except Exception:
                pass
        except Exception:
            print(f"[SUBSCRIPTION] Error HTTP creando suscripcion: {e_http.response.status_code}")
        return False
    except Exception as e:
        try:
            safe_email = str(account_email).encode('ascii', 'ignore').decode('ascii')
            safe_error = str(e).encode('ascii', 'ignore').decode('ascii')
            print(f"[SUBSCRIPTION] Error inesperado creando suscripcion para {safe_email}: {safe_error}")
        except Exception:
            print(f"[SUBSCRIPTION] Error inesperado creando suscripcion")
        return False


def renew_ms_graph_subscriptions():
    """
    Renueva las suscripciones de Microsoft Graph que están próximas a expirar.
    Esta función debe ejecutarse periódicamente (por ejemplo, cada hora).
    """
    try:
        print("[RENEW] Verificando suscripciones de Microsoft Graph que necesitan renovación...")
        now = datetime.now(timezone.utc)
        renew_threshold = now + timedelta(hours=MS_GRAPH_RENEW_BEFORE_EXPIRY_HOURS)
        
        renewed_count = 0
        failed_count = 0
        
        for acct in ACCOUNTS:
            if acct.get("provider") != "outlook_oauth":
                continue
            
            subscription_id = acct.get("ms_subscription_id")
            expires_at_str = acct.get("ms_subscription_expires_at")
            
            if not subscription_id or not expires_at_str:
                continue
            
            try:
                expires_at = datetime.fromisoformat(expires_at_str)
                if expires_at.tzinfo is None:
                    expires_at = expires_at.replace(tzinfo=timezone.utc)
                else:
                    expires_at = expires_at.astimezone(timezone.utc)
            except (ValueError, TypeError):
                print(f"[RENEW] Error parseando fecha de expiración para {acct.get('email')}")
                continue
            
            # Si la suscripción expira dentro del umbral, renovarla
            if expires_at <= renew_threshold:
                account_email = acct.get("email")
                print(f"[RENEW] Renovando suscripción para {account_email} (expira: {expires_at})")
                
                # Obtener token
                token = get_graph_token_for_account(account_email)
                if not token:
                    print(f"[RENEW] No se pudo obtener token para {account_email}")
                    failed_count += 1
                    continue
                
                # Crear nueva suscripción
                if setup_ms_graph_subscription(account_email, token):
                    # Eliminar la suscripción antigua (opcional, Microsoft la elimina automáticamente al expirar)
                    try:
                        graph_api_base_url = "https://graph.microsoft.com/v1.0"
                        delete_url = f"{graph_api_base_url}/subscriptions/{subscription_id}"
                        headers = {'Authorization': f'Bearer {token}'}
                        requests.delete(delete_url, headers=headers, timeout=10)
                    except Exception:
                        pass  # No crítico si falla la eliminación
                    
                    renewed_count += 1
                    print(f"[RENEW] Suscripción renovada exitosamente para {account_email}")
                else:
                    failed_count += 1
                    print(f"[RENEW] Error al renovar suscripción para {account_email}")
        
        print(f"[RENEW] Renovación completada: {renewed_count} renovadas, {failed_count} fallidas")
        return renewed_count, failed_count
        
    except Exception as e:
        print(f"[RENEW] Error en renovación de suscripciones: {e}")
        import traceback
        traceback.print_exc()
        return 0, 0


def process_new_email_from_notification(account_email: str, email_id: str):
    """
    Procesa automáticamente un nuevo email cuando llega una notificación de Microsoft Graph.
    """
    try:
        try:
            safe_email = str(account_email).encode('ascii', 'ignore').decode('ascii')
            safe_id = str(email_id)[:50].encode('ascii', 'ignore').decode('ascii')
            print(f"[AUTO-PROCESS] Procesando email automaticamente - Cuenta: {safe_email}, ID: {safe_id}...")
        except Exception:
            print("[AUTO-PROCESS] Procesando email automaticamente")
        
        # Verificar si ya está procesado
        if is_already_processed(account_email, email_id):
            try:
                safe_id = str(email_id)[:50].encode('ascii', 'ignore').decode('ascii')
                print(f"[AUTO-PROCESS] Email ya procesado - ID: {safe_id}...")
            except Exception:
                print("[AUTO-PROCESS] Email ya procesado")
            return
        
        # Obtener el token
        token = get_graph_token_for_account(account_email)
        if not token:
            try:
                safe_email = str(account_email).encode('ascii', 'ignore').decode('ascii')
                print(f"[AUTO-PROCESS] No se pudo obtener token para {safe_email}")
            except Exception:
                print("[AUTO-PROCESS] No se pudo obtener token")
            return
        
        # Obtener el email completo desde Microsoft Graph
        headers = {"Authorization": f"Bearer {token}", "Prefer": 'outlook.body-content-type="text"'}
        email_url = f"https://graph.microsoft.com/v1.0/me/messages/{email_id}?$select=id,subject,from,body,receivedDateTime"
        
        r = requests.get(email_url, headers=headers)
        if r.status_code >= 300:
            try:
                safe_id = str(email_id)[:50].encode('ascii', 'ignore').decode('ascii')
                print(f"[AUTO-PROCESS] Error al obtener email {safe_id}...: {r.status_code}")
            except Exception:
                print(f"[AUTO-PROCESS] Error al obtener email: {r.status_code}")
            return
        
        email_data = r.json()
        
        # Extraer información del email
        subj = email_data.get("subject") or "Sin asunto"
        from_obj = email_data.get("from", {}).get("emailAddress", {})
        from_addr = f"{from_obj.get('name','')} <{from_obj.get('address','')}>".strip()
        body_obj = email_data.get("body", {})
        if body_obj.get("contentType", "").lower() == "html":
            body = strip_html(body_obj.get("content", ""))
        else:
            body = body_obj.get("content", "")
        
        # Obtener contenido de adjuntos
        attachments_content = ""
        try:
            attachments_content = get_attachments_content(token, email_id)
        except Exception as e:
            try:
                safe_error = str(e).encode('ascii', 'ignore').decode('ascii')
                print(f"[AUTO-PROCESS] Error al obtener adjuntos: {safe_error}")
            except Exception:
                print("[AUTO-PROCESS] Error al obtener adjuntos")
        
        # Combinar texto para clasificación
        classification_text = f"Asunto: {subj}\n\n"
        if body:
            classification_text += f"Cuerpo del email:\n{body}\n\n"
        if attachments_content:
            classification_text += f"Contenido de adjuntos:\n{attachments_content}"
        
        if not classification_text.strip() or classification_text.strip() == f"Asunto: {subj}\n\n":
            classification_text = subj
        
        # Clasificar el email
        label = classify_email(classification_text)
        
        # Si no se clasificó, marcar como procesado y salir
        if label == "Sin etiqueta":
            mark_processed(account_email, email_id, label)
            try:
                safe_id = str(email_id)[:50].encode('ascii', 'ignore').decode('ascii')
                print(f"[AUTO-PROCESS] Email sin etiqueta - ID: {safe_id}...")
            except Exception:
                print("[AUTO-PROCESS] Email sin etiqueta")
            return
        
        # Obtener la categoría completa
        category = get_category_by_name(label)
        
        # Aplicar etiqueta (mover a carpeta)
        label_success = apply_label_graph(email_id, label, account_email)
        
        # Enviar respuesta automática si está configurada
        reply_sent = False
        if label_success and category and category.get('autoReply') and category.get('replyText'):
            reply_text = category.get('replyText', '')
            try:
                reply_sent = send_reply_graph(email_id, from_addr, subj, reply_text, account_email)
            except Exception as e:
                try:
                    safe_error = str(e).encode('ascii', 'ignore').decode('ascii')
                    print(f"[AUTO-PROCESS] Error al enviar respuesta automatica: {safe_error}")
                except Exception:
                    print("[AUTO-PROCESS] Error al enviar respuesta automatica")
        
        # Marcar como procesado
        mark_processed(account_email, email_id, label)
        try:
            safe_id = str(email_id)[:50].encode('ascii', 'ignore').decode('ascii')
            safe_label = str(label).encode('ascii', 'ignore').decode('ascii')
            reply_status = "Enviada" if reply_sent else "No enviada"
            print(f"[AUTO-PROCESS] Email procesado automaticamente - ID: {safe_id}... | Etiqueta: {safe_label} | Respuesta: {reply_status}")
        except Exception:
            print("[AUTO-PROCESS] Email procesado automaticamente")
        
    except Exception as e:
        try:
            safe_error = str(e).encode('ascii', 'ignore').decode('ascii')
            print(f"[AUTO-PROCESS] Error procesando email automaticamente: {safe_error}")
        except Exception:
            print("[AUTO-PROCESS] Error procesando email automaticamente")
        import traceback
        try:
            traceback_str = traceback.format_exc()
            safe_traceback = traceback_str.encode('ascii', 'ignore').decode('ascii')
            print(safe_traceback)
        except Exception:
            traceback.print_exc()


def get_or_create_folder_id(access_token: str, display_name: str):
    headers = {"Authorization": f"Bearer {access_token}"}
    try:
        # Buscar la carpeta primero
        r = requests.get("https://graph.microsoft.com/v1.0/me/mailFolders?$select=id,displayName", headers=headers)
        if r.status_code >= 300:
            try:
                safe_error = str(r.text[:200]).encode('ascii', 'ignore').decode('ascii')
                print(f"Error al obtener carpetas: {r.status_code} - {safe_error}")
            except Exception:
                print(f"Error al obtener carpetas: {r.status_code}")
            return None
        
        # Buscar la carpeta por nombre (case-insensitive)
        display_name_clean = display_name.strip().lower()
        for f in r.json().get("value", []):
            folder_name = f.get("displayName", "").strip().lower()
            if folder_name == display_name_clean:
                return f.get("id")
        
        # Si no existe, intentar crearla
        r2 = requests.post("https://graph.microsoft.com/v1.0/me/mailFolders", 
                          headers={**headers, "Content-Type": "application/json"}, 
                          json={"displayName": display_name})
        
        # Si el error es 409 (ya existe), buscar de nuevo (puede haber sido creada entre tanto)
        if r2.status_code == 409:
            # Buscar de nuevo por si fue creada por otro proceso
            r3 = requests.get("https://graph.microsoft.com/v1.0/me/mailFolders?$select=id,displayName", headers=headers)
            if r3.status_code < 300:
                for f in r3.json().get("value", []):
                    folder_name = f.get("displayName", "").strip().lower()
                    if folder_name == display_name_clean:
                        return f.get("id")
            # Si aún no se encuentra, devolver None
            try:
                safe_name = str(display_name).encode('ascii', 'ignore').decode('ascii')
                print(f"La carpeta '{safe_name}' ya existe pero no se pudo obtener su ID")
            except Exception:
                print("La carpeta ya existe pero no se pudo obtener su ID")
            return None
        
        if r2.status_code >= 300:
            try:
                safe_name = str(display_name).encode('ascii', 'ignore').decode('ascii')
                safe_error = str(r2.text[:200]).encode('ascii', 'ignore').decode('ascii')
                print(f"Error al crear carpeta '{safe_name}': {r2.status_code} - {safe_error}")
            except Exception:
                print(f"Error al crear carpeta: {r2.status_code}")
            return None
        
        return r2.json().get("id")
    except Exception as e:
        try:
            safe_name = str(display_name).encode('ascii', 'ignore').decode('ascii')
            safe_error = str(e).encode('ascii', 'ignore').decode('ascii')
            print(f"Excepcion al crear/obtener carpeta '{safe_name}': {safe_error}")
        except Exception:
            print("Excepcion al crear/obtener carpeta")
        return None


def classify_email(text: str) -> str:
    """Devuelve el nombre de la categoria configurada que mejor se ajusta al mensaje, o "Sin etiqueta" si no coincide con ninguna."""
    import sys
    try:
        safe_text = str(text[:100]).encode('ascii', 'ignore').decode('ascii')
        print(f"[CLASSIFY] Iniciando clasificacion. Texto recibido: {safe_text}...", file=sys.stderr, flush=True)
    except Exception:
        print(f"[CLASSIFY] Iniciando clasificacion.", file=sys.stderr, flush=True)
    
    categories = CATEGORIES if CATEGORIES else []
    print(f"[CLASSIFY] Categorias disponibles: {len(categories)}", file=sys.stderr, flush=True)
    
    if not categories:
        print("[CLASSIFY] No hay categorias configuradas", file=sys.stderr, flush=True)
        return "Sin etiqueta"
    
    names = [cat.get('name') for cat in categories if cat.get('name')]
    print(f"[CLASSIFY] Nombres de categorias: {names}", file=sys.stderr, flush=True)
    
    if not names:
        print("[CLASSIFY] No hay nombres de categorias validos", file=sys.stderr, flush=True)
        return "Sin etiqueta"

    guidance_lines = [
        'Actuas como un clasificador de correos estricto pero inteligente.',
        'IMPORTANTE: Solo debes devolver una etiqueta si el mensaje REALMENTE coincide con la descripcion de esa categoria.',
        'Si el mensaje NO coincide con NINGUNA de las categorias disponibles, debes responder EXACTAMENTE con: "Sin etiqueta"',
        'No clasifiques mensajes que no tengan relacion con ninguna categoria.',
        'Debes devolver una unica etiqueta EXACTAMENTE igual a una de las listadas, o "Sin etiqueta" si no coincide.',
        'Compara el contenido del mensaje (asunto y cuerpo) con la descripcion de cada categoria y elige la etiqueta mas especifica que aplique.',
        'Si una categoria describe explicitamente el caso, prefierela frente a otras mas genericas.',
        'Presta atencion a palabras clave: si el mensaje contiene palabras como "reserva", "reservar", "reservacion", "booking", "hotel", "vuelo", "viaje", "alojamiento", "rent a car" o similares, y hay una categoria que menciona reservas, DEBES clasificarlo en esa categoria.',
        'Si el asunto o el cuerpo del mensaje contiene la palabra "reserva" o palabras relacionadas con reservas/viajes, y existe una categoria de "Reservas", DEBES clasificarlo como "Reservas".',
        'No inventes etiquetas nuevas ni respondas con texto adicional.',
        'Categorias disponibles:',
    ]
    for cat in categories:
        prompt = cat.get('prompt') or 'Sin descripcion adicional.'
        # Asegurar que el prompt esté correctamente codificado
        if isinstance(prompt, bytes):
            prompt = prompt.decode('utf-8', errors='ignore')
        guidance_lines.append(f"- {cat['name']}: {prompt}")
    guidance_lines.append('\nRecuerda: Si el mensaje no coincide con ninguna categoria, responde EXACTAMENTE: "Sin etiqueta"')
    guidance = "\n".join(guidance_lines)
    
    # Log del prompt completo para debug
    import sys
    try:
        safe_guidance = str(guidance[:500]).encode('ascii', 'ignore').decode('ascii')
        print(f"[CLASSIFY] Prompt completo (primeros 500 chars): {safe_guidance}...", file=sys.stderr, flush=True)
    except Exception:
        print(f"[CLASSIFY] Prompt completo", file=sys.stderr, flush=True)

    few_shots: list[dict[str, str]] = []
    # Agregar ejemplos específicos para cada categoría
    for cat in categories[:2]:
        prompt_text = cat.get('prompt', '')
        # Crear ejemplos más específicos basados en el prompt
        if 'reserva' in prompt_text.lower() or 'viaje' in prompt_text.lower():
            # Ejemplos para reservas
            examples = [
                f"Mensaje:\nreserva hotel\n\nNecesito reservar un hotel para mis vacaciones.",
                f"Mensaje:\nconfirmacion de reserva de viaje\n\nTu reserva ha sido confirmada.",
                f"Mensaje:\nreserva de hotel\n\nHemos recibido tu solicitud de reserva."
            ]
            for example in examples[:2]:  # Usar 2 ejemplos
                few_shots.append({'role': 'user', 'content': example})
                few_shots.append({'role': 'assistant', 'content': cat['name']})
        else:
            # Para otras categorías, usar el formato original
            example_text = f"Mensaje:\nEste correo encaja claramente con la categoria \"{cat['name']}\". {prompt_text}".strip()
            few_shots.append({'role': 'user', 'content': example_text})
            few_shots.append({'role': 'assistant', 'content': cat['name']})
    
    # Agregar ejemplo de "Sin etiqueta"
    few_shots.append({'role': 'user', 'content': 'Mensaje:\nEste es un correo de notificacion del sistema que no tiene relacion con ninguna categoria configurada.'})
    few_shots.append({'role': 'assistant', 'content': 'Sin etiqueta'})

    # Asegurar que el texto esté correctamente codificado
    if isinstance(text, bytes):
        text = text.decode('utf-8', errors='ignore')
    
    user_message = f"Mensaje:\n{text}\n\nResponde solo con una de las etiquetas o 'Sin etiqueta' si no coincide con ninguna."
    messages = [{'role': 'system', 'content': guidance}, *few_shots,
                {'role': 'user', 'content': user_message}]

    try:
        import sys
        print(f"[CLASSIFY] Enviando a OpenAI. Categorias: {names}", file=sys.stderr, flush=True)
        try:
            safe_text = str(text[:200]).encode('ascii', 'ignore').decode('ascii')
            print(f"[CLASSIFY] Texto completo del mensaje: {safe_text}...", file=sys.stderr, flush=True)
        except Exception:
            print(f"[CLASSIFY] Texto completo del mensaje", file=sys.stderr, flush=True)
        
        resp = openai.chat.completions.create(
            model='gpt-4o-mini',
            messages=messages,
            temperature=0,
            max_tokens=20,
        )
        raw_label = (resp.choices[0].message.content or '').strip()
        print(f"[CLASSIFY] Respuesta de OpenAI: '{raw_label}'", file=sys.stderr, flush=True)
        
        # Si devuelve "Sin etiqueta", retornarlo directamente
        if raw_label.lower() in ['sin etiqueta', 'sin etiqueta.', 'no etiqueta', 'ninguna']:
            print(f"[CLASSIFY] OpenAI devolvio 'Sin etiqueta'", file=sys.stderr, flush=True)
            return "Sin etiqueta"
        
        # Verificar si coincide exactamente con alguna etiqueta
        if raw_label in names:
            print(f"[CLASSIFY] Coincidencia exacta encontrada: '{raw_label}'", file=sys.stderr, flush=True)
            return raw_label
        
        # Verificar coincidencia case-insensitive
        lowered = {name.lower(): name for name in names}
        lowered_key = raw_label.lower() if raw_label else ''
        if lowered_key in lowered:
            matched = lowered[lowered_key]
            print(f"[CLASSIFY] Coincidencia case-insensitive: '{raw_label}' -> '{matched}'", file=sys.stderr, flush=True)
            return matched
        
        # Si no coincide, devolver "Sin etiqueta"
        print(f"[CLASSIFY] ERROR: '{raw_label}' no coincide con ninguna categoria. Disponibles: {names}", file=sys.stderr, flush=True)
        return "Sin etiqueta"
        
    except Exception as e:
        import sys
        print(f'[CLASSIFY] Error al clasificar con OpenAI: {e}', file=sys.stderr, flush=True)
        import traceback
        traceback.print_exc(file=sys.stderr)
        return "Sin etiqueta"

def get_category_by_name(label_name: str):
    """Obtiene la categoría completa por nombre, incluyendo autoReply y replyText."""
    for cat in CATEGORIES:
        if cat.get('name') == label_name:
            return cat
    return None

def send_reply_imap(email_id: str, to_email: str, subject: str, reply_text: str, 
                    account_email: str, account_password: str, provider: str):
    """Envía una respuesta automática usando SMTP (IMAP)."""
    try:
        smtp_server, smtp_port = SMTP_SERVERS.get(provider or "gmail", SMTP_SERVERS["gmail"])
        # Extraer email del remitente si viene en formato "Nombre <email>"
        to_addr = parseaddr(to_email)[1] if to_email else None
        if not to_addr:
            print(f"No se pudo extraer direccion de email de: {to_email}")
            return False
        
        # Crear el mensaje de respuesta
        msg = MIMEMultipart()
        msg['From'] = account_email
        msg['To'] = to_addr
        msg['Subject'] = f"Re: {subject}" if not subject.startswith("Re:") else subject
        
        msg.attach(MIMEText(reply_text, 'plain', 'utf-8'))
        
        # Enviar el email
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(account_email, account_password)
            server.send_message(msg)
        
        print(f"Respuesta automatica enviada a {to_addr}")
        return True
    except Exception as e:
        print(f"Error al enviar respuesta automatica (IMAP): {e}")
        return False

def send_reply_graph(email_id: str, to_email: str, subject: str, reply_text: str, 
                     account_email: str):
    """Envía una respuesta automática usando Microsoft Graph API."""
    try:
        token = get_graph_token_for_account(account_email)
        if not token:
            print("No token Graph disponible para enviar respuesta")
            return False
        
        # Extraer email del remitente
        to_addr = parseaddr(to_email)[1] if to_email else None
        if not to_addr:
            print(f"No se pudo extraer direccion de email de: {to_email}")
            return False
        
        # Crear el mensaje de respuesta
        reply_subject = f"Re: {subject}" if not subject.startswith("Re:") else subject
        
        message = {
            "message": {
                "subject": reply_subject,
                "body": {
                    "contentType": "Text",
                    "content": reply_text
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": to_addr
                        }
                    }
                ]
            }
        }
        
        # Enviar el email
        url = "https://graph.microsoft.com/v1.0/me/sendMail"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        r = requests.post(url, headers=headers, json=message)
        
        if r.status_code >= 300:
            try:
                safe_error = str(r.text[:200]).encode('ascii', 'ignore').decode('ascii')
                print(f"Error al enviar respuesta automatica (Graph): {r.status_code} - {safe_error}")
            except Exception:
                print(f"Error al enviar respuesta automatica (Graph): {r.status_code}")
            return False
        
        print(f"Respuesta automatica enviada a {to_addr}")
        return True
    except Exception as e:
        print(f"Error al enviar respuesta automatica (Graph): {e}")
        return False

def apply_label_graph(email_id, label, account_email):
    try:
        print(f"[DEBUG] apply_label_graph - ID: {email_id[:50]}... | Carpeta: {label}")
        token = get_graph_token_for_account(account_email)
        if not token:
            print(f"[DEBUG] No token Graph disponible - ID: {email_id[:50]}...")
            return False
        folder_id = get_or_create_folder_id(token, label)
        if not folder_id:
            print(f"[DEBUG] No se pudo asegurar carpeta {label} - ID: {email_id[:50]}...")
            return False
        print(f"[DEBUG] Carpeta obtenida - ID: {email_id[:50]}... | Folder ID: {folder_id[:30]}...")
        url = f"https://graph.microsoft.com/v1.0/me/messages/{email_id}/copy"
        r = requests.post(url, headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"}, json={"destinationId": folder_id})
        if r.status_code >= 300:
            safe_error = str(r.text[:200]).encode('ascii', 'ignore').decode('ascii')
            print(f"[DEBUG] Graph copy fallo - ID: {email_id[:50]}... | Status: {r.status_code} | Error: {safe_error}")
            return False
        print(f"[DEBUG] Graph copy exitoso - ID: {email_id[:50]}... | Carpeta: {label}")
        return True
    except Exception as e:
        safe_error = str(e).encode('ascii', 'ignore').decode('ascii')
        print(f"[DEBUG] apply_label_graph error - ID: {email_id[:50]}... | Error: {safe_error}")
        return False

def apply_label(email_id, label, account_email, account_password, provider: str):
    """
    En Gmail: aplica etiqueta con X-GM-LABELS.
    En Outlook: copia el mensaje a una carpeta con el nombre de la etiqueta (si no existe la crea).
    """
    try:
        server = IMAP_SERVERS.get(provider or "gmail", IMAP_SERVERS["gmail"])
        with imaplib.IMAP4_SSL(server) as mail:
            mail.login(account_email, account_password)
            mail.select("INBOX")

            if (provider or "gmail").lower() == "gmail":
                result, _ = mail.store(email_id, '+X-GM-LABELS', f'"{label}"')
                if result != 'OK':
                    print(f"No se pudo aplicar etiqueta {label} a {email_id}")
                    return False
                return True
            else:
                # Outlook/IMAP: usar carpetas. Crear si no existe y copiar.
                try:
                    # Intentar crear carpeta (si ya existe, IMAP devuelve error benigno)
                    mail.create(label)
                except Exception:
                    pass
                # Asegurar que el nombre esté entre comillas para carpetas con espacios
                dest_folder = f'"{label}"'
                # Copiar a la carpeta
                result = mail.copy(email_id, dest_folder)
                if result[0] != 'OK':
                    print(f"No se pudo copiar el correo {email_id} a la carpeta {label}")
                    return False
                return True
    except Exception as e:
        print("apply_label error:", e)
        return False


@app.route('/')
def home():
    return render_template('index.html')


@app.route('/categories', methods=['GET', 'POST'])
def categories_endpoint():
    global CATEGORIES
    if request.method == 'GET':
        return jsonify(CATEGORIES)
    payload = request.get_json(silent=True)
    if isinstance(payload, list):
        submitted = payload
    elif isinstance(payload, dict):
        submitted = payload.get('categories')
    else:
        submitted = None
    if not isinstance(submitted, list):
        return jsonify({'error': "Se esperaba una lista en 'categories'."}), 400
    normalized = normalize_categories(submitted)
    if not normalized:
        return jsonify({'error': 'Define al menos una categoria con nombre.'}), 400
    CATEGORIES = normalized
    save_categories(CATEGORIES)
    return jsonify({'success': True, 'categories': CATEGORIES})


@app.route('/accounts', methods=['GET'])
def get_accounts():
    # Normalizar provider para cuentas antiguas - ahora solo soportamos outlook_oauth
    global ACCOUNTS
    needs_save = False
    normalized = []
    for acct in ACCOUNTS:
        # Crear una copia del diccionario para no modificar el original directamente
        acct_copy = dict(acct)
        # Asegurar que tenga el campo email (requerido)
        if not acct_copy.get("email"):
            continue  # Saltar cuentas sin email
        # Normalizar provider a outlook_oauth
        if not acct_copy.get("provider") or acct_copy.get("provider").lower() != "outlook_oauth":
            acct_copy["provider"] = "outlook_oauth"
            needs_save = True
        normalized.append(acct_copy)
    # Si hemos normalizado algo, actualizar ACCOUNTS y persistir
    if needs_save:
        ACCOUNTS = normalized
        try:
            save_json(ACCOUNTS_FILE, ACCOUNTS)
        except Exception as e:
            print(f"Error al guardar cuentas normalizadas: {e}")
    # Siempre devolver las cuentas normalizadas
    return jsonify(normalized)


@app.route('/add-account', methods=['POST'])
def add_account():
    data = request.get_json()
    new_email = data.get("email")
    new_provider = (data.get("provider") or "outlook_oauth").lower()
    
    if not new_email:
        return jsonify({"error": "Se requiere email."}), 400
    
    if new_provider != "outlook_oauth":
        return jsonify({"error": "Solo se soporta Outlook OAuth."}), 400
    
    if any(acct.get("email") == new_email for acct in ACCOUNTS):
        return jsonify({"error": "La cuenta ya existe."}), 400
    
    if not (MS_CLIENT_ID and MS_CLIENT_SECRET):
        return jsonify({"error": "Falta configuración OAuth de Microsoft en .env"}), 500
    
    app_msal = _ms_app()
    auth_url = app_msal.get_authorization_request_url(
        scopes=["Mail.Read","Mail.ReadWrite","Mail.Send"],
        redirect_uri=_effective_redirect_uri(),
        state=new_email,
        login_hint=new_email,
        prompt="select_account",
    )
    return jsonify({"success": True, "authUrl": auth_url, "message": "Redirigiendo a Microsoft para autorizar."})


@app.route('/fetch-emails', methods=['GET'])
def fetch_emails():
    selected_email = request.args.get('accountEmail')
    if not selected_email:
        return jsonify({"error": "No se proporcionó cuenta."}), 400

    account = next((acct for acct in ACCOUNTS if acct.get("email") == selected_email), None)
    if not account:
        return jsonify({"error": "Cuenta no encontrada."}), 400

    provider = (account.get("provider") or "outlook_oauth").lower()
    if provider != "outlook_oauth":
        return jsonify({"error": "Solo se soporta Outlook OAuth."}), 400

    emails = []
    try:
        token = get_graph_token_for_account(account["email"])
        if not token:
            return jsonify({"error": "No autorizado en Microsoft. Vuelve a conectar la cuenta."}), 401
        
        headers = {"Authorization": f"Bearer {token}", "Prefer": 'outlook.body-content-type="text"'}
        
        # Calcular fecha de hace 1 semana para el filtro
        one_week_ago = get_one_week_ago_datetime()
        # Formato ISO 8601 para Microsoft Graph API
        filter_date = one_week_ago.strftime("%Y-%m-%dT%H:%M:%SZ")
        
        # INBOX - Obtener emails y filtrar por fecha en el código (más confiable que filtro OData)
        # Obtener más emails de los que necesitamos y filtrar localmente
        inbox_url = "https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages?$top=100&$select=id,subject,from,body,receivedDateTime&$orderby=receivedDateTime desc"
        
        print(f"[DEBUG] Fetching emails from INBOX")
        print(f"[DEBUG] One week ago: {one_week_ago}")
        print(f"[DEBUG] Filter date string: {filter_date}")
        
        r = requests.get(inbox_url, headers=headers)
        if r.status_code >= 300:
            error_text = r.text[:500] if r.text else "No error message"
            print(f"[ERROR] Graph API error INBOX {r.status_code}: {error_text}")
            return jsonify({"error": f"Graph error INBOX {r.status_code}: {error_text}"}), 500
        
        for m in r.json().get("value", []):
            # Extraer fecha recibida
            received_date_str = m.get("receivedDateTime")
            received_date = parse_email_date(received_date_str) if received_date_str else None
            
            # Filtrar solo emails de la última semana
            if not received_date or not is_within_last_week(received_date):
                continue
            
            subj = m.get("subject") or "Sin asunto"
            from_obj = m.get("from", {}).get("emailAddress", {})
            from_addr = f"{from_obj.get('name','')} <{from_obj.get('address','')}>".strip()
            body_obj = m.get("body", {})
            if body_obj.get("contentType", "").lower() == "html":
                body = strip_html(body_obj.get("content", ""))
            else:
                body = body_obj.get("content", "")
            
            emails.append({
                "id": m.get("id"),
                "subject": subj,
                "body": body or "Sin contenido",
                "folder": "inbox",
                "from": from_addr or "Unknown",
                "date": received_date_str,  # Fecha en formato ISO para el frontend
                "dateFormatted": format_email_date(received_date) if received_date else "Fecha no disponible",
            })
        
        # Junk/Spam carpeta conocida - Obtener emails y filtrar por fecha en el código
        spam_url = "https://graph.microsoft.com/v1.0/me/mailFolders/JunkEmail/messages?$top=100&$select=id,subject,from,body,receivedDateTime&$orderby=receivedDateTime desc"
        rj = requests.get(spam_url, headers=headers)
        if rj.status_code < 300:
            for m in rj.json().get("value", []):
                # Extraer fecha recibida
                received_date_str = m.get("receivedDateTime")
                received_date = parse_email_date(received_date_str) if received_date_str else None
                
                # Filtrar solo emails de la última semana
                if not received_date or not is_within_last_week(received_date):
                    continue
                
                subj = m.get("subject") or "Sin asunto"
                from_obj = m.get("from", {}).get("emailAddress", {})
                from_addr = f"{from_obj.get('name','')} <{from_obj.get('address','')}>".strip()
                body_obj = m.get("body", {})
                if body_obj.get("contentType", "").lower() == "html":
                    body = strip_html(body_obj.get("content", ""))
                else:
                    body = body_obj.get("content", "")
                
                emails.append({
                    "id": m.get("id"),
                    "subject": subj,
                    "body": body or "Sin contenido",
                    "folder": "spam",
                    "from": from_addr or "Unknown",
                    "date": received_date_str,  # Fecha en formato ISO para el frontend
                    "dateFormatted": format_email_date(received_date) if received_date else "Fecha no disponible",
                })
        
        return jsonify(emails)
    except Exception as e:
        try:
            safe_error = str(e).encode('ascii', 'ignore').decode('ascii')
            print(f"Error al obtener emails: {safe_error}")
        except Exception:
            print("Error al obtener emails")
        return jsonify({"error": f"Error al obtener emails: {str(e)}"}), 500


@app.route('/classify', methods=['POST'])
def classify():
    import sys
    global CATEGORIES
    
    # Recargar categorias para asegurar que esten actualizadas
    CATEGORIES = load_categories()
    print(f"[CLASSIFY ENDPOINT] Categorias cargadas: {len(CATEGORIES)}", file=sys.stderr, flush=True)
    for cat in CATEGORIES:
        try:
            safe_name = str(cat.get('name', '')).encode('ascii', 'ignore').decode('ascii')
            safe_prompt = str(cat.get('prompt', '')[:50]).encode('ascii', 'ignore').decode('ascii')
            print(f"[CLASSIFY ENDPOINT] - {safe_name}: {safe_prompt}...", file=sys.stderr, flush=True)
        except Exception:
            print(f"[CLASSIFY ENDPOINT] - categoria", file=sys.stderr, flush=True)
    
    data = request.get_json()
    text = data.get("text")
    email_id = data.get("id")
    account_email = data.get("accountEmail")
    from_addr = data.get("from", "")
    original_subject = data.get("originalSubject", "Sin asunto")

    try:
        safe_subject = str(original_subject[:50]).encode('ascii', 'ignore').decode('ascii')
        safe_text = str(text[:100]).encode('ascii', 'ignore').decode('ascii')
        print(f"[CLASSIFY ENDPOINT] Recibido - Asunto: {safe_subject}... | Texto: {safe_text}...", file=sys.stderr, flush=True)
    except Exception:
        print(f"[CLASSIFY ENDPOINT] Recibido", file=sys.stderr, flush=True)

    if not text or not email_id or not account_email:
        print(f"[CLASSIFY ENDPOINT] ERROR: Faltan campos. text={bool(text)}, id={bool(email_id)}, account={bool(account_email)}", file=sys.stderr, flush=True)
        return jsonify({"error": "Faltan campos."}), 400

    # Log para debugging - ID completo sin truncar para verificar si son diferentes
    try:
        safe_subject = original_subject.encode('ascii', 'ignore').decode('ascii')[:50]
        safe_from = from_addr.encode('ascii', 'ignore').decode('ascii')[:50]
        # Log ID completo SIN TRUNCAR para verificar si realmente son iguales
        safe_id = email_id.encode('ascii', 'ignore').decode('ascii')
        safe_msg = f"[DEBUG] Clasificando email - ID COMPLETO (sin truncar): {safe_id} | Asunto: {safe_subject}... | De: {safe_from}..."
        print(safe_msg)
    except Exception:
        try:
            safe_id = email_id.encode('ascii', 'ignore').decode('ascii')
            print(f"[DEBUG] Clasificando email - ID COMPLETO (sin truncar): {safe_id}")
        except Exception:
            pass  # Si incluso esto falla, continuar sin log

    # Verificar si ya está procesado ANTES de procesar
    # IMPORTANTE: Recargar desde archivo para evitar condiciones de carrera
    if is_already_processed(account_email, email_id):
        # Log para debugging
        try:
            safe_id = email_id.encode('ascii', 'ignore').decode('ascii')
            print(f"[DEBUG] Email YA PROCESADO detectado - ID COMPLETO: {safe_id}")
        except Exception:
            pass
        prev_label = get_processed_label(account_email, email_id)
        # Si está en procesamiento, esperar un momento y reintentar
        if prev_label == "EN_PROCESO":
            try:
                safe_id = email_id.encode('ascii', 'ignore').decode('ascii')
                print(f"[DEBUG] Email en procesamiento - ID completo: {safe_id} | Esperando...")
            except Exception:
                pass
            # Esperar un poco más y reintentar varias veces
            for i in range(3):
                time.sleep(0.3)  # Esperar 300ms cada vez
                global PROCESSED_DATA
                PROCESSED_DATA = load_json(PROCESSED_FILE, {})
                if email_id in PROCESSED_DATA.get(account_email, {}):
                    final_label = PROCESSED_DATA[account_email][email_id]
                    if final_label != "EN_PROCESO":
                        try:
                            safe_id = email_id.encode('ascii', 'ignore').decode('ascii')
                            print(f"[DEBUG] Email procesado (despues de espera) - ID completo: {safe_id} | Etiqueta: {final_label}")
                        except Exception:
                            pass
                        return jsonify({"label": final_label, "alreadyProcessed": True})
            # Si después de esperar sigue en procesamiento, devolver error
            try:
                safe_id = email_id.encode('ascii', 'ignore').decode('ascii')
                print(f"[DEBUG] Email todavia en procesamiento despues de espera - ID completo: {safe_id}")
            except Exception:
                pass
            return jsonify({"label": "EN_PROCESO", "alreadyProcessed": True, "error": "Email en procesamiento"})
        else:
            try:
                safe_id = email_id.encode('ascii', 'ignore').decode('ascii')
                print(f"[DEBUG] Email ya procesado - ID completo: {safe_id} | Etiqueta previa: {prev_label}")
            except Exception:
                pass
            return jsonify({"label": prev_label, "alreadyProcessed": True})

    # IMPORTANTE: Marcar como "en procesamiento" ANTES de procesar para evitar duplicados
    # Si otro request llega con el mismo ID, será detectado como ya procesado
    try:
        safe_id = email_id.encode('ascii', 'ignore').decode('ascii')
        print(f"[DEBUG] Marcando como EN_PROCESO - ID completo: {safe_id}")
    except Exception:
        pass
    mark_processed(account_email, email_id, "EN_PROCESO")

    account = next((acct for acct in ACCOUNTS if acct.get("email") == account_email), None)
    if not account:
        # Si no hay cuenta, desmarcar el procesamiento
        PROCESSED_DATA.get(account_email, {}).pop(email_id, None)
        save_json(PROCESSED_FILE, PROCESSED_DATA)
        return jsonify({"error": "Cuenta no encontrada."}), 400

    # Obtener el cuerpo del email si no se proporcionó
    email_body = data.get("body", "")
    if not email_body and text:
        # Si text contiene el asunto y el cuerpo, separarlos
        parts = text.split("\n\n", 1)
        if len(parts) > 1:
            email_body = parts[1]
        else:
            email_body = text
    
    # Obtener el contenido de los adjuntos
    attachments_content = ""
    try:
        import sys
        print(f"[DEBUG] Intentando obtener adjuntos para email ID: {email_id[:50]}...", file=sys.stderr, flush=True)
        token = get_graph_token_for_account(account_email)
        if token:
            print(f"[DEBUG] Token obtenido, buscando adjuntos...", file=sys.stderr, flush=True)
            attachments_content = get_attachments_content(token, email_id)
            if attachments_content:
                try:
                    safe_attachments_preview = attachments_content[:500].encode('ascii', 'ignore').decode('ascii')
                    print(f"[DEBUG] Contenido de adjuntos obtenido (primeros 500 chars): {safe_attachments_preview}...", file=sys.stderr, flush=True)
                except Exception:
                    print(f"[DEBUG] Contenido de adjuntos obtenido (longitud: {len(attachments_content)} caracteres)", file=sys.stderr, flush=True)
            else:
                print(f"[DEBUG] No se encontró contenido de adjuntos o el email no tiene adjuntos", file=sys.stderr, flush=True)
        else:
            print(f"[DEBUG] No se pudo obtener token para buscar adjuntos", file=sys.stderr, flush=True)
    except Exception as e:
        try:
            safe_error = str(e).encode('ascii', 'ignore').decode('ascii')
            print(f"[DEBUG] Error al obtener adjuntos: {safe_error}", file=sys.stderr, flush=True)
        except Exception:
            print(f"[DEBUG] Error al obtener adjuntos", file=sys.stderr, flush=True)
    
    # Combinar asunto, cuerpo y adjuntos para la clasificación
    classification_text = f"Asunto: {original_subject}\n\n"
    if email_body:
        classification_text += f"Cuerpo del email:\n{email_body}\n\n"
    if attachments_content:
        classification_text += f"Contenido de adjuntos:\n{attachments_content}"
    
    # Si no hay texto para clasificar, usar el texto original
    if not classification_text.strip() or classification_text.strip() == f"Asunto: {original_subject}\n\n":
        classification_text = text or original_subject

    import sys
    # Mostrar el texto completo que se enviará a la IA
    try:
        safe_text_full = classification_text.encode('ascii', 'ignore').decode('ascii')
        print(f"[CLASSIFY ENDPOINT] TEXTO COMPLETO que se enviará a la IA (longitud: {len(classification_text)} caracteres):", file=sys.stderr, flush=True)
        print(f"[CLASSIFY ENDPOINT] ========== INICIO TEXTO COMPLETO ==========", file=sys.stderr, flush=True)
        print(safe_text_full, file=sys.stderr, flush=True)
        print(f"[CLASSIFY ENDPOINT] ========== FIN TEXTO COMPLETO ==========", file=sys.stderr, flush=True)
    except Exception:
        try:
            safe_text_preview = classification_text[:500].encode('ascii', 'ignore').decode('ascii')
            print(f"[CLASSIFY ENDPOINT] Llamando a classify_email con texto (primeros 500 chars): {safe_text_preview}...", file=sys.stderr, flush=True)
        except Exception:
            print(f"[CLASSIFY ENDPOINT] Llamando a classify_email (longitud: {len(classification_text)} caracteres)", file=sys.stderr, flush=True)
    
    label = classify_email(classification_text)
    
    try:
        safe_id = email_id.encode('ascii', 'ignore').decode('ascii')[:50]
        print(f"[CLASSIFY ENDPOINT] Clasificacion resultante - ID: {safe_id}... | Etiqueta: {label}", file=sys.stderr, flush=True)
    except Exception as e:
        print(f"[CLASSIFY ENDPOINT] Error al loguear resultado: {e}", file=sys.stderr, flush=True)

    # Si no se clasificó en ninguna etiqueta, marcar como procesado y devolver
    if label == "Sin etiqueta":
        try:
            print(f"[DEBUG] Email sin etiqueta - ID: {email_id[:50]}... | Marcando como procesado")
        except Exception:
            pass
        mark_processed(account_email, email_id, label)
        return jsonify({"label": "Sin etiqueta", "skipped": True})

    # Obtener la categoría completa para verificar si debe responder
    category = get_category_by_name(label)
    
    # Aplicar etiquetado según proveedor - solo soportamos outlook_oauth
    provider = (account.get("provider") or "outlook_oauth").lower()
    if provider != "outlook_oauth":
        print(f"[ERROR] Proveedor no soportado: {provider}. Solo se soporta outlook_oauth.")
        provider = "outlook_oauth"  # Forzar outlook_oauth como fallback
    
    try:
        print(f"[DEBUG] Aplicando etiqueta - ID: {email_id[:50]}... | Etiqueta: {label} | Proveedor: {provider}")
    except Exception:
        pass
    
    # Solo usar apply_label_graph ya que solo soportamos outlook_oauth
    label_success = apply_label_graph(email_id, label, account["email"])
    
    try:
        print(f"[DEBUG] Resultado guardado en carpeta - ID: {email_id[:50]}... | Exito: {label_success}")
    except Exception:
        pass

    # Solo enviar respuesta automática si se guardó correctamente en la carpeta
    # El envío de respuesta no debe bloquear el registro del procesado
    reply_sent = False
    if label_success and category and category.get('autoReply') and category.get('replyText'):
        reply_text = category.get('replyText', '')
        original_subject = data.get("originalSubject", "Sin asunto")
        
        try:
            safe_from = from_addr.encode('ascii', 'ignore').decode('ascii')[:50]
            safe_msg = f"[DEBUG] Enviando respuesta automatica - ID: {email_id[:50]}... | A: {safe_from}..."
            print(safe_msg)
        except Exception:
            try:
                print(f"[DEBUG] Enviando respuesta automatica - ID: {email_id[:50]}...")
            except Exception:
                pass
        try:
            # Solo usar send_reply_graph ya que solo soportamos outlook_oauth
            reply_sent = send_reply_graph(email_id, from_addr, original_subject, reply_text, account["email"])
            try:
                print(f"[DEBUG] Respuesta enviada - ID: {email_id[:50]}... | Exito: {reply_sent}")
            except Exception:
                pass
        except Exception as e:
            # El error en el envío no debe impedir marcar como procesado
            try:
                safe_error = str(e).encode('ascii', 'ignore').decode('ascii')
                print(f"Error al enviar respuesta (no bloquea el proceso): {safe_error}")
            except Exception:
                pass

    # Registrar procesado SIEMPRE (independiente de si se guardó en carpeta o no)
    # Esto evita reprocesar emails que ya fueron analizados
    try:
        print(f"[DEBUG] Marcando como procesado - ID: {email_id[:50]}... | Etiqueta: {label} | Guardado en carpeta: {label_success}")
    except Exception:
        pass
    mark_processed(account_email, email_id, label)

    return jsonify({"label": label})


@app.route('/microsoft/callback')
def microsoft_callback():
    auth_code = request.args.get('code')
    state_email = request.args.get('state')
    if not auth_code or not state_email:
        return "Faltan parámetros en callback.", 400
    if not (MS_CLIENT_ID and MS_CLIENT_SECRET):
        return "Falta configuración OAuth de Microsoft.", 500
    cache = msal.SerializableTokenCache()
    app_msal = _ms_app(cache)
    try:
        result = app_msal.acquire_token_by_authorization_code(
            auth_code,
            scopes=MS_SCOPES,
            redirect_uri=_effective_redirect_uri(),
        )
    except Exception as e:
        return f"Error al canjear código: {e}", 500
    if not result or 'access_token' not in result:
        return f"No se obtuvo access_token: {result}", 500
    serialized = cache.serialize()
    access_token = result.get('access_token')
    
    # Actualizar o crear cuenta
    idx = next((i for i, a in enumerate(ACCOUNTS) if a.get('email') == state_email), None)
    acct_data = {"email": state_email, "provider": "outlook_oauth", "ms_cache": serialized}
    if idx is not None:
        # Preservar otros campos útiles, pero quitar password si hay
        existing = ACCOUNTS[idx]
        existing.pop('password', None)
        ACCOUNTS[idx] = {**existing, **acct_data}
    else:
        ACCOUNTS.append(acct_data)
    save_json(ACCOUNTS_FILE, ACCOUNTS)
    
    # Configurar suscripción de Microsoft Graph para recibir notificaciones automáticas
    if access_token:
        subscription_success = setup_ms_graph_subscription(state_email, access_token)
        if subscription_success:
            print(f"[CALLBACK] Suscripcin de Microsoft Graph configurada para {state_email}")
        else:
            print(f"[CALLBACK] Advertencia: No se pudo configurar suscripción para {state_email}, pero la cuenta se vinculó correctamente")
    
    return "<script>window.location.href='/'</script>"


@app.route('/microsoft/renew-subscriptions', methods=['POST'])
def renew_subscriptions_endpoint():
    """
    Endpoint para renovar manualmente las suscripciones de Microsoft Graph.
    Útil para testing o si no tienes un scheduler configurado.
    """
    try:
        renewed, failed = renew_ms_graph_subscriptions()
        return jsonify({
            "success": True,
            "renewed": renewed,
            "failed": failed,
            "message": f"Renovación completada: {renewed} renovadas, {failed} fallidas"
        })
    except Exception as e:
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


@app.route('/webhook/microsoft-graph', methods=['POST', 'GET'])
def microsoft_graph_webhook():
    """
    Endpoint para recibir notificaciones de Microsoft Graph cuando llegan nuevos emails.
    Microsoft Graph envía notificaciones cuando hay cambios en los emails.
    Implementado siguiendo el mismo patrón que el software de referencia.
    """
    # Microsoft Graph valida el webhook con un GET que incluye un query parameter 'validationToken'
    validation_token = request.args.get('validationToken')
    if validation_token:
        try:
            safe_msg = f"[WEBHOOK] Recibida solicitud de validacion de suscripcion. Token: {validation_token}"
            print(safe_msg.encode('ascii', 'ignore').decode('ascii'))
        except Exception:
            print("[WEBHOOK] Recibida solicitud de validacion de suscripcion")
        # Microsoft Graph espera que devolvamos el validationToken como texto plano
        return validation_token, 200, {'Content-Type': 'text/plain'}
    
    # POST: Notificación de cambio
    try:
        notification_data = request.get_json()
        if notification_data is None:
            print("[WEBHOOK] Payload de notificación no es JSON válido o está vacío.")
            return jsonify({"error": "Payload de notificación inválido, se esperaba JSON."}), 400
        
        try:
            safe_notif = str(notification_data)[:500].encode('ascii', 'ignore').decode('ascii')
            print(f"[WEBHOOK] Notificacion de MS Graph recibida: {safe_notif}")
        except Exception:
            print("[WEBHOOK] Notificacion de MS Graph recibida")
        
        if not isinstance(notification_data.get("value"), list):
            print("[WEBHOOK] Payload de notificacion invalido o falta el array 'value'.")
            return jsonify({"error": "Payload de notificación inválido, falta 'value'."}), 400
        
        for notification_item in notification_data["value"]:
            if not isinstance(notification_item, dict):
                print("[WEBHOOK] Item de notificacion no es un diccionario, saltando.")
                continue
            
            subscription_id = notification_item.get("subscriptionId")
            resource_data = notification_item.get("resourceData", {})
            change_type = notification_item.get("changeType")
            resource_id = resource_data.get("id")
            
            if change_type == "created" and resource_id and subscription_id:
                try:
                    safe_id = str(resource_id)[:50].encode('ascii', 'ignore').decode('ascii')
                    safe_sub = str(subscription_id)[:50].encode('ascii', 'ignore').decode('ascii')
                    print(f"[WEBHOOK] Nuevo email creado (ID: {safe_id}...) para suscripcion {safe_sub}...")
                except Exception:
                    print("[WEBHOOK] Nuevo email creado")
                
                # Buscar la cuenta asociada a esta suscripción
                # Primero por subscription_id, luego por cualquier cuenta con suscripción activa
                global ACCOUNTS
                account_email = None
                for acct in ACCOUNTS:
                    if acct.get("ms_subscription_id") == subscription_id:
                        account_email = acct.get("email")
                        break
                
                # Si no se encontró por subscription_id exacto, buscar cualquier cuenta con suscripción
                # (puede haber múltiples suscripciones o la suscripción se renovó)
                if not account_email:
                    # Recargar ACCOUNTS para asegurar datos actualizados
                    ACCOUNTS = load_json(ACCOUNTS_FILE, [])
                    # Buscar cualquier cuenta outlook_oauth (puede tener suscripción activa aunque el ID cambió)
                    for acct in ACCOUNTS:
                        if acct.get("provider") == "outlook_oauth":
                            account_email = acct.get("email")
                            # Actualizar el subscription_id en la cuenta para futuras notificaciones
                            acct["ms_subscription_id"] = subscription_id
                            # Actualizar también la fecha de expiración si viene en la notificación
                            expiration_str = notification_item.get("subscriptionExpirationDateTime")
                            if expiration_str:
                                try:
                                    if expiration_str.endswith('Z'):
                                        parsed_expiry = datetime.fromisoformat(expiration_str[:-1] + '+00:00')
                                    else:
                                        parsed_expiry = datetime.fromisoformat(expiration_str)
                                    if parsed_expiry.tzinfo is None:
                                        parsed_expiry = parsed_expiry.replace(tzinfo=timezone.utc)
                                    else:
                                        parsed_expiry = parsed_expiry.astimezone(timezone.utc)
                                    acct["ms_subscription_expires_at"] = parsed_expiry.isoformat()
                                except (ValueError, TypeError):
                                    pass
                            save_json(ACCOUNTS_FILE, ACCOUNTS)
                            try:
                                safe_email = str(account_email).encode('ascii', 'ignore').decode('ascii')
                                safe_sub = str(subscription_id)[:30].encode('ascii', 'ignore').decode('ascii')
                                print(f"[WEBHOOK] Encontrada cuenta {safe_email} y actualizado subscription_id: {safe_sub}...")
                            except Exception:
                                print("[WEBHOOK] Encontrada cuenta y actualizado subscription_id")
                            break
                
                if account_email:
                    # Obtener información del email para logging antes de procesar
                    try:
                        token = get_graph_token_for_account(account_email)
                        if token:
                            email_url = f"https://graph.microsoft.com/v1.0/me/messages/{resource_id}?$select=id,subject,from,body,receivedDateTime"
                            headers = {"Authorization": f"Bearer {token}", "Prefer": 'outlook.body-content-type="text"'}
                            r = requests.get(email_url, headers=headers, timeout=10)
                            if r.status_code < 300:
                                email_data = r.json()
                                subj = email_data.get("subject") or "Sin asunto"
                                from_obj = email_data.get("from", {}).get("emailAddress", {})
                                from_addr = from_obj.get('address', 'Desconocido')
                                body_preview = ""
                                body_obj = email_data.get("body", {})
                                if body_obj.get("content"):
                                    body_preview = str(body_obj.get("content", ""))[:200]
                                
                                try:
                                    safe_subj = str(subj).encode('ascii', 'ignore').decode('ascii')
                                    safe_from = str(from_addr).encode('ascii', 'ignore').decode('ascii')
                                    safe_body = str(body_preview).encode('ascii', 'ignore').decode('ascii')
                                    safe_email = str(account_email).encode('ascii', 'ignore').decode('ascii')
                                    safe_id = str(resource_id)[:50].encode('ascii', 'ignore').decode('ascii')
                                    print(f"[WEBHOOK] ===== EMAIL RECIBIDO =====")
                                    print(f"[WEBHOOK] Cuenta: {safe_email}")
                                    print(f"[WEBHOOK] Email ID: {safe_id}...")
                                    print(f"[WEBHOOK] De: {safe_from}")
                                    print(f"[WEBHOOK] Asunto: {safe_subj}")
                                    print(f"[WEBHOOK] Cuerpo (primeros 200 chars): {safe_body}...")
                                    print(f"[WEBHOOK] ==========================")
                                except Exception:
                                    print("[WEBHOOK] Email recibido (error al formatear log)")
                    except Exception as e:
                        try:
                            safe_error = str(e).encode('ascii', 'ignore').decode('ascii')
                            print(f"[WEBHOOK] Error al obtener info del email para log: {safe_error}")
                        except Exception:
                            print("[WEBHOOK] Error al obtener info del email para log")
                    
                    # Procesar en un hilo separado para no bloquear la respuesta
                    thread = threading.Thread(
                        target=process_new_email_from_notification,
                        args=(account_email, resource_id)
                    )
                    thread.daemon = True
                    thread.start()
                    try:
                        safe_email = str(account_email).encode('ascii', 'ignore').decode('ascii')
                        safe_id = str(resource_id)[:50].encode('ascii', 'ignore').decode('ascii')
                        print(f"[WEBHOOK] Procesando nuevo email en background - Cuenta: {safe_email}, ID: {safe_id}...")
                    except Exception:
                        print("[WEBHOOK] Procesando nuevo email en background")
                else:
                    try:
                        safe_sub = str(subscription_id).encode('ascii', 'ignore').decode('ascii')
                        print(f"[WEBHOOK] No se encontro cuenta para subscription_id: {safe_sub}")
                    except Exception:
                        print("[WEBHOOK] No se encontro cuenta para subscription_id")
            else:
                try:
                    safe_change = str(change_type).encode('ascii', 'ignore').decode('ascii')
                    safe_rid = str(resource_id).encode('ascii', 'ignore').decode('ascii') if resource_id else "None"
                    safe_sid = str(subscription_id).encode('ascii', 'ignore').decode('ascii') if subscription_id else "None"
                    print(f"[WEBHOOK] Notificacion de tipo '{safe_change}' (recurso ID: {safe_rid}, suscripcion: {safe_sid}) no procesada")
                except Exception:
                    print("[WEBHOOK] Notificacion no procesada")
        
        return jsonify({"status": "Notification received"}), 202
        
    except Exception as e:
        print(f"[WEBHOOK] Error inesperado: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": "Error interno del servidor"}), 500


if __name__ == '__main__':
    app.run(debug=True)

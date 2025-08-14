import imaplib
import email
from email.header import decode_header
from datetime import datetime
import os
import re
import shutil
import hashlib

# === CONFIGURACIÓN GENERAL ===
USUARIO = "partes@policiafederal.gov.ar"
CONTRASEÑA = "Novedades03"
SERVIDOR_IMAP = "10.1.150.25"

CARPETA_NEXTCLOUD = os.path.expanduser("~/nextcloud/PARTES")
CARPETA_TEMP = os.path.expanduser("~/descarga_temp")
os.makedirs(CARPETA_TEMP, exist_ok=True)

# Extensiones que no se deben descargar (imágenes y videos)
EXTENSIONES_IGNORADAS = ['.jpg', '.jpeg', '.png', '.mp4', '.avi', '.mov', '.mkv']

# === FUNCIONES AUXILIARES ===

def limpiar_nombre_archivo(nombre):
    return re.sub(r'[\\/*?:"<>|\r\n]+', '', nombre).strip()

def calcular_hash(contenido):
    return hashlib.md5(contenido).hexdigest()

def es_pdf(contenido):
    """Verifica si el contenido corresponde a un PDF mirando su firma inicial."""
    return contenido.startswith(b'%PDF')

def descargar_adjuntos(turno):
    print(f"🕒 Iniciando descarga del {turno}...")

    conn = imaplib.IMAP4_SSL(SERVIDOR_IMAP)
    conn.login(USUARIO, CONTRASEÑA)
    conn.select("INBOX")

    estado, mensajes = conn.search(None, '(UNSEEN)')
    if estado != "OK":
        print("❌ No se pudieron buscar mensajes.")
        conn.logout()
        return

    ids = mensajes[0].split()
    print(f"📨 Correos no leídos: {len(ids)}")

    archivos_guardados = []
    hashes_vistos = set()

    for num in ids:
        estado, datos = conn.fetch(num, "(RFC822)")
        if estado != "OK":
            continue

        raw_email = datos[0][1]
        msg = email.message_from_bytes(raw_email)

        asunto = decode_header(msg["Subject"])[0][0]
        if isinstance(asunto, bytes):
            asunto = asunto.decode(errors="ignore")

        print(f"✉ Procesando: {asunto}")

        for parte in msg.walk():
            if parte.get_content_maintype() == "multipart":
                continue
            if parte.get("Content-Disposition") is None:
                continue

            nombre_archivo = parte.get_filename()
            if nombre_archivo:
                nombre_archivo = decode_header(nombre_archivo)[0][0]
                if isinstance(nombre_archivo, bytes):
                    nombre_archivo = nombre_archivo.decode(errors="ignore")
                nombre_archivo = limpiar_nombre_archivo(nombre_archivo)

                _, extension = os.path.splitext(nombre_archivo)
                if extension.lower() in EXTENSIONES_IGNORADAS:
                    print(f"🚫 Archivo ignorado por extensión: {nombre_archivo}")
                    continue

                contenido = parte.get_payload(decode=True)

                # Si el archivo es PDF pero no tiene extensión correcta, renombrarlo
                if es_pdf(contenido) and not nombre_archivo.lower().endswith('.pdf'):
                    print(f"📄 Detectado PDF sin extensión correcta: {nombre_archivo}")
                    nombre_archivo = os.path.splitext(nombre_archivo)[0] + '.pdf'

                hash_actual = calcular_hash(contenido)
                if hash_actual in hashes_vistos:
                    print(f"⚠️ Duplicado omitido (hash coincidente): {nombre_archivo}")
                    continue
                hashes_vistos.add(hash_actual)

                ruta_archivo = os.path.join(CARPETA_NEXTCLOUD, nombre_archivo)
                base, extension = os.path.splitext(ruta_archivo)
                contador = 1
                while os.path.exists(ruta_archivo):
                    ruta_archivo = f"{base}_{contador}{extension}"
                    contador += 1

                with open(ruta_archivo, "wb") as f:
                    f.write(contenido)
                archivos_guardados.append(os.path.basename(ruta_archivo))

    conn.logout()

    if archivos_guardados:
        ahora = datetime.now().strftime("%Y-%m-%d_%H-%M")
        # Crear log
        nombre_log = f"log_{ahora}.txt"
        ruta_log = os.path.join(CARPETA_NEXTCLOUD, nombre_log)
        with open(ruta_log, "w") as f:
            for nombre in archivos_guardados:
                f.write(f"{nombre}\n")

        print(f"✅ Archivos guardados directamente en la nube.")
        print(f"📝 Log creado: {ruta_log}")
    else:
        print("⚠️ No se encontraron adjuntos para guardar.")

    # Limpiar carpeta temporal (por si queda algo)
    shutil.rmtree(CARPETA_TEMP)
    os.makedirs(CARPETA_TEMP, exist_ok=True)

# === LLAMADA PARA CRONTAB O EJECUCIÓN MANUAL ===
if __name__ == "__main__":
    descargar_adjuntos("turno mañana")

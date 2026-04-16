"""
Humand — Automatización Llamado de Atención
Corre automáticamente via GitHub Actions cada 30 minutos.
"""

import os
import json
import shutil
import zipfile
import re
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path

import requests

# --- Configuración (se lee desde variables de entorno en GitHub Actions) ---
API_BASE_URL    = "https://api-prod.humand.co/public/api/v1"
API_KEY         = os.environ.get("HUMAND_API_KEY", "")
SERVICE_ITEM_ID = os.environ.get("HUMAND_SERVICE_ITEM_ID", "01KNS891HBAP4EVH6S0ESEWPP4")
FOLDER_ID       = int(os.environ.get("HUMAND_FOLDER_ID", "412799"))
CREATED_FROM    = os.environ.get("HUMAND_CREATED_FROM", "2025-04-16T00:00:00.000Z")

TEMPLATE_PATH  = "template_llamado_atencion_v3.docx"
PROCESSED_FILE = "processed_tasks.json"

HEADERS = {
    "Authorization": f"Basic {API_KEY}",
    "Accept": "application/json"
}

FIELD_TITLES = {
    "¿Qué colaborador está siendo sancionado/apercibido o se le está llamando la atención? (Ingresar DNI)": "dni",
    "¿Querés dejar algún otro comentario?": "comentario",
    "Fecha de la notificación": "fecha",
    "Descripción / Motivo": "descripcion",
    "Nombre del responsable": "nombre_responsable"
}


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def load_processed_tasks():
    if os.path.exists(PROCESSED_FILE):
        with open(PROCESSED_FILE, "r") as f:
            return set(json.load(f))
    return set()


def save_processed_tasks(processed: set):
    with open(PROCESSED_FILE, "w") as f:
        json.dump(list(processed), f, indent=2)


def extract_form_fields(task: dict) -> dict:
    fields = {}
    sections = task.get("formAnswer", {}).get("sections", [])
    for section in sections:
        for answer in section.get("answers", []):
            title = answer.get("title", "")
            raw_answer = answer.get("answer")
            if title in FIELD_TITLES:
                key = FIELD_TITLES[title]
                if isinstance(raw_answer, dict):
                    fields[key] = raw_answer.get("fieldValue", "")
                elif isinstance(raw_answer, list):
                    fields[key] = ", ".join(raw_answer)
                else:
                    fields[key] = str(raw_answer) if raw_answer else ""
    return fields


# ============================================================
# API DE HUMAND
# ============================================================

def get_all_tasks() -> list:
    all_tasks = []
    page = 1
    while True:
        url = f"{API_BASE_URL}/service-management/service-items/{SERVICE_ITEM_ID}/tasks"
        params = {"page": page, "limit": 50, "createdAtFrom": CREATED_FROM}
        response = requests.get(url, headers=HEADERS, params=params)
        response.raise_for_status()
        data = response.json()
        items = data.get("items", [])
        all_tasks.extend(items)
        total = data.get("count", 0)
        print(f"  Página {page}: {len(items)} tasks (acumulado: {len(all_tasks)}/{total})")
        if len(all_tasks) >= total or len(items) == 0:
            break
        page += 1
    return all_tasks


def get_user_by_employee_id(employee_internal_id: str):
    url = f"{API_BASE_URL}/users"
    params = {"search": employee_internal_id, "page": 1, "limit": 10}
    response = requests.get(url, headers=HEADERS, params=params)
    response.raise_for_status()
    users = response.json().get("users", [])
    for user in users:
        if user.get("employeeInternalId") == employee_internal_id:
            return user
    return users[0] if users else None


def upload_document(employee_internal_id: str, pdf_path: str, document_name: str) -> dict:
    url = f"{API_BASE_URL}/users/{employee_internal_id}/documents/files"
    with open(pdf_path, "rb") as pdf_file:
        files = {"file": (os.path.basename(pdf_path), pdf_file, "application/pdf")}
        data = {
            "folderId": FOLDER_ID,
            "name": document_name,
            "sendNotification": "true",
            "signatureStatus": "PENDING",
            "allowDisagreement": "false"
        }
        response = requests.post(
            url,
            headers={"Authorization": f"Basic {API_KEY}"},
            files=files,
            data=data
        )
    response.raise_for_status()
    return response.json() if response.text else {"status": "ok"}


# ============================================================
# GENERACIÓN DE DOCUMENTOS
# ============================================================

def fill_template(template_path: str, output_docx_path: str, replacements: dict):
    shutil.copy2(template_path, output_docx_path)
    with zipfile.ZipFile(output_docx_path, "r") as zin:
        doc_xml = zin.read("word/document.xml").decode("utf-8")
        other_files = {item: zin.read(item) for item in zin.namelist() if item != "word/document.xml"}

    for placeholder, value in replacements.items():
        value_escaped = (
            value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                 .replace('"', "&quot;").replace("'", "&apos;")
        )
        pattern_direct = f"&lt;&lt;{placeholder}&gt;&gt;"
        if pattern_direct in doc_xml:
            doc_xml = doc_xml.replace(pattern_direct, value_escaped)
            continue
        xml_between_runs = r'</w:t></w:r><w:r[^>]*>(?:<w:rPr>.*?</w:rPr>)?<w:t[^>]*>'
        pattern_split = r'&lt;&lt;' + xml_between_runs + re.escape(placeholder) + r'&gt;&gt;'
        match = re.search(pattern_split, doc_xml, re.DOTALL)
        if match:
            doc_xml = doc_xml[:match.start()] + value_escaped + doc_xml[match.end():]

    with zipfile.ZipFile(output_docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item_name, item_data in other_files.items():
            zout.writestr(item_name, item_data)
        zout.writestr("word/document.xml", doc_xml.encode("utf-8"))


def convert_docx_to_pdf(docx_path: str, output_dir: str) -> str:
    result = subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", output_dir, docx_path],
        capture_output=True, text=True, timeout=120
    )
    if result.returncode != 0:
        raise RuntimeError(f"Error al convertir a PDF: {result.stderr}")
    pdf_path = os.path.join(output_dir, Path(docx_path).stem + ".pdf")
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"No se generó el PDF en: {pdf_path}")
    return pdf_path


# ============================================================
# PROCESO PRINCIPAL
# ============================================================

def process_task(task: dict, processed: set) -> bool:
    task_id = task["id"]
    task_number = task.get("taskNumber", "?")

    if task_id in processed:
        print(f"  ⏭️  Task #{task_number} ya procesado. Saltando.")
        return False

    print(f"\n  📋 Procesando Task #{task_number} ({task_id})...")

    fields = extract_form_fields(task)
    if not fields:
        print(f"  ⚠️  Sin campos. Saltando.")
        return False

    employee_id = fields.get("dni", "")
    if not employee_id:
        print(f"  ⚠️  Sin DNI. Saltando.")
        return False

    print(f"    → DNI: {employee_id}")

    user = get_user_by_employee_id(employee_id)
    if not user:
        print(f"  ⚠️  Usuario con DNI '{employee_id}' no encontrado. Saltando.")
        return False

    user_full_name = f"{user.get('firstName', '')} {user.get('lastName', '')}".strip()
    user_puesto = user.get("jobTitle", "") or user.get("position", "") or ""
    user_area = user.get("department", "") or user.get("area", "") or ""

    print(f"    → Usuario: {user_full_name} | {user_puesto} | {user_area}")

    replacements = {
        "Nombre": user_full_name,
        "DNI": employee_id,
        "Puesto": user_puesto,
        "Area": user_area,
        "Fecha": fields.get("fecha", ""),
        "NombreResponsable": fields.get("nombre_responsable", ""),
        "Descripcion": fields.get("descripcion", "")
    }

    with tempfile.TemporaryDirectory() as tmp_dir:
        output_docx = os.path.join(tmp_dir, f"LlamadoAtencion_{employee_id}.docx")
        fill_template(TEMPLATE_PATH, output_docx, replacements)

        pdf_path = convert_docx_to_pdf(output_docx, tmp_dir)
        print(f"    → PDF generado: {pdf_path}")

        document_name = f"Llamado de Atención - {user_full_name} - {fields.get('fecha', '')}"
        result = upload_document(employee_id, pdf_path, document_name)
        print(f"    ✅ Subido: {document_name}")

    processed.add(task_id)
    save_processed_tasks(processed)
    return True


def main():
    print("=" * 60)
    print("  LLAMADO DE ATENCIÓN")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    if not API_KEY:
        print("❌ HUMAND_API_KEY no configurada.")
        return

    if not os.path.exists(TEMPLATE_PATH):
        print(f"❌ Template no encontrado: {TEMPLATE_PATH}")
        return

    processed = load_processed_tasks()
    print(f"\n📁 Tasks previamente procesados: {len(processed)}")

    print("\n🔍 Obteniendo tasks...")
    tasks = get_all_tasks()

    pending = [t for t in tasks if t["id"] not in processed]
    print(f"📋 Pendientes: {len(pending)} / {len(tasks)}")

    if not pending:
        print("\n✅ No hay tasks nuevos.")
        return

    success = error = 0
    for task in pending:
        try:
            if process_task(task, processed):
                success += 1
        except Exception as e:
            print(f"  ❌ Error en task #{task.get('taskNumber', '?')}: {e}")
            error += 1

    print("\n" + "=" * 60)
    print(f"  ✅ Procesados: {success} | ❌ Errores: {error}")
    print("=" * 60)


if __name__ == "__main__":
    main()

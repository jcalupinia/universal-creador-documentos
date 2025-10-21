from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse, FileResponse
from pydantic import BaseModel
import os
import uuid
import pandas as pd
from fpdf import FPDF
from pptx import Presentation
from docx import Document
from openpyxl import Workbook

# --- Inicialización ---
app = FastAPI(
    title="Universal Artifact Generator",
    description="API para generar documentos Excel, Word, PDF, PPT y más, con soporte para Render + UptimeRobot.",
    version="1.0.0"
)

# --- Directorio de resultados ---
RESULTS_DIR = "resultados"
os.makedirs(RESULTS_DIR, exist_ok=True)

# --- MODELO GENERAL ---
class ArtifactResponse(BaseModel):
    url: str

# ===========================
# 🔹 HEALTHCHECK para Render / UptimeRobot
# ===========================
@app.get(os.getenv("HEALTHCHECK_PATH", "/healthz"))
async def health_check():
    """
    Endpoint de salud usado por Render y UptimeRobot para mantener el servicio activo.
    """
    return {"status": "ok"}

# ===========================
# 🔹 GENERAR EXCEL
# ===========================
@app.post("/generate_excel", response_model=ArtifactResponse)
async def generate_excel(payload: dict):
    try:
        titulo = payload.get("titulo", "archivo")
        data = payload["data"]
        headers = data["headers"]
        rows = data["rows"]

        wb = Workbook()
        ws = wb.active
        ws.title = "Detalle"

        ws.append(headers)
        for row in rows:
            ws.append(row)

        file_id = str(uuid.uuid4())
        filename = f"{titulo}_{file_id}.xlsx"
        filepath = os.path.join(RESULTS_DIR, filename)
        wb.save(filepath)

        return {"url": f"/{RESULTS_DIR}/{filename}"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# ===========================
# 🔹 GENERAR WORD
# ===========================
@app.post("/generate_word", response_model=ArtifactResponse)
async def generate_word(payload: dict):
    try:
        doc = Document()
        doc.add_heading(payload.get("placeholders", {}).get("titulo", "Documento"), level=1)
        doc.add_paragraph(payload.get("placeholders", {}).get("subtitulo", "Generado automáticamente"))

        file_id = str(uuid.uuid4())
        filename = f"word_{file_id}.docx"
        filepath = os.path.join(RESULTS_DIR, filename)
        doc.save(filepath)

        return {"url": f"/{RESULTS_DIR}/{filename}"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# ===========================
# 🔹 GENERAR POWERPOINT
# ===========================
@app.post("/generate_ppt", response_model=ArtifactResponse)
async def generate_ppt(payload: dict):
    try:
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        title = slide.shapes.title
        subtitle = slide.placeholders[1]

        title.text = payload.get("title", "Presentación")
        subtitle.text = payload.get("subtitle", "Generada automáticamente")

        file_id = str(uuid.uuid4())
        filename = f"ppt_{file_id}.pptx"
        filepath = os.path.join(RESULTS_DIR, filename)
        prs.save(filepath)

        return {"url": f"/{RESULTS_DIR}/{filename}"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# ===========================
# 🔹 GENERAR PDF
# ===========================
@app.post("/generate_pdf", response_model=ArtifactResponse)
async def generate_pdf(payload: dict):
    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        title = payload.get("title", "Informe PDF")
        pdf.cell(200, 10, txt=title, ln=True, align="C")

        sections = payload.get("sections", [])
        for sec in sections:
            if sec.get("type") == "p":
                pdf.multi_cell(0, 10, sec.get("text", ""))
            elif sec.get("type") == "h1":
                pdf.set_font("Arial", style="B", size=14)
                pdf.cell(0, 10, sec.get("text", ""), ln=True)
                pdf.set_font("Arial", size=12)

        file_id = str(uuid.uuid4())
        filename = f"pdf_{file_id}.pdf"
        filepath = os.path.join(RESULTS_DIR, filename)
        pdf.output(filepath)

        return {"url": f"/{RESULTS_DIR}/{filename}"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# ===========================
# 🔹 DESCARGA DE ARCHIVOS
# ===========================
@app.get("/{folder}/{filename}")
async def download_file(folder: str, filename: str):
    filepath = os.path.join(folder, filename)
    if not os.path.exists(filepath):
        raise HTTPException(status_code=404, detail="Archivo no encontrado")
    return FileResponse(filepath)

# ===========================
# 🔹 LIMPIEZA AUTOMÁTICA DE RESULTADOS
# ===========================
@app.on_event("startup")
async def cleanup_results():
    """
    Elimina archivos antiguos (>1 hora) para mantener el entorno limpio.
    """
    import time
    max_age = 60 * 60  # 1 hora
    now = time.time()

    for f in os.listdir(RESULTS_DIR):
        fp = os.path.join(RESULTS_DIR, f)
        if os.path.isfile(fp) and now - os.path.getmtime(fp) > max_age:
            os.remove(fp)

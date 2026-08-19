import os
import pandas as pd
from fastapi import FastAPI, Header, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from docx import Document
from pptx import Presentation

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# مفتاح الوصول المعتمد
ACTIVE_SUBSCRIPTIONS = ["sulaiman_vip_2026"]

def verify_sub(key: str):
    if key not in ACTIVE_SUBSCRIPTIONS:
        raise HTTPException(status_code=401, detail="مفتاح الوصول غير صحيح - NOX Security")

@app.post("/generate-word/")
def generate_word(topic: str, x_api_key: str = Header(...)):
    verify_sub(x_api_key)
    doc = Document()
    doc.add_heading(f"NOX Report: {topic}", 0)
    doc.add_paragraph(f"تم توليد هذا التقرير بواسطة منظومة NOX الذكية لموضوع: {topic}.")
    file_name = f"NOX_Report_{os.urandom(3).hex()}.docx"
    doc.save(file_name)
    return FileResponse(file_name, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename=file_name)

@app.post("/generate-ppt/")
def generate_ppt(presentation_title: str, x_api_key: str = Header(...)):
    verify_sub(x_api_key)
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = presentation_title
    slide.placeholders[1].text = "Powered by NOX Engine"
    file_name = f"NOX_Presentation_{os.urandom(3).hex()}.pptx"
    prs.save(file_name)
    return FileResponse(file_name, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", filename=file_name)

@app.post("/generate-excel/")
def generate_excel(item_name: str, value: int = 100, x_api_key: str = Header(...)):
    verify_sub(x_api_key)
    df = pd.DataFrame({"العنصر": [item_name, "مؤشر NOX"], "القيمة": [value, value * 1.5]})
    file_name = f"NOX_Data_{os.urandom(3).hex()}.xlsx"
    df.to_excel(file_name, index=False)
    return FileResponse(file_name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=file_name)

@app.post("/analyze-report/")
def analyze_report(x_api_key: str = Header(...)):
    verify_sub(x_api_key)
    df = pd.DataFrame({"الشهر": ["يناير", "فبراير"], "المبيعات": [50000, 80000], "التكاليف": [30000, 40000]})
    df['الربح'] = df['المبيعات'] - df['التكاليف']
    doc = Document()
    doc.add_heading("NOX | التحليل المالي", 0)
    table = doc.add_table(rows=1, cols=4)
    for i, h in enumerate(["الشهر", "المبيعات", "التكاليف", "الربح"]):
        table.rows[0].cells[i].text = h
    for _, row in df.iterrows():
        cells = table.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = str(val)
    file_name = f"NOX_Financial_{os.urandom(2).hex()}.docx"
    doc.save(file_name)
    return FileResponse(file_name, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename=file_name)
from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
import google.generativeai as genai
import docx
from pptx import Presentation
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = FastAPI()

# HTML Frontend
html_content = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <title>NOX AI Studio | منظومة سليمان ماهر المتكاملة</title>
    <style>
        :root { --bg-main: #212121; --bg-input: #2f2f2f; --accent-color: #1a73e8; --text-main: #ececec; }
        body { font-family: sans-serif; background-color: var(--bg-main); color: var(--text-main); margin: 0; display: flex; height: 100vh; }
        .sidebar { width: 280px; background-color: #171717; padding: 15px; border-left: 1px solid #383838; }
        .main-content { flex-grow: 1; display: flex; flex-direction: column; }
        .chat-messages { flex-grow: 1; overflow-y: auto; padding: 20px; }
        .input-area { padding: 20px; background: #212121; }
        textarea { width: 100%; background: var(--bg-input); color: white; border: none; padding: 10px; border-radius: 8px; }
        .export-btn { background: #334155; color: white; border: none; padding: 5px 10px; margin: 5px; cursor: pointer; border-radius: 5px; }
    </style>
</head>
<body>
    <div class="sidebar">
        <h3>الإعدادات</h3>
        <input type="password" id="apiKey" placeholder="ادخل مفتاح API هنا..." style="width:100%">
    </div>
    <div class="main-content">
        <div class="chat-messages" id="chatMessages"></div>
        <div class="input-area">
            <textarea id="promptInput" rows="2" placeholder="اكتب طلبك هنا..."></textarea>
            <button onclick="sendMessage()" style="margin-top:10px">إرسال</button>
        </div>
    </div>
    <script>
        async function sendMessage() {
            const text = document.getElementById('promptInput').value;
            const apiKey = document.getElementById('apiKey').value;
            const chat = document.getElementById('chatMessages');
            chat.innerHTML += `<div><b>أنت:</b> ${text}</div>`;
            const response = await fetch('/chat', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ prompt: text, apiKey })
            });
            const data = await response.json();
            chat.innerHTML += `<div><b>AI:</b> ${data.reply}</div>`;
            chat.innerHTML += `
                <div>
                    <button class="export-btn" onclick="exportFile('${text}', 'word')">📥 Word</button>
                    <button class="export-btn" onclick="exportFile('${text}', 'excel')">📥 Excel</button>
                </div>`;
        }
        async function exportFile(topic, type) {
            const apiKey = document.getElementById('apiKey').value;
            const response = await fetch('/generate', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ topic, apiKey, serviceType: type })
            });
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = type === 'excel' ? 'Report.xlsx' : 'Report.docx';
            a.click();
        }
    </script>
</body>
</html>
"""


def get_best_model(api_key):
    genai.configure(api_key=api_key)
    for m in genai.list_models():
        if 'generateContent' in m.supported_generation_methods and 'flash' in m.name:
            return m.name
    return 'gemini-1.5-flash'


@app.get("/", response_class=HTMLResponse)
async def read_index(): return html_content


@app.post("/chat")
async def chat_endpoint(request: Request):
    data = await request.json()
    model = genai.GenerativeModel(get_best_model(data['apiKey']))
    return {"reply": model.generate_content(data['prompt']).text}


@app.post("/generate")
async def generate_endpoint(request: Request):
    data = await request.json()
    model = genai.GenerativeModel(get_best_model(data['apiKey']))
    content = model.generate_content(f"اكتب تقرير مفصل عن: {data['topic']}").text

    if data['serviceType'] == 'excel':
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["البند", "التفاصيل"])
        ws.append(["موضوع التقرير", data['topic']])
        path = "Report.xlsx"
        wb.save(path)
        return FileResponse(path, filename="Report.xlsx")
    else:
        doc = docx.Document()
        doc.add_paragraph(content)
        path = "Report.docx"
        doc.save(path)
        return FileResponse(path, filename="Report.docx")
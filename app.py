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
        :root {
            --bg-sidebar: #171717;
            --bg-main: #212121;
            --bg-input: #2f2f2f;
            --border-color: #383838;
            --text-main: #ececec;
            --text-muted: #9b9b9b;
            --accent-color: #1a73e8;
            --accent-hover: #1557b0;
        }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: var(--bg-main);
            color: var(--text-main);
            margin: 0;
            display: flex;
            height: 100vh;
            overflow: hidden;
        }
        .sidebar {
            width: 280px;
            background-color: var(--bg-sidebar);
            border-left: 1px solid var(--border-color);
            display: flex;
            flex-direction: column;
            padding: 15px;
            box-sizing: border-box;
        }
        .settings-group { margin-bottom: 15px; }
        .settings-group label { font-size: 12px; color: var(--text-muted); display: block; margin-bottom: 5px; }
        .settings-group input {
            width: 100%;
            background: var(--bg-input);
            border: 1px solid var(--border-color);
            color: #fff;
            padding: 8px;
            border-radius: 6px;
            font-size: 13px;
            box-sizing: border-box;
            outline: none;
        }
        .main-content {
            flex-grow: 1;
            display: flex;
            flex-direction: column;
            position: relative;
            height: 100vh;
        }
        .chat-messages {
            flex-grow: 1;
            overflow-y: auto;
            padding: 20px;
            display: flex;
            flex-direction: column;
            gap: 20px;
            max-width: 800px;
            width: 100%;
            margin: 0 auto;
            box-sizing: border-box;
        }
        .message {
            background: #262626;
            padding: 15px;
            border-radius: 12px;
            border: 1px solid var(--border-color);
            line-height: 1.6;
        }
        .input-area {
            padding: 20px;
            background: var(--bg-main);
            display: flex;
            justify-content: center;
        }
        .input-box-wrapper {
            background-color: var(--bg-input);
            border: 1px solid var(--border-color);
            border-radius: 16px;
            padding: 12px 15px;
            width: 100%;
            max-width: 800px;
            display: flex;
            flex-direction: column;
            gap: 10px;
        }
        textarea {
            width: 100%;
            background: transparent;
            border: none;
            color: var(--text-main);
            font-size: 15px;
            outline: none;
            resize: none;
            height: 45px;
            font-family: inherit;
        }
        .send-btn {
            background-color: var(--accent-color);
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 8px;
            cursor: pointer;
            font-weight: 600;
        }
        .export-buttons {
            display: flex;
            gap: 8px;
            margin-top: 10px;
        }
        .export-btn {
            background: #334155;
            color: #fff;
            border: none;
            padding: 6px 12px;
            border-radius: 6px;
            font-size: 12px;
            cursor: pointer;
        }
        .export-btn:hover { background: #475569; }
    </style>
</head>
<body>

    <div class="sidebar">
        <h3>الإعدادات</h3>
        <div class="settings-group">
            <label>مفتاح Gemini API Key</label>
            <input type="password" id="apiKey" placeholder="AIzaSy...">
        </div>
        <div style="margin-top: auto; font-size: 13px; color: var(--text-muted);">
            <div><b>سليمان ماهر</b></div>
            <div>منظومة العمل الذكية</div>
        </div>
    </div>

    <div class="main-content">
        <div class="chat-messages" id="chatMessages">
            <div class="message">
                <strong>أهلاً بك يا سليمان.</strong>
                <p style="color: var(--text-muted); margin: 5px 0 0 0;">قم بإدخال مفتاح الـ API بالأعلى وابدأ كتابة طلبك لتحليل أو إنشاء الملفات مباشرة.</p>
            </div>
        </div>

        <div class="input-area">
            <div class="input-box-wrapper">
                <textarea id="promptInput" placeholder="اكتب طلبك هنا (مثال: ابي جدول احترافي اكسل لربع الاول للسنة)..." onkeydown="handleKey(event)"></textarea>
                <div style="display: flex; justify-content: flex-end;">
                    <button class="send-btn" onclick="sendMessage()">إرسال ➔</button>
                </div>
            </div>
        </div>
    </div>

    <script>
        function handleKey(e) {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                sendMessage();
            }
        }

        async function sendMessage() {
            const text = document.getElementById('promptInput').value.trim();
            const apiKey = document.getElementById('apiKey').value;
            const chat = document.getElementById('chatMessages');
            if (!text) return;
            if (!apiKey) { alert('الرجاء إدخال مفتاح API أولاً'); return; }

            chat.innerHTML += `<div class="message"><b>أنت:</b> ${text}</div>`;
            document.getElementById('promptInput').value = '';

            const res = await fetch('/chat', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ prompt: text, apiKey })
            });
            const data = await res.json();

            chat.innerHTML += `
                <div class="message">
                    <b>AI:</b>
                    <div style="white-space: pre-wrap; margin-top: 8px;">${data.reply}</div>
                    <div class="export-buttons">
                        <button class="export-btn" onclick="exportFile('${encodeURIComponent(text)}', 'excel')">📥 تصدير Excel (Q1)</button>
                        <button class="export-btn" onclick="exportFile('${encodeURIComponent(text)}', 'word')">📥 تصدير Word</button>
                    </div>
                </div>
            `;
            chat.scrollTop = chat.scrollHeight;
        }

        async function exportFile(topic, type) {
            const apiKey = document.getElementById('apiKey').value;
            alert('جاري تجهيز وتنزيل الملف...');
            const res = await fetch('/generate', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ topic: decodeURIComponent(topic), apiKey, serviceType: type })
            });
            if (res.ok) {
                const blob = await res.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = type === 'excel' ? 'Q1_Professional_Sheet.xlsx' : 'Gemini_Report.docx';
                document.body.appendChild(a);
                a.click();
                a.remove();
            } else {
                alert('فشل التصدير.');
            }
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
async def read_index():
    return html_content


@app.post("/chat")
async def chat_endpoint(request: Request):
    data = await request.json()
    try:
        model = genai.GenerativeModel(get_best_model(data['apiKey']))
        response = model.generate_content(data['prompt'])
        return {"reply": response.text}
    except Exception as e:
        return JSONResponse(content={"detail": str(e)}, status_code=400)


@app.post("/generate")
async def generate_endpoint(request: Request):
    data = await request.json()
    topic = data.get("topic")
    api_key = data.get("apiKey")
    service_type = data.get("serviceType")

    try:
        model = genai.GenerativeModel(get_best_model(api_key))
        ai_content = model.generate_content(f"اكتب تفاصيل حول: {topic}").text
    except:
        ai_content = f"محتوى خاص بـ: {topic}"

    if service_type == "excel":
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "ميزانية الربع الأول Q1"
        ws.sheet_view.rightToLeft = True

        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        regular_font = Font(name="Segoe UI", size=10)
        bold_font = Font(name="Segoe UI", size=10, bold=True)
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
        )

        headers = ["البند / القسم", "يناير", "فبراير", "مارس", "الإجمالي الربعي"]
        ws.append(headers)

        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        q1_rows = [
            ["الإيرادات والمبيعات", 50000, 55000, 60000, "=SUM(B2:D2)"],
            ["التكاليف التشغيلية", 20000, 22000, 21000, "=SUM(B3:D3)"],
            ["الرواتب والأجور", 15000, 15000, 15000, "=SUM(B4:D4)"],
            ["المصروفات النثرية", 5000, 4000, 4500, "=SUM(B5:D5)"],
            ["صافي الربح التشغيلي", "=B2-SUM(B3:B5)", "=C2-SUM(C3:C5)", "=D2-SUM(D3:D5)", "=E2-SUM(E3:E5)"]
        ]

        for row_idx, row_data in enumerate(q1_rows, start=2):
            ws.append(row_data)
            for col_idx in range(1, len(row_data) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = bold_font if row_idx == 6 else regular_font
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 5, 18)

        path = "Q1_Professional_Sheet.xlsx"
        wb.save(path)
        return FileResponse(path, filename="Q1_Professional_Sheet.xlsx")

    else:
        doc = docx.Document()
        doc.add_heading(f'تقرير: {topic}', 0)
        for line in ai_content.split('\n'):
            if line.strip(): doc.add_paragraph(line)
        path = "Gemini_Report.docx"
        doc.save(path)
        return FileResponse(path, filename="Gemini_Report.docx")
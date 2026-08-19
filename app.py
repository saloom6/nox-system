from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
import os
import docx
from pptx import Presentation
import openpyxl

app = FastAPI()

html_content = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <title>NOX | المنظومة الذكية</title>
    <style>
        body { font-family: Tahoma, sans-serif; background-color: #0b0f19; color: #f8fafc; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }
        .card { background: #131b2e; padding: 30px; border-radius: 16px; width: 400px; box-shadow: 0 8px 24px rgba(0,0,0,0.6); border: 1px solid #1e293b; text-align: center; }
        input, select { width: 100%; padding: 12px; margin: 10px 0; border-radius: 8px; background: #0b0f19; border: 1px solid #334155; color: #fff; box-sizing: border-box; text-align: right; }
        select option { background: #131b2e; color: #fff; direction: rtl; }
        button { width: 100%; background: #4f46e5; color: white; border: none; padding: 12px; border-radius: 8px; cursor: pointer; font-weight: bold; margin-top: 10px; }
        #status { margin-top: 15px; font-size: 13px; }
    </style>
</head>
<body>
    <div class="card">
        <h2>NOX ENGINE v2.0</h2>
        <input type="text" id="topic" placeholder="موضوع الملف أو التقرير">
        <input type="password" id="apiKey" placeholder="مفتاح الوصول (API Key)">

        <select id="serviceType">
            <option value="word">تقرير نصي - Word</option>
            <option value="powerpoint">عرض تقديمي - PowerPoint</option>
            <option value="excel">جدول بيانات - Excel</option>
        </select>

        <button onclick="generateFile()">تنفيذ وتوليد الملف</button>
        <div id="status"></div>
    </div>

    <script>
        async function generateFile() {
            const topic = document.getElementById('topic').value;
            const apiKey = document.getElementById('apiKey').value;
            const serviceType = document.getElementById('serviceType').value;
            const statusDiv = document.getElementById('status');

            if (!topic) {
                statusDiv.innerHTML = '<span style="color: #f87171;">⚠️ الرجاء كتابة موضوع الملف</span>';
                return;
            }

            statusDiv.innerHTML = '<span style="color: #38bdf8;">⏳ جاري توليد الملف...</span>';

            try {
                const response = await fetch('/generate', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ topic, apiKey, serviceType })
                });

                if (response.ok) {
                    const blob = await response.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;

                    if (serviceType === 'word') a.download = 'NOX_Report.docx';
                    else if (serviceType === 'powerpoint') a.download = 'NOX_Presentation.pptx';
                    else a.download = 'NOX_Sheet.xlsx';

                    document.body.appendChild(a);
                    a.click();
                    a.remove();

                    statusDiv.innerHTML = '<span style="color: #4ade80;">✅ تم توليد وتنزيل الملف بنجاح!</span>';
                } else {
                    const err = await response.json();
                    statusDiv.innerHTML = '<span style="color: #f87171;">❌ ' + (err.detail || 'حدث خطأ') + '</span>';
                }
            } catch (error) {
                statusDiv.innerHTML = '<span style="color: #f87171;">❌ تعذر الاتصال بالسيرفر.</span>';
            }
        }
    </script>
</body>
</html>
"""


@app.get("/", response_class=HTMLResponse)
async def read_index():
    return html_content


@app.post("/generate")
async def generate_endpoint(request: Request):
    data = await request.json()
    topic = data.get("topic")
    service_type = data.get("serviceType")

    if service_type == "word":
        doc = docx.Document()
        doc.add_heading(f'تقرير منظومة NOX', 0)
        doc.add_paragraph(f'موضوع التقرير: {topic}')
        file_path = "NOX_Report.docx"
        doc.save(file_path)
        return FileResponse(file_path,
                            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                            filename="NOX_Report.docx")

    elif service_type == "powerpoint":
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "منظومة NOX الذكية"
        slide.placeholders[1].text = f"موضوع العرض: {topic}"
        file_path = "NOX_Presentation.pptx"
        prs.save(file_path)
        return FileResponse(file_path,
                            media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                            filename="NOX_Presentation.pptx")

    elif service_type == "excel":
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "NOX Data"
        ws['A1'] = "موضوع التقرير"
        ws['B1'] = topic
        ws['A2'] = "الحالة"
        ws['B2'] = "تم التوليد بنجاح عبر منظومة NOX"
        file_path = "NOX_Sheet.xlsx"
        wb.save(file_path)
        return FileResponse(file_path, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            filename="NOX_Sheet.xlsx")

    return JSONResponse(content={"detail": "نوع الخدمة غير مدعوم"}, status_code=400)
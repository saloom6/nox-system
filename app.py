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
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>NOX AI | المنظومة الذكية</title>
    <style>
        :root {
            --bg-sidebar: #171717;
            --bg-main: #212121;
            --bg-input: #2f2f2f;
            --border-color: #383838;
            --text-main: #ececec;
            --text-muted: #9b9b9b;
            --accent-color: #10a37f;
            --accent-hover: #1a7f64;
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

       /* الشريط الجانبي على طريقة OpenAI */
        .sidebar {
            width: 260px;
            background-color: var(--bg-sidebar);
            border-left: 1px solid var(--border-color);
            display: flex;
            flex-direction: column;
            padding: 12px;
            box-sizing: border-box;
            transition: transform 0.3s ease;
        }

        .new-chat-btn {
            background: transparent;
            border: 1px solid var(--border-color);
            color: var(--text-main);
            padding: 10px 14px;
            border-radius: 8px;
            cursor: pointer;
            text-align: right;
            font-size: 14px;
            display: flex;
            align-items: center;
            gap: 10px;
            margin-bottom: 20px;
            transition: background 0.2s;
        }
        .new-chat-btn:hover { background-color: var(--bg-input); }

        .history-title {
            font-size: 12px;
            color: var(--text-muted);
            margin-bottom: 10px;
            padding-right: 8px;
        }

        .history-list {
            flex-grow: 1;
            overflow-y: auto;
        }

        .history-item {
            padding: 10px 12px;
            border-radius: 6px;
            font-size: 13px;
            color: var(--text-main);
            cursor: pointer;
            margin-bottom: 4px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .history-item:hover { background-color: var(--bg-input); }

        /* منطقة العمل الرئيسية */
        .main-content {
            flex-grow: 1;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            position: relative;
            padding: 20px;
            box-sizing: border-box;
        }

        .chat-container {
            width: 100%;
            max-width: 650px;
            text-align: center;
        }

        .chat-container h1 {
            font-size: 28px;
            font-weight: 600;
            margin-bottom: 25px;
            color: #fff;
        }

        /* صندوق الإدخال الذكي */
        .input-box-wrapper {
            background-color: var(--bg-input);
            border: 1px solid var(--border-color);
            border-radius: 16px;
            padding: 15px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            text-align: right;
        }

        .input-box-wrapper input, 
        .input-box-wrapper select {
            width: 100%;
            background: transparent;
            border: none;
            color: var(--text-main);
            font-size: 15px;
            outline: none;
            padding: 8px 0;
            box-sizing: border-box;
        }

        .input-box-wrapper select option {
            background-color: var(--bg-input);
            color: #fff;
        }

        .divider {
            height: 1px;
            background-color: var(--border-color);
            margin: 10px 0;
        }

        .controls-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-top: 10px;
        }

        .send-btn {
            background-color: var(--accent-color);
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 8px;
            cursor: pointer;
            font-weight: 600;
            font-size: 14px;
            transition: background 0.2s;
        }
        .send-btn:hover { background-color: var(--accent-hover); }

        #status {
            margin-top: 15px;
            font-size: 13px;
        }

        /* قسم الحساب في أسفل القائمة */
        .user-profile {
            border-top: 1px solid var(--border-color);
            padding-top: 12px;
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 13px;
            color: var(--text-main);
        }
        .user-avatar {
            width: 32px;
            height: 32px;
            background: var(--accent-color);
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: bold;
        }
    </style>
</head>
<body>

    <!-- الشريط الجانبي (Sidebar) -->
    <div class="sidebar">
        <button class="new-chat-btn" onclick="location.reload()">
            <span>＋</span> مشروع / تقرير جديد
        </button>

        <div class="history-title">الملفات السابقة</div>
        <div class="history-list">
            <div class="history-item">📄 تقرير ميزانية الربع الأول</div>
            <div class="history-item">📊 عرض تقديمي للمشروع</div>
            <div class="history-item">📈 جدول الأداء المالي</div>
        </div>

        <div class="user-profile">
            <div class="user-avatar">س</div>
            <div>
                <div style="font-weight: 600;">سليمان ماهر</div>
                <div style="font-size: 11px; color: var(--text-muted);">حساب نشط (Pro)</div>
            </div>
        </div>
    </div>

    <!-- المساحة الرئيسية -->
    <div class="main-content">
        <div class="chat-container">
            <h1>ما الذي تريد إنجازه اليوم يا سليمان؟</h1>

            <div class="input-box-wrapper">
                <input type="text" id="topic" placeholder="اكتب موضوع التقرير أو الملف (مثال: خطة تسويق مصنع تاكون كولد)...">

                <div class="divider"></div>

                <input type="password" id="apiKey" placeholder="أدخل مفتاح الوصول (API Key)...">

                <div class="divider"></div>

                <div class="controls-row">
                    <select id="serviceType" style="width: 60%;">
                        <option value="word">📄 تقرير نصي - Word</option>
                        <option value="powerpoint">📊 عرض تقديمي - PowerPoint</option>
                        <option value="excel">📈 جدول بيانات - Excel</option>
                    </select>

                    <button class="send-btn" onclick="generateFile()">توليد الملف ➔</button>
                </div>
            </div>

            <div id="status"></div>
        </div>
    </div>

    <script>
        async function generateFile() {
            const topic = document.getElementById('topic').value;
            const apiKey = document.getElementById('apiKey').value;
            const serviceType = document.getElementById('serviceType').value;
            const statusDiv = document.getElementById('status');

            if (!topic) {
                statusDiv.innerHTML = '<span style="color: #ef4444;">⚠️ الرجاء إدخال موضوع الملف أولاً</span>';
                return;
            }

            statusDiv.innerHTML = '<span style="color: #38bdf8;">⏳ جاري المعالجة وتوليد الملف الذكي...</span>';

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

                    if (serviceType === 'word') a.download = 'NOX_AI_Report.docx';
                    else if (serviceType === 'powerpoint') a.download = 'NOX_AI_Presentation.pptx';
                    else a.download = 'NOX_AI_Sheet.xlsx';

                    document.body.appendChild(a);
                    a.click();
                    a.remove();

                    statusDiv.innerHTML = '<span style="color: #10a37f;">✅ تم إنشاء وتنزيل الملف بنجاح!</span>';
                } else {
                    const err = await response.json();
                    statusDiv.innerHTML = '<span style="color: #ef4444;">❌ ' + (err.detail || 'حدث خطأ') + '</span>';
                }
            } catch (error) {
                statusDiv.innerHTML = '<span style="color: #ef4444;">❌ تعذر الاتصال بالسيرفر.</span>';
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

    # محرك توليد الملفات الذكي
    if service_type == "word":
        doc = docx.Document()
        doc.add_heading(f'تقرير منظومة NOX AI', 0)
        doc.add_paragraph(f'الموضوع: {topic}')
        doc.add_paragraph('تم التوليد أوتوماتيكياً بواسطة نموذج NOX الذكي.')
        file_path = "NOX_AI_Report.docx"
        doc.save(file_path)
        return FileResponse(file_path,
                            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                            filename="NOX_AI_Report.docx")

    elif service_type == "powerpoint":
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "NOX AI Engine"
        slide.placeholders[1].text = f"الموضوع: {topic}"
        file_path = "NOX_AI_Presentation.pptx"
        prs.save(file_path)
        return FileResponse(file_path,
                            media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                            filename="NOX_AI_Presentation.pptx")

    elif service_type == "excel":
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "NOX AI Data"
        ws['A1'] = "الموضوع"
        ws['B1'] = topic
        ws['A2'] = "الحالة"
        ws['B2'] = "تم التوليد بنجاح"
        file_path = "NOX_AI_Sheet.xlsx"
        wb.save(file_path)
        return FileResponse(file_path, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            filename="NOX_AI_Sheet.xlsx")

    return JSONResponse(content={"detail": "نوع الخدمة غير مدعوم"}, status_code=400)
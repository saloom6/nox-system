from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
import google.generativeai as genai
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
    <title>NOX AI Studio | منظومة Gemini المتكاملة</title>
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

        .settings-group {
            margin-bottom: 15px;
        }
        .settings-group label {
            font-size: 12px;
            color: var(--text-muted);
            display: block;
            margin-bottom: 5px;
        }
        .settings-group select, .settings-group input {
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

        .history-title {
            font-size: 12px;
            color: var(--text-muted);
            margin: 15px 0 8px 0;
            padding-right: 4px;
        }

        .history-list {
            flex-grow: 1;
            overflow-y: auto;
        }

        .history-item {
            padding: 8px 10px;
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

        .main-content {
            flex-grow: 1;
            display: flex;
            flex-direction: column;
            position: relative;
            height: 100vh;
            box-sizing: border-box;
        }

        .chat-messages {
            flex-grow: 1;
            overflow-y: auto;
            padding: 20px;
            display: flex;
            flex-direction: column;
            gap: 20px;
            box-sizing: border-box;
            max-width: 800px;
            width: 100%;
            margin: 0 auto;
        }

        .message {
            display: flex;
            gap: 15px;
            max-width: 100%;
            line-height: 1.6;
            font-size: 15px;
        }
        .message.ai {
            align-self: flex-start;
            background-color: transparent;
            padding: 0;
            width: 100%;
        }

        .avatar {
            width: 32px;
            height: 32px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: bold;
            font-size: 13px;
            flex-shrink: 0;
        }
        .user-av { background: #5436da; color: #fff; }
        .ai-av { background: var(--accent-color); color: #fff; }

        .input-area {
            padding: 20px;
            background: linear-gradient(to top, var(--bg-main) 80%, transparent);
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
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            display: flex;
            flex-direction: column;
            gap: 10px;
        }

        .input-box-wrapper textarea {
            width: 100%;
            background: transparent;
            border: none;
            color: var(--text-main);
            font-size: 15px;
            outline: none;
            resize: none;
            height: 45px;
            box-sizing: border-box;
            font-family: inherit;
        }

        .input-controls {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-top: 1px solid var(--border-color);
            padding-top: 8px;
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

        .export-buttons {
            display: flex;
            gap: 8px;
            margin-top: 8px;
        }
        .export-btn {
            background: #334155;
            color: #fff;
            border: none;
            padding: 5px 10px;
            border-radius: 6px;
            font-size: 12px;
            cursor: pointer;
            transition: background 0.2s;
        }
        .export-btn:hover { background: #475569; }

        .user-profile {
            border-top: 1px solid var(--border-color);
            padding-top: 12px;
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 13px;
            color: var(--text-main);
            margin-top: auto;
        }
    </style>
</head>
<body>

    <div class="sidebar">
        <button class="new-chat-btn" onclick="location.reload()">
            <span>＋</span> محادثة جديدة
        </button>

        <div class="settings-group">
            <label>مفتاح Gemini API Key</label>
            <input type="password" id="apiKey" placeholder="AIzaSy...">
        </div>

        <div class="settings-group">
            <label>نموذج الذكاء الاصطناعي</label>
            <select id="modelType">
                <option value="gemini-2.0-flash">Gemini 2.0 Flash (سريع ومجاني)</option>
            </select>
        </div>

        <div class="history-title">سجل العمليات</div>
        <div class="history-list">
            <div class="history-item">📄 تقارير المشاريع والتحليل</div>
        </div>

        <div class="user-profile">
            <div class="avatar user-av">س</div>
            <div>
                <div style="font-weight: 600;">سليمان ماهر</div>
                <div style="font-size: 11px; color: var(--text-muted);">مطوّر المنظومة</div>
            </div>
        </div>
    </div>

    <div class="main-content">
        <div class="chat-messages" id="chatMessages">
            <div class="message ai">
                <div style="display: flex; gap: 15px;">
                    <div class="avatar ai-av">AI</div>
                    <div>
                        <strong>أهلاً بك يا سليمان في استوديو Gemini الذكي.</strong>
                        <p style="color: var(--text-muted); margin: 5px 0 0 0;">قم بإدخال مفتاحك المجاني من Google AI Studio بالأعلى، وابدأ بكتابة طلبك أو استفسارك لنبدأ الإنجاز الفوري.</p>
                    </div>
                </div>
            </div>
        </div>

        <div class="input-area">
            <div class="input-box-wrapper">
                <textarea id="promptInput" placeholder="اكتب رسالتك أو موضوع التقرير هنا (Enter للإرسال)..." rows="1" onkeydown="handleKeyPress(event)"></textarea>
                <div class="input-controls">
                    <span style="font-size: 12px; color: var(--text-muted);">اضغط إرسال أو اختر تصدير مباشر للملفات</span>
                    <button class="send-btn" onclick="sendMessage()">إرسال ➔</button>
                </div>
            </div>
        </div>
    </div>

    <script>
        function handleKeyPress(event) {
            if (event.key === 'Enter' && !event.shiftKey) {
                event.preventDefault();
                sendMessage();
            }
        }

        async function sendMessage() {
            const promptInput = document.getElementById('promptInput');
            const apiKey = document.getElementById('apiKey').value;
            const chatMessages = document.getElementById('chatMessages');

            const text = promptInput.value.trim();
            if (!text) return;

            if (!apiKey) {
                alert('الرجاء إدخال مفتاح Gemini API Key في القائمة الجانبية أولاً.');
                return;
            }

            chatMessages.innerHTML += `
                <div class="message" style="align-self: flex-start; background: #2f2f2f; padding: 12px 18px; border-radius: 12px; border: 1px solid var(--border-color);">
                    <div style="display: flex; gap: 10px; align-items: center;">
                        <div class="avatar user-av">س</div>
                        <div>${text}</div>
                    </div>
                </div>
            `;
            promptInput.value = '';
            chatMessages.scrollTop = chatMessages.scrollHeight;

            const loadingId = 'loading-' + Date.now();
            chatMessages.innerHTML += `
                <div class="message ai" id="${loadingId}">
                    <div style="display: flex; gap: 15px;">
                        <div class="avatar ai-av">AI</div>
                        <div style="color: #38bdf8;">⏳ جاري المعالجة عبر Gemini...</div>
                    </div>
                </div>
            `;
            chatMessages.scrollTop = chatMessages.scrollHeight;

            try {
                const response = await fetch('/chat', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ prompt: text, apiKey })
                });

                const data = await response.json();
                document.getElementById(loadingId).remove();

                if (response.ok) {
                    chatMessages.innerHTML += `
                        <div class="message ai">
                            <div style="display: flex; gap: 15px; width: 100%;">
                                <div class="avatar ai-av">AI</div>
                                <div style="width: 100%;">
                                    <div style="white-space: pre-wrap; background: #262626; padding: 15px; border-radius: 10px; border: 1px solid var(--border-color);">${data.reply}</div>
                                    <div class="export-buttons">
                                        <button class="export-btn" onclick="exportFile('${encodeURIComponent(text)}', 'word')">📥 تصدير Word</button>
                                        <button class="export-btn" onclick="exportFile('${encodeURIComponent(text)}', 'powerpoint')">📥 تصدير PowerPoint</button>
                                        <button class="export-btn" onclick="exportFile('${encodeURIComponent(text)}', 'excel')">📥 تصدير Excel</button>
                                    </div>
                                </div>
                            </div>
                        </div>
                    `;
                } else {
                    chatMessages.innerHTML += `<div class="message ai" style="color: #ef4444;">❌ خطأ: ${data.detail}</div>`;
                }
            } catch (err) {
                document.getElementById(loadingId).remove();
                chatMessages.innerHTML += `<div class="message ai" style="color: #ef4444;">❌ تعذر الاتصال بالسيرفر.</div>`;
            }
            chatMessages.scrollTop = chatMessages.scrollHeight;
        }

        async function exportFile(topic, type) {
            const apiKey = document.getElementById('apiKey').value;
            alert('جاري تجهيز وتنزيل ملف الـ ' + type.toUpperCase() + '...');

            const response = await fetch('/generate', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ topic: decodeURIComponent(topic), apiKey, serviceType: type })
            });

            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = type === 'word' ? 'Gemini_Report.docx' : type === 'powerpoint' ? 'Gemini_Presentation.pptx' : 'Gemini_Sheet.xlsx';
                document.body.appendChild(a);
                a.click();
                a.remove();
            } else {
                alert('فشل توليد الملف.');
            }
        }
    </script>
</body>
</html>
"""


@app.get("/", response_class=HTMLResponse)
async def read_index():
    return html_content


@app.post("/chat")
async def chat_endpoint(request: Request):
    data = await request.json()
    prompt = data.get("prompt")
    api_key = data.get("apiKey")

    if not api_key:
        return JSONResponse(content={"detail": "مفتاح API مطلوب"}, status_code=400)

    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-2.0-flash')
        response = model.generate_content(prompt)
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
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-2.0-flash')
        response = model.generate_content(f"اكتب تقرير مفصل ومنظم حول موضوع: {topic}")
        ai_content = response.text
    except Exception as e:
        ai_content = f"محتوى افتراضي للموضوع: {topic}"

    if service_type == "word":
        doc = docx.Document()
        doc.add_heading(f'تقرير Gemini: {topic}', 0)
        for line in ai_content.split('\n'):
            if line.strip(): doc.add_paragraph(line)
        path = "Gemini_Report.docx"
        doc.save(path)
        return FileResponse(path, filename="Gemini_Report.docx")

    elif service_type == "powerpoint":
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = topic
        slide.placeholders[1].text = "منظومة Gemini AI Studio"

        slide2 = prs.slides.add_slide(prs.slide_layouts[1])
        slide2.shapes.title.text = "التفاصيل الأساسية"
        slide2.placeholders[1].text = ai_content[:600]

        path = "Gemini_Presentation.pptx"
        prs.save(path)
        return FileResponse(path, filename="Gemini_Presentation.pptx")

    elif service_type == "excel":
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Gemini Data"
        ws['A1'] = "الموضوع"
        ws['B1'] = topic
        ws['A2'] = "التحليل المستخرج"
        ws['B2'] = ai_content[:400]
        path = "Gemini_Sheet.xlsx"
        wb.save(path)
        return FileResponse(path, filename="Gemini_Sheet.xlsx")

    return JSONResponse(content={"detail": "خطأ في نوع الملف"}, status_code=400)
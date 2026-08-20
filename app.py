from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
import os
import docx
from pptx import Presentation
import openpyxl
import openai

app = FastAPI()

html_content = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>NOX AI Studio | المنظومة المتكاملة</title>
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

        /* الشريط الجانبي */
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

        /* منطقة المحادثة الرئيسية */
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
        .message.user {
            align-self: flex-start;
            background-color: #2f2f2f;
            padding: 12px 18px;
            border-radius: 12px;
            border: 1px solid var(--border-color);
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

        /* منطقة الإدخال بالأسفل */
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

    <!-- الشريط الجانبي للإعدادات والموديلات -->
    <div class="sidebar">
        <button class="new-chat-btn" onclick="location.reload()">
            <span>＋</span> محادثة جديدة
        </button>

        <div class="settings-group">
            <label>مفتاح OpenAI API Key</label>
            <input type="password" id="apiKey" placeholder="sk-...">
        </div>

        <div class="settings-group">
            <label>اختر نموذج الذكاء الاصطناعي</label>
            <select id="modelType">
                <option value="gpt-4o-mini">GPT-4o-mini (سريع وذكي)</option>
                <option value="gpt-4o">GPT-4o (شامل وعالي الدقة)</option>
            </select>
        </div>

        <div class="settings-group">
            <label>شخصية الذكاء الاصطناعي</label>
            <select id="persona">
                <option value="مساعد أعمال ومشاريع محترف">خبير أعمال وتقارير</option>
                <option value="محلل بيانات ومالية">محلل مالي وبيانات</option>
                <option value="كاتب محتوى تسويقي إبداعي">خبير تسويق ومحتوى</option>
            </select>
        </div>

        <div class="history-title">سجل العمليات</div>
        <div class="history-list" id="historyList">
            <div class="history-item">📄 تحليل خطة العمل</div>
            <div class="history-item">📊 تقرير المبيعات</div>
        </div>

        <div class="user-profile">
            <div class="avatar user-av">س</div>
            <div>
                <div style="font-weight: 600;">سليمان ماهر</div>
                <div style="font-size: 11px; color: var(--text-muted);">مطور المنظومة</div>
            </div>
        </div>
    </div>

    <!-- منطقة الشات والدردشة الرئيسية -->
    <div class="main-content">
        <div class="chat-messages" id="chatMessages">
 رسالة الترحيب الافتراضية -->
            <div class="message ai">
                <div style="display: flex; gap: 15px;">
                    <div class="avatar ai-av">AI</div>
                    <div>
                        <strong>أهلاً بك يا سليمان في استوديو NOX الذكي.</strong>
                        <p style="color: var(--text-muted); margin: 5px 0 0 0;">قم بإدخال مفتاح الـ API الخاص بك من القائمة الجانبية، واكتب طلبك أو استفسارك بالأعلى لنبدأ الإنجاز الفوري وتوليد التقارير.</p>
                    </div>
                </div>
            </div>
        </div>

        <!-- صندوق الإدخال السفلي -->
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
            const modelType = document.getElementById('modelType').value;
            const persona = document.getElementById('persona').value;
            const chatMessages = document.getElementById('chatMessages');

            const text = promptInput.value.trim();
            if (!text) return;

            if (!apiKey) {
                alert('الرجاء إدخال مفتاح OpenAI API Key من القائمة الجانبية أولاً.');
                return;
            }

            // عرض رسالة المستخدم
            chatMessages.innerHTML += `
                <div class="message user">
                    <div style="display: flex; gap: 10px; align-items: center;">
                        <div class="avatar user-av">س</div>
                        <div>${text}</div>
                    </div>
                </div>
            `;
            promptInput.value = '';
            chatMessages.scrollTop = chatMessages.scrollHeight;

            // رسالة الانتظار
            const loadingId = 'loading-' + Date.now();
            chatMessages.innerHTML += `
                <div class="message ai" id="${loadingId}">
                    <div style="display: flex; gap: 15px;">
                        <div class="avatar ai-av">AI</div>
                        <div style="color: #38bdf8;">⏳ جاري التفكير ومعالجة الطلب...</div>
                    </div>
                </div>
            `;
            chatMessages.scrollTop = chatMessages.scrollHeight;

            try {
                const response = await fetch('/chat', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ prompt: text, apiKey, modelType, persona })
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
            const modelType = document.getElementById('modelType').value;

            alert('جاري تجهيز وتنزيل ملف الـ ' + type.toUpperCase() + '...');

            const response = await fetch('/generate', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ topic: decodeURIComponent(topic), apiKey, modelType, serviceType: type })
            });

            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = type === 'word' ? 'NOX_Report.docx' : type === 'powerpoint' ? 'NOX_Presentation.pptx' : 'NOX_Sheet.xlsx';
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
    model_type = data.get("modelType", "gpt-4o-mini")
    persona = data.get("persona", "مساعد أعمال محترف")

    if not api_key:
        return JSONResponse(content={"detail": "مفتاح API مطلوب"}, status_code=400)

    try:
        client = openai.OpenAI(api_key=api_key)
        completion = client.chat.completions.create(
            model=model_type,
            messages=[
                {"role": "system",
                 "content": f"أنت ذكاء اصطناعي تعمل بصفتك: {persona}. أجب باحترافية وبتنسيق دقيق باللغة العربية."},
                {"role": "user", "content": prompt}
            ]
        )
        reply = completion.choices[0].message.content
        return {"reply": reply}
    except Exception as e:
        return JSONResponse(content={"detail": str(e)}, status_code=400)


@app.post("/generate")
async def generate_endpoint(request: Request):
    data = await request.json()
    topic = data.get("topic")
    api_key = data.get("apiKey")
    model_type = data.get("modelType", "gpt-4o-mini")
    service_type = data.get("serviceType")

    try:
        client = openai.OpenAI(api_key=api_key)
        completion = client.chat.completions.create(
            model=model_type,
            messages=[
                {"role": "system", "content": "أنت مساعد ذكاء اصطناعي لتجهيز محتوى الملفات والتقارير التنفيذية."},
                {"role": "user", "content": f"اكتب محتوى تفصيلي ومنظم لإنشاء ملف حول: {topic}"}
            ]
        )
        ai_content = completion.choices[0].message.content
    except Exception as e:
        ai_content = f"محتوى افتراضي للموضوع: {topic}"

    if service_type == "word":
        doc = docx.Document()
        doc.add_heading(f'تقرير: {topic}', 0)
        for line in ai_content.split('\n'):
            if line.strip(): doc.add_paragraph(line)
        path = "NOX_Report.docx"
        doc.save(path)
        return FileResponse(path, filename="NOX_Report.docx")

    elif service_type == "powerpoint":
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = topic
        slide.placeholders[1].text = "منظومة NOX AI Studio"

        slide2 = prs.slides.add_slide(prs.slide_layouts[1])
        slide2.shapes.title.text = "التفاصيل الأساسية"
        slide2.placeholders[1].text = ai_content[:600]

        path = "NOX_Presentation.pptx"
        prs.save(path)
        return FileResponse(path, filename="NOX_Presentation.pptx")

    elif service_type == "excel":
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "NOX Data"
        ws['A1'] = "الموضوع"
        ws['B1'] = topic
        ws['A2'] = "التحليل المستخرج"
        ws['B2'] = ai_content[:400]
        path = "NOX_Sheet.xlsx"
        wb.save(path)
        return FileResponse(path, filename="NOX_Sheet.xlsx")

    return JSONResponse(content={"detail": "خطأ في نوع الملف"}, status_code=400)
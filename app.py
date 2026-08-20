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
    <title>NOX AI Studio | منظومة Gemini الذكية</title>
    <style>
        :root { --bg-sidebar: #171717; --bg-main: #212121; --bg-input: #2f2f2f; --border-color: #383838; --text-main: #ececec; --accent-color: #1a73e8; }
        body { font-family: sans-serif; background-color: var(--bg-main); color: var(--text-main); margin: 0; display: flex; height: 100vh; }
        .sidebar { width: 280px; background-color: var(--bg-sidebar); border-left: 1px solid var(--border-color); padding: 15px; }
        .settings-group { margin-bottom: 15px; }
        .settings-group label { font-size: 12px; color: #9b9b9b; display: block; margin-bottom: 5px; }
        .settings-group input { width: 100%; background: var(--bg-input); border: 1px solid var(--border-color); color: #fff; padding: 8px; border-radius: 6px; }
        .main-content { flex-grow: 1; display: flex; flex-direction: column; }
        .chat-messages { flex-grow: 1; overflow-y: auto; padding: 20px; display: flex; flex-direction: column; gap: 20px; max-width: 800px; margin: 0 auto; }
        .input-area { padding: 20px; display: flex; justify-content: center; }
        .input-box-wrapper { background-color: var(--bg-input); border: 1px solid var(--border-color); border-radius: 16px; padding: 12px; width: 100%; max-width: 800px; }
        textarea { width: 100%; background: transparent; border: none; color: #fff; outline: none; resize: none; height: 45px; }
        .send-btn { background-color: var(--accent-color); color: white; border: none; padding: 8px 16px; border-radius: 8px; cursor: pointer; }
        .export-btn { background: #334155; color: #fff; border: none; padding: 5px 10px; border-radius: 6px; font-size: 12px; cursor: pointer; margin-top: 5px; }
    </style>
</head>
<body>
    <div class="sidebar">
        <div class="settings-group"><label>مفتاح Gemini API Key</label><input type="password" id="apiKey" placeholder="AIzaSy..."></div>
    </div>
    <div class="main-content">
        <div class="chat-messages" id="chatMessages"></div>
        <div class="input-area">
            <div class="input-box-wrapper">
                <textarea id="promptInput" placeholder="اكتب طلبك هنا..."></textarea>
                <button class="send-btn" onclick="sendMessage()">إرسال ➔</button>
            </div>
        </div>
    </div>
    <script>
        async function sendMessage() {
            const promptInput = document.getElementById('promptInput');
            const apiKey = document.getElementById('apiKey').value;
            const chatMessages = document.getElementById('chatMessages');
            const text = promptInput.value.trim();
            if (!text || !apiKey) return;

            chatMessages.innerHTML += `<div class="message" style="background:#2f2f2f; padding:10px; border-radius:10px;">سليمان: ${text}</div>`;
            promptInput.value = '';

            const response = await fetch('/chat', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ prompt: text, apiKey })
            });
            const data = await response.json();
            chatMessages.innerHTML += `<div class="message" style="background:#262626; padding:10px; border-radius:10px;">AI: ${data.reply}
                <br><button class="export-btn" onclick="exportFile('${encodeURIComponent(text)}', 'word')">📥 تصدير Word</button></div>`;
        }
        async function exportFile(topic, type) {
            const apiKey = document.getElementById('apiKey').value;
            const response = await fetch('/generate', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ topic: decodeURIComponent(topic), apiKey, serviceType: type })
            });
            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url; a.download = 'Gemini_Report.docx';
                document.body.appendChild(a); a.click(); a.remove();
            }
        }
    </script>
</body>
</html>
"""


@app.get("/", response_class=HTMLResponse)
async def read_index(): return html_content


@app.post("/chat")
async def chat_endpoint(request: Request):
    data = await request.json()
    genai.configure(api_key=data.get("apiKey"))
    model = genai.GenerativeModel('gemini-1.5-flash-latest')
    response = model.generate_content(data.get("prompt"))
    return {"reply": response.text}


@app.post("/generate")
async def generate_endpoint(request: Request):
    data = await request.json()
    genai.configure(api_key=data.get("apiKey"))
    model = genai.GenerativeModel('gemini-1.5-flash-latest')
    response = model.generate_content(f"اكتب تقرير مفصل حول: {data.get('topic')}")
    doc = docx.Document()
    doc.add_paragraph(response.text)
    doc.save("Gemini_Report.docx")
    return FileResponse("Gemini_Report.docx", filename="Gemini_Report.docx")
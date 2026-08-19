from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse
import os

app = FastAPI()


# نقطة الدخول الرئيسية: تعرض صفحة الـ HTML
@app.get("/", response_class=HTMLResponse)
async def read_index():
    file_path = "index.html"
    if os.path.exists(file_path):
        with open(file_path, "r", encoding="utf-8") as f:
            return f.read()
    return "<h1>عذراً، ملف index.html غير موجود!</h1>"


# مسار معالجة الطلبات (مهم جداً لربط الواجهة بالسيرفر)
@app.post("/generate")
async def generate_endpoint(request: Request):
    data = await request.json()
    topic = data.get("topic")
    service_type = data.get("serviceType")

    # هنا ستضع مستقبلاً كود الربط بالذكاء الاصطناعي
    return JSONResponse(content={
        "status": "success",
        "message": f"تم استلام طلبك لـ ({topic}) وسيتم توليد {service_type}."
    })
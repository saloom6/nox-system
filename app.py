from fastapi import FastAPI
from fastapi.responses import HTMLResponse

app = FastAPI()

html_content = """




    منظومة سليمان ماهر - NOX




        أهلاً بك يا سليمان في منظومة NOX الذكية 🚀
        الموقع يعمل بنجاح تام على سحابة Render!



"""


@app.get("/", response_class=HTMLResponse)
async def main_page():
    return html_content
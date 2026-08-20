@app.post("/generate")
async def generate_endpoint(request: Request):
    data = await request.json()
    topic = data.get("topic")
    api_key = data.get("apiKey")
    service_type = data.get("serviceType")

    # توجيه ذكي عام لأي موضوع يطلبه المستخدم لجلب تفاصيل احترافية
    try:
        model = genai.GenerativeModel(get_best_model(api_key))
        if service_type == "powerpoint":
            ai_prompt = f"""
            أنت خبير عروض تقديمية. أنشئ محتوى عرض تقديمي احترافي من 4 شرائح حول: '{topic}'.
            التنسيق بدقة:
            شريحة: العنوان الأول
            - نقطة تفصيلية
            - نقطة تفصيلية
            شريحة: العنوان الثاني
            - نقطة تفصيلية
            - نقطة تفصيلية
            شريحة: العنوان الثالث
            - نقطة تفصيلية
            - نقطة تفصيلية
            شريحة: الخاتمة والتوصيات
            - نقطة تفصيلية
            - نقطة تفصيلية
            """
        elif service_type == "excel":
            ai_prompt = f"""
            أنت محلل بيانات وخبير إكسل. بناءً على موضوع الطلب: '{topic}', 
            اقترح 4 بنود رئيسية أو أقسام يمكن وضعها في جدول بيانات إكسل، بحيث يعطى كل بند 3 أعمدة قيم رقمية أو وصفية تناسب سياق الموضوع، 
            واكتب النتائج في أسطر تبدأ بـ 'السطر:' وتحتوي على البيانات مفصولة بشرطة عمودية (|).
            مثال:
            السطر: العنصر الأول | القيمة أ | القيمة ب | القيمة ج
            """
        else:
            ai_prompt = f"اكتب تقريراً احترافياً ومفصلاً وشاملاً حول الموضوع التالي بناءً على طلب المستخدم: '{topic}'"

        ai_content = model.generate_content(ai_prompt).text
    except:
        ai_content = f"محتوى تفصيلي خاص بمشروع: {topic}"

    # 1. تصدير ملف إكسل ديناميكي يعتمد على طلب المستخدم
    if service_type == "excel":
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "تحليل المشروع الذكي"
        ws.sheet_view.rightToLeft = True

        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        regular_font = Font(name="Segoe UI", size=10)
        bold_font = Font(name="Segoe UI", size=10, bold=True)
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
        )

        headers = ["البند / القسم الرئيسي", "البيان الأول", "البيان الثاني", "البيان الثالث", "التقييم / المجموع"]
        ws.append(headers)

        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # محاولة استخراج الأسطر الذكية التي ولّدها النموذج أو وضع بيانات ديناميكية افتراضية بناءً على موضوع المستخدم
        rows_data = []
        for line in ai_content.split('\n'):
            if "السطر:" in line:
                parts = [p.strip() for p in line.replace("السطر:", "").split("|")]
                if len(parts) >= 2:
                    while len(parts) < 5:
                        parts.append("موافق")
                    rows_data.append(parts[:5])

        if not rows_data:
            # بيانات ديناميكية افتراضية ذكية تتكيف مع أي موضوع يتم إدخاله
            rows_data = [
                [f"تحليل {topic}", "الركيزة الأساسية", "مستوى التنفيذ", "جاهزية التشغيل", "ممتاز"],
                ["الموارد والتكلفة", "دراسة الميزانية", "التكلفة التقديرية", "تحسين الموارد", "مكتمل"],
                ["مؤشرات الأداء (KPIs)", "معدل الإنجاز", "جودة المخرجات", "تقييم الأداء", "عالي"],
                ["التوصيات والملاحظات", "الخطة الزمنية", "إدارة المخاطر", "المتابعة المستمرة", "نشط"]
            ]

        for row_idx, row_vals in enumerate(rows_data, start=2):
            ws.append(row_vals)
            for col_idx in range(1, len(row_vals) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = regular_font
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 5, 20)

        path = "Dynamic_Project_Sheet.xlsx"
        wb.save(path)
        return FileResponse(path, filename="Dynamic_Project_Sheet.xlsx")

    # 2. تصدير ملف وورد ديناميكي
    elif service_type == "word":
        doc = docx.Document()
        doc.add_heading(f'تقرير استراتيجي: {topic}', 0)
        for line in ai_content.split('\n'):
            if line.strip(): doc.add_paragraph(line)
        path = "Dynamic_Report.docx"
        doc.save(path)
        return FileResponse(path, filename="Dynamic_Report.docx")

    # 3. تصدير بوربوينت ديناميكي
    elif service_type == "powerpoint":
        prs = Presentation()
        slide1 = prs.slides.add_slide(prs.slide_layouts[0])
        slide1.shapes.title.text = topic
        slide1.placeholders[1].text = f"إعداد: منظومة سليمان ماهر الذكية\nموضوع المشروع: {topic}"

        sections = ai_content.split("شريحة:")
        for sec in sections:
            if not sec.strip(): continue
            lines = [l.strip() for l in sec.split('\n') if l.strip()]
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = lines[0] if lines else f"مشروع: {topic}"
            bullet_points = "\n".join(lines[1:]) if len(lines) > 1 else f"- تفاصيل وتقارير شاملة حول {topic}"
            slide.placeholders[1].text = bullet_points

        path = "Dynamic_Presentation.pptx"
        prs.save(path)
        return FileResponse(path, filename="Dynamic_Presentation.pptx")

    # 4. تصدير نصي افتراضي
    else:
        path = "Dynamic_Document.txt"
        with open(path, "w", encoding="utf-8") as f:
            f.write(ai_content)
        return FileResponse(path, filename="Dynamic_Document.txt")
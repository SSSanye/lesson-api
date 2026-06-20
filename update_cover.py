from pathlib import Path

path = Path("/opt/lesson-api/main.py")
text = path.read_text(encoding="utf-8")

cover_func = r'''
def set_cover_paragraph(paragraph, label: str, value: str):
    """
    填充封面普通段落，例如：课程名称、编写教师、授课班级等。
    """
    paragraph.clear()

    label_run = paragraph.add_run(f"{label}：")
    label_run.bold = True
    label_run.font.name = "黑体"
    label_run._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")
    label_run.font.size = Pt(16)

    value_run = paragraph.add_run(f"  {value or ''}  ")
    value_run.underline = True
    value_run.font.name = "宋体"
    value_run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    value_run.font.size = Pt(16)


def fill_cover_info(doc, data: LessonDocxRequest):
    """
    填充封面信息。封面一般是普通段落，不是表格。
    """
    cover_map = {
        "课程名称": value_or_default(data.course_name),
        "编写教师": value_or_default(data.teacher_name),
        "授课班级": value_or_default(data.class_name),
        "授课学期": value_or_default(data.semester),
        "教研室（组）": value_or_default(getattr(data, "teaching_group", "")),
        "二级学院": value_or_default(data.college)
    }

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        for label, value in cover_map.items():
            if label in text:
                set_cover_paragraph(paragraph, label, value)
                break
'''

# 1. 给 LessonDocxRequest 增加 teaching_group 字段
if "teaching_group: str | None = None" not in text:
    text = text.replace(
        "    college: str | None = None\n    lesson_date: str | None = None",
        "    college: str | None = None\n    teaching_group: str | None = None\n    lesson_date: str | None = None"
    )

# 2. 插入封面填充函数
if "def fill_cover_info(doc, data: LessonDocxRequest):" not in text:
    marker = "def fill_template_docx(data: LessonDocxRequest) -> str:"
    if marker not in text:
        raise SystemExit("未找到 fill_template_docx 函数，无法插入封面填充函数。")
    text = text.replace(marker, cover_func + "\n\n" + marker)

# 3. 在打开模板后调用 fill_cover_info
if "fill_cover_info(doc, data)" not in text:
    text = text.replace(
        "    doc = Document(TEMPLATE_PATH)\n\n    if len(doc.tables) == 0:",
        "    doc = Document(TEMPLATE_PATH)\n\n    fill_cover_info(doc, data)\n\n    if len(doc.tables) == 0:"
    )

# 4. 更新版本号
text = text.replace('version="0.3.0"', 'version="0.4.0"')
text = text.replace('"version": "0.3.0"', '"version": "0.4.0"')

path.write_text(text, encoding="utf-8")
print("更新完成：已支持填充封面信息。")

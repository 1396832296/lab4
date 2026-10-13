# -*- coding: utf-8 -*-
"""
Build the Lab 4 report (XML + JSON experiments) for 乔润喆 / 247050.

The script:
  1. trims the white space at the bottom of the browser screenshots;
  2. reads the real source files and embeds them as the "core code";
  3. writes the report in the layout of the course lab-report template.
"""

import re
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from PIL import Image, ImageChops

STUDENT_NAME = "乔润喆"
STUDENT_ID = "247050"
GITHUB_NOTE = "https://1396832296.github.io/lab4/"

ROOT = Path(r"C:\Users\13968\Desktop\web developement\Lab4-XML-JSON")
SHOTS = ROOT / "reports" / "screenshots"
DOCX = ROOT / "reports" / f"Lab4-{STUDENT_ID}-{STUDENT_NAME}.docx"

BODY_FONT = "Times New Roman"
CJK_FONT = "宋体"
HEAD_FONT = "Arial"
CJK_HEAD_FONT = "黑体"
CODE_FONT = "Consolas"


# --------------------------------------------------------------------------- #
# 1. figures
# --------------------------------------------------------------------------- #
def trim(path: Path) -> None:
    """Crop the uniform page background from the right/bottom of a screenshot."""
    img = Image.open(path).convert("RGB")
    bg = Image.new("RGB", img.size, (244, 244, 244))
    box = ImageChops.difference(img, bg).getbbox()
    if box:
        left, top, right, bottom = box
        img = img.crop((0, 0, min(img.width, right + 25), min(img.height, bottom + 25)))
        img.save(path)
    print(f"  figure {path.name}: {img.width}x{img.height}")


# --------------------------------------------------------------------------- #
# 2. python-docx helpers
# --------------------------------------------------------------------------- #
def style_run(run, font=BODY_FONT, cjk=CJK_FONT, size=11, bold=False,
              italic=False, color=None):
    run.font.name = font
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = RGBColor(*color)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), cjk)
    return run


def body(doc, text="", size=11, bold=False, italic=False, indent=0.0, space_after=6,
         align=None, font=BODY_FONT, cjk=CJK_FONT, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    if align is not None:
        p.alignment = align
    if text:
        style_run(p.add_run(text), font=font, cjk=cjk, size=size, bold=bold,
                  italic=italic, color=color)
    return p


def bullet(doc, text, indent=0.75):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Cm(indent + 0.4)
    p.paragraph_format.space_after = Pt(2)
    style_run(p.add_run(text))
    return p


def heading_cn(doc, text):
    """一、实验目的 style heading."""
    return body(doc, text, size=14, bold=True, space_after=8,
                font=HEAD_FONT, cjk=CJK_HEAD_FONT)


def sub_heading(doc, text, size=11.5, space_before=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(4)
    style_run(p.add_run(text), bold=True, size=size)
    return p


def shade(paragraph, fill="F5F5F5"):
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pPr.append(shd)


def code(doc, source, caption=None):
    """Insert a monospaced code block, one paragraph per line."""
    if caption:
        body(doc, caption, size=10, bold=True, space_after=3)
    lines = source.rstrip("\n").split("\n")
    for i, line in enumerate(lines):
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(4 if i == 0 else 0)
        pf.space_after = Pt(4 if i == len(lines) - 1 else 0)
        pf.left_indent = Cm(0.6)
        pf.line_spacing = 1.0
        style_run(p.add_run(line if line else " "), font=CODE_FONT, cjk=CODE_FONT, size=8.5)
        shade(p)
    return None


def figure(doc, image_path: Path, caption: str, max_height_in=4.6, max_width_in=6.2):
    """Insert a centred figure, keeping it inside max_height_in x max_width_in."""
    with Image.open(image_path) as img:
        ratio = img.width / img.height
    width_in = min(max_width_in, max_height_in * ratio)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.add_run().add_picture(str(image_path), width=Inches(width_in))
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_after = Pt(10)
    style_run(cap.add_run(caption), size=9, italic=True, cjk=CJK_FONT)


def script_of(html_text: str) -> str:
    """Return the content of the (single) <script> block of an HTML file."""
    match = re.search(r"<script>(.*?)</script>", html_text, re.S)
    return match.group(1).strip("\n") if match else html_text


# --------------------------------------------------------------------------- #
# 3. build the report
# --------------------------------------------------------------------------- #
def build():
    print("Trimming figures ...")
    for name in ("xml-lab", "json-lab"):
        trim(SHOTS / f"{name}.png")

    xml_src = (ROOT / "xml-lab" / "book.xml").read_text(encoding="utf-8")
    xml_html = (ROOT / "xml-lab" / "index.html").read_text(encoding="utf-8")
    json_src = (ROOT / "json-lab" / "userProfiles.json").read_text(encoding="utf-8")
    json_html = (ROOT / "json-lab" / "index.html").read_text(encoding="utf-8")

    doc = Document()

    # page setup: A4 with 2 cm margins
    for section in doc.sections:
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)

    normal = doc.styles["Normal"]
    normal.font.name = BODY_FONT
    normal.font.size = Pt(11)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), CJK_FONT)

    # ------------------------------------------------------------------ title
    body(doc, "Lab 4   XML Lab: Create an XML document", size=16, bold=True,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4,
         font=HEAD_FONT, cjk=CJK_HEAD_FONT)
    body(doc, f"姓名：{STUDENT_NAME}          学号：{STUDENT_ID}", size=11,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14)

    # ================================================== experiment 1: XML ====
    body(doc, "实验一  XML Lab: Create an XML document", size=13, bold=True,
         space_after=8, font=HEAD_FONT, cjk=CJK_HEAD_FONT)

    heading_cn(doc, "一、实验目的")
    body(doc,
         "In this lab, you will create an XML document representing a simple book catalogue. "
         "You will then validate the XML to ensure it's well-formed and finally display the XML "
         "data in an HTML page.")

    heading_cn(doc, "二、实验内容")
    sub_heading(doc, "Create the XML Document:", space_before=0)
    bullet(doc, "Create a new XML document named book.xml.")
    bullet(doc, "Each book should have an id attribute.")
    bullet(doc, "Each book should contain child elements: title, author, and price.")
    sub_heading(doc, "Validate the XML:")
    bullet(doc, "Ensure that your XML is well-formed.")
    bullet(doc, "Use online XML validators like XML Validation to check the structure.")
    sub_heading(doc, "Display the XML Data in an HTML Page:")
    bullet(doc, "Fetch the XML data using fetch API in JavaScript.")
    bullet(doc, "Parse the XML string \u201cxmlText\u201d using:")
    code(doc,
         "const parser = new DOMParser();\n"
         "const xmlDoc = parser.parseFromString(xmlText, 'text/xml');")
    bullet(doc, "Display each book's title, author, and price in the HTML page.")
    bullet(doc, "Publish your webpage on GitHub Pages and share the link in your submission.")

    heading_cn(doc, "三、实验结果")
    sub_heading(doc, "核心代码", space_before=0)
    body(doc, "【调试成功的核心源程序代码】", size=10, italic=True, space_after=4,
         color=(0x1F, 0x49, 0x7D))
    code(doc, xml_src, caption="（1）book.xml —— XML 文档（每个 <book> 都有 id 属性，并包含 title、author、price 子元素）")
    code(doc, script_of(xml_html),
         caption="（2）xml-lab/index.html —— 使用 fetch API 获取并用 DOMParser 解析 XML 的核心代码")

    sub_heading(doc, "运行效果：")
    body(doc, "【用截图方式辅助说明程序运行结果】", size=10, italic=True, space_after=2,
         color=(0x1F, 0x49, 0x7D))
    bullet(doc, "book.xml 通过在线 XML 校验器检查，结构良好（well-formed），浏览器控制台无 parsererror。")
    bullet(doc, "页面加载后提示 \u201cThe XML document is well-formed. 5 books found.\u201d，"
                "5 本书的 id、title、author、price 全部正确显示。")
    figure(doc, SHOTS / "xml-lab.png",
           "图 1  XML 实验运行效果（线上地址：https://1396832296.github.io/lab4/xml-lab/）："
           "成功读取并显示 book.xml 中的 5 本书")
    body(doc, f"GitHub Pages 链接：{GITHUB_NOTE}xml-lab/", size=10, space_after=14)

    doc.add_page_break()

    # ================================================= experiment 2: JSON ====
    heading_cn(doc, "实验二  JSON Lab: Creating a Simple User Profile")
    body(doc,
         "In this lab, you will craft a JSON document representing a set of user profiles. "
         "After creating the document, you will validate its structure to ensure its correctness. "
         "Finally, you will display the user profile data on an HTML page.")

    heading_cn(doc, "二、实验内容")
    sub_heading(doc, "Create the JSON Document:", space_before=0)
    bullet(doc, "Create a new JSON document named userProfiles.json.")
    bullet(doc, "Each user profile should have the following data: firstName, lastName, email, "
                "dateOfBirth, and user preferences (which includes theme, language, notifications).")
    bullet(doc, "Add at least 3 user profiles to the document, each with its unique details "
                "and preferences.")
    sub_heading(doc, "Validate the JSON:")
    bullet(doc, "Confirm that your JSON structure is well-formed.")
    bullet(doc, "Use online JSON validators like JSONFORMATTER&VALIDATOR to check the structure. "
                "Simply paste your JSON data into the validator and check for any errors.")
    sub_heading(doc, "Extend the JSON:")
    bullet(doc, "Add an address field to the profile section of each user. The address should "
                "include street, city, state, and zipCode.")
    bullet(doc, "Add a subscription field to the preferences section indicating whether the user "
                "has subscribed to the newsletter (true or false).")
    sub_heading(doc, "Display the User Profiles in an HTML Page:")
    bullet(doc, "Fetch and parse the JSON data employing fetch API in JavaScript.")
    code(doc,
         "fetch('xxx.json')\n"
         "    .then(response => {\n"
         "        if (!response.ok)\n"
         "            throw new Error('Fetch error: ' + response.status)\n"
         "        else\n"
         "            return response.json();\n"
         "    })\n"
         "    .then(json => {...})   // things you are going to do with the parsed json\n"
         "    .catch(err => console.error('Fetch error: ' + err.message));")
    bullet(doc, "Exhibit each user's firstName, lastName, email, dateOfBirth, and other attributes "
                "on the HTML page.")
    bullet(doc, "Publish your webpage on GitHub Pages and share the link in your submission.")

    heading_cn(doc, "三、实验结果")
    sub_heading(doc, "核心代码", space_before=0)
    body(doc, "【调试成功的核心源程序代码】", size=10, italic=True, space_after=4,
         color=(0x1F, 0x49, 0x7D))
    code(doc, json_src,
         caption="（1）userProfiles.json —— 4 个用户资料，包含 firstName、lastName、email、dateOfBirth、"
                 "address（street / city / state / zipCode）与 preferences（theme / language / "
                 "notifications / subscription）")
    code(doc, script_of(json_html),
         caption="（2）json-lab/index.html —— 使用 fetch API 获取并解析 JSON 的核心代码")

    sub_heading(doc, "运行效果：")
    body(doc, "【用截图方式辅助说明程序运行结果】", size=10, italic=True, space_after=2,
         color=(0x1F, 0x49, 0x7D))
    bullet(doc, "userProfiles.json 通过在线 JSON 校验器检查，JSON 结构良好，无语法错误。")
    bullet(doc, "页面加载后提示 \u201c4 user profiles loaded.\u201d，4 位用户的姓名、邮箱、出生日期、"
                "地址以及 theme、language、notifications、subscription 全部正确显示。")
    figure(doc, SHOTS / "json-lab.png",
           "图 2  JSON 实验运行效果（线上地址：https://1396832296.github.io/lab4/json-lab/）："
           "成功读取并显示 userProfiles.json 中的 4 个用户资料")
    body(doc, f"GitHub Pages 链接：{GITHUB_NOTE}json-lab/", size=10, space_after=14)

    body(doc, f"备注：上传文件命名要求：Lab4-学号-姓名.doc   ——   本文件："
              f"Lab4-{STUDENT_ID}-{STUDENT_NAME}.doc",
         size=10.5, bold=True, color=(0xC0, 0x00, 0x00))

    doc.save(DOCX)
    print(f"written: {DOCX}")


if __name__ == "__main__":
    build()


import zipfile
import os
import xml.etree.ElementTree as ET
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import streamlit as st
import tempfile
import re


def parse_formatted_html(text_frame, html_content, style=None):
    import html
    from html.parser import HTMLParser
    from pptx.dml.color import RGBColor
    from pptx.util import Pt
    from pptx.enum.text import PP_ALIGN

    px_to_pt = {8: 6, 10: 7, 12: 9, 14: 10, 16: 12, 18: 14, 20: 15, 22: 17, 25: 19, 30: 22, 35: 25, 40: 30, 45: 34, 50: 36}

    class PPTXHTMLParser(HTMLParser):
        def __init__(self, text_frame):
            super().__init__()
            self.text_frame = text_frame
            self.p = text_frame.paragraphs[0]
            self.current_style = {"bold": False, "italic": False, "underline": False}
            self.new_paragraph = False
            self.default_font = style.get("font", "Tahoma") if style else "Tahoma"
            self.default_size = int(style.get("fontsize", "20")) if style else 20
            self.default_color = style.get("fontcolor", "#000000") if style else "#000000"

        def apply_style(self, run):
            run.font.bold = self.current_style["bold"]
            run.font.italic = self.current_style["italic"]
            run.font.underline = self.current_style["underline"]
            run.font.name = self.default_font
            run.font.size = Pt(px_to_pt.get(self.default_size, self.default_size))
            try:
                r, g, b = tuple(int(self.default_color.lstrip("#")[i:i+2], 16) for i in (0, 2, 4))
                run.font.color.rgb = RGBColor(r, g, b)
            except:
                pass

        def handle_starttag(self, tag, attrs):
            if tag == "b":
                self.current_style["bold"] = True
            elif tag == "i":
                self.current_style["italic"] = True
            elif tag == "u":
                self.current_style["underline"] = True
            elif tag == "br":
                self.p = self.text_frame.add_paragraph()
            elif tag == "p":
                self.p = self.text_frame.add_paragraph()

        def handle_endtag(self, tag):
            if tag == "b":
                self.current_style["bold"] = False
            elif tag == "i":
                self.current_style["italic"] = False
            elif tag == "u":
                self.current_style["underline"] = False

        def handle_data(self, data):
            run = self.p.add_run()
            run.text = html.unescape(data)
            self.apply_style(run)

    html_content = html_content or ""
    parser = PPTXHTMLParser(text_frame)
    parser.feed(html_content)
    
        elif alignment == "right":
            p.alignment = PP_ALIGN.RIGHT
        else:
            p.alignment = PP_ALIGN.LEFT
    elif alignment == "right":
        for p in text_frame.paragraphs:
            p.alignment = PP_ALIGN.RIGHT
    else:
        for p in text_frame.paragraphs:
            p.alignment = PP_ALIGN.LEFT


def from_course(val, axis):
    if axis == "y":
        corrected = float(val) + 10.917
        px = corrected / 152.838 * 700
    else:
        px = float(val) / 149.351 * 1150
    return px * 0.01043

def from_look(val):
    return float(val) * 0.01043

def extract_xml_from_zip(zip_path, extract_to):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)

def find_file(root_dir, filename):
    for root, _, files in os.walk(root_dir):
        if filename in files:
            return os.path.join(root, filename)
    return None

def parse_xml(file_path):
    tree = ET.parse(file_path)
    return tree.getroot()

def add_textbox(slide, text, left, top, width, height, style, design_el=None):
    textbox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    text_frame = textbox.text_frame
    text_frame.clear()
    p = text_frame.paragraphs[0]
    run = p.add_run()
    print(f"ALIGN DEBUG → align =", style.get("align"))
    parse_formatted_html(text_frame, text or "", style)
    # Appliquer l'alignement
    alignment = style.get("align") if design_el is None else design_el.attrib.get("align")
    if alignment == "center":
        p.alignment = PP_ALIGN.CENTER
    elif alignment == "right":
        p.alignment = PP_ALIGN.RIGHT
    else:
        p.alignment = PP_ALIGN.LEFT


    font_name = design_el.attrib.get("font") if design_el is not None and "font" in design_el.attrib else style.get("font", "Arial")
    font_size = int(design_el.attrib.get("fontsize")) if design_el is not None and "fontsize" in design_el.attrib else int(style.get("fontsize", 18))
    bold = design_el.attrib.get("bold", "0") == "1" if design_el is not None else style.get("bold", "0") == "1"
    italic = design_el.attrib.get("italic", "0") == "1" if design_el is not None else style.get("italic", "0") == "1"
    underline = design_el.attrib.get("underline", "0") == "1" if design_el is not None else style.get("underline", "0") == "1"
    font_color = design_el.attrib.get("fontcolor") if design_el is not None and "fontcolor" in design_el.attrib else style.get("fontcolor", "#000000")

    font = run.font
    font.name = font_name
    
    # Table de conversion PX → PT
    px_to_pt = {8: 6, 10: 7, 12: 9, 14: 10, 16: 12, 18: 14, 20: 15, 22: 17, 25: 19, 30: 22, 35: 25, 40: 30, 45: 34, 50: 36}
    adjusted_size = px_to_pt.get(font_size, font_size)
    font.size = Pt(adjusted_size)
    
        elif alignment == "right":
            p.alignment = PP_ALIGN.RIGHT
        else:
            p.alignment = PP_ALIGN.LEFT

    font.bold = bold
    font.italic = italic
    font.underline = underline
    try:
        font.color.rgb = RGBColor.from_string(font_color.replace("#", ""))
    except:
        pass

def build_presentation(zip_file):
    with tempfile.TemporaryDirectory() as tmpdir:
        zip_path = os.path.join(tmpdir, "scorm.zip")
        with open(zip_path, "wb") as f:
            f.write(zip_file.read())

        extract_xml_from_zip(zip_path, tmpdir)

        look_path = find_file(tmpdir, "look.xml")
        course_path = find_file(tmpdir, "course.xml")

        if not look_path or not course_path:
            raise FileNotFoundError("look.xml ou course.xml introuvable dans l’archive.")

        look_root = parse_xml(look_path)
        course_root = parse_xml(course_path)

        
        style_map = {}
        for el in look_root.findall(".//text"):
            text_id = el.attrib.get("id")
            design_el = el.find("design")
            if text_id and design_el is not None:
                style_map[text_id] = design_el.attrib


        prs = Presentation()
        prs.slide_width = Inches(12)
        prs.slide_height = Inches(7.3)
        blank_slide_layout = prs.slide_layouts[6]

        for node in course_root.findall(".//node"):
            title_el = node.find(".//title")
            title_text = title_el.text if title_el is not None else ""
            slide = prs.slides.add_slide(blank_slide_layout)
            title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(11), Inches(1))
            title_frame = title_box.text_frame
            title_frame.text = title_text

            for text_el in node.findall(".//text"):
                text_id = text_el.attrib.get("id")
                text_content_el = text_el.find(".//content")
                text_content = text_content_el.text if text_content_el is not None else ""
                design_el = text_el.find(".//design")
                style = style_map.get(text_id, {})

                top = from_course(design_el.attrib["top"], "y") if design_el is not None and "top" in design_el.attrib else from_look(float(style.get("top", "0")))
                left = from_course(design_el.attrib["left"], "x") if design_el is not None and "left" in design_el.attrib else from_look(float(style.get("left", "0")))
                width = from_course(design_el.attrib["width"], "x") if design_el is not None and "width" in design_el.attrib else from_look(float(style.get("width", "140")))
                height = from_course(design_el.attrib["height"], "y") if design_el is not None and "height" in design_el.attrib else from_look(float(style.get("height", "10")))

                st.text(f"Ajout box at → top={top}, left={left}, width={width}, height={height}")
                
                st.text(f"text_id = {text_id} → style = {style}")
                st.text(f"Ajout box at → top={top}, left={left}, width={width}, height={height}")
                add_textbox(slide, text_content, left, top, width, height, style, design_el)

    return prs

st.title("Convertisseur ECMG vers PowerPoint")
uploaded_file = st.file_uploader("Déposez votre archive SCORM (ZIP)", type="zip")

if uploaded_file is not None:
    ppt = build_presentation(uploaded_file)
    output_path = os.path.join(tempfile.gettempdir(), "output.pptx")
    ppt.save(output_path)

    with open(output_path, "rb") as f:
        st.download_button("📥 Télécharger le PowerPoint", f.read(), file_name="module_converti.pptx")

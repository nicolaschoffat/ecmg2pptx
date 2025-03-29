
import zipfile
import os
import xml.etree.ElementTree as ET
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from html.parser import HTMLParser
import re

# === CONFIGURATION ===
ECMG_WIDTH_PX = 1150
ECMG_HEIGHT_PX = 700
PPT_WIDTH_CM = 30.48
PPT_HEIGHT_CM = 18.543

def extract_zip(zip_path, extract_to="/tmp/ecmg_extract"):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)
    return extract_to

def parse_xml(path):
    tree = ET.parse(path)
    return tree.getroot()

def convert_ecmg_to_inches(value, axis="x"):
    px = float(value)
    if axis == "x":
        return (px + 1.299) * (PPT_WIDTH_CM / ECMG_WIDTH_PX) / 2.54
    else:
        return (px + 10.917) * (PPT_HEIGHT_CM / ECMG_HEIGHT_PX) / 2.54

def px_to_pt(px):
    return float(px) * 0.75

class SimpleHTMLParser(HTMLParser):
    def __init__(self, text_frame, default_style):
        super().__init__()
        self.text_frame = text_frame
        self.default_style = default_style
        self.current_run = self.text_frame.paragraphs[0].add_run()
        self.bold = False
        self.italic = False
        self.underline = False

    def handle_starttag(self, tag, attrs):
        if tag in ("b", "strong"):
            self.bold = True
        elif tag in ("i", "em"):
            self.italic = True
        elif tag == "u":
            self.underline = True
        elif tag == "br":
            self.text_frame.add_paragraph()
            self.current_run = self.text_frame.paragraphs[-1].add_run()

    def handle_endtag(self, tag):
        if tag in ("b", "strong"):
            self.bold = False
        elif tag in ("i", "em"):
            self.italic = False
        elif tag == "u":
            self.underline = False

    def handle_data(self, data):
        run = self.text_frame.paragraphs[-1].add_run()
        run.text = data
        run.font.bold = self.bold or self.default_style.get("bold") == "1"
        run.font.italic = self.italic or self.default_style.get("italic") == "1"
        run.font.underline = self.underline or self.default_style.get("underline") == "1"
        run.font.size = Pt(px_to_pt(self.default_style.get("fontsize", "20")))

def get_text_design_position(style, design_el, key, axis="x"):
    if design_el is not None and design_el.attrib.get(key) not in [None, ""]:
        return Inches(convert_ecmg_to_inches(design_el.attrib[key], axis))
    return Inches(convert_ecmg_to_inches(style.get(key, "0"), axis))

def add_textbox(slide, text, left, top, width, height, style, design_el):
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    text_frame.clear()
    if text:
        parser = SimpleHTMLParser(text_frame, style)
        parser.feed(text)
    align = style.get("align", "left")
    if align == "center":
        for para in text_frame.paragraphs:
            para.alignment = PP_ALIGN.CENTER
    elif align == "right":
        for para in text_frame.paragraphs:
            para.alignment = PP_ALIGN.RIGHT

def extract_styles(root, tag="text"):
    return {el.attrib.get("id"): el.find("design").attrib for el in root.findall(tag)}

def extract_text_content(el):
    content_el = el.find("content")
    if content_el is not None and content_el.text:
        return content_el.text
    return ""

def build_presentation(zip_file):
    temp_dir = extract_zip(zip_file)
    course_path = os.path.join(temp_dir, "course.xml")
    look_path = os.path.join(temp_dir, "look.xml")

    course_root = parse_xml(course_path)
    look_root = parse_xml(look_path)

    course_texts = {el.attrib.get("id"): el for el in course_root.findall(".//text")}
    look_styles = extract_styles(look_root)

    prs = Presentation()
    prs.slide_width = Inches(PPT_WIDTH_CM / 2.54)
    prs.slide_height = Inches(PPT_HEIGHT_CM / 2.54)

    for text_id, text_el in course_texts.items():
        design_el = text_el.find("design")
        style = look_styles.get(text_id, {})

        # Override look with course design if available
        if design_el is not None:
            for k in design_el.attrib:
                style[k] = design_el.attrib[k]

        content = extract_text_content(text_el)
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        top = get_text_design_position(style, design_el, "top", "y")
        left = get_text_design_position(style, design_el, "left", "x")
        width = get_text_design_position(style, design_el, "width", "x")
        height = get_text_design_position(style, design_el, "height", "y")

        print(f"text_id = {text_id} → style = {style}")
        print(f"Ajout box at → top={round(top.inches*2.54, 2)} cm, left={round(left.inches*2.54, 2)} cm, width={round(width.inches*2.54, 2)} cm, height={round(height.inches*2.54, 2)} cm")

        add_textbox(slide, content, left, top, width, height, style, design_el)

    output_path = "/mnt/data/ecmg_presentation_final.pptx"
    prs.save(output_path)
    return output_path

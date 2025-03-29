import zipfile
import os
import xml.etree.ElementTree as ET
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import streamlit as st
import tempfile

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

def parse_xml(file_path):
    tree = ET.parse(file_path)
    return tree.getroot()

def add_textbox(slide, text, left, top, width, height, style, design_el=None):
    textbox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    text_frame = textbox.text_frame
    text_frame.clear()
    p = text_frame.paragraphs[0]
    run = p.add_run()
    run.text = text

    font_name = design_el.attrib.get("font") if design_el is not None and "font" in design_el.attrib else style.get("font", "Arial")
    font_size = int(design_el.attrib.get("fontsize")) if design_el is not None and "fontsize" in design_el.attrib else int(style.get("fontsize", 18))
    bold = design_el.attrib.get("bold", "0") == "1" if design_el is not None else style.get("bold", "0") == "1"
    italic = design_el.attrib.get("italic", "0") == "1" if design_el is not None else style.get("italic", "0") == "1"
    underline = design_el.attrib.get("underline", "0") == "1" if design_el is not None else style.get("underline", "0") == "1"
    font_color = design_el.attrib.get("fontcolor") if design_el is not None and "fontcolor" in design_el.attrib else style.get("fontcolor", "#000000")

    font = run.font
    font.name = font_name
    font.size = Pt(font_size)
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

        look_path = os.path.join(tmpdir, "look.xml")
        course_path = os.path.join(tmpdir, "course.xml")

        look_root = parse_xml(look_path)
        course_root = parse_xml(course_path)

        style_map = {}
        for el in look_root.iter():
            if "id" in el.attrib:
                style_map[el.attrib["id"]] = el.attrib

        prs = Presentation()
        blank_slide_layout = prs.slide_layouts[6]

        for node in course_root.findall(".//node"):
            title_el = node.find(".//title")
            title_text = title_el.text if title_el is not None else ""
            slide = prs.slides.add_slide(blank_slide_layout)
            slide.shapes.title.text = title_text

            for text_el in node.findall(".//text"):
                text_id = text_el.attrib.get("id")
                text_content_el = text_el.find(".//content")
                text_content = text_content_el.text if text_content_el is not None else ""
                design_el = text_el.find(".//design")
                style = style_map.get(text_id, {})

                top = from_course(design_el.attrib.get("top", 0), "y") if design_el is not None else from_look(style.get("top", 0))
                left = from_course(design_el.attrib.get("left", 0), "x") if design_el is not None else from_look(style.get("left", 0))
                width = from_course(design_el.attrib.get("width", 140), "x") if design_el is not None else from_look(style.get("width", 140))
                height = from_course(design_el.attrib.get("height", 10), "y") if design_el is not None else from_look(style.get("height", 10))

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

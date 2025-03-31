import streamlit as st
import zipfile
import tempfile
import os
from bs4 import BeautifulSoup
from xml.etree import ElementTree as ET
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from html.parser import HTMLParser

st.set_page_config(page_title="ECMG to PowerPoint Converter")
st.title("📤 Convertisseur ECMG vers PowerPoint")

uploaded_file = st.file_uploader("Upload un module ECMG (zip SCORM)", type="zip")

PX_TO_PT = {
    10: 7, 11: 8, 12: 9, 14: 10, 16: 12, 18: 13,
    20: 15, 22: 17, 24: 18, 26: 20, 28: 21, 30: 22,
    32: 24, 35: 26, 40: 30, 45: 34, 50: 36, 55: 40,
    60: 45, 65: 50, 70: 55, 75: 60, 80: 64
}

class HTMLtoPPTX(HTMLParser):
    def __init__(self, text_frame, style=None):
        super().__init__()
        self.tf = text_frame
        self.p = text_frame.paragraphs[0]
        self.run = self.p.add_run()
        self.style = {"bold": False, "italic": False}
        self.default_size = int(style.get("fontsize", "20")) if style else 20
        self.default_color = style.get("fontcolor", "#000000") if style else "#000000"

    def handle_starttag(self, tag, attrs):
        if tag == "b": self.style["bold"] = True
        if tag == "i": self.style["italic"] = True
        self.run = self.p.add_run()
        self.apply_style(self.run)

    def handle_endtag(self, tag):
        if tag == "b": self.style["bold"] = False
        if tag == "i": self.style["italic"] = False
        self.run = self.p.add_run()
        self.apply_style(self.run)

    def handle_data(self, data):
        self.run.text += data

    def apply_style(self, run):
        run.font.bold = self.style["bold"]
        run.font.italic = self.style["italic"]
        run.font.size = Pt(PX_TO_PT.get(self.default_size, self.default_size))
        hex_color = self.default_color.lstrip("#")
        run.font.color.rgb = RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))

def from_course(val, axis):
    if axis == "y":
        corrected = float(val) + 10.917
        px = corrected / 152.838 * 700
    else:
        px = float(val) / 149.351 * 1150
    return px * 0.01043

def from_look(val):
    return float(val) * 0.01043

if uploaded_file:
    with tempfile.TemporaryDirectory() as tmpdir:
        zip_path = os.path.join(tmpdir, "module.zip")
        with open(zip_path, "wb") as f:
            f.write(uploaded_file.read())
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(tmpdir)

        course_path, look_path, author_path = None, None, None
        for root, dirs, files in os.walk(tmpdir):
            if "course.xml" in files:
                course_path = os.path.join(root, "course.xml")
            if "look.xml" in files:
                look_path = os.path.join(root, "look.xml")
            if "author.xml" in files:
                author_path = os.path.join(root, "author.xml")

        if not course_path or not look_path or not author_path:
            st.error("Fichiers course.xml, look.xml ou author.xml introuvables.")
            st.stop()

        tree = ET.parse(course_path)
        root = tree.getroot()
        nodes = root.findall(".//node")

        look_tree = ET.parse(look_path)
        look_root = look_tree.getroot()
        style_map = {}
        for el in look_root.findall(".//*[@id]"):
            design = el.find("design")
            if design is not None:
                style_map[el.attrib["id"]] = design.attrib
                if "author_id" in el.attrib:
                    style_map[el.attrib["author_id"]] = design.attrib

        prs = Presentation()
        prs.slide_width = Inches(12)
        prs.slide_height = Inches(7.3)

        for node in nodes:
            title_el = node.find("./metadata/title")
            title_text = title_el.text.strip() if title_el is not None else "Sans titre"
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = title_text

            page = node.find(".//page")
            screen = page.find("screen") if page is not None else None
            if not screen:
                continue

            for el in screen.findall("text"):
                content_el = el.find("content")
                if content_el is None or not content_el.text:
                    continue
                text_id = el.attrib.get("id") or el.attrib.get("author_id")
                style = style_map.get(text_id, {})
                design_el = el.find("design")
                top = from_course(design_el.attrib["top"], "y") if design_el is not None and "top" in design_el.attrib else from_look(float(style.get("top", "0")))
                left = from_course(design_el.attrib["left"], "x") if design_el is not None and "left" in design_el.attrib else from_look(float(style.get("left", "0")))
                width = from_course(design_el.attrib["width"], "x") if design_el is not None and "width" in design_el.attrib else from_look(float(style.get("width", "140")))
                height = from_course(design_el.attrib["height"], "y") if design_el is not None and "height" in design_el.attrib else from_look(float(style.get("height", "10")))
                box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
                tf = box.text_frame
                tf.clear()
                tf.word_wrap = True
                if "align" in style and style["align"] == "center":
                    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
                elif style.get("align") == "right":
                    tf.paragraphs[0].alignment = PP_ALIGN.RIGHT
                else:
                    tf.paragraphs[0].alignment = PP_ALIGN.LEFT
                parser = HTMLtoPPTX(tf, style)
                parser.feed(content_el.text)

        output_path = os.path.join(tmpdir, "converted.pptx")
        prs.save(output_path)
        with open(output_path, "rb") as f:
            st.download_button("📥 Télécharger le PowerPoint", data=f, file_name="module_ecmg_converti.pptx")

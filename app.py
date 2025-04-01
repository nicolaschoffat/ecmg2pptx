import streamlit as st
import zipfile
import tempfile
import os
from bs4 import BeautifulSoup
from xml.etree import ElementTree as ET
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from html.parser import HTMLParser

px_to_pt = {
    20: 15, 25: 18, 30: 22, 35: 26, 40: 30, 45: 34, 50: 38
}

st.set_page_config(page_title="ECMG to PowerPoint Converter")
st.title("📤 Convertisseur ECMG vers PowerPoint")

uploaded_file = st.file_uploader("Upload un module ECMG (zip SCORM)", type="zip")

class HTMLtoPPTX(HTMLParser):
    def __init__(self, text_frame, style=None):
        super().__init__()
        self.tf = text_frame
        self.p = text_frame.paragraphs[0]
        self.style = {"bold": False, "italic": False}
        self.default_style = style or {}
        self.run = self.p.add_run()
        self.apply_style()

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag == "b": self.style["bold"] = True
        elif tag == "i": self.style["italic"] = True
        elif tag == "font":
            if "face" in attrs: self.default_style["font"] = attrs["face"]
            if "color" in attrs: self.default_style["fontcolor"] = attrs["color"]
            if "size" in attrs:
                try: self.default_style["fontsize"] = int(attrs["size"])
                except: pass
        self.run = self.p.add_run()
        self.apply_style()

    def handle_endtag(self, tag):
        if tag == "b": self.style["bold"] = False
        elif tag == "i": self.style["italic"] = False
        elif tag == "font":
            self.default_style.pop("font", None)
            self.default_style.pop("fontcolor", None)
            self.default_style.pop("fontsize", None)
        self.run = self.p.add_run()
        self.apply_style()

    def handle_data(self, data):
        self.run.text += data

    def apply_style(self):
        font = self.run.font
        font.bold = self.style["bold"]
        font.italic = self.style["italic"]
        if "font" in self.default_style:
            font.name = self.default_style["font"]
        if "fontcolor" in self.default_style:
            color = self.default_style["fontcolor"].lstrip("#")
            if len(color) == 6:
                try:
                    font.color.rgb = RGBColor.from_string(color.upper())
                except ValueError: pass
        if "fontsize" in self.default_style:
            try:
                px = int(self.default_style["fontsize"])
                pt = px_to_pt.get(px, int(px * 0.75))
                font.size = Pt(pt)
            except: pass

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
    st.info("📦 Traitement du fichier...")

    with tempfile.TemporaryDirectory() as tmpdir:
        zip_path = os.path.join(tmpdir, "module.zip")
        with open(zip_path, "wb") as f:
            f.write(uploaded_file.read())
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
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
            st.error("Fichiers manquants dans le module.")
            st.stop()

        tree = ET.parse(course_path)
        root = tree.getroot()
        nodes = root.findall(".//node")

        prs = Presentation()
        prs.slide_width = Inches(12)
        prs.slide_height = Inches(7.3)

        for i, node in enumerate(nodes):
            title_el = node.find("./metadata/title")
            title_text = title_el.text.strip() if title_el is not None else f"Slide {i+1}"
            st.text(f"➡️ Slide {i+1}: {title_text}")
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = title_text

        output_path = os.path.join(tmpdir, "converted.pptx")
        prs.save(output_path)

        with open(output_path, "rb") as f:
            st.download_button("📥 Télécharger le PowerPoint", data=f, file_name="module_ecmg_converti.pptx")

import streamlit as st
import zipfile
import tempfile
import os
import shutil
from bs4 import BeautifulSoup
from xml.etree import ElementTree as ET
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from html.parser import HTMLParser

st.set_page_config(page_title="ECMG to PowerPoint Converter")
st.title("📤 Convertisseur ECMG vers PowerPoint")

# Upload du fichier zip
uploaded_file = st.file_uploader("Upload un module ECMG (zip SCORM)", type="zip")

# HTML Parser PowerPoint simplifié
class HTMLtoPPTX(HTMLParser):
    def __init__(self, text_frame):
        super().__init__()
        self.tf = text_frame
        self.p = text_frame.paragraphs[0]
        self.run = self.p.add_run()
        self.style = {"bold": False, "italic": False}

    def handle_starttag(self, tag, attrs):
        if tag == "b":
            self.style["bold"] = True
        if tag == "i":
            self.style["italic"] = True
        self.run = self.p.add_run()
        self.apply_style()

    def handle_endtag(self, tag):
        if tag == "b":
            self.style["bold"] = False
        if tag == "i":
            self.style["italic"] = False
        self.run = self.p.add_run()
        self.apply_style()

    def handle_data(self, data):
        self.run.text += data

    def apply_style(self):
        self.run.font.bold = self.style["bold"]
        self.run.font.italic = self.style["italic"]

# Conversion pixels ECMG → inches
def to_inches(px):
    try:
        return float(px) * 0.0264
    except:
        return 1.0

if uploaded_file:
    with tempfile.TemporaryDirectory() as tmpdir:
        zip_path = os.path.join(tmpdir, "module.zip")
        with open(zip_path, "wb") as f:
            f.write(uploaded_file.read())
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(tmpdir)

        # Rechercher les fichiers XML
        course_path, look_path = None, None
        for root, dirs, files in os.walk(tmpdir):
            if "course.xml" in files:
                course_path = os.path.join(root, "course.xml")
            if "look.xml" in files:
                look_path = os.path.join(root, "look.xml")

        if not course_path or not look_path:
            st.error("Fichiers course.xml ou look.xml introuvables dans le zip")
            st.stop()

        # Parse les fichiers
        tree = ET.parse(course_path)
        root = tree.getroot()
        nodes = root.findall(".//node")

        look_tree = ET.parse(look_path)
        look_root = look_tree.getroot()
        style_map = {
            el.attrib["id"]: el.attrib
            for el in look_root.findall(".//screen/*[@id]")
        }
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

            y = 1.5
            for el in screen.findall("text"):
                content_el = el.find("content")
                if content_el is None or not content_el.text:
                    continue

                style = style_map.get(el.attrib.get("id", ""), {})
                height = to_inches(style.get("height", 10))
                box = slide.shapes.add_textbox(Inches(1), Inches(y), Inches(10), Inches(height))
                tf = box.text_frame
                tf.clear()
                parser = HTMLtoPPTX(tf)
                parser.feed(content_el.text)
                y += height + 0.2

        # Export du fichier
        output_path = os.path.join(tmpdir, "converted.pptx")
        prs.save(output_path)

        with open(output_path, "rb") as f:
            st.download_button(
                label="📄 Télécharger le PowerPoint",
                data=f,
                file_name="module_ecmg_converti.pptx"
            )

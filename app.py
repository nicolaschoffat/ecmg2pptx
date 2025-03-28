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

class HTMLtoPPTX(HTMLParser):
    def __init__(self, text_frame):
        super().__init__()
        self.tf = text_frame
        self.p = text_frame.paragraphs[0]
        self.run = self.p.add_run()
        self.style = {"bold": False, "italic": False}

    def handle_starttag(self, tag, attrs):
        if tag == "b": self.style["bold"] = True
        if tag == "i": self.style["italic"] = True
        self.run = self.p.add_run()
        self.apply_style()

    def handle_endtag(self, tag):
        if tag == "b": self.style["bold"] = False
        if tag == "i": self.style["italic"] = False
        self.run = self.p.add_run()
        self.apply_style()

    def handle_data(self, data):
        self.run.text += data

    def apply_style(self):
        self.run.font.bold = self.style["bold"]
        self.run.font.italic = self.style["italic"]

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

        course_path, look_path = None, None
        for root, dirs, files in os.walk(tmpdir):
            if "course.xml" in files:
                course_path = os.path.join(root, "course.xml")
            if "look.xml" in files:
                look_path = os.path.join(root, "look.xml")

        if not course_path or not look_path:
            st.error("Fichiers course.xml ou look.xml introuvables.")
            st.stop()

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

            video_file = None
            for content_el in screen.findall(".//content"):
                if "file" in content_el.attrib and content_el.attrib["file"].endswith(".mp4"):
                    video_file = content_el.attrib["file"]
                    break
            if video_file:
                box = slide.shapes.add_textbox(Inches(3), Inches(3), Inches(6), Inches(1))
                tf = box.text_frame
                tf.text = f"🎥 Vidéo : {video_file} à intégrer"
                tf.paragraphs[0].alignment = PP_ALIGN.CENTER
                continue

            # Gestion des cartes (page_cards)
            cards_blocks = screen.findall(".//cards")
            if cards_blocks:
                notes = slide.notes_slide.notes_text_frame
                feedback_texts = []
                for cards in cards_blocks:
                    for card in cards.findall("card"):
                        head = card.find("head").text.strip() if card.find("head") is not None else ""
                        face_html = card.find("face").text if card.find("face") is not None else ""
                        back_html = card.find("back").text if card.find("back") is not None else ""
                        face = BeautifulSoup(face_html or "", "html.parser").get_text(separator=" ").strip()
                        back = BeautifulSoup(back_html or "", "html.parser").get_text(separator=" ").strip()
                        feedback_texts.append(f"Carte : {head}\nFace : {face}\nBack : {back}")
                if feedback_texts:
                    notes.clear()
                    notes.text = "\n---\n".join(feedback_texts)
                continue

            elfe = screen.find("elfe")
            if elfe is not None and elfe.find("content") is not None and elfe.find("content").attrib.get("type") == "MCQText":
                items = elfe.find("content/items")
                question_el = screen.find("question")
                if question_el is not None:
                    question_text = BeautifulSoup(question_el.find("content").text, "html.parser").get_text()
                    box = slide.shapes.add_textbox(Inches(1), Inches(y), Inches(10), Inches(1))
                    box.text_frame.text = f"❓ {question_text}"
                    y += 1.0
                for item in items.findall("item"):
                    score = item.attrib.get("score", "0")
                    label = "✅" if score == "100" else "⬜"
                    box = slide.shapes.add_textbox(Inches(1.2), Inches(y), Inches(9.5), Inches(0.5))
                    box.text_frame.text = f"{label} {item.text.strip()}"
                    y += 0.5
                feedbacks = page.findall(".//feedbacks/correc/fb/screen/feedback")
                notes = slide.notes_slide.notes_text_frame
                feedback_texts = []
                for fb in feedbacks:
                    fb_content = fb.find("content")
                    if fb_content is not None and fb_content.text:
                        soup = BeautifulSoup(fb_content.text, "html.parser")
                        feedback_texts.append(soup.get_text(separator="\n"))
                if feedback_texts:
                    notes.clear()
                    notes.text = "\n---\n".join(feedback_texts)
                continue

            text_blocks = []
            for el in screen.findall("text"):
                content_el = el.find("content")
                if content_el is None or not content_el.text:
                    continue
                style = style_map.get(el.attrib.get("id", ""), {})
                top = 1000.0
                design_el = el.find("design")
                if design_el is not None and "top" in design_el.attrib:
                    top = float(design_el.attrib["top"])
                elif "top" in style:
                    top = float(style["top"])
                text_blocks.append((top, el, content_el.text, style))

            text_blocks.sort(key=lambda x: x[0])

            for _, el, text, style in text_blocks:
                height = to_inches(style.get("height", 10))
                box = slide.shapes.add_textbox(Inches(1), Inches(y), Inches(10), Inches(height))
                tf = box.text_frame
                tf.clear()
                parser = HTMLtoPPTX(tf)
                parser.feed(text)
                y += height + 0.2

        output_path = os.path.join(tmpdir, "converted.pptx")
        prs.save(output_path)

        with open(output_path, "rb") as f:
            st.download_button("📥 Télécharger le PowerPoint", data=f, file_name="module_ecmg_converti.pptx")

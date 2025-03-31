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
    20: 15,
    25: 18,
    30: 22,
    35: 26,
    40: 30,
    45: 34,
    50: 38
}

st.set_page_config(page_title="ECMG to PowerPoint Converter")
st.title("\U0001F4E4 Convertisseur ECMG vers PowerPoint")

uploaded_file = st.file_uploader("Upload un module ECMG (zip SCORM)", type="zip")

from html.parser import HTMLParser
from pptx.dml.color import RGBColor
from pptx.util import Pt

px_to_pt = {
    20: 15,
    25: 18,
    30: 22,
    35: 26,
    40: 30,
    45: 34,
    50: 38
}

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

        if tag == "b":
            self.style["bold"] = True
        elif tag == "i":
            self.style["italic"] = True
        elif tag == "font":
            if "face" in attrs:
                self.default_style["font"] = attrs["face"]
            if "color" in attrs:
                self.default_style["fontcolor"] = attrs["color"]
            if "size" in attrs:
                try:
                    px = int(attrs["size"])
                    self.default_style["fontsize"] = px
                except:
                    pass

        self.run = self.p.add_run()
        self.apply_style()

    def handle_endtag(self, tag):
        if tag == "b":
            self.style["bold"] = False
        elif tag == "i":
            self.style["italic"] = False
        elif tag == "font":
            # On efface le style de <font>
            self.default_style.pop("font", None)
            self.default_style.pop("fontcolor", None)
            self.default_style.pop("fontsize", None)
        self.run = self.p.add_run()
        self.apply_style()

    def handle_data(self, data):
        self.run.text += data

    def apply_style(self):
        font = self.run.font
        font.bold = self.style.get("bold", False)
        font.italic = self.style.get("italic", False)

        if "font" in self.default_style:
            font.name = self.default_style["font"]
        if "fontcolor" in self.default_style:
            color = self.default_style["fontcolor"].lstrip("#")
            if len(color) == 6:
                try:
                    font.color.rgb = RGBColor.from_string(color.upper())
                except ValueError:
                    pass
        if "fontsize" in self.default_style:
            try:
                px = int(self.default_style["fontsize"])
                pt = px_to_pt.get(px, int(px * 0.75))
                font.size = Pt(pt)
            except:
                pass

def from_course(val, axis):
    if axis == "y":
        corrected = float(val) + 10.917
        px = corrected / 152.838 * 700
    else:
        px = float(val) / 149.351 * 1150
    return px * 0.01043

def from_look(val):
    return float(val) * 0.01043

def to_inches(px):
    try:
        return float(str(px).strip()) * 0.0104
    except:
        return 1.0

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

        author_tree = ET.parse(author_path)
        author_root = author_tree.getroot()
        author_map = {
            el.attrib.get("id"): el.findtext("description")
            for el in author_root.findall(".//item")
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
                tf.text = f" Vidéo : {video_file} à intégrer"
                tf.paragraphs[0].alignment = PP_ALIGN.CENTER
                continue

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

            sound_blocks = screen.findall(".//sound")
            if sound_blocks:
                notes = slide.notes_slide.notes_text_frame
                audio_notes = []
                for snd in sound_blocks:
                    author_id = snd.attrib.get("author_id")
                    content = snd.find("content")
                    filename = content.attrib.get("file") if content is not None else None
                    audio_text = author_map.get(author_id)
                    if filename:
                        audio_notes.append(f"Audio : {filename}\nTexte lu : {audio_text or '[non trouvé]'}")
                if audio_notes:
                    notes.text += "\n\n" + "\n---\n".join(audio_notes)

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
                    notes.text += "\n---\n" + "\n---\n".join(feedback_texts)
                continue

            for el in screen.findall("text"):
                content_el = el.find("content")
                if content_el is None or not content_el.text:
                    continue
                text_id = el.attrib.get("id") or el.attrib.get("author_id")
                style = style_map.get(text_id, {})
                design_el = el.find("design")

                top = from_course(design_el.attrib.get("top", 0), "y") if design_el is not None else from_look(style.get("top", 0))
                left = from_course(design_el.attrib.get("left", 0), "x") if design_el is not None else from_look(style.get("left", 0))
                width = from_course(design_el.attrib.get("width", 140), "x") if design_el is not None else from_look(style.get("width", 140))
                height = from_course(design_el.attrib.get("height", 10), "y") if design_el is not None else from_look(style.get("height", 10))

                st.text(f"Ajout box at → top={top}, left={left}, width={width}, height={height}")
                box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
                tf = box.text_frame
                tf.clear()
                tf.word_wrap = True

                alignment = style.get("align", "").lower()
                if alignment == "center":
                    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
                elif alignment == "right":
                    tf.paragraphs[0].alignment = PP_ALIGN.RIGHT
                else:
                    tf.paragraphs[0].alignment = PP_ALIGN.LEFT

                if "valign" in style:
                    valign = style.get("valign", "").lower()
                    if valign == "middle":
                        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                    elif valign == "bottom":
                        tf.vertical_anchor = MSO_ANCHOR.BOTTOM
                    else:
                        tf.vertical_anchor = MSO_ANCHOR.TOP

                parser = HTMLtoPPTX(tf, style)
                parser.feed(content_el.text)

        output_path = os.path.join(tmpdir, "converted.pptx")
        prs.save(output_path)

        with open(output_path, "rb") as f:
            st.download_button("📥 Télécharger le PowerPoint", data=f, file_name="module_ecmg_converti.pptx")

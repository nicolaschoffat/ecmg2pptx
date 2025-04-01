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
from PIL import Image

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
                try:
                    font.color.rgb = RGBColor.from_string(color.upper())
                except ValueError:
                    pass
        if "fontsize" in self.default_style:
            try:
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

# Le reste du code sera complété après ajout du style de titre

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

        for i, node in enumerate(nodes):
            title_el = node.find("./metadata/title")
            title_text = title_el.text.strip() if title_el is not None else "Sans titre"
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            title_style = style_map.get("titre_activite")
            # 📐 Repositionnement/Redimensionnement de la zone de titre
            try:
                left = from_look(float(title_style.get("left", 0)))
                width = from_look(float(title_style.get("width", 800)))
                height = from_look(float(title_style.get("height", 50)))
                title_shape = slide.shapes.title
                title_shape.left = Inches(left)
                title_shape.top = Inches(top)
                title_shape.width = Inches(width)
                title_shape.height = Inches(height)
            except Exception as e:
                st.warning(f"❗ Erreur redimension titre: {e}")
            if title_style:
                title_shape = slide.shapes.title
                tf = title_shape.text_frame
                p = tf.paragraphs[0]
                run = p.add_run()
                run.text = title_text
                font = run.font
                font.name = title_style.get("font", "Tahoma")
                try:
                    font.size = Pt(px_to_pt.get(fontsize, int(fontsize * 0.75)))
                except:
                    font.size = Pt(16.5)
                font.bold = title_style.get("bold", "0") == "1"
                font.italic = title_style.get("italic", "0") == "1"
                color = title_style.get("fontcolor", "#000000").lstrip("#")
                if len(color) == 6:
                    try:
                    except ValueError:
                        pass
                align = title_style.get("align", "left").lower()
                if align == "center":
                    p.alignment = PP_ALIGN.CENTER
                elif align == "right":
                    p.alignment = PP_ALIGN.RIGHT
                else:
                    p.alignment = PP_ALIGN.LEFT
            page = node.find(".//page")
            screen = page.find("screen") if page is not None else None
            if not screen:
                continue
            y = 1.5
            # 🎥 Vidéo (si présente)
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
            # 🔊 Sons (audio)
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
            # ❓ QCM (MCQText)
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
            # 🖼️ Images dans le screen
            for img_el in screen.findall("image"):
                content = img_el.find("content")
                if content is None or not content.attrib.get("file"):
                    continue
                img_file = content.attrib["file"]
                image_id = img_el.attrib.get("id") or img_el.attrib.get("author_id")
                style = style_map.get(image_id, {})
                design_el = img_el.find("design")
        
                def has_position_attrs(d):
                    return (
                        d is not None and any(
                            attr in d.attrib and float(d.attrib[attr]) > 0
                            for attr in ["top", "left", "width", "height"]
                        )
                    )

                if has_position_attrs(design_el):
                    top_px = float(design_el.attrib.get("top", 0))
                    left_px = float(design_el.attrib.get("left", 0))
                    width_px = float(design_el.attrib.get("width", 200))
                    height_px = float(design_el.attrib.get("height", 200))
                    top = from_course(top_px, "y")
                    left = from_course(left_px, "x")
                    width = from_course(width_px, "x")
                    height = from_course(height_px, "y")
                else:
                    top_px = float(style.get("top", 0))
                    left_px = float(style.get("left", 0))
                    width_px = float(style.get("width", 200))
                    height_px = float(style.get("height", 200))
                    top = from_look(top_px)
                    left = from_look(left_px)
                    width = from_look(width_px)
                    height = from_look(height_px)

                image_dir = os.path.dirname(course_path)
                image_dir = os.path.dirname(course_path)
                image_path = os.path.join(image_dir, os.path.basename(img_file))
                if not os.path.exists(image_path):
                    image_dir = os.path.dirname(look_path)
                    image_path = os.path.join(image_dir, os.path.basename(img_file))
                st.text(f'🔍 Image path testé : {image_path}')
                if os.path.exists(image_path):
                    try:
                    with Image.open(image_path) as im:
                        img_width_px, img_height_px = im.size
                    # Conversion ECMG en pixels
                    target_width_px = width * 96  # 1 inch = 96 px
                    target_height_px = height * 96
                    # Calcul des échelles
                    scale_w = target_width_px / img_width_px
                    scale_h = target_height_px / img_height_px
                    scale = min(scale_w, scale_h)
                    final_width = img_width_px * scale
                    final_height = img_height_px * scale
                    slide.shapes.add_picture(
                        image_path,
                        Inches(left),
                        Inches(top),
                        width=Inches(final_width / 96),
                        height=Inches(final_height / 96)
                    )
                except Exception as e:
                    st.warning(f"⚠️ Erreur ajout image avec ratio {img_file} : {e}")
                        slide.shapes.add_picture(image_path, Inches(left), Inches(top), width=Inches(width), height=Inches(height))
                    except Exception as e:
                        st.warning(f"⚠️ Erreur ajout image {img_file} : {e}")
                else:
                    st.warning(f"❌ Image non trouvée : {img_file}")
            for el in screen.findall("text"):
                content_el = el.find("content")
                if content_el is None or not content_el.text:
                    continue
                text_id = el.attrib.get("id") or el.attrib.get("author_id")
                style = style_map.get(text_id, {})
                design_el = el.find("design")
                def has_position_attrs(d):
                    return (
                        d is not None and any(
                            attr in d.attrib and float(d.attrib[attr]) > 0
                            for attr in ["top", "left", "width", "height"]
                        )
                    )
                if has_position_attrs(design_el):
                    source = "course.xml"
                    top_px = float(design_el.attrib.get("top", 0))
                    left_px = float(design_el.attrib.get("left", 0))
                    width_px = float(design_el.attrib.get("width", 140))
                    height_px = float(design_el.attrib.get("height", 10))
                    top = from_course(top_px, "y")
                    left = from_course(left_px, "x")
                    width = from_course(width_px, "x")
                    height = from_course(height_px, "y")
                else:
                    source = "look.xml"
                    top_px = float(style.get("top", 0))
                    left_px = float(style.get("left", 0))
                    width_px = float(style.get("width", 140))
                    height_px = float(style.get("height", 10))
                    top = from_look(top_px)
                    left = from_look(left_px)
                    width = from_look(width_px)
                    height = from_look(height_px)
                    f"[{source}] Texte ID='{text_id}' → px: (l={left_px}, t={top_px}, w={width_px}, h={height_px}) | pouces: (l={left:.2f}, t={top:.2f}, w={width:.2f}, h={height:.2f})"
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
            st.download_button("📅 Télécharger le PowerPoint", data=f, file_name="module_ecmg_converti.pptx")

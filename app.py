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

for i, node in enumerate(nodes):
            title_el = node.find("./metadata/title")
            title_text = title_el.text.strip() if title_el is not None else "Sans titre"
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = title_text
            page = node.find(".//page")
            screen = page.find("screen") if page is not None else None
            if not screen:
                continue
            y = 1.5

            # ➕ Zone de texte stylée depuis titre_activite de look.xml
            title_style = style_map.get("titre_activite")
            if title_style:
                top = from_look(float(title_style.get("top", 20)))
                left = from_look(float(title_style.get("left", 10)))
                width = from_look(float(title_style.get("width", 800)))
                height = from_look(float(title_style.get("height", 50)))
                box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
                tf = box.text_frame
                tf.clear()
                tf.word_wrap = True
                p = tf.paragraphs[0]
                run = p.add_run()
                run.text = title_text
                font = run.font
                font.name = title_style.get("font", "Tahoma")
                try:
                    fontsize = int(title_style.get("fontsize", 20))
                    pt = px_to_pt.get(fontsize, int(fontsize * 0.75))
                    font.size = Pt(pt)
                except:
                    pass
                font.bold = title_style.get("bold", "0") == "1"
                font.italic = title_style.get("italic", "0") == "1"
                color = title_style.get("fontcolor", "#000000").lstrip("#")
                if len(color) == 6:
                    font.color.rgb = RGBColor.from_string(color.upper())
                align = title_style.get("align", "").lower()
                if align == "center":
                    p.alignment = PP_ALIGN.CENTER
                elif align == "right":
                    p.alignment = PP_ALIGN.RIGHT
                else:
                    p.alignment = PP_ALIGN.LEFT

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
                continue

            # 🃏 Cartes (cards)
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
                continue

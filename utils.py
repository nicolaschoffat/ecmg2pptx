from pptx import Presentation
from pptx.util import Inches
import xml.etree.ElementTree as ET
import os
from bs4 import BeautifulSoup

def px_to_inches(px):
    return float(px) / 96.0  # 1 inch = 96 pixels

def extract_title(node):
    meta = node.find("metadata")
    if meta is not None:
        title = meta.find("title")
        if title is not None and title.text:
            return title.text.strip()
    return "Sans titre"

def parse_xml_to_slides(xml_file, media_dir):
    tree = ET.parse(xml_file)
    root = tree.getroot()
    course_structure = []

    for node in root.findall(".//node"):
        subnodes = node.findall(".//node")
        if subnodes:
            # SECTION parent (s’il a des sous-NODEs)
            section_title = extract_title(node)
            section = {"section": section_title, "slides": []}
            for sub in subnodes:
                section["slides"].append(extract_slide_data(sub, media_dir))
            course_structure.append(section)
        else:
            # Slide simple (pas de hiérarchie)
            slide = extract_slide_data(node, media_dir)
            course_structure.append({"section": None, "slides": [slide]})
    
    return course_structure

def extract_slide_data(node, media_dir):
    slide_data = {'title': extract_title(node), 'texts': [], 'images': []}
    screen = node.find(".//screen")
    if screen is None:
        return slide_data

    # Images
    for img in screen.findall("image"):
        content = img.find("content")
        if content is not None and "file" in content.attrib:
            file_path = content.attrib['file'].replace("@/", "")
            design = img.find("design")
            if design is not None:
                slide_data['images'].append({
                    "file": os.path.join(media_dir, file_path),
                    "left": float(design.attrib.get("left", 0)),
                    "top": float(design.attrib.get("top", 0)),
                    "width": float(design.attrib.get("width", 1)),
                    "height": float(design.attrib.get("height", 1))
                })

    # Textes
    for txt in screen.findall("text"):
        content = txt.find("content")
        if content is not None:
            raw_html = content.text or ''
            soup = BeautifulSoup(raw_html, "html.parser")
            text = soup.get_text().strip()
            design = txt.find("design")
            if design is not None:
                slide_data['texts'].append({
                    "text": text,
                    "left": float(design.attrib.get("left", 0)),
                    "top": float(design.attrib.get("top", 0)),
                    "width": float(design.attrib.get("width", 5)),
                    "height": float(design.attrib.get("height", 1))
                })

    return slide_data

def generate_pptx(structure, output_path):
    prs = Presentation()
    prs.slide_width = Inches(11.98)
    prs.slide_height = Inches(7.29)
    blank_layout = prs.slide_layouts[6]
    section_slide_map = []

    for group in structure:
        # Enregistrer l'index du premier slide de la section (si présente)
        section_start_idx = len(prs.slides)
        for slide_data in group['slides']:
            slide = prs.slides.add_slide(blank_layout)
            slide.name = slide_data['title'] or "Slide"

            for img in slide_data['images']:
                if os.path.exists(img['file']):
                    slide.shapes.add_picture(
                        img['file'],
                        Inches(px_to_inches(img['left'])),
                        Inches(px_to_inches(img['top'])),
                        width=Inches(px_to_inches(img['width'])),
                        height=Inches(px_to_inches(img['height']))
                    )

            for txt in slide_data['texts']:
                textbox = slide.shapes.add_textbox(
                    Inches(px_to_inches(txt['left'])),
                    Inches(px_to_inches(txt['top'])),
                    Inches(px_to_inches(txt['width'])),
                    Inches(px_to_inches(txt['height']))
                )
                tf = textbox.text_frame
                tf.text = txt['text']

        # Créer la section (si label existant)
        if group['section']:
            section_slide_map.append((section_start_idx, group['section']))

    # Injecter les sections via _prs._element (hack)
    # ATTENTION : python-pptx ne supporte pas officiellement add_section()
    for idx, name in section_slide_map:
        prs.slides._sldIdLst.insert(idx, prs.slides._sldIdLst[idx])  # duplicate ID
        sec_tag = prs.slides._sldIdLst[idx]
        sec_tag.set("name", name)

    prs.save(output_path)

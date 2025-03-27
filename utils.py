from pptx import Presentation
from pptx.util import Inches
import xml.etree.ElementTree as ET
import os
from bs4 import BeautifulSoup

# 📏 Conversion pixels ➜ pouces (1 pouce = 96 px)
def px_to_inches(px):
    return float(px) / 96.0

def parse_xml_to_slides(xml_file, media_dir):
    tree = ET.parse(xml_file)
    root = tree.getroot()
    slides = []

    for node in root.findall(".//node"):
        for subnode in node.findall(".//node"):
            slide_content = {'texts': [], 'images': []}
            screen = subnode.find(".//screen")
            if screen is None:
                continue

            # 🖼️ Extraction des images
            for img in screen.findall("image"):
                content = img.find("content")
                if content is not None and "file" in content.attrib:
                    file_path = content.attrib['file'].replace("@/", "")
                    design = img.find("design")
                    if design is not None:
                        slide_content['images'].append({
                            "file": os.path.join(media_dir, file_path),
                            "left": float(design.attrib.get("left", 0)),
                            "top": float(design.attrib.get("top", 0)),
                            "width": float(design.attrib.get("width", 1)),
                            "height": float(design.attrib.get("height", 1))
                        })

            # 📝 Extraction des textes
            for txt in screen.findall("text"):
                content = txt.find("content")
                if content is not None:
                    raw_html = content.text or ''
                    soup = BeautifulSoup(raw_html, "html.parser")
                    text = soup.get_text().strip()
                    design = txt.find("design")
                    if design is not None:
                        slide_content['texts'].append({
                            "text": text,
                            "left": float(design.attrib.get("left", 0)),
                            "top": float(design.attrib.get("top", 0)),
                            "width": float(design.attrib.get("width", 5)),
                            "height": float(design.attrib.get("height", 1))
                        })

            slides.append(slide_content)

    return slides

# 🎯 Génération PowerPoint avec slide 1150x700px (11.98in x 7.29in)
def generate_pptx(slides, output_path):
    prs = Presentation()
    prs.slide_width = Inches(11.98)
    prs.slide_height = Inches(7.29)

    blank_slide_layout = prs.slide_layouts[6]

    for slide_data in slides:
        slide = prs.slides.add_slide(blank_slide_layout)

        # 🎯 Images
        for img in slide_data['images']:
            if os.path.exists(img['file']):
                slide.shapes.add_picture(
                    img['file'],
                    left=Inches(px_to_inches(img['left'])),
                    top=Inches(px_to_inches(img['top'])),
                    width=Inches(px_to_inches(img['width'])),
                    height=Inches(px_to_inches(img['height']))
                )

        # 🎯 Textes
        for txt in slide_data['texts']:
            textbox = slide.shapes.add_textbox(
                Inches(px_to_inches(txt['left'])),
                Inches(px_to_inches(txt['top'])),
                Inches(px_to_inches(txt['width'])),
                Inches(px_to_inches(txt['height']))
            )
            tf = textbox.text_frame
            tf.text = txt['text']

    prs.save(output_path)

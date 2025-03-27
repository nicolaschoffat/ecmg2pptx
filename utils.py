from pptx import Presentation
from pptx.util import Inches
import xml.etree.ElementTree as ET
import os
from bs4 import BeautifulSoup

# Conversion pixels ➜ pouces (1150px = 10 po, 700px = 6.1 po)
def px_to_inches(px, axis='x'):
    return float(px) * (10 / 1150) if axis == 'x' else float(px) * (6.1 / 700)

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

            for img in screen.findall("image"):
                content = img.find("content")
                if content is not None and "file" in content.attrib:
                    file_path = content.attrib['file'].replace("@/", "")
                    design = img.find("design")
                    if design is not None:
                        slide_content['images'].append({
                            "file": os.path.join(media_dir, file_path),
                            "left": px_to_inches(design.attrib.get("left", 0), 'x'),
                            "top": px_to_inches(design.attrib.get("top", 0), 'y'),
                            "width": px_to_inches(design.attrib.get("width", 1), 'x'),
                            "height": px_to_inches(design.attrib.get("height", 1), 'y')
                        })

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
                            "left": px_to_inches(design.attrib.get("left", 0), 'x'),
                            "top": px_to_inches(design.attrib.get("top", 0), 'y'),
                            "width": px_to_inches(design.attrib.get("width", 5), 'x'),
                            "height": px_to_inches(design.attrib.get("height", 1), 'y')
                        })

            slides.append(slide_content)

    return slides

def generate_pptx(slides, output_path):
    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6]

    for slide_data in slides:
        slide = prs.slides.add_slide(blank_slide_layout)

        for img in slide_data['images']:
            if os.path.exists(img['file']):
                slide.shapes.add_picture(
                    img['file'],
                    Inches(img['left']),
                    Inches(img['top']),
                    width=Inches(img['width']),
                    height=Inches(img['height'])
                )

        for txt in slide_data['texts']:
            textbox = slide.shapes.add_textbox(
                Inches(txt['left']),
                Inches(txt['top']),
                Inches(txt['width']),
                Inches(txt['height'])
            )
            tf = textbox.text_frame
            tf.text = txt['text']

    prs.save(output_path)

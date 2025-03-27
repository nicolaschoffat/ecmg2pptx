import streamlit as st
import xml.etree.ElementTree as ET
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import os
from io import BytesIO
import zipfile

# === Constants ===
CANVAS_WIDTH = 1150
CANVAS_HEIGHT = 700
PPT_WIDTH = Inches(13)
PPT_HEIGHT = Inches(7.91)

def percent_to_inches(val, total):
    return Inches((float(val) / total) * (PPT_WIDTH.inches if total == CANVAS_WIDTH else PPT_HEIGHT.inches))

# === Main App ===
st.set_page_config(page_title="ECMG ➜ PPT Slide 1 Converter", layout="centered")
st.title("🎓 ECMG ➜ PowerPoint Slide Converter")

uploaded_zip = st.file_uploader("📦 Upload ECMG zip (course.xml + media)", type="zip")

if uploaded_zip:
    with zipfile.ZipFile(uploaded_zip) as zip_ref:
        zip_ref.extractall("temp_ecmg")  # Safe temp dir
        st.success("✅ Archive extracted!")

        xml_path = os.path.join("temp_ecmg", "course.xml")
        if not os.path.exists(xml_path):
            st.error("❌ course.xml not found.")
        else:
            # === Parse XML
            tree = ET.parse(xml_path)
            root = tree.getroot()
            screen = root.find(".//screen")

            prs = Presentation()
            prs.slide_width = PPT_WIDTH
            prs.slide_height = PPT_HEIGHT
            slide = prs.slides.add_slide(prs.slide_layouts[6])

            # === Images
            for image in screen.findall("image"):
                design = image.find("design")
                content = image.find("content")
                src = content.attrib["file"].replace("@/", "").strip()

                left = percent_to_inches(design.attrib["left"], CANVAS_WIDTH)
                top = percent_to_inches(design.attrib["top"], CANVAS_HEIGHT)
                width = percent_to_inches(design.attrib["width"], CANVAS_WIDTH)
                height = percent_to_inches(design.attrib["height"], CANVAS_HEIGHT)

                img_path = os.path.join("temp_ecmg", src)
                if os.path.exists(img_path):
                    slide.shapes.add_picture(img_path, left, top, width=width, height=height)
                else:
                    st.warning(f"⚠️ Image not found: {src}")

            # === Text
            for text in screen.findall("text"):
                design = text.find("design")
                content = text.find("content")

                left = percent_to_inches(design.attrib["left"], CANVAS_WIDTH)
                top = percent_to_inches(design.attrib["top"], CANVAS_HEIGHT)
                width = percent_to_inches(design.attrib["width"], CANVAS_WIDTH)
                height = percent_to_inches(design.attrib["height"], CANVAS_HEIGHT)

                textbox = slide.shapes.add_textbox(left, top, width, height)
                tf = textbox.text_frame
                tf.clear()

                p = tf.paragraphs[0]
                p.text = "DÉONTOLOGIE - LE DÉCRET DANS SES GRANDES LIGNES ET L'ARTICLE 2"
                run = p.runs[0]
                run.font.size = Pt(24)
                run.font.bold = True
                run.font.name = "Tahoma"
                run.font.color.rgb = RGBColor(0x13, 0xAB, 0xB5)

            # === Output to Memory
            output = BytesIO()
            prs.save(output)
            st.success("🎉 PowerPoint slide created!")

            st.download_button(
                label="📥 Download slide1.pptx",
                data=output.getvalue(),
                file_name="slide1.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

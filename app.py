import streamlit as st
import xml.etree.ElementTree as ET
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import os
from io import BytesIO
import zipfile

# Constants
CANVAS_WIDTH = 1150
CANVAS_HEIGHT = 700
PPT_WIDTH = Inches(13)
PPT_HEIGHT = Inches(7.91)

def percent_of(value, base):
    return float(value) / 100.0 * base

def scale_to_ppt(px, total_px, ppt_total):
    return Inches((px / total_px) * ppt_total.inches)

def load_template_dimensions(look_root, screen_type='author_page_intro'):
    screen_node = look_root.find(f".//screen[@id='{screen_type}']")
    if screen_node is not None:
        width = float(screen_node.attrib['width'])
        height = float(screen_node.attrib['height'])
        return width, height
    return 770, 458  # fallback default

# UI
st.set_page_config(page_title="ECMG ➜ PPT Slide 1 Converter", layout="centered")
st.title("🎓 ECMG ➜ PowerPoint Slide Converter")

uploaded_zip = st.file_uploader("📦 Upload ZIP (course.xml + look.xml + media)", type="zip")

if uploaded_zip:
    with zipfile.ZipFile(uploaded_zip) as zip_ref:
        zip_ref.extractall("temp_ecmg")
        st.success("✅ Archive extracted")

        # File checks
        course_path = os.path.join("temp_ecmg", "course.xml")
        look_path = os.path.join("temp_ecmg", "look.xml")

        if not os.path.exists(course_path) or not os.path.exists(look_path):
            st.error("❌ Missing course.xml or look.xml in archive.")
        else:
            # Load look.xml
            look_tree = ET.parse(look_path)
            look_root = look_tree.getroot()

            # Load course.xml
            course_tree = ET.parse(course_path)
            course_root = course_tree.getroot()
            screen = course_root.find(".//screen")
            screen_id = screen.attrib.get('id', 'author_page_intro')

            template_width, template_height = load_template_dimensions(look_root, screen_id)

            # Prepare PPT
            prs = Presentation()
            prs.slide_width = PPT_WIDTH
            prs.slide_height = PPT_HEIGHT
            slide = prs.slides.add_slide(prs.slide_layouts[6])

            # Add images
            for image in screen.findall("image"):
                design = image.find("design")
                content = image.find("content")
                src = content.attrib["file"].replace("@/", "").strip()
                img_path = os.path.join("temp_ecmg", src)

                try:
                    left_px = percent_of(design.attrib["left"], template_width)
                    top_px = percent_of(design.attrib["top"], template_height)
                    width_px = percent_of(design.attrib["width"], template_width)
                    height_px = percent_of(design.attrib["height"], template_height)

                    if os.path.exists(img_path):
                        slide.shapes.add_picture(
                            img_path,
                            scale_to_ppt(left_px, template_width, PPT_WIDTH),
                            scale_to_ppt(top_px, template_height, PPT_HEIGHT),
                            width=scale_to_ppt(width_px, template_width, PPT_WIDTH),
                            height=scale_to_ppt(height_px, template_height, PPT_HEIGHT)
                        )
                    else:
                        st.warning(f"⚠️ Image not found: {src}")
                except Exception as e:
                    st.error(f"❌ Error with image: {src}\n{str(e)}")

            # Add text
            for text in screen.findall("text"):
                design = text.find("design")
                content = text.find("content")

                left_px = percent_of(design.attrib["left"], template_width)
                top_px = percent_of(design.attrib["top"], template_height)
                width_px = percent_of(design.attrib["width"], template_width)
                height_px = percent_of(design.attrib["height"], template_height)

                textbox = slide.shapes.add_textbox(
                    scale_to_ppt(left_px, template_width, PPT_WIDTH),
                    scale_to_ppt(top_px, template_height, PPT_HEIGHT),
                    scale_to_ppt(width_px, template_width, PPT_WIDTH),
                    scale_to_ppt(height_px, template_height, PPT_HEIGHT)
                )
                tf = textbox.text_frame
                tf.clear()

                p = tf.paragraphs[0]
                p.text = "DÉONTOLOGIE - LE DÉCRET DANS SES GRANDES LIGNES ET L'ARTICLE 2"
                run = p.runs[0]
                run.font.size = Pt(24)
                run.font.bold = True
                run.font.name = "Tahoma"
                run.font.color.rgb = RGBColor(0x13, 0xAB, 0xB5)

            # Save output
            output = BytesIO()
            prs.save(output)
            st.success("🎉 slide1.pptx generated!")

            st.download_button(
                label="📥 Download PowerPoint",
                data=output.getvalue(),
                file_name="slide1.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

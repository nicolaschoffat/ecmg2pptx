import streamlit as st
from utils import parse_xml_to_slides, generate_pptx
from zipfile import ZipFile
import tempfile
import os

st.set_page_config(page_title="SCORM ➜ PowerPoint", layout="wide")
st.title("🧠 Convertisseur SCORM (.zip) ➜ PowerPoint")

scorm_zip = st.file_uploader("📦 Dépose ton export SCORM (.zip)", type=["zip"])

if scorm_zip:
    with tempfile.TemporaryDirectory() as tmpdir:
        with ZipFile(scorm_zip, 'r') as zip_ref:
            zip_ref.extractall(tmpdir)

        # Auto-scan récursif : trouver course.xml et son dossier
        course_xml_path = None
        for root, dirs, files in os.walk(tmpdir):
            if "course.xml" in files:
                course_xml_path = os.path.join(root, "course.xml")
                media_dir = root  # le dossier contenant course.xml
                break

        if course_xml_path and os.path.exists(course_xml_path):
            st.success("✅ Fichier 'course.xml' détecté automatiquement")
            slides = parse_xml_to_slides(course_xml_path, media_dir)

            output_path = os.path.join(tmpdir, "presentation.pptx")
            generate_pptx(slides, output_path)

            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 Télécharger le PowerPoint",
                    data=f,
                    file_name="presentation_convertie.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
        else:
            st.error("❌ Fichier 'course.xml' introuvable dans le ZIP.")

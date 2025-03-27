import streamlit as st
from utils import parse_xml_to_slides, generate_pptx
from zipfile import ZipFile
import tempfile
import os

st.set_page_config(page_title="Convertisseur XML ➜ PowerPoint", layout="wide")
st.title("🧠 Convertisseur XML + Médias ➜ PowerPoint")

# Upload zone
xml_file = st.file_uploader("📂 Dépose ton fichier XML", type=["xml"])
media_zip = st.file_uploader("📦 Fichiers médias (ZIP contenant images, vidéos...)", type=["zip"])

if xml_file and media_zip:
    with tempfile.TemporaryDirectory() as tmpdir:
        # Dézipper les fichiers
        media_dir = os.path.join(tmpdir, "media")
        os.makedirs(media_dir, exist_ok=True)

        with ZipFile(media_zip, 'r') as zip_ref:
            zip_ref.extractall(media_dir)

        # Lire et parser le XML
        st.success("✅ XML et fichiers médias reçus, traitement en cours...")
        slides = parse_xml_to_slides(xml_file, media_dir)

        # Génération PowerPoint
        output_path = os.path.join(tmpdir, "presentation.pptx")
        generate_pptx(slides, output_path)

        with open(output_path, "rb") as f:
            st.download_button(
                label="📥 Télécharger le PowerPoint",
                data=f,
                file_name="converti.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

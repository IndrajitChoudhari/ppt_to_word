import streamlit as st
from pptx import Presentation
from docx import Document
from io import BytesIO

def ppt_to_word(ppt_file):
    # Load the PowerPoint file
    presentation = Presentation(ppt_file)
    # Create a Word document
    doc = Document()

    for slide in presentation.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                doc.add_paragraph(paragraph.text)

    word_file = BytesIO()
    doc.save(word_file)
    word_file.seek(0)
    return word_file

st.title("PPT to Word Converter")

uploaded_file = st.file_uploader("Choose a PPT file", type="pptx")

if uploaded_file is not None:
    with st.spinner("Converting..."):
        word_file = ppt_to_word(uploaded_file)

    st.success("Conversion successful!")

    st.download_button(
        label="Download Word file",
        data=word_file,
        file_name="converted.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
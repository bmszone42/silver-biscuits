# app.py

import streamlit as st
import ast
import os
from pptx.util import Inches, Pt
from pptx import Presentation
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import base64

def create_presentation(slides_content, company_name, presentation_name, presenter):
    # Initialize a Presentation object
    presentation = Presentation()

    # Add title slide
    slide_layout = presentation.slide_layouts[0]  # Slide layout 0 is a title slide
    slide = presentation.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = presentation_name
    subtitle = slide.placeholders[1]
    subtitle.text = presenter

    # Add other slides
    for i, slide_content in enumerate(slides_content["slides"], start=2):
        # Add a new slide with a title and content layout
        slide_layout = presentation.slide_layouts[1]
        slide = presentation.slides.add_slide(slide_layout)

        # Set the title
        title = slide.shapes.title
        title.text = slide_content["title"]

        # Add bullet points
        content = slide.placeholders[1]
        tf = content.text_frame
        tf.text = slide_content["bulletPoints"][0]
        for point in slide_content["bulletPoints"][1:]:
            p = tf.add_paragraph()
            p.text = point

        # Add takeaway message at the bottom left with larger font
        left = Inches(1)
        top = Inches(6) 
        width = Inches(6)
        height = Inches(2)
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = slide_content["takeawayMessage"]
        for paragraph in tf.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(24)  # Larger font size

        # Add notes to the slide
        notes_slide = slide.notes_slide
        notes_tf = notes_slide.notes_text_frame
        for note in slide_content["talkingPoints"]:
            notes_tf.add_paragraph().text = note

    # Save the presentation
    presentation.save("MitoSense for Space.pptx")

def get_download_link(file_path):
    with open(file_path, 'rb') as f:
        data = f.read()
    b64 = base64.b64encode(data).decode()
    return f'<a href="data:application/octet-stream;base64,{b64}" download="{file_path}">Download file</a>'

def main():
    st.title('PowerPoint Presentation Creator')

    st.write('Please paste your Python dictionary here.')
    user_input = st.text_area("Paste your dictionary here", "{}")
    
    # Add the Make Presentation button
    button_clicked = st.button('Make Presentation')

    if button_clicked:
        try:
            # The content is evaluated as a Python dictionary
            slides_content = ast.literal_eval(user_input)

            company_name = "MitoSense"
            presentation_name = "Mitochondrial Frontiers: Exploring the Role of Mitochondria Organelle Transplantation in Spaceflight and Neurodegeneration"
            presenter = "Brent Segal, PhD"

            create_presentation(slides_content, company_name, presentation_name, presenter)

            st.success('Presentation created successfully!')

            # Provide download link
            download_link = get_download_link("MitoSense for Space.pptx")
            st.markdown(download_link, unsafe_allow_html=True)

        except Exception as e:
            st.error(f'Error parsing dictionary: {e}')

if __name__ == "__main__":
    main()


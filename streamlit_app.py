# app.py

import streamlit as st
import ast
import os
from pptx.util import Inches, Pt
from pptx import Presentation
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import base64
import openai

# Set OpenAI API key
openai.api_key = st.secrets['OPENAIAPI_KEY']

def generate_outline(topic, num_slides):
    response = openai.Completion.create(
      engine="text-davinci-003",
      prompt=f"Create an outline for a presentation on the topic of '{topic}' with {num_slides} slides.",
      temperature=0.5,
      max_tokens=60
    )

    outline = response.choices[0].text.strip().split('\n')
    return outline

def generate_presentation(slide_titles):
    # In a real-world scenario, you might want to use the OpenAI API here to generate the content for each slide
    presentation = {
        "slides": [
            {
                "title": title,
                "bulletPoints": ["Bullet point 1", "Bullet point 2", "Bullet point 3"],
                "takeawayMessage": "Takeaway message",
                "talkingPoints": ["Talking point 1", "Talking point 2", "Talking point 3", "Talking point 4", "Talking point 5"]
            }
            for title in slide_titles
        ]
    }
    return presentation

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
    
    # Ask for the topic and number of slides
topic = st.sidebar.text_input('Topic')
num_slides = st.sidebar.number_input('Number of slides', min_value=1)

# When the user clicks the "Generate Outline" button, generate the outline
if st.sidebar.button('Generate Outline'):
    outline = generate_outline(topic, num_slides)
    
    # Store the outline in a session state variable so it persists across runs
    st.session_state['outline'] = outline

    # Display checkboxes for each item in the outline
    st.session_state['include_slide'] = []
    for i, slide_title in enumerate(outline):
        include_slide = st.checkbox(f'Include "{slide_title}" in the presentation?', key=f'include_slide_{i}')
        st.session_state['include_slide'].append(include_slide)
        
        # Allow the user to edit the title
        edited_title = st.text_input('Edit the slide title', slide_title, key=f'edited_title_{i}')
        if edited_title != slide_title:
            outline[i] = edited_title
            
    # When the user clicks the "Generate Presentation" button, generate the presentation
    if st.button('Generate Presentation'):
        # Use only the titles that the user checked
        slide_titles = [title for include, title in zip(st.session_state['include_slide'], outline) if include]
        
        # Generate the presentation dictionary
        presentation = generate_presentation(slide_titles)
        
        # Display the presentation dictionary
        st.code(presentation, language='json')

    # Add input fields in sidebar
    st.sidebar.title('Presentation Details')
    company_name = st.sidebar.text_input('Company name', 'Company')
    presentation_name = st.sidebar.text_input('Presentation name', 'Presentation')
    presenter = st.sidebar.text_input('Presenter', 'Presenter')

    st.write('Please paste your Python dictionary here.')
    user_input = st.text_area("Paste your dictionary here", "{}")
    
    # Add the Make Presentation button
    button_clicked = st.button('Make Presentation')

    if button_clicked:
        try:
            # The content is evaluated as a Python dictionary
            slides_content = ast.literal_eval(user_input)

            create_presentation(slides_content, company_name, presentation_name, presenter)

            st.success('Presentation created successfully!')

            # Provide download link
            download_link = get_download_link("MitoSense for Space.pptx")
            st.markdown(download_link, unsafe_allow_html=True)

        except Exception as e:
            st.error(f'Error parsing dictionary: {e}')

if __name__ == "__main__":
    main()


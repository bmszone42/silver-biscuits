# app.py

import streamlit as st
import ast
from pptx.util import Inches, Pt
from pptx import Presentation
import base64
import openai
from collections.abc import Iterable

# Set OpenAI API key
openai.api_key = st.secrets['OPENAI_KEY']

def generate_slide_content(title):
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=f"Generate slide content for the title: '{title}'\n\nShort crisp title:",
        temperature=0.5,
        max_tokens=10,
        n=1
    )
    crisp_title = response.choices[0].text.strip()

    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=f"Generate slide content for the title: '{title}'\n\nThree poignant and useful bullets:\n1.",
        temperature=0.5,
        max_tokens=30,
        n=3
    )
    bullets = [bullet.strip() for bullet in response.choices]

    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=f"Generate slide content for the title: '{title}'\n\nShort takeaway message (8 words or less):",
        temperature=0.5,
        max_tokens=10,
        n=1
    )
    takeaway_message = response.choices[0].text.strip()

    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=f"Generate slide content for the title: '{title}'\n\nFive detailed talking points (50 words each):\n1.",
        temperature=0.5,
        max_tokens=60,
        n=5
    )
    talking_points = [point.strip() for point in response.choices]

    slide_content = {
        "title": title,
        "crisp_title": crisp_title,
        "bullets": bullets,
        "takeaway_message": takeaway_message,
        "talking_points": talking_points
    }

    return slide_content

def generate_outline(topic, num_slides):
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=f"Generate {num_slides} slide titles for a presentation on the topic: '{topic}'.\n\n",
        temperature=0.5,
        max_tokens=100,
        n=1
    )
    
    # Extract the text from the single completion choice
    generated_text = response.choices[0].text.strip()
    
    # Split the generated text into individual slide titles
    # This assumes that the generated text contains slide titles separated by newlines
    outline = generated_text.split('\n')
    
    # Make sure that the number of slide titles matches num_slides
    if len(outline) != num_slides:
        print(f"Warning: Expected {num_slides} slide titles but received {len(outline)}")
    
    return outline


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
    for slide_content in slides_content:
        # Add a new slide with a title and content layout
        slide_layout = presentation.slide_layouts[1]
        slide = presentation.slides.add_slide(slide_layout)

        # Set the title
        title = slide.shapes.title
        title.text = slide_content["crisp_title"]

        # Add bullet points
        content = slide.placeholders[1]
        tf = content.text_frame
        for bullet in slide_content["bullets"]:
            p = tf.add_paragraph()
            p.text = bullet

        # Add takeaway message at the bottom left with larger font
        left = Inches(1)
        top = Inches(6)
        width = Inches(6)
        height = Inches(2)
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = slide_content["takeaway_message"]
        for paragraph in tf.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(24)  # Larger font size

        # Add notes to the slide
        notes_slide = slide.notes_slide
        notes_tf = notes_slide.notes_text_frame
        for talking_point in slide_content["talking_points"]:
            notes_tf.add_paragraph().text = talking_point

    # Save the presentation
    presentation.save("SlideDeck.pptx")

def get_download_link(file_path):
    with open(file_path, 'rb') as f:
        data = f.read()
    b64 = base64.b64encode(data).decode()
    return f'<a href="data:application/octet-stream;base64,{b64}" download="{file_path}">Download file</a>'

def main():
    # Step 1: Allow user to enter a topic
    if 'topic' not in st.session_state:
        st.session_state['topic'] = ''
    topic = st.sidebar.text_input('Topic', value=st.session_state['topic'])
    st.session_state['topic'] = topic

    # Step 2: Allow the user to select n charts for the outline
    if 'num_slides' not in st.session_state:
        st.session_state['num_slides'] = 1
    num_slides = st.sidebar.number_input('Number of slides', min_value=1, value=st.session_state['num_slides'])
    st.session_state['num_slides'] = num_slides

    # Step 3: Generate the outline upon pressing Generate Outline
    if st.sidebar.button('Generate Outline'):
        st.session_state['outline'] = generate_outline(topic, num_slides)

    # Check if the outline was successfully generated
    if 'outline' in st.session_state and st.session_state['outline'] and isinstance(st.session_state['outline'], list):
        # Step 4: Display the outline in the sidebar
        st.sidebar.write('Generated Outline:')
        for slide_title in st.session_state['outline']:
            st.sidebar.write(f'- {slide_title}')

        # Step 5: Allow the user to approve the Outline
        if 'approved' not in st.session_state:
            st.session_state['approved'] = False
        st.session_state['approved'] = st.sidebar.checkbox('Approve Outline', value=st.session_state['approved'])

        # Step 6: If the Outline is approved, generate a python dictionary with content for the presentation
        if st.session_state['approved']:
            slides_content = []
            for slide_title in st.session_state['outline']:
                st.write(f"Generating slide content for: {slide_title}")
                slide_content = generate_slide_content(slide_title)
                slides_content.append(slide_content)

                # Display the slide content dictionary
                st.write(f"Slide Content: {slide_content}")

            # Step 8: Prompt the user to enter their presenter name, presentation title, and company name
            st.sidebar.title('Presentation Details')
            company_name = st.sidebar.text_input('Company name', 'Company')
            presentation_name = st.sidebar.text_input('Presentation name', 'Presentation')
            presenter = st.sidebar.text_input('Presenter', 'Presenter')

            # Step 9: Show the "Create Presentation" button
            if st.sidebar.button('Create Presentation'):
                create_presentation(slides_content, company_name, presentation_name, presenter)
                st.success('Presentation created successfully!')

                # Step 10: Allow the user to download the presentation with a link
                download_link = get_download_link("SlideDeck.pptx")
                st.markdown(download_link, unsafe_allow_html=True)

if __name__ == "__main__":
    main()








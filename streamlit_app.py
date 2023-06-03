# app.py

import streamlit as st
import ast
from pptx.util import Inches, Pt
from pptx import Presentation
import base64
import openai

# Set OpenAI API key
openai.api_key = st.secrets['OPENAI_KEY']

def generate_slide_content(title):
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=f"Generate slide content for the title: '{title}'\n\nShort crisp title:",
        temperature=0.5,
        max_tokens=100,
        n=1
    )
    crisp_title = response.choices[0].text.strip()

    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=f"Generate slide content for the title: '{title}'\n\nThree poignant and useful bullets:\n1.",
        temperature=0.5,
        max_tokens=50,
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
    outline = []
    for i in range(num_slides):
        response = openai.Completion.create(
            engine="text-davinci-003",
            prompt=f"Generate outline for slide {i+1} on the topic: '{topic}'\n\nSlide title:",
            temperature=0.5,
            max_tokens=10,
            n=1
        )
        slide_title = response.choices[0].text.strip()
        outline.append(slide_title)

    return outline

def generate_presentation(outline):
    slides_content = []
    for title in outline:
        slide_content = generate_slide_content(title)
        slides_content.append(slide_content)

    presentation = {
        "slides": slides_content
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

    # When the user clicks the "Generate Presentation" button, generate the presentation
    if st.button('Generate Presentation'):
        # Use only the titles that the user checked
        slide_titles = [title for include, title in zip(st.session_state['include_slide'], st.session_state['outline']) if include]

        # Generate the presentation dictionary
        presentation = generate_presentation(slide_titles)

        # Display the presentation dictionary
        st.code(presentation, language='json')

    # Add input fields in sidebar
    st.sidebar.title('Presentation Details')
    company_name = st.sidebar.text_input('Company name', 'Company')
    presentation_name = st.sidebar.text_input('Presentation name', 'Presentation')
    presenter = st.sidebar.text_input('Presenter', 'Presenter')

    # Add the Make Presentation button if required fields are filled
    if company_name and presentation_name and presenter:
        button_clicked = st.button('Make Presentation')

        if button_clicked:
            try:
                # Create slides content based on the generated presentation dictionary
                slides_content = []
                for slide_title in presentation['slides']:
                    slide_content = generate_slide_content(slide_title)
                    slides_content.append(slide_content)

                # Display the slides content
                st.code(slides_content, language='json')

                # Ask for user approval
                approved = st.checkbox('I approve the slide content and want to create the presentation.')

                if approved:
                    create_presentation(slides_content, company_name, presentation_name, presenter)

                    st.success('Presentation created successfully!')

                    # Provide download link
                    download_link = get_download_link("SlideDeck.pptx")
                    st.markdown(download_link, unsafe_allow_html=True)

            except Exception as e:
                st.error(f'Error generating presentation: {e}')
    else:
        st.warning('Please fill in all required fields.')

if __name__ == "__main__":
    main()

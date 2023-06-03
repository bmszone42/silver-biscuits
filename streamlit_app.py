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
    outline = [choice.text.strip() for choice in response.choices]
    
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
    st.title('PowerPoint Presentation Creator')

    # Step 1: Allow user to enter a topic
    topic = st.sidebar.text_input('Topic')

    # Step 2: Allow the user to select n charts for the outline
    num_slides = st.sidebar.number_input('Number of slides', min_value=1)

    # Step 3: Generate the outline upon pressing Generate Outline
    if st.sidebar.button('Generate Outline'):
        outline = generate_outline(topic, num_slides)

        # Step 4: Show the Outline to the User
        st.write('Generated Outline:')
        for slide_title in outline:
            st.write(f'- {slide_title}')

        # Step 5: Allow the user to approve the Outline or generate a new outline
        approved = st.button('Approve Outline')
        if not approved:
            st.write("Outline not approved. Please generate a new outline.")
            return

        # Step 6: If the Outline is approved, generate a python dictionary with content for the presentation
        slides_content = []
        for slide_title in outline:
            st.write(f"Generating slide content for: {slide_title}")
            slide_content = generate_slide_content(slide_title)
            slides_content.append(slide_content)
            # Display the slide content dictionary
            st.write(f"Slide Content: {slide_content}")

        # Step 7: Prompt the user to enter their presenter name, presentation title, and company name
        st.sidebar.title('Presentation Details')
        company_name = st.sidebar.text_input('Company name', 'Company')
        presentation_name = st.sidebar.text_input('Presentation name', 'Presentation')
        presenter = st.sidebar.text_input('Presenter', 'Presenter')

        # Step 8: If the user approves the outline, create a presentation in pptx format
        if st.button('Create Presentation'):
            create_presentation(slides_content, company_name, presentation_name, presenter)
            st.success('Presentation created successfully!')

            # Step 9: Allow the user to download the presentation with a link
            download_link = get_download_link("SlideDeck.pptx")
            st.markdown(download_link, unsafe_allow_html=True)

            return  # Exit the function to prevent further execution

    # Additional code logic here if the outline is not approved or no action is taken
    # ...

if __name__ == "__main__":
    main()

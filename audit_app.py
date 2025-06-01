
import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from docx import Document
from datetime import datetime
import os

# Load the Word document
doc = Document("AW08.02-Site safety inspection sheet V2.docx")

# Extract questions from the Word document
questions = []
for para in doc.paragraphs:
    if para.text.strip():
        questions.append(para.text.strip())

# Streamlit app
st.title("Health and Safety Audit Form")

# Project selection
projects = ["Project A", "Project B", "Project C"]  # Placeholder projects
project = st.selectbox("Select your project", projects)

# Initialize response dictionary
responses = {"Project": project, "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

# Display questions and collect responses
for question in questions:
    st.write(question)
    response = st.radio("", ["Yes", "No", "N/A"], key=question)
    comment = st.text_area("Comments", key=question + "_comment")
    responses[question] = {"Response": response, "Comment": comment}

# File uploader for media
uploaded_files = st.file_uploader("Upload media", accept_multiple_files=True)
media_files = []
if uploaded_files:
    for uploaded_file in uploaded_files:
        media_files.append(uploaded_file.name)
        with open(os.path.join("uploads", uploaded_file.name), "wb") as f:
            f.write(uploaded_file.getbuffer())

# Save responses to Excel
if st.button("Submit"):
    df = pd.DataFrame.from_dict(responses, orient="index")
    df.to_excel("audit_responses.xlsx")

    # Generate PDF report
    pdf_filename = "audit_report.pdf"
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((50, 50), "Health and Safety Audit Report")
    page.insert_text((50, 80), f"Project: {project}")
    page.insert_text((50, 110), f"Date: {responses['Date']}")

    y = 140
    for question, answer in responses.items():
        if question not in ["Project", "Date"]:
            page.insert_text((50, y), f"{question}: {answer['Response']}")
            y += 30
            if answer["Comment"]:
                page.insert_text((50, y), f"Comment: {answer['Comment']}")
                y += 30

    # Save media files in the PDF
    for media_file in media_files:
        page.insert_text((50, y), f"Media: {media_file}")
        y += 30

    doc.save(pdf_filename)
    st.success("Audit submitted successfully!")
    st.write("Responses saved to audit_responses.xlsx")
    st.write("PDF report generated: audit_report.pdf")

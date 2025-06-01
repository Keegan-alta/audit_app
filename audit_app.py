
import streamlit as st
import pandas as pd
import docx
import fitz  # PyMuPDF
import openpyxl

# Load the Word document
doc = docx.Document("AW08.02-Site safety inspection sheet V2.docx")

# Extract questions from the Word document
questions = []
for para in doc.paragraphs:
    if para.text.strip():
        questions.append(para.text.strip())

# Streamlit app
st.title("Health and Safety Audit Form")

# Project selection dropdown (placeholder projects)
projects = ["Project A", "Project B", "Project C"]
selected_project = st.selectbox("Select your project", projects)

# Form to collect responses
responses = {}
for question in questions:
    st.subheader(question)
    response = st.radio("Response", ["Yes", "No", "N/A"], key=question)
    comment = st.text_area("Comments", key=f"{question}_comment")
    flag_issue = st.checkbox("Flag Issue", key=f"{question}_flag")
    action = st.text_area("Action", key=f"{question}_action")
    media = st.file_uploader("Upload Media", type=["jpg", "jpeg", "png", "pdf"], key=f"{question}_media")
    
    responses[question] = {
        "response": response,
        "comment": comment,
        "flag_issue": flag_issue,
        "action": action,
        "media": media
    }

# Save responses to Excel
if st.button("Submit"):
    df = pd.DataFrame(responses).T
    df.to_excel("audit_responses.xlsx", engine='openpyxl')
    st.success("Responses saved to audit_responses.xlsx")

    # Generate PDF report
    pdf = fitz.open()
    page = pdf.new_page()
    page.insert_text((50, 50), f"Health and Safety Audit Report
Project: {selected_project}
")
    y = 100
    for question, response in responses.items():
        page.insert_text((50, y), f"Question: {question}
Response: {response['response']}
Comments: {response['comment']}
Flag Issue: {response['flag_issue']}
Action: {response['action']}
")
        y += 100
    pdf.save("audit_report.pdf")
    st.success("PDF report generated: audit_report.pdf")

import streamlit as st
import pandas as pd
import requests
from docx import Document
from io import BytesIO
import base64
import os
from typing import Dict, List
import tempfile
import csv
from PyPDF2 import PdfReader
from pptx import Presentation

# Configuration
DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"
ALLOWED_EXTENSIONS = ["docx", "pdf", "pptx"]

def read_file_content(uploaded_file) -> str:
    """Read content from different file formats"""
    content = ""
    
    if uploaded_file.name.endswith('.docx'):
        doc = Document(BytesIO(uploaded_file.getvalue()))
        content = "\n".join([para.text for para in doc.paragraphs])
    elif uploaded_file.name.endswith('.pdf'):
        reader = PdfReader(BytesIO(uploaded_file.getvalue()))
        content = "\n".join([page.extract_text() for page in reader.pages])
    elif uploaded_file.name.endswith('.pptx'):
        prs = Presentation(BytesIO(uploaded_file.getvalue()))
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    content += shape.text + "\n"
    return content

def call_deepseek_api(prompt: str, system_prompt: str, model: str = "deepseek-r1") -> str:
    """Call DeepSeek API with given prompts"""
    headers = {
        "Authorization": f"Bearer {st.secrets['DEEPSEEK_API_KEY']}",
        "Content-Type": "application/json"
    }
    
    data = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.3
    }
    
    response = requests.post(DEEPSEEK_API_URL, json=data, headers=headers)
    response.raise_for_status()
    return response.json()["choices"][0]["message"]["content"]

def generate_feedback_document(rubric_df: pd.DataFrame, overall_comments: str, feedforward: str) -> bytes:
    """Generate Word document with feedback"""
    doc = Document()
    
    # Add rubric table
    doc.add_heading('Assessment Rubric', 1)
    table = doc.add_table(rows=1, cols=len(rubric_df.columns))
    table.style = 'Table Grid'
    
    # Header row
    for i, col in enumerate(rubric_df.columns):
        table.cell(0, i).text = str(col)
    
    # Data rows
    for _, row in rubric_df.iterrows():
        cells = table.add_row().cells
        for i, value in enumerate(row):
            cells[i].text = str(value)
            if rubric_df.columns[i] in ['Score', 'Comment']:
                for paragraph in cells[i].paragraphs:
                    paragraph.runs[0].font.highlight_color = WD_COLOR_INDEX.GREEN
    
    # Add overall comments and feedforward
    doc.add_heading('Overall Comments', 2)
    doc.add_paragraph(overall_comments)
    
    doc.add_heading('Feedforward Suggestions', 2)
    for point in feedforward.split('\n'):
        doc.add_paragraph(point, style='ListBullet')
    
    # Save to bytes buffer
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

def main():
    st.title("Automated Assignment Grading")
    
    # Password protection
    if 'authenticated' not in st.session_state:
        password = st.text_input("Enter application password:", type='password')
        if password == st.secrets["APP_PASSWORD"]:
            st.session_state.authenticated = True
        elif password:
            st.error("Incorrect password")
        return
    
    # Main application
    with st.sidebar:
        st.header("Configuration")
        rubric_file = st.file_uploader("Upload Grading Rubric (CSV)", type=['csv'])
        assignment_task = st.text_area("Assignment Task Description")
        level = st.selectbox("Academic Level", [
            "Undergraduate Level 4", "Undergraduate Level 5", "Undergraduate Level 6",
            "Masters Level 7", "PhD Level 8"
        ])
        assessment_type = st.selectbox("Assessment Type", [
            "Essay", "Report", "Presentation", "Practical Work"
        ])
        additional_instructions = st.text_area("Additional Instructions")
    
    st.header("Student Submissions")
    student_files = st.file_uploader(
        "Upload Student Assignments",
        type=ALLOWED_EXTENSIONS,
        accept_multiple_files=True
    )
    
    if st.button("Run Marking") and rubric_file and student_files:
        # Read and parse rubric
        rubric_df = pd.read_csv(rubric_file)
        
        for uploaded_file in student_files:
            with st.spinner(f"Processing {uploaded_file.name}..."):
                try:
                    # Process student submission
                    content = read_file_content(uploaded_file)
                    
                    # Prepare API prompts
                    system_prompt = f"""
                    You are an academic assessment expert. Evaluate the student's work based on:
                    - Academic level: {level}
                    - Assessment type: {assessment_type}
                    - Assignment task: {assignment_task}
                    - Additional instructions: {additional_instructions}
                    Use the provided rubric and return structured feedback.
                    """
                    
                    user_prompt = f"""
                    Student Submission:
                    {content}
                    
                    Grading Rubric:
                    {rubric_df.to_csv(index=False)}
                    
                    Provide:
                    1. Rubric scores and comments for each criterion
                    2. Overall comments
                    3. Feedforward suggestions
                    """
                    
                    # Call DeepSeek API
                    response = call_deepseek_api(user_prompt, system_prompt)
                    
                    # Process API response
                    # (This would need custom parsing based on API response structure)
                    # For demonstration, we'll mock this part
                    feedback_doc = generate_feedback_document(
                        rubric_df, 
                        "Overall comments placeholder", 
                        "Feedforward placeholder"
                    )
                    
                    # Download button
                    st.download_button(
                        label=f"Download Feedback for {uploaded_file.name}",
                        data=feedback_doc,
                        file_name=f"feedback_{uploaded_file.name.split('.')[0]}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                
                except Exception as e:
                    st.error(f"Error processing {uploaded_file.name}: {str(e)}")

if __name__ == "__main__":
    main()

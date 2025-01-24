# app.py
import streamlit as st
import pandas as pd
import requests
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from io import BytesIO
from PyPDF2 import PdfReader
from pptx import Presentation

# Configuration
DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"
ALLOWED_EXTENSIONS = ["docx", "pdf", "pptx"]
MODEL_CONFIG = {
    "default": "deepseek-reasoner",
    "available_models": ["deepseek-reasoner"]  # Add other models if available
}

def read_file_content(uploaded_file) -> str:
    """Read content from different file formats"""
    content = ""
    
    try:
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
    except Exception as e:
        st.error(f"Error reading {uploaded_file.name}: {str(e)}")
        raise

def call_deepseek_api(prompt: str, system_prompt: str) -> str:
    """Call DeepSeek API with reasoning-optimized parameters"""
    headers = {
        "Authorization": f"Bearer {st.secrets['DEEPSEEK_API_KEY']}",
        "Content-Type": "application/json",
        "Accept": "application/json"
    }
    
    data = {
        "model": MODEL_CONFIG["default"],
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.2,  # Lower temperature for more focused responses
        "top_p": 0.85,
        "max_tokens": 3000,  # Increased for detailed feedback
        "frequency_penalty": 0.2,
        "presence_penalty": 0.1
    }
    
    try:
        response = requests.post(DEEPSEEK_API_URL, json=data, headers=headers)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except requests.exceptions.HTTPError as err:
        error_data = response.json().get("error", {})
        error_msg = (f"API Error ({response.status_code}): "
                    f"{error_data.get('message', 'Unknown error')}")
        st.error(error_msg)
        st.json(error_data)  # Display full error details
        raise
    except Exception as e:
        st.error(f"API Call Failed: {str(e)}")
        raise

def generate_feedback_document(rubric_df: pd.DataFrame, overall_comments: str, feedforward: str) -> bytes:
    """Generate Word document with feedback"""
    try:
        doc = Document()
        
        # Add rubric table
        doc.add_heading('Assessment Rubric', 1)
        table = doc.add_table(rows=1, cols=len(rubric_df.columns))
        table.style = 'Table Grid'
        
        # Header row
        header_cells = table.rows[0].cells
        for i, col in enumerate(rubric_df.columns):
            header_cells[i].text = str(col)
        
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
            if point.strip():
                doc.add_paragraph(point.strip(), style='ListBullet')
        
        # Save to bytes buffer
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer.getvalue()
    except Exception as e:
        st.error(f"Error generating document: {str(e)}")
        raise

def main():
    st.set_page_config(page_title="AutoGrader Pro", layout="wide")
    st.title("📚 Automated Assignment Grading System")
    
    # Password protection
    if 'authenticated' not in st.session_state:
        with st.container():
            st.markdown("## Secure Login")
            password = st.text_input("Enter application password:", type='password')
            if st.button("Authenticate"):
                if password == st.secrets["APP_PASSWORD"]:
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error("Incorrect password")
        return
    
    # Main application
    with st.sidebar:
        st.header("⚙️ Configuration")
        rubric_file = st.file_uploader("📝 Upload Grading Rubric (CSV)", type=['csv'])
        assignment_task = st.text_area("📋 Assignment Task Description", height=150,
                                      help="Clearly describe the assignment requirements")
        level = st.selectbox("🎓 Academic Level", [
            "Undergraduate Level 4", "Undergraduate Level 5", 
            "Undergraduate Level 6", "Masters Level 7", "PhD Level 8"
        ])
        assessment_type = st.selectbox("📄 Assessment Type", [
            "Essay", "Report", "Presentation", "Practical Work"
        ])
        additional_instructions = st.text_area("🔧 Additional Instructions", height=100,
                                              help="Special considerations or marking guidelines")
    
    st.header("📤 Student Submissions")
    student_files = st.file_uploader(
        "Upload Student Assignments",
        type=ALLOWED_EXTENSIONS,
        accept_multiple_files=True,
        help="Supported formats: DOCX, PDF, PPTX"
    )
    
    if st.button("🚀 Run Automated Grading") and rubric_file and student_files:
        try:
            rubric_df = pd.read_csv(rubric_file)
            required_columns = ['Criteria', 'Description', 'Max Score']
            if not all(col in rubric_df.columns for col in required_columns):
                st.error(f"Rubric CSV must contain: {', '.join(required_columns)}")
                return
            
            for uploaded_file in student_files:
                with st.expander(f"Processing {uploaded_file.name}", expanded=True):
                    try:
                        content = read_file_content(uploaded_file)
                        
                        # System prompt optimized for deepseek-reasoner
                        system_prompt = f"""
                        You are an expert academic assessor specializing in {assessment_type} evaluations. 
                        Conduct a comprehensive analysis of the student submission considering:

                        Academic Level: {level}
                        Assessment Type: {assessment_type}
                        Assignment Task: {assignment_task}
                        Additional Instructions: {additional_instructions}

                        Rubric Structure:
                        {rubric_df.to_csv(index=False)}

                        Required Output Format:
                        SCORES:
                        - [Criterion Name]: [Awarded Score]/[Max Score], [Concise Justification]
                        ...
                        OVERALL_COMMENTS:
                        [Structured evaluation covering strengths/weaknesses]
                        FEEDFORWARD:
                        - [Actionable Improvement Suggestion 1]
                        - [Actionable Improvement Suggestion 2]
                        ...
                        """
                        
                        user_prompt = f"""
                        STUDENT SUBMISSION CONTENT:
                        {content[:10000]}... [truncated if exceeding length]

                        ANALYSIS INSTRUCTIONS:
                        1. Perform criterion-by-criterion evaluation
                        2. Maintain strict alignment with rubric metrics
                        3. Provide specific examples from the text
                        4. Balance conciseness with depth
                        5. Prioritize objective, measurable feedback
                        """
                        
                        with st.spinner("🔍 Conducting in-depth analysis..."):
                            response = call_deepseek_api(
                                prompt=user_prompt,
                                system_prompt=system_prompt
                            )
                        
                        # Parse response
                        scores = {}
                        overall_comments = []
                        feedforward = []
                        current_section = None

                        for line in response.split('\n'):
                            line = line.strip()
                            if line.startswith("SCORES:"):
                                current_section = 'scores'
                            elif line.startswith("OVERALL_COMMENTS:"):
                                current_section = 'overall'
                            elif line.startswith("FEEDFORWARD:"):
                                current_section = 'feedforward'
                            else:
                                if current_section == 'scores' and line.startswith('-'):
                                    parts = line[2:].split(':', 1)
                                    if len(parts) == 2:
                                        criterion, evaluation = parts
                                        score_part, comment = evaluation.split(',', 1)
                                        scores[criterion.strip()] = {
                                            'Score': score_part.strip(),
                                            'Comment': comment.strip()
                                        }
                                elif current_section == 'overall':
                                    overall_comments.append(line)
                                elif current_section == 'feedforward' and line.startswith('-'):
                                    feedforward.append(line[2:].strip())

                        # Update rubric dataframe
                        for criterion, data in scores.items():
                            mask = rubric_df['Criteria'] == criterion
                            if not mask.any():
                                st.warning(f"Criterion '{criterion}' not found in rubric")
                                continue
                            rubric_df.loc[mask, 'Score'] = data['Score']
                            rubric_df.loc[mask, 'Comment'] = data['Comment']
                        
                        # Generate feedback document
                        feedback_doc = generate_feedback_document(
                            rubric_df,
                            "\n".join(overall_comments).strip(),
                            "\n".join(feedforward)
                        )
                        
                        # Download button
                        st.download_button(
                            label=f"📥 Download Feedback for {uploaded_file.name}",
                            data=feedback_doc,
                            file_name=f"feedback_{uploaded_file.name.split('.')[0]}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    
                    except Exception as e:
                        st.error(f"Error processing {uploaded_file.name}: {str(e)}")
        
        except Exception as e:
            st.error(f"System error: {str(e)}")

if __name__ == "__main__":
    main()
    

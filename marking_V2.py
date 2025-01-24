import streamlit as st
import pandas as pd
import requests
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from io import BytesIO
from PyPDF2 import PdfReader
from pptx import Presentation
import re

# Configuration
DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"
ALLOWED_EXTENSIONS = ["docx", "pdf", "pptx"]

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

def process_rubric(uploaded_file):
    """Process percentage-based rubric format"""
    try:
        df = pd.read_csv(uploaded_file, skip_blank_lines=True).dropna(how='all')
        df = df[df['Criteria'].notna()].reset_index(drop=True)
        df = df.iloc[:, :9]  # Keep relevant columns
        
        # Extract max score from criteria name
        df['Max Score'] = df['Criteria'].str.extract(r'\((\d+)%\)').astype(float) / 100
        df['Criteria'] = df['Criteria'].str.replace(r'\s*\(\d+%\)', '', regex=True)
        
        # Rename columns
        df.columns = [
            'Criteria', '80-100%', '70-79%', '60-69%', 
            '50-59%', '40-49%', '0-39%', 'Score', 'Comment', 'Max Score'
        ]
        return df
    except Exception as e:
        st.error(f"Error processing rubric: {str(e)}")
        raise

def call_deepseek_api(prompt: str, system_prompt: str) -> str:
    """Call DeepSeek API"""
    headers = {
        "Authorization": f"Bearer {st.secrets['DEEPSEEK_API_KEY']}",
        "Content-Type": "application/json"
    }
    
    data = {
        "model": "deepseek-reasoner",
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.2,
        "max_tokens": 3000
    }
    
    try:
        response = requests.post(DEEPSEEK_API_URL, json=data, headers=headers)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except requests.exceptions.HTTPError as err:
        st.error(f"API Error: {err}\nResponse: {response.text}")
        raise

def generate_feedback_document(rubric_df: pd.DataFrame, overall_comments: str, feedforward: str) -> bytes:
    """Generate feedback document with banded rubric"""
    try:
        doc = Document()
        doc.add_heading('Assessment Rubric', 1)
        
        # Create table with original rubric structure
        cols = ['Criteria', '80-100%', '70-79%', '60-69%', 
               '50-59%', '40-49%', '0-39%', 'Score', 'Comment']
        table = doc.add_table(rows=1, cols=len(cols))
        table.style = 'Table Grid'
        
        # Header row
        for i, col in enumerate(cols):
            table.cell(0, i).text = col
        
        # Data rows
        for _, row in rubric_df.iterrows():
            cells = table.add_row().cells
            for i, col in enumerate(cols):
                cells[i].text = str(row[col]) if pd.notna(row[col]) else ''
                # Highlight score and comment cells
                if col in ['Score', 'Comment']:
                    for paragraph in cells[i].paragraphs:
                        paragraph.runs[0].font.highlight_color = WD_COLOR_INDEX.GREEN
        
        # Add overall comments and feedforward
        doc.add_heading('Overall Comments', 2)
        doc.add_paragraph(overall_comments)
        
        doc.add_heading('Feedforward Suggestions', 2)
        for point in feedforward.split('\n'):
            if point.strip():
                doc.add_paragraph(point.strip(), style='ListBullet')
        
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
    
    # Authentication
    if 'authenticated' not in st.session_state:
        with st.container():
            password = st.text_input("Enter application password:", type='password')
            if st.button("Authenticate"):
                if password == st.secrets["APP_PASSWORD"]:
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error("Incorrect password")
        return
    
    # Main interface
    with st.sidebar:
        st.header("⚙️ Configuration")
        rubric_file = st.file_uploader("📝 Upload Grading Rubric (CSV)", type=['csv'])
        assignment_task = st.text_area("📋 Assignment Task Description", height=150)
        level = st.selectbox("🎓 Academic Level", [
            "Undergraduate Level 4", "Undergraduate Level 5", 
            "Undergraduate Level 6", "Masters Level 7", "PhD Level 8"
        ])
    
    st.header("📤 Student Submissions")
    student_files = st.file_uploader(
        "Upload Student Assignments",
        type=ALLOWED_EXTENSIONS,
        accept_multiple_files=True
    )
    
    if st.button("🚀 Run Automated Grading") and rubric_file and student_files:
        try:
            rubric_df = process_rubric(rubric_file)
            
            required_columns = ['Criteria', '80-100%', '70-79%', '60-69%',
                               '50-59%', '40-49%', '0-39%', 'Max Score']
            if not all(col in rubric_df.columns for col in required_columns):
                st.error("Invalid rubric format. Please use the template CSV.")
                return
            
            for uploaded_file in student_files:
                with st.expander(f"Processing {uploaded_file.name}", expanded=True):
                    try:
                        content = read_file_content(uploaded_file)
                        
                        system_prompt = f"""
                        As an academic expert, evaluate submissions using this rubric:
                        
                        Academic Level: {level}
                        Assignment Task: {assignment_task}
                        
                        Rubric Structure:
                        {rubric_df.to_csv(index=False)}
                        
                        Provide response in EXACT format:
                        SCORES:
                        - [Criteria Name]: [Selected Band], [Score]/[Max Score], [Comment]
                        OVERALL_COMMENTS:
                        [Concise evaluation]
                        FEEDFORWARD:
                        - [Suggestion 1]
                        - [Suggestion 2]
                        """
                        
                        user_prompt = f"""
                        STUDENT SUBMISSION CONTENT:
                        {content[:10000]}... [truncated if long]
                        
                        ANALYSIS INSTRUCTIONS:
                        1. Match work to appropriate percentage band
                        2. Reference specific rubric descriptors
                        3. Maintain academic rigor
                        4. Provide detailed justification
                        """
                        
                        with st.spinner("🔍 Analyzing submission..."):
                            response = call_deepseek_api(user_prompt, system_prompt)
                        
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
                                    parts = re.split(r':\s*', line[2:], 1)
                                    if len(parts) == 2:
                                        criterion, rest = parts
                                        band, score_comment = rest.split(',', 1)
                                        score, comment = score_comment.split('/', 1)
                                        scores[criterion.strip()] = {
                                            'Selected Band': band.strip(),
                                            'Score': score.strip(),
                                            'Comment': comment.strip()
                                        }
                                elif current_section == 'overall':
                                    overall_comments.append(line)
                                elif current_section == 'feedforward' and line.startswith('-'):
                                    feedforward.append(line[2:].strip())
                        
                        # Update rubric dataframe
                        for criterion, data in scores.items():
                            mask = rubric_df['Criteria'] == criterion
                            if mask.any():
                                rubric_df.loc[mask, 'Score'] = data['Score']
                                rubric_df.loc[mask, 'Comment'] = data['Comment']
                        
                        # Generate feedback
                        feedback_doc = generate_feedback_document(
                            rubric_df.fillna(''),
                            "\n".join(overall_comments).strip(),
                            "\n".join(feedforward)
                        )
                        
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
    

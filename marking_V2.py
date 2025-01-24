import streamlit as st
import pandas as pd
import requests
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Inches
from docx.oxml.shared import OxmlElement, nsdecls
from io import BytesIO
from PyPDF2 import PdfReader
from pptx import Presentation
import re

# Configuration
DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"
ALLOWED_EXTENSIONS = ["docx", "pdf", "pptx"]

def set_document_format(doc):
    """Configure document layout with correct 2cm margins"""
    section = doc.sections[0]
    section.orientation = 1  # Landscape
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    
    # Set all margins to 2cm (≈0.79 inches)
    section.top_margin = Inches(0.79)
    section.bottom_margin = Inches(0.79)
    section.left_margin = Inches(0.79)
    section.right_margin = Inches(0.79)
    
    # Set default font
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(12)

def read_file_content(uploaded_file) -> str:
    """Extract text from supported file formats"""
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
    """Process rubric while preserving original column names"""
    try:
        # Read and clean CSV
        df = pd.read_csv(uploaded_file, skip_blank_lines=False)
        df = df.dropna(how='all', axis=0)
        df.columns = [col.strip() for col in df.columns]

        # Validate required columns
        required_columns = [
            'Criteria', 'Criteria weighting', '80-100%', '70-79%',
            '60-69%', '50-59%', '40-49%', '0-39%',
            'Criteria Score', 'Brief Comment'
        ]
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            raise ValueError(f"Missing columns: {', '.join(missing)}")

        # Process weighting column
        if df['Criteria weighting'].dtype == object:
            df['Weighting'] = df['Criteria weighting'].str.replace('%', '').astype(float) / 100
        else:
            df['Weighting'] = df['Criteria weighting'].astype(float)

        # Validate weighting sum
        total_weight = round(df['Weighting'].sum(), 2)
        if total_weight != 1.0:
            raise ValueError(f"Total weighting must be 100% (current: {total_weight*100}%)")

        return df.fillna('')

    except Exception as e:
        st.error(f"Rubric Error: {str(e)}")
        st.markdown("""
        **Required CSV Format:**
        - Must maintain all columns even if empty
        - Preserve exact column names and order
        - Keep Criteria Score and Brief Comment columns present
        """)
        st.stop()
        raise

def call_deepseek_api(prompt: str, system_prompt: str) -> str:
    """Execute API call with error handling"""
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

def calculate_overall_score(rubric_df):
    """Compute weighted total score"""
    try:
        rubric_df['Numerical Score'] = rubric_df['Criteria Score'].str.extract(r'(\d+)').astype(float)
        total = (rubric_df['Numerical Score'] * rubric_df['Weighting']).sum()
        return round(total, 1)
    except Exception as e:
        st.error(f"Score calculation error: {str(e)}")
        return 0.0

def add_shading(cell):
    """Add light green shading to a table cell"""
    shading = OxmlElement('w:shd')
    shading.set(nsdecls('w'), 'fill', '90EE90')
    cell._tc.get_or_add_tcPr().append(shading)

def generate_feedback_document(rubric_df: pd.DataFrame, overall_comments: str, feedforward: str, overall_score: float) -> bytes:
    """Generate formatted feedback document with all elements"""
    try:
        doc = Document()
        set_document_format(doc)
        
        # Student information header
        doc.add_heading('Student Feedback', 0)
        para = doc.add_paragraph()
        para.add_run("Name: ").bold = True
        para.add_run("[Student Name]")
        
        # Rubric table
        doc.add_heading('Assessment Rubric', 1)
        cols = [
            'Criteria', 'Criteria weighting', '80-100%', '70-79%',
            '60-69%', '50-59%', '40-49%', '0-39%', 
            'Criteria Score', 'Brief Comment'
        ]
        
        table = doc.add_table(rows=1, cols=len(cols))
        table.style = 'Table Grid'
        
        # Set column widths
        col_widths = [Inches(1.5), Inches(0.7)] + [Inches(1.2)]*6 + [Inches(0.7), Inches(1.5)]
        for i, width in enumerate(col_widths):
            table.columns[i].width = width

        # Header row
        for i, col in enumerate(cols):
            cell = table.cell(0, i)
            cell.text = col
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.paragraphs[0].runs[0].font.bold = True

        # Data rows
        for _, row in rubric_df.iterrows():
            row_cells = table.add_row().cells
            for i, col in enumerate(cols):
                cell = row_cells[i]
                cell.text = str(row[col])
                
                # Add shading to score and comment cells
                if col in ['Criteria Score', 'Brief Comment'] and str(row[col]).strip():
                    add_shading(cell)

        # Overall score section
        doc.add_heading('Overall Mark', 1)
        para = doc.add_paragraph()
        para.add_run(f"Final Grade: {overall_score}%").bold = True
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Feedback sections
        doc.add_heading('Overall Comments', 1)
        doc.add_paragraph(overall_comments)

        doc.add_heading('Feedforward Suggestions', 1)
        for point in feedforward.split('\n'):
            if point.strip():
                doc.add_paragraph(point.strip(), style='ListBullet')

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer.getvalue()
    except Exception as e:
        st.error(f"Document generation failed: {str(e)}")
        raise

def main():
    st.set_page_config(page_title="AutoGrader Pro", layout="wide")
    st.title("📚 Automated Assignment Grading System")
    
    # Authentication
    if 'authenticated' not in st.session_state:
        with st.container():
            password = st.text_input("Enter password:", type='password')
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
        rubric_file = st.file_uploader("Upload Rubric CSV", type=['csv'])
        assignment_task = st.text_area("Assignment Description", height=150)
        level = st.selectbox("Academic Level", [
            "Undergraduate Level 4", "Undergraduate Level 5", 
            "Undergraduate Level 6", "Masters Level 7", "PhD Level 8"
        ])
    
    st.header("📤 Student Submissions")
    student_files = st.file_uploader(
        "Upload Assignments",
        type=ALLOWED_EXTENSIONS,
        accept_multiple_files=True
    )
    
    if st.button("🚀 Start Grading") and rubric_file and student_files:
        try:
            rubric_df = process_rubric(rubric_file)
            
            for uploaded_file in student_files:
                with st.expander(f"Processing {uploaded_file.name}", expanded=True):
                    try:
                        content = read_file_content(uploaded_file)
                        
                        system_prompt = f"""
                        As an academic assessor, evaluate using:
                        - Level: {level}
                        - Task: {assignment_task}
                        - Rubric:
                        {rubric_df.to_csv(index=False)}
                        
                        Required Format:
                        SCORES:
                        - [Criterion]: [Selected Band], [Score%], [Comment]
                        OVERALL_COMMENTS:
                        [Comprehensive evaluation]
                        FEEDFORWARD:
                        - [Actionable suggestion 1]
                        - [Actionable suggestion 2]
                        """
                        
                        user_prompt = f"""
                        Submission Content:
                        {content[:10000]}... [truncated if long]
                        
                        Analysis Guidelines:
                        1. Match exact percentage band descriptors
                        2. Provide percentage scores in 'Criteria Score' column
                        3. Add specific comments in 'Brief Comment' column
                        4. Reference concrete examples from the text
                        5. Maintain academic rigor in feedback
                        """
                        
                        with st.spinner("Analyzing..."):
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
                                    match = re.match(r"- (.+?): (.+?), (\d+)%, (.+)", line)
                                    if match:
                                        criterion, band, score, comment = match.groups()
                                        scores[criterion.strip()] = {
                                            'Criteria Score': f"{score}%",
                                            'Brief Comment': comment.strip()
                                        }
                                elif current_section == 'overall':
                                    overall_comments.append(line)
                                elif current_section == 'feedforward' and line.startswith('-'):
                                    feedforward.append(line[2:].strip())
                        
                        # Update rubric dataframe
                        for criterion, data in scores.items():
                            mask = rubric_df['Criteria'] == criterion
                            if mask.any():
                                rubric_df.loc[mask, 'Criteria Score'] = data['Criteria Score']
                                rubric_df.loc[mask, 'Brief Comment'] = data['Brief Comment']
                        
                        # Calculate overall score
                        overall_score = calculate_overall_score(rubric_df)
                        
                        # Generate document
                        feedback_doc = generate_feedback_document(
                            rubric_df,
                            "\n".join(overall_comments).strip(),
                            "\n".join(feedforward),
                            overall_score
                        )
                        
                        st.download_button(
                            label=f"Download Feedback - {uploaded_file.name}",
                            data=feedback_doc,
                            file_name=f"feedback_{uploaded_file.name.split('.')[0]}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    
                    except Exception as e:
                        st.error(f"Processing failed: {str(e)}")
        
        except Exception as e:
            st.error(f"Grading error: {str(e)}")

if __name__ == "__main__":
    main()
    
    

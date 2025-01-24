import streamlit as st
import pandas as pd
import requests
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import nsdecls
from io import BytesIO
from PyPDF2 import PdfReader
from pptx import Presentation
import re

# Configuration
DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"
ALLOWED_EXTENSIONS = ["docx", "pdf", "pptx"]

def set_document_format(doc):
    """Configure document layout with correct margins and font"""
    section = doc.sections[0]
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    
    # Set margins to 2cm
    margins = Inches(0.79)
    section.top_margin = margins
    section.bottom_margin = margins
    section.left_margin = margins
    section.right_margin = margins
    
    # Set default font
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(12)

def read_file_content(uploaded_file) -> str:
    """Extract text from supported file formats"""
    content = ""
    try:
        file_bytes = BytesIO(uploaded_file.getvalue())
        if uploaded_file.name.endswith('.docx'):
            doc = Document(file_bytes)
            content = "\n".join([para.text for para in doc.paragraphs])
        elif uploaded_file.name.endswith('.pdf'):
            reader = PdfReader(file_bytes)
            content = "\n".join([page.extract_text() or "" for page in reader.pages])
        elif uploaded_file.name.endswith('.pptx'):
            prs = Presentation(file_bytes)
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        content += shape.text + "\n"
        return content.strip()
    except Exception as e:
        st.error(f"Error reading {uploaded_file.name}: {str(e)}")
        raise

def process_rubric(uploaded_file):
    """Process and validate rubric CSV"""
    try:
        df = pd.read_csv(uploaded_file)
        df = df.dropna(how='all', axis=0).reset_index(drop=True)
        df.columns = [col.strip() for col in df.columns]

        required_columns = [
            'Criteria', 'Criteria weighting', '80-100%', '70-79%',
            '60-69%', '50-59%', '40-49%', '0-39%',
            'Criteria Score', 'Brief Comment'
        ]
        
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            raise ValueError(f"Missing columns: {', '.join(missing)}")

        df['Weighting'] = (df['Criteria weighting']
                           .astype(str)
                           .str.replace('%', '')
                           .astype(float) / 100)
        
        total_weight = round(df['Weighting'].sum(), 2)
        if total_weight != 1.0:
            raise ValueError(f"Total weighting must be 100% (current: {total_weight*100}%)")

        return df.fillna('')

    except Exception as e:
        st.error(f"Rubric Error: {str(e)}")
        st.markdown("""
        **Required CSV Format:**
        - Maintain all original columns even if empty
        - Preserve exact column names and order
        - Include Criteria Score and Brief Comment columns
        """)
        st.stop()

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
    except requests.exceptions.RequestException as err:
        st.error(f"API Error: {err}\nResponse: {response.text if response else 'No response'}")
        raise

def calculate_overall_score(rubric_df):
    """Compute weighted total score"""
    try:
        rubric_df['Numerical Score'] = (rubric_df['Criteria Score']
                                        .str.extract(r'(\d+)', expand=False)
                                        .astype(float))
        total = (rubric_df['Numerical Score'] * rubric_df['Weighting']).sum()
        return min(max(round(total, 1), 0), 100)
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
        
        # Header
        header = doc.add_heading('Student Feedback', 0)
        header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Student info table
        info_table = doc.add_table(rows=1, cols=2)
        info_table.style = 'Table Grid'
        info_row = info_table.rows[0].cells
        info_row[0].text = "Student Name:\nSubmission Date:\nCourse:"
        info_row[1].text = "[Name]\n[Date]\n[Course]"

        # Rubric table
        doc.add_heading('Assessment Rubric', 1)
        cols = [
            'Criteria', 'Criteria weighting', '80-100%', '70-79%',
            '60-69%', '50-59%', '40-49%', '0-39%', 
            'Criteria Score', 'Brief Comment'
        ]
        
        table = doc.add_table(rows=1, cols=len(cols))
        table.style = 'Table Grid'
        table.autofit = False

        # Set column widths
        col_widths = [Inches(1.5), Inches(0.7)] + [Inches(1.2)]*6 + [Inches(0.7), Inches(1.5)]
        for i, width in enumerate(col_widths):
            table.columns[i].width = width

        # Header row
        header_row = table.rows[0]
        for i, col in enumerate(cols):
            cell = header_row.cells[i]
            cell.text = col
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.paragraphs[0].runs[0].font.bold = True

        # Data rows with shading
        for _, row in rubric_df.iterrows():
            new_row = table.add_row().cells
            for i, col in enumerate(cols):
                cell = new_row[i]
                cell.text = str(row[col])
                if col in ['Criteria Score', 'Brief Comment'] and str(row[col]).strip():
                    add_shading(cell)

        # Overall score section
        doc.add_heading('Overall Mark', 1)
        score_para = doc.add_paragraph()
        score_run = score_para.add_run(f"Final Grade: {overall_score}%")
        score_run.bold = True
        score_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Feedback sections
        sections = [
            ('Overall Comments', overall_comments),
            ('Feedforward Suggestions', feedforward)
        ]
        
        for section, content in sections:
            doc.add_heading(section, 1)
            if section == 'Feedforward Suggestions':
                for point in filter(None, content.split('\n')):
                    doc.add_paragraph(point.strip(), style='ListBullet')
            else:
                p = doc.add_paragraph()
                p.add_run(content.strip())

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
    
    if rubric_file and student_files and st.button("🚀 Start Grading"):
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
                        - [Criterion Name]: [Selected Band], [Score%], [Comment]
                        OVERALL_COMMENTS:
                        [Comprehensive evaluation in 3-5 paragraphs]
                        FEEDFORWARD:
                        - [Actionable suggestion 1]
                        - [Actionable suggestion 2]
                        - [Actionable suggestion 3]
                        """
                        
                        user_prompt = f"""
                        Submission Content:
                        {content[:10000]}

                        Analysis Guidelines:
                        1. Match exact percentage band descriptors
                        2. Provide specific examples from the text
                        3. Maintain academic rigor in feedback
                        4. Use British English spelling
                        5. Reference university guidelines where appropriate
                        """
                        
                        with st.spinner("Generating detailed feedback..."):
                            response = call_deepseek_api(user_prompt, system_prompt)
                        
                        # Enhanced response parsing
                        scores = {}
                        sections = {
                            'SCORES:': 'scores',
                            'OVERALL_COMMENTS:': 'overall',
                            'FEEDFORWARD:': 'feedforward'
                        }
                        current_section = None
                        overall_comments = []
                        feedforward = []

                        for line in response.split('\n'):
                            line = line.strip()
                            if line in sections:
                                current_section = sections[line]
                            elif current_section == 'scores' and line.startswith('-'):
                                match = re.match(r"- (.+?): (.+?), (\d+)%, (.+)", line)
                                if match:
                                    criterion, band, score, comment = match.groups()
                                    scores[criterion.strip()] = {
                                        'Criteria Score': f"{score}%",
                                        'Brief Comment': comment.strip()
                                    }
                            elif current_section == 'overall':
                                if line:  # Skip empty lines
                                    overall_comments.append(line)
                            elif current_section == 'feedforward' and line.startswith('-'):
                                feedforward.append(line[2:].strip())

                        # Update rubric dataframe
                        rubric_df['Criteria Score'] = ''
                        rubric_df['Brief Comment'] = ''
                        for criterion, data in scores.items():
                            mask = rubric_df['Criteria'].str.strip().str.lower() == criterion.strip().lower()
                            if mask.any():
                                idx = rubric_df[mask].index[0]
                                rubric_df.at[idx, 'Criteria Score'] = data['Criteria Score']
                                rubric_df.at[idx, 'Brief Comment'] = data['Brief Comment']

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
                            label=f"📥 Download Feedback - {uploaded_file.name}",
                            data=feedback_doc,
                            file_name=f"feedback_{uploaded_file.name.split('.')[0]}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    
                    except Exception as e:
                        st.error(f"Processing failed for {uploaded_file.name}: {str(e)}")

        except Exception as e:
            st.error(f"Grading process failed: {str(e)}")

if __name__ == "__main__":
    main()
    

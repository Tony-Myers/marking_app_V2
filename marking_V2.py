import streamlit as st
import pandas as pd
import requests
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import nsdecls
from io import BytesIO, StringIO
from PyPDF2 import PdfReader
import re
import tiktoken
import os

# Configuration
DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"
ALLOWED_EXTENSIONS = ["docx", "pdf"]
MAX_TOKENS = 7000
PROMPT_BUFFER = 1000

# Initialize encoding
try:
    encoding = tiktoken.encoding_for_model("deepseek-chat")
except KeyError:
    encoding = tiktoken.get_encoding("cl100k_base")

def count_tokens(text):
    return len(encoding.encode(text))

def truncate_text(text, max_tokens):
    tokens = encoding.encode(text)
    return encoding.decode(tokens[:max_tokens]) if len(tokens) > max_tokens else text

def call_deepseek_api(prompt, system_prompt):
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
        "temperature": 0.3,
        "max_tokens": 3000
    }
    
    try:
        response = requests.post(DEEPSEEK_API_URL, json=data, headers=headers)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"].strip()
    except Exception as e:
        st.error(f"API Error: {str(e)}")
        return None

def extract_text_from_docx(file):
    try:
        doc = Document(file)
        return '\n'.join([para.text for para in doc.paragraphs])
    except Exception as e:
        st.error(f"DOCX Error: {str(e)}")
        return None

def extract_text_from_pdf(file):
    try:
        reader = PdfReader(file)
        return '\n'.join([page.extract_text() for page in reader.pages])
    except Exception as e:
        st.error(f"PDF Error: {str(e)}")
        return None

def parse_csv_section(csv_text):
    """Enhanced CSV parsing with column validation"""
    try:
        df = pd.read_csv(StringIO(csv_text), quotechar='"', skipinitialspace=True)
        
        # Normalize column names
        df.columns = df.columns.str.strip().str.lower()
        
        # Validate required columns
        required_columns = {'criterion', 'score', 'comment'}
        missing = required_columns - set(df.columns)
        if missing:
            st.error(f"CSV missing columns: {', '.join(missing)}")
            return None
            
        return df
    except Exception as e:
        st.error(f"CSV Parse Error: {str(e)}")
        return None

def extract_weight(criterion_name):
    match = re.search(r'\((\d+)%\)', criterion_name)
    return float(match.group(1)) if match else 0.0

def add_shading(cell):
    shading = OxmlElement('w:shd')
    shading.set(nsdecls('w'), 'fill', 'D9EAD3')
    cell._tc.get_or_add_tcPr().append(shading)

def generate_feedback_doc(student_name, rubric_df, overall_comments, feedforward, total_mark):
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    
    # Header
    doc.add_heading(f"Feedback for {student_name}", 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Rubric Table
    table = doc.add_table(rows=1, cols=len(rubric_df.columns))
    table.style = 'Table Grid'
    
    # Header row
    hdr_cells = table.rows[0].cells
    for i, col in enumerate(rubric_df.columns):
        hdr_cells[i].text = str(col)
        hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Data rows with shading
    for _, row in rubric_df.iterrows():
        row_cells = table.add_row().cells
        for i, value in enumerate(row):
            cell = row_cells[i]
            cell.text = str(value)
            if 'score' in str(rubric_df.columns[i]).lower():
                try:
                    score = float(value)
                    for range_col in [col for col in rubric_df.columns if '-' in col and '%' in col]:
                        lower, upper = map(float, range_col.replace('%','').split('-'))
                        if lower <= score <= upper:
                            add_shading(cell)
                            break
                except ValueError:
                    pass
    
    # Feedback sections
    doc.add_heading('Overall Comments', 1)
    doc.add_paragraph(overall_comments)
    
    doc.add_heading('Feedforward', 1)
    for point in feedforward.split('\n'):
        if point.strip().startswith('-'):
            doc.add_paragraph(point.strip()[2:], style='ListBullet')
    
    doc.add_heading('Total Mark', 1)
    doc.add_paragraph(f"{total_mark:.2f}%").bold = True
    
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def main():
    st.set_page_config(page_title="AutoGrader Pro", layout="wide")
    st.title("✏️ Automated Assignment Grading © Tony Myers (DeepSeek Version)")
    
    # Authentication
    if 'authenticated' not in st.session_state:
        password = st.text_input("Enter password:", type="password")
        if st.button("Authenticate") and password == st.secrets["APP_PASSWORD"]:
            st.session_state.authenticated = True
            st.rerun()
        return
    
    st.header("Assignment Configuration")
    assignment_task = st.text_area("Assignment Task & Level", height=150)
    
    st.header("Upload Files")
    rubric_file = st.file_uploader("Rubric (CSV)", type=['csv'])
    submissions = st.file_uploader("Submissions", type=ALLOWED_EXTENSIONS, accept_multiple_files=True)
    
    if rubric_file and submissions and st.button("Start Marking"):
        try:
            # Read and process rubric
            rubric_df = pd.read_csv(rubric_file)
            
            # Clean and normalize columns
            rubric_df.columns = rubric_df.columns.str.strip().str.lower()
            rubric_df['criterion'] = rubric_df['criterion'].astype(str)
            rubric_df['weight'] = rubric_df['criterion'].apply(extract_weight)
            rubric_df['criterion'] = rubric_df['criterion'].apply(lambda x: re.sub(r'\s*\(\d+%\)', '', x).strip())
            
            percentage_columns = [col for col in rubric_df.columns if '%' in col]
            criteria_string = '\n'.join(rubric_df['criterion'].tolist())
            
            for submission in submissions:
                student_name = os.path.splitext(submission.name)[0]
                
                # Extract text
                if submission.type == "application/pdf":
                    text = extract_text_from_pdf(submission)
                else:
                    text = extract_text_from_docx(submission)
                
                if not text:
                    st.error(f"Failed to extract text from {submission.name}")
                    continue
                
                # Truncate if needed
                if count_tokens(text) > MAX_TOKENS * 0.6:
                    text = truncate_text(text, int(MAX_TOKENS * 0.6))
                
                # Prepare prompts
                system_prompt = f"""You are an experienced UK academic. Provide strict feedback using:
- British English spelling
- Birmingham Newman University guidelines
- Second person narrative"""
                
                user_prompt = f"""
**Task:** {assignment_task}
**Submission:** {text[:10000]}
**Rubric Criteria:** {criteria_string}

Generate feedback with:
1. CSV starting with 'Criterion,Score,Comment'
2. Overall Comments (150 words)
3. Feedforward (bullet points)

**Example:**
Criterion,Score,Comment
"Linking Theory",75,"Good analysis but needs more depth"
...
"""
                # API Call
                response = call_deepseek_api(user_prompt, system_prompt)
                if not response:
                    continue
                
                # Parse response
                try:
                    csv_part = response.split('Overall Comments:')[0].strip()
                    comments_part = response.split('Overall Comments:')[1].split('Feedforward:')
                    overall_comments = comments_part[0].strip()
                    feedforward = comments_part[1].strip()
                except IndexError:
                    st.error("Invalid response format")
                    continue
                
                # Process scores
                scores_df = parse_csv_section(csv_part)
                if scores_df is None or scores_df.empty:
                    st.error("Invalid scores data")
                    continue
                
                # Merge dataframes
                try:
                    merged_df = rubric_df.merge(
                        scores_df[['criterion', 'score', 'comment']],
                        on='criterion',
                        how='left'
                    )
                    merged_df['weighted'] = merged_df['weight'] * merged_df['score'] / 100
                    total_mark = merged_df['weighted'].sum()
                except KeyError as e:
                    st.error(f"Merge error: {str(e)}")
                    continue
                
                # Generate document
                doc_buffer = generate_feedback_doc(
                    student_name,
                    merged_df[['criterion'] + percentage_columns + ['score', 'comment']],
                    overall_comments,
                    feedforward,
                    total_mark
                )
                
                st.download_button(
                    label=f"Download {student_name} Feedback",
                    data=doc_buffer,
                    file_name=f"{student_name}_feedback.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        except Exception as e:
            st.error(f"Processing Error: {str(e)}")

if __name__ == "__main__":
    main()

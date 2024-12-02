import streamlit as st
import tempfile
from pathlib import Path
from pptx import Presentation
import language_tool_python
import csv
import re

# LanguageTool API initialization
def initialize_language_tool():
    try:
        return language_tool_python.LanguageToolPublicAPI('en-US')  # Use Public API mode
    except Exception as e:
        st.error(f"LanguageTool initialization failed: {e}")
        return None

grammar_tool = initialize_language_tool()

# Function to validate punctuation issues
def detect_punctuation_issues(text):
    excessive_punctuation_pattern = r"([!?.:,;]{2,})"  # Two or more punctuation marks
    repeated_word_pattern = r"\b(\w+)\s+\1\b"  # Repeated words (e.g., "the the")
    
    if re.search(excessive_punctuation_pattern, text):
        return "Punctuation Issue", "Excessive punctuation marks detected"
    if re.search(repeated_word_pattern, text, flags=re.IGNORECASE):
        return "Punctuation Issue", "Repeated words detected"
    
    return None, None

# Function to validate grammar and spelling using LanguageTool
def validate_language_tool(input_ppt):
    presentation = Presentation(input_ppt)
    language_issues = []

    for slide_index, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        text = run.text.strip()
                        if text:
                            # Check for punctuation issues first
                            punctuation_issue, punctuation_correction = detect_punctuation_issues(text)
                            if punctuation_issue:
                                language_issues.append({
                                    'slide': slide_index,
                                    'issue': punctuation_issue,
                                    'text': text,
                                    'corrected': punctuation_correction
                                })
                                continue  # Skip further checks for punctuation issues

                            # Use LanguageTool to check for grammar and spelling issues
                            if grammar_tool:
                                matches = grammar_tool.check(text)
                                if matches:
                                    corrected_text = language_tool_python.utils.correct(text, matches)
                                    language_issues.append({
                                        'slide': slide_index,
                                        'issue': 'Grammar or Spelling Error',
                                        'text': text,
                                        'corrected': corrected_text
                                    })
    return language_issues

# Function to validate fonts in a presentation
def validate_fonts(input_ppt, default_font):
    presentation = Presentation(input_ppt)
    font_issues = []

    for slide_index, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.text.strip():
                            if run.font.name != default_font:
                                font_issues.append({
                                    'slide': slide_index,
                                    'issue': 'Inconsistent Font',
                                    'text': run.text,
                                    'corrected': f"Expected font: {default_font}"
                                })
    return font_issues

# Function to save issues to CSV
def save_to_csv(issues, output_csv):
    with open(output_csv, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=['slide', 'issue', 'text', 'corrected'])
        writer.writeheader()
        writer.writerows(issues)

# Main Streamlit app
def main():
    st.title("PPT Validator")

    uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["pptx"])
    font_options = ["Arial", "Calibri", "Times New Roman", "Verdana", "Helvetica"]
    default_font = st.selectbox("Select the default font for validation", font_options)

    if uploaded_file and st.button("Run Validation"):
        with tempfile.TemporaryDirectory() as tmpdir:
            temp_ppt_path = Path(tmpdir) / "uploaded_ppt.pptx"
            with open(temp_ppt_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            csv_output_path = Path(tmpdir) / "validation_report.csv"

            font_issues = validate_fonts(temp_ppt_path, default_font)
            language_issues = validate_language_tool(temp_ppt_path)

            all_issues = font_issues + language_issues
            save_to_csv(all_issues, csv_output_path)

            st.success("Validation completed!")
            st.download_button("Download Validation Report (CSV)", csv_output_path.read_bytes(),
                               file_name="validation_report.csv")

if __name__ == "__main__":
    main()

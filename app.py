import streamlit as st
import tempfile
from pathlib import Path
from pptx import Presentation
import language_tool_python
import csv
import re
import string

# Initialize LanguageTool
def initialize_language_tool():
    try:
        return language_tool_python.LanguageToolPublicAPI('en-US')  # Use Public API mode
    except Exception as e:
        st.error(f"LanguageTool initialization failed: {e}")
        return None

grammar_tool = initialize_language_tool()

import re

# Updated validate_combined function
def validate_combined(input_ppt):
    presentation = Presentation(input_ppt)
    combined_issues = []

    # Define patterns for grammatical corrections
    grammar_patterns = [
        (r"\bcan overused\b", "can be overused"),
        (r"\bsentense\b", "sentence"),
        (r"\bspleling\b", "spelling"),
        # Tambahkan pola lain jika perlu
    ]

    for slide_index, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        text = run.text.strip()
                        if text:
                            # Grammar Check using LanguageTool
                            if grammar_tool:
                                matches = grammar_tool.check(text)
                                corrected = language_tool_python.utils.correct(text, matches)
                                if corrected != text:
                                    combined_issues.append({
                                        'slide': slide_index,
                                        'issue': 'Grammar or Spelling Error',
                                        'text': text,
                                        'corrected': corrected
                                    })

                            # Additional Regex-Based Grammar Fixes
                            for pattern, replacement in grammar_patterns:
                                corrected_text = re.sub(pattern, replacement, text)
                                if corrected_text != text:  # If changes are made
                                    combined_issues.append({
                                        'slide': slide_index,
                                        'issue': 'Grammar Error',
                                        'text': text,
                                        'corrected': corrected_text
                                    })
    return combined_issues


# Validate fonts in the presentation
def validate_fonts(input_ppt, default_font):
    presentation = Presentation(input_ppt)
    font_issues = []

    for slide_index, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.text.strip() and run.font.name != default_font:
                            font_issues.append({
                                'slide': slide_index,
                                'issue': 'Inconsistent Font',
                                'text': run.text,
                                'corrected': f"Expected font: {default_font}"
                            })
    return font_issues

# Save issues to a CSV file
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
            combined_issues = validate_combined(temp_ppt_path)

            all_issues = font_issues + combined_issues
            save_to_csv(all_issues, csv_output_path)

            st.success("Validation completed!")
            st.download_button("Download Validation Report (CSV)", csv_output_path.read_bytes(),
                               file_name="validation_report.csv")

if __name__ == "__main__":
    main()

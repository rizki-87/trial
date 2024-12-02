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

# Combined grammar validation function
def validate_combined(input_ppt):
    presentation = Presentation(input_ppt)
    combined_issues = []

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
                                if matches:
                                    corrected = language_tool_python.utils.correct(text, matches)
                                    if corrected != text:  # Only log if correction is made
                                        combined_issues.append({
                                            'slide': slide_index,
                                            'issue': 'Grammar Error',
                                            'text': text,
                                            'corrected': corrected
                                        })
    return combined_issues

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

    if uploaded_file and st.button("Run Validation"):
        with tempfile.TemporaryDirectory() as tmpdir:
            temp_ppt_path = Path(tmpdir) / "uploaded_ppt.pptx"
            with open(temp_ppt_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            csv_output_path = Path(tmpdir) / "validation_report.csv"

            combined_issues = validate_combined(temp_ppt_path)

            save_to_csv(combined_issues, csv_output_path)

            st.success("Validation completed!")
            st.download_button("Download Validation Report (CSV)", csv_output_path.read_bytes(),
                               file_name="validation_report.csv")

if __name__ == "__main__":
    main()

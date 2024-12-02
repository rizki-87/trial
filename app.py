import streamlit as st
import tempfile
from pathlib import Path
from pptx import Presentation
import language_tool_python
import csv
import re
import string


# LanguageTool API initialization
def initialize_language_tool():
    try:
        return language_tool_python.LanguageToolPublicAPI('en-US')  # Use Public API mode
    except Exception as e:
        st.error(f"LanguageTool initialization failed: {e}")
        return None


grammar_tool = initialize_language_tool()


# Fallback grammar check rules
def fallback_grammar_check(text):
    fallback_rules = [
        (r'\b(can|could|should|would|may|might|must)\s+(\w+)\b(?!\s+be)', r'\1 \2 be'),  # Missing 'be' after modals
        (r'\b(doesn\'t|don\'t|didn\'t|isn\'t|aren\'t|weren\'t)\s+(\w+)\b', r'\1 \2'),  # Contraction errors
        (r'\b(is|are|was|were|has|have|had|does|do|did)\s*(?!\w)', r'\1 '),  # Missing subject after auxiliaries
    ]

    for pattern, replacement in fallback_rules:
        if re.search(pattern, text):
            corrected_text = re.sub(pattern, replacement, text)
            return f"Grammar Error detected", corrected_text

    return None, None


# Function to validate combined issues
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
                            # Grammar check with LanguageTool
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

                            # Fallback grammar check
                            fallback_issue, fallback_suggestion = fallback_grammar_check(text)
                            if fallback_issue:
                                combined_issues.append({
                                    'slide': slide_index,
                                    'issue': fallback_issue,
                                    'text': text,
                                    'corrected': fallback_suggestion
                                })

    return combined_issues


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


# Function to validate punctuation
def validate_punctuation(input_ppt):
    presentation = Presentation(input_ppt)
    punctuation_issues = []

    excessive_punctuation_pattern = r"([!?.:,;]{2,})"

    for slide_index, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        text = run.text.strip()
                        if text:
                            match = re.search(excessive_punctuation_pattern, text)
                            if match:
                                punctuation_marks = match.group(1)
                                punctuation_issues.append({
                                    'slide': slide_index,
                                    'issue': 'Punctuation Marks',
                                    'text': text,
                                    'corrected': f"Excessive punctuation marks detected ({punctuation_marks})"
                                })

    return punctuation_issues


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
            punctuation_issues = validate_punctuation(temp_ppt_path)
            combined_issues = validate_combined(temp_ppt_path)

            all_issues = font_issues + punctuation_issues + combined_issues
            save_to_csv(all_issues, csv_output_path)

            st.success("Validation completed!")
            st.download_button("Download Validation Report (CSV)", csv_output_path.read_bytes(),
                               file_name="validation_report.csv")


if __name__ == "__main__":
    main()

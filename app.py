import streamlit as st
import tempfile
from pathlib import Path
from pptx import Presentation
import language_tool_python
import csv
import re

# LanguageTool initialization
def initialize_language_tool():
    try:
        return language_tool_python.LanguageToolPublicAPI('en-US')
    except Exception as e:
        st.error(f"LanguageTool initialization failed: {e}")
        return None

grammar_tool = initialize_language_tool()

# Fallback rules for grammar
def fallback_grammar_check(text):
    issues = []

    # Rule 1: Modal + Verb Auxiliary (e.g., "She can go" vs. "She can going")
    modal_pattern = r'\b(can|could|should|would|might|must)\s+\b(\w+ing)\b'
    match = re.search(modal_pattern, text, flags=re.IGNORECASE)
    if match:
        issues.append({
            'issue': 'Grammar Error',
            'text': text,
            'corrected': text.replace(match.group(2), match.group(2).rstrip('ing'))
        })

    # Rule 2: Subject-Verb Agreement (e.g., "He don't" vs. "He doesn't")
    singular_subject_pattern = r'\b(he|she|it)\s+don\'t\b'
    match = re.search(singular_subject_pattern, text, flags=re.IGNORECASE)
    if match:
        issues.append({
            'issue': 'Grammar Error',
            'text': text,
            'corrected': text.replace("don't", "doesn't")
        })

    # Rule 3: Missing Auxiliary Verbs (e.g., "She going" vs. "She is going")
    missing_aux_pattern = r'\b(she|he|they|we|i)\s+(\w+ing)\b'
    match = re.search(missing_aux_pattern, text, flags=re.IGNORECASE)
    if match:
        issues.append({
            'issue': 'Grammar Error',
            'text': text,
            'corrected': f"{match.group(1)} is {match.group(2)}"
        })

    # Rule 4: Common Verb Conjugation Errors (e.g., "He do" vs. "He does")
    verb_conjugation_pattern = r'\b(he|she|it)\s+do\b'
    match = re.search(verb_conjugation_pattern, text, flags=re.IGNORECASE)
    if match:
        issues.append({
            'issue': 'Grammar Error',
            'text': text,
            'corrected': text.replace("do", "does")
        })

    return issues

# Combined grammar and spelling validation function
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
                                corrected_text = language_tool_python.utils.correct(text, matches)
                                if corrected_text != text:
                                    combined_issues.append({
                                        'slide': slide_index,
                                        'issue': 'Grammar Error',
                                        'text': text,
                                        'corrected': corrected_text
                                    })

                            # Fallback Grammar Check
                            fallback_issues = fallback_grammar_check(text)
                            for issue in fallback_issues:
                                issue['slide'] = slide_index
                                combined_issues.append(issue)

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
                        if run.text.strip() and run.font.name != default_font:
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
                                punctuation_issues.append({
                                    'slide': slide_index,
                                    'issue': 'Punctuation Marks',
                                    'text': text,
                                    'corrected': f"Excessive punctuation marks detected ({match.group(1)})"
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

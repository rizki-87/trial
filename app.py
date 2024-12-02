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

# Validate combined logic (grammar, spelling, punctuation)
def validate_combined(input_ppt):
    presentation = Presentation(input_ppt)
    combined_issues = []

    # Auxiliary verbs for grammar correction
    auxiliary_verbs = ['is', 'are', 'was', 'were', 'has', 'have', 'had', 'be', 'being', 'been', 'does', 'do', 'did']

    for slide_index, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        text = run.text.strip()
                        if text:
                            # Grammar or Spelling Check
                            if grammar_tool:
                                matches = grammar_tool.check(text)
                                corrected = language_tool_python.utils.correct(text, matches)
                                if corrected != text:  # Only log if correction is made
                                    combined_issues.append({
                                        'slide': slide_index,
                                        'issue': 'Grammar or Spelling Error',
                                        'text': text,
                                        'corrected': corrected
                                    })

                            # Missing Auxiliary Verbs Check
                            missing_auxiliary_pattern = r"\b(This|That|These|Those|It|They|We|I|You|He|She|Someone)\s+[a-zA-Z]+(\s+[a-zA-Z]+)*(\s+(a|an|the)\s+[a-zA-Z]+)?$"
                            missing_auxiliary_match = re.search(missing_auxiliary_pattern, text)
                            if missing_auxiliary_match:
                                suggested_correction = f"{text.split()[0]} {auxiliary_verbs[0]} {' '.join(text.split()[1:])}"
                                combined_issues.append({
                                    'slide': slide_index,
                                    'issue': 'Grammar Error: Missing Auxiliary Verb',
                                    'text': text,
                                    'corrected': suggested_correction
                                })

                            # Custom Rules for Missing "be" or "has"
                            custom_patterns = [
                                (r"\b(can|could|should|would|may|might|must)\s+\b(\w+)", r"\1 be \2"),  # Modal missing 'be'
                                (r"\b(has|have|had)\s+\b(\w+)", r"\1 been \2"),  # 'Has' missing 'been'
                                (r"\b(This|That|These|Those|He|She|It|They|We|You|I)\s+[a-zA-Z]+\b", r"\1 is \2")  # Generic missing 'is'
                            ]
                            for pattern, replacement in custom_patterns:
                                match = re.search(pattern, text)
                                if match:
                                    corrected_text = re.sub(pattern, replacement, text)
                                    combined_issues.append({
                                        'slide': slide_index,
                                        'issue': 'Grammar Error: Missing Verb',
                                        'text': text,
                                        'corrected': corrected_text
                                    })

                            # Punctuation Check (Excessive Punctuation)
                            excessive_punctuation_pattern = r"([!?.:,;]{2,})"
                            match = re.search(excessive_punctuation_pattern, text)
                            if match:
                                punctuation_marks = match.group(1)
                                combined_issues.append({
                                    'slide': slide_index,
                                    'issue': 'Punctuation Marks',
                                    'text': text,
                                    'corrected': f"Excessive punctuation marks detected ({punctuation_marks})"
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

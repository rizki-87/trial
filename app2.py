# app.py

import streamlit as st
import tempfile
from pathlib import Path
from pptx import Presentation
from spellchecker import SpellChecker
import language_tool_python
import csv
import re
import string
from pptx.dml.color import RGBColor
import logging
from pydantic import BaseModel
from concurrent.futures import ThreadPoolExecutor  # Tambahkan impor ini
from utils.validation import highlight_ppt, save_to_csv
from utils.font_validation import validate_fonts_slide
from utils.grammar_validation import initialize_language_tool, validate_grammar_slide
from utils.spelling_validation import is_exempted, validate_spelling_slide
from utils.decimal_validation import validate_decimal_consistency
from utils.million_notation_validation import validate_million_notations
from config import PREDEFINED_PASSWORD, TECHNICAL_TERMS, NUMERIC_TERMS

# Initialize LanguageTool
grammar_tool = initialize_language_tool()

# Initialize SpellChecker
spell = SpellChecker()
spell.word_frequency.load_words(TECHNICAL_TERMS.union(NUMERIC_TERMS))

# Password Protection
def password_protection():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if not st.session_state.authenticated:
        with st.form("password_form", clear_on_submit=True):
            password_input = st.text_input("Enter Password", type="password")
            submitted = st.form_submit_button("Submit")
            if submitted and password_input == PREDEFINED_PASSWORD:
                st.session_state.authenticated = True
                st.success("Access Granted! Please click 'Submit' again to proceed.")
            elif submitted:
                st.error("Incorrect Password")
        return False
    return True

def main():
    if not password_protection():
        return

    st.title("PPT Validator")
    uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["pptx"])
    font_options = ["Arial", "Calibri", "Times New Roman", "Verdana", "Helvetica", "EYInterstate"]
    default_font = st.selectbox("Select the default font for validation", font_options)
    validation_option = st.radio("Validation Option:", ["All Slides", "Custom Range"])

    if uploaded_file:
        with tempfile.TemporaryDirectory() as tmpdir:
            temp_ppt_path = Path(tmpdir) / "uploaded_ppt.pptx"
            with open(temp_ppt_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            presentation = Presentation(temp_ppt_path)
            total_slides = len(presentation.slides)

            # Slide Range Selection
            start_slide, end_slide = 1, total_slides
            if validation_option == "Custom Range":
                start_slide = st.number_input("From Slide", min_value=1, max_value=total_slides, value=1)
                end_slide_default = min(total_slides, 100)  # Ensure default value does not exceed total slides
                end_slide = st.number_input("To Slide", min_value=start_slide, max_value=total_slides, value=end_slide_default)

            if st.button("Run Validation"):
                progress_bar = st.progress(0)
                progress_text = st.empty()
                issues = []
                reference_decimal_points = None

                # Parallel Processing
                with ThreadPoolExecutor() as executor:
                    futures = []
                    for slide_index in range(start_slide - 1, end_slide):
                        slide = presentation.slides[slide_index]
                        futures.append(executor.submit(validate_spelling_slide, slide, slide_index + 1, spell))
                        futures.append(executor.submit(validate_fonts_slide, slide, slide_index + 1, default_font))
                        futures.append(executor.submit(validate_grammar_slide, slide, slide_index + 1, grammar_tool))
                        futures.append(executor.submit(validate_decimal_consistency, slide, slide_index + 1, reference_decimal_points))
                        futures.append(executor.submit(validate_million_notations, slide, slide_index + 1))  # Added function call

                    for i, future in enumerate(futures):
                        result = future.result()
                        if isinstance(result, tuple):
                            slide_issues, reference_decimal_points = result
                            issues.extend(slide_issues)
                        else:
                            issues.extend(result)
                        progress_percent = int((i + 1) / len(futures) * 100)
                        progress_text.text(f"Progress: {progress_percent}%")
                        progress_bar.progress(progress_percent / 100)

                # Save Results
                csv_output_path = Path(tmpdir) / "validation_report.csv"
                highlighted_ppt_path = Path(tmpdir) / "highlighted_presentation.pptx"
                save_to_csv(issues, csv_output_path)
                highlight_ppt(temp_ppt_path, highlighted_ppt_path, issues)

                # Store results in session state
                st.session_state['csv_output'] = csv_output_path.read_bytes()
                st.session_state['ppt_output'] = highlighted_ppt_path.read_bytes()
                st.success("Validation completed!")

                # Display Download Buttons
                if 'csv_output' in st.session_state:
                    st.download_button("Download Validation Report (CSV)", st.session_state['csv_output'], file_name="validation_report.csv")
                if 'ppt_output' in st.session_state:
                    st.download_button("Download Highlighted PPT", st.session_state['ppt_output'], file_name="highlighted_presentation.pptx")

if __name__ == "__main__":
    main()
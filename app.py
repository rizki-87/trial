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

# LanguageTool API initialization
def initialize_language_tool():
    try:
        return language_tool_python.LanguageToolPublicAPI('en-US')  # Use Public API mode
    except Exception as e:
        st.error(f"LanguageTool initialization failed: {e}")
        return None

grammar_tool = initialize_language_tool()

# Function to highlight issues in a PPT
def highlight_ppt(input_ppt, output_ppt, issues):
    presentation = Presentation(input_ppt)
    for issue in issues:
        slide_index = issue['slide'] - 1  # Slide index starts at 0
        slide = presentation.slides[slide_index]
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if issue['text'] in run.text:
                            run.font.color.rgb = RGBColor(255, 255, 0)  # Highlight text in yellow
    presentation.save(output_ppt)

# Function to validate grammar issues
def validate_grammar(input_ppt):
    presentation = Presentation(input_ppt)
    grammar_issues = []

    for slide_index, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        text = run.text.strip()
                        if text:
                            if grammar_tool:
                                matches = grammar_tool.check(text)
                                if matches:
                                    corrected = language_tool_python.utils.correct(text, matches)
                                    if corrected != text:  # Log only if correction is made
                                        grammar_issues.append({
                                            'slide': slide_index,
                                            'issue': 'Grammar Error',
                                            'text': text,
                                            'corrected': corrected
                                        })
    return grammar_issues

# Function to detect and correct misspellings
def validate_spelling(input_ppt):
    presentation = Presentation(input_ppt)
    spelling_issues = []
    spell = SpellChecker()

    for slide_index, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.text.strip():
                            words = run.text.split()
                            for word in words:
                                clean_word = word.strip(string.punctuation)
                                if clean_word and clean_word.lower() not in spell:
                                    correction = spell.correction(clean_word)
                                    if correction:  # Only log valid corrections
                                        spelling_issues.append({
                                            'slide': slide_index,
                                            'issue': 'Misspelling',
                                            'text': f"Original: {word}",
                                            'corrected': f"Suggestion: {correction}"
                                        })
    return spelling_issues

# Function to validate fonts
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
                                'corrected': f"Detected: {run.font.name}, Expected: {default_font}"
                            })
    return font_issues

# Function to validate punctuation
def validate_punctuation(input_ppt):
    presentation = Presentation(input_ppt)
    punctuation_issues = []

    excessive_punctuation_pattern = r"([!?.:,;]{2,})"
    repeated_word_pattern = r"\b(\w+)\s+\1\b"

    for slide_index, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.text.strip():
                            text = run.text
                            if re.search(excessive_punctuation_pattern, text):
                                punctuation_issues.append({
                                    'slide': slide_index,
                                    'issue': 'Punctuation Marks',
                                    'text': text,
                                    'corrected': "Excessive punctuation marks detected"
                                })
                            if re.search(repeated_word_pattern, text, flags=re.IGNORECASE):
                                punctuation_issues.append({
                                    'slide': slide_index,
                                    'issue': 'Punctuation Marks',
                                    'text': text,
                                    'corrected': "Repeated words detected"
                                })
    return punctuation_issues

# Save issues to CSV
def save_to_csv(issues, output_csv):
    with open(output_csv, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=['slide', 'issue', 'text', 'corrected'])
        writer.writeheader()
        writer.writerows(issues)

# Set predefined password
PREDEFINED_PASSWORD = "securepassword123"

# Function to handle password protection
def password_protection():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if not st.session_state.authenticated:
        with st.form("password_form", clear_on_submit=True):
            password_input = st.text_input("Enter Password", type="password")
            submitted = st.form_submit_button("Submit")
            if submitted:
                if password_input == PREDEFINED_PASSWORD:
                    st.session_state.authenticated = True
                    st.experimental_set_query_params(authenticated="true")  # Set query parameters
                    st.success("Access Granted! Loading...")
                    st.experimental_rerun()  # Reload the app to bypass the password form
                else:
                    st.error("Incorrect Password")
        return False
    return True



# Main function
def main():
    # Proteksi password di awal
    if not password_protection():
        return  # Hentikan eksekusi jika password salah atau belum dimasukkan

    # CSS untuk menyembunyikan footer Streamlit
    hide_streamlit_style = """
    <style>
    footer {visibility: hidden;}
    [title~="View analytics"] {visibility: hidden;}
    </style>
    """
    st.markdown(hide_streamlit_style, unsafe_allow_html=True)

    st.title("PPT Validator")

    # File uploader dan pilihan font default
    uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["pptx"])
    font_options = ["Arial", "Calibri", "Times New Roman", "Verdana", "Helvetica", "EYInterstate"]
    default_font = st.selectbox("Select the default font for validation", font_options)

    if uploaded_file:
        if "uploaded_file" not in st.session_state or st.session_state.uploaded_file != uploaded_file:
            st.session_state.uploaded_file = uploaded_file
            st.session_state.pop('csv_path', None)
            st.session_state.pop('ppt_path', None)

        if st.button("Run Validation"):
            with tempfile.TemporaryDirectory() as tmpdir:
                temp_ppt_path = Path(tmpdir) / "uploaded_ppt.pptx"
                with open(temp_ppt_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                csv_output_path = Path(tmpdir) / "validation_report.csv"
                highlighted_ppt_path = Path(tmpdir) / "highlighted_presentation.pptx"

                # Initialize progress bar and percentage display
                progress_container = st.empty()
                progress_bar = st.progress(0)
                total_steps = 4  # Total validation tasks

                # Helper function to update progress
                def update_progress(task_name, current_step):
                    percentage = int((current_step / total_steps) * 100)
                    progress_container.markdown(f"**Progress: {percentage}% - {task_name}**")
                    progress_bar.progress(current_step / total_steps)

                # Run validations
                update_progress("Running grammar validation...", 1)
                grammar_issues = validate_grammar(temp_ppt_path)

                update_progress("Running punctuation validation...", 2)
                punctuation_issues = validate_punctuation(temp_ppt_path)

                update_progress("Running spelling validation...", 3)
                spelling_issues = validate_spelling(temp_ppt_path)

                update_progress("Running font validation...", 4)
                font_issues = validate_fonts(temp_ppt_path, default_font)

                # Combine results and save output
                combined_issues = grammar_issues + punctuation_issues + spelling_issues + font_issues

                save_to_csv(combined_issues, csv_output_path)
                highlight_ppt(temp_ppt_path, highlighted_ppt_path, combined_issues)

                st.session_state['csv_path'] = csv_output_path.read_bytes()
                st.session_state['ppt_path'] = highlighted_ppt_path.read_bytes()

                st.success("Validation completed!")

    # Display download buttons
    if 'csv_path' in st.session_state:
        st.download_button("Download Validation Report (CSV)", st.session_state['csv_path'],
                           file_name="validation_report.csv")

    if 'ppt_path' in st.session_state:
        st.download_button("Download Highlighted PPT", st.session_state['ppt_path'],
                           file_name="highlighted_presentation.pptx")


if __name__ == "__main__":
    main()





#########################################################################################

# import streamlit as st
# import tempfile
# from pathlib import Path
# from pptx import Presentation
# from spellchecker import SpellChecker
# import language_tool_python
# import csv
# import re
# import string

# # LanguageTool API initialization
# def initialize_language_tool():
#     try:
#         return language_tool_python.LanguageToolPublicAPI('en-US')  # Use Public API mode
#     except Exception as e:
#         st.error(f"LanguageTool initialization failed: {e}")
#         return None

# grammar_tool = initialize_language_tool()

# # Function to detect grammar issues
# def validate_grammar(input_ppt):
#     presentation = Presentation(input_ppt)
#     grammar_issues = []

#     for slide_index, slide in enumerate(presentation.slides, start=1):
#         for shape in slide.shapes:
#             if shape.has_text_frame:
#                 for paragraph in shape.text_frame.paragraphs:
#                     for run in paragraph.runs:
#                         text = run.text.strip()
#                         if text:
#                             # Grammar Check using LanguageTool
#                             if grammar_tool:
#                                 matches = grammar_tool.check(text)
#                                 if matches:
#                                     corrected = language_tool_python.utils.correct(text, matches)
#                                     if corrected != text:  # Only log if correction is made
#                                         grammar_issues.append({
#                                             'slide': slide_index,
#                                             'issue': 'Grammar Error',
#                                             'text': text,
#                                             'corrected': corrected
#                                         })
#     return grammar_issues

# # Function to detect and correct misspellings
# def detect_misspellings(text):
#     spell = SpellChecker()
#     words = text.split()
#     misspellings = {}

#     for word in words:
#         # Remove punctuation from the word for checking
#         clean_word = word.strip(string.punctuation)
        
#         # Check if the word is misspelled
#         if clean_word and clean_word.lower() not in spell:
#             correction = spell.correction(clean_word)
#             if correction:  # Only suggest if there's a valid correction
#                 misspellings[word] = correction

#     return misspellings

# # Function to validate spelling
# def validate_spelling(input_ppt):
#     presentation = Presentation(input_ppt)
#     spelling_issues = []

#     for slide_index, slide in enumerate(presentation.slides, start=1):
#         for shape in slide.shapes:
#             if shape.has_text_frame:
#                 for paragraph in shape.text_frame.paragraphs:
#                     for run in paragraph.runs:
#                         if run.text.strip():
#                             misspellings = detect_misspellings(run.text)
#                             for original_word, correction in misspellings.items():
#                                 spelling_issues.append({
#                                     'slide': slide_index,
#                                     'issue': 'Misspelling',
#                                     'text': f"Original: {original_word}",
#                                     'corrected': f"Suggestion: {correction}"
#                                 })

#     return spelling_issues

# # Function to validate fonts in a presentation
# def validate_fonts(input_ppt, default_font):
#     presentation = Presentation(input_ppt)
#     issues = []

#     for slide_index, slide in enumerate(presentation.slides, start=1):
#         for shape in slide.shapes:
#             if shape.has_text_frame:
#                 for paragraph in shape.text_frame.paragraphs:
#                     for run in paragraph.runs:
#                         if run.text.strip():  # Skip empty text
#                             # Check for inconsistent fonts
#                             if run.font.name != default_font:
#                                 issues.append({
#                                     'slide': slide_index,
#                                     'issue': 'Inconsistent Font',
#                                     'text': run.text,
#                                     'corrected': f"Detected: {run.font.name}, Expected: {default_font}"
#                                 })
#     return issues

# # Function to detect punctuation issues
# def validate_punctuation(input_ppt):
#     presentation = Presentation(input_ppt)
#     punctuation_issues = []

#     # Define patterns for punctuation problems
#     excessive_punctuation_pattern = r"([!?.:,;]{2,})"  # Two or more punctuation marks
#     repeated_word_pattern = r"\b(\w+)\s+\1\b"  # Repeated words (e.g., "the the")

#     for slide_index, slide in enumerate(presentation.slides, start=1):
#         for shape in slide.shapes:
#             if shape.has_text_frame:
#                 for paragraph in shape.text_frame.paragraphs:
#                     for run in paragraph.runs:
#                         if run.text.strip():
#                             text = run.text

#                             # Check excessive punctuation
#                             match = re.search(excessive_punctuation_pattern, text)
#                             if match:
#                                 punctuation_marks = match.group(1)  # Extract punctuation
#                                 punctuation_issues.append({
#                                     'slide': slide_index,
#                                     'issue': 'Punctuation Marks',
#                                     'text': text,
#                                     'corrected': f"Excessive punctuation marks detected ({punctuation_marks})"
#                                 })

#                             # Check repeated words
#                             if re.search(repeated_word_pattern, text, flags=re.IGNORECASE):
#                                 punctuation_issues.append({
#                                     'slide': slide_index,
#                                     'issue': 'Punctuation Marks',
#                                     'text': text,
#                                     'corrected': "Repeated words detected"
#                                 })

#     return punctuation_issues

# # Function to save issues to CSV
# def save_to_csv(issues, output_csv):
#     with open(output_csv, mode='w', newline='', encoding='utf-8') as file:
#         writer = csv.DictWriter(file, fieldnames=['slide', 'issue', 'text', 'corrected'])
#         writer.writeheader()
#         writer.writerows(issues)

# # Main Streamlit app
# def main():
#     # CSS to hide Streamlit footer and profile menu
#     hide_streamlit_style = """
#     <style>
#     footer {visibility: hidden;}
#     [title~="View analytics"] {visibility: hidden;}
#     </style>
#     """
#     st.markdown(hide_streamlit_style, unsafe_allow_html=True)

#     st.title("PPT Validator")

#     uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["pptx"])

#     font_options = ["Arial", "Calibri", "Times New Roman", "Verdana", "Helvetica", "EYInterstate"]
#     default_font = st.selectbox("Select the default font for validation", font_options)

#     if uploaded_file and st.button("Run Validation"):
#         with tempfile.TemporaryDirectory() as tmpdir:
#             # Save uploaded file temporarily
#             temp_ppt_path = Path(tmpdir) / "uploaded_ppt.pptx"
#             with open(temp_ppt_path, "wb") as f:
#                 f.write(uploaded_file.getbuffer())

#             # Output path
#             csv_output_path = Path(tmpdir) / "validation_report.csv"

#             # Run validations
#             font_issues = validate_fonts(temp_ppt_path, default_font)
#             punctuation_issues = validate_punctuation(temp_ppt_path)
#             spelling_issues = validate_spelling(temp_ppt_path)
#             grammar_issues = validate_grammar(temp_ppt_path)

#             # Combine issues and save to CSV
#             combined_issues = font_issues + punctuation_issues + spelling_issues + grammar_issues
#             save_to_csv(combined_issues, csv_output_path)

#             # Display success and download link
#             st.success("Validation completed!")
#             st.download_button("Download Validation Report (CSV)", csv_output_path.read_bytes(),
#                                file_name="validation_report.csv")

# if __name__ == "__main__":
#     main()

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

# Initialize LanguageTool
def initialize_language_tool():
    try:
        return language_tool_python.LanguageToolPublicAPI('en-US')
    except Exception as e:
        st.error(f"LanguageTool initialization failed: {e}")
        return None

grammar_tool = initialize_language_tool()

# Custom dictionary
TECHNICAL_TERMS = {
    "TensorFlow", "Keras", "Scikit-learn", "NumPy", "Pandas", "Matplotlib", "OpenAI",
    "GPT-3", "Deep Learning", "Neural Network", "Data Science", "Seaborn", "Jupyter",
    "Anaconda", "Reinforcement Learning", "Supervised Learning", "Unsupervised Learning",
    "Natural Language Processing", "Computer Vision", "Big Data", "Data Mining",
    "Feature Engineering", "Hyperparameter", "Gradient Descent", "Convolutional Neural Network",
    "Recurrent Neural Network", "Support Vector Machine", "Decision Tree", "Random Forest",
    "Ensemble Learning", "Clustering", "Dimensionality Reduction", "Principal Component Analysis",
    "Exploratory Data Analysis", "Model Evaluation", "Cross-Validation", "Overfitting",
    "Underfitting", "Batch Normalization", "Dropout", "Activation Function", "Loss Function",
    "Backpropagation", "Transfer Learning", "Generative Adversarial Network", "Autoencoder",
    "Tokenization", "Embedding", "Word2Vec", "BERT", "OpenCV", "Flask", "Django",
    "REST API", "GraphQL", "SQL", "NoSQL", "MongoDB", "PostgreSQL", "MySQL", "Firebase",
    "Cloud Computing", "AWS", "Azure", "Google Cloud", "Docker", "Kubernetes", "CI/CD",
    "DevOps", "Agile", "Scrum", "Kanban", "Git", "GitHub", "Bitbucket", "Version Control",
    "API", "SDK", "Microservices", "Blockchain", "Cryptocurrency", "IoT", "Edge Computing",
    "Quantum Computing", "Augmented Reality", "Virtual Reality", "3D Printing", "Cybersecurity",
    "Penetration Testing", "Phishing", "Malware", "Ransomware", "Firewall", "VPN", "SSL",
    "Encryption", "Decryption", "Hashing", "Digital Signature", "Data Privacy", "GDPR",
    "1+", "2+", "3+", "4+", "5+", "6+", "7+", "8+", "9+", "10+", "11+", "12+", "13+",
    "14+", "15+", "16+", "17+", "18+", "19+", "20+", "21+", "22+", "23+", "24+", "25+",
    "26+", "27+", "28+", "29+", "30+", "31+", "32+", "33+", "34+", "35+", "36+", "37+",
    "38+", "39+", "40+", "41+", "42+", "43+", "44+", "45+", "46+", "47+", "48+", "49+",
    "50+", "51+", "52+", "53+", "54+", "55+", "56+", "57+", "58+", "59+", "60+", "61+",
    "62+", "63+", "64+", "65+", "66+", "67+", "68+", "69+", "70+", "71+", "72+", "73+",
    "74+", "75+", "76+", "77+", "78+", "79+", "80+", "81+", "82+", "83+", "84+", "85+",
    "86+", "87+", "88+", "89+", "90+", "91+", "92+", "93+", "94+", "95+", "96+", "97+",
    "98+", "99+", "100+", "+1", "+2", "+3", "+4", "+5", "+6", "+7", "+8", "+9", "+10",
    "+11", "+12", "+13", "+14", "+15", "+16", "+17", "+18", "+19", "+20", "+21", "+22",
    "+23", "+24", "+25", "+26", "+27", "+28", "+29", "+30", "+31", "+32", "+33", "+34",
    "+35", "+36", "+37", "+38", "+39", "+40", "+41", "+42", "+43", "+44", "+45", "+46",
    "+47", "+48", "+49", "+50", "+51", "+52", "+53", "+54", "+55", "+56", "+57", "+58",
    "+59", "+60", "+61", "+62", "+63", "+64", "+65", "+66", "+67", "+68", "+69", "+70",
    "+71", "+72", "+73", "+74", "+75", "+76", "+77", "+78", "+79", "+80", "+81", "+82",
    "+83", "+84", "+85", "+86", "+87", "+88", "+89", "+90", "+91", "+92", "+93", "+94",
    "+95", "+96", "+97", "+98", "+99", "+100"
}
NUMERIC_TERMS = {f"{i}+" for i in range(1, 101)}

# Initialize SpellChecker
spell = SpellChecker()
spell.word_frequency.load_words(TECHNICAL_TERMS.union(NUMERIC_TERMS))

# Skip validation for technical terms and numeric patterns
def is_exempted(word):
    return word in TECHNICAL_TERMS or re.match(r"^\d+\+$", word)

# Spelling Validation
def validate_spelling(input_ppt, progress_callback):
    presentation = Presentation(input_ppt)
    spelling_issues = []
    total_slides = len(presentation.slides)

    for slide_index, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        words = run.text.split()
                        for word in words:
                            clean_word = word.strip(string.punctuation)
                            if is_exempted(clean_word):
                                continue
                            if clean_word.lower() not in spell:
                                correction = spell.correction(clean_word)
                                if correction and correction != clean_word:
                                    spelling_issues.append({
                                        'slide': slide_index,
                                        'issue': 'Misspelling',
                                        'text': word,
                                        'corrected': correction
                                    })
        progress_callback(slide_index, total_slides, "Spelling Validation")
    return spelling_issues

# Font Validation
def validate_fonts(input_ppt, default_font, progress_callback):
    presentation = Presentation(input_ppt)
    font_issues = []
    total_slides = len(presentation.slides)

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
                                'corrected': f"Expected: {default_font}, Found: {run.font.name}"
                            })
        progress_callback(slide_index, total_slides, "Font Validation")
    return font_issues

# Highlight issues in PPT
def highlight_ppt(input_ppt, output_ppt, issues):
    presentation = Presentation(input_ppt)
    for issue in issues:
        slide_index = issue['slide'] - 1
        slide = presentation.slides[slide_index]
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if issue['text'] in run.text:
                            run.font.color.rgb = RGBColor(255, 255, 0)
    presentation.save(output_ppt)

# Save issues to CSV
def save_to_csv(issues, output_csv):
    with open(output_csv, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=['slide', 'issue', 'text', 'corrected'])
        writer.writeheader()
        writer.writerows(issues)

# Password Protection
PREDEFINED_PASSWORD = "securepassword123"

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
                    st.success("Access Granted! Click 'Run Validation' to proceed.")
                else:
                    st.error("Incorrect Password")
        return False
    return True

# Main Function
def main():
    if not password_protection():
        return

    st.title("PPT Validator")
    uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["pptx"])
    font_options = ["Arial", "Calibri", "Times New Roman", "Verdana", "Helvetica", "EYInterstate"]
    default_font = st.selectbox("Select the default font for validation", font_options)

    if uploaded_file and st.button("Run Validation"):
        with tempfile.TemporaryDirectory() as tmpdir:
            temp_ppt_path = Path(tmpdir) / "uploaded_ppt.pptx"
            with open(temp_ppt_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            csv_output_path = Path(tmpdir) / "validation_report.csv"
            highlighted_ppt_path = Path(tmpdir) / "highlighted_presentation.pptx"
            progress_bar = st.progress(0)
            progress_text = st.empty()

            def update_progress(current, total, task_name):
                percentage = int((current / total) * 100)
                progress_bar.progress(percentage / 100)
                progress_text.text(f"{task_name}: {percentage}%")

            # Run validations
            spelling_issues = validate_spelling(temp_ppt_path, update_progress)
            font_issues = validate_fonts(temp_ppt_path, default_font, update_progress)

            combined_issues = spelling_issues + font_issues
            save_to_csv(combined_issues, csv_output_path)
            highlight_ppt(temp_ppt_path, highlighted_ppt_path, combined_issues)

            st.session_state["csv_path"] = csv_output_path.read_bytes()
            st.session_state["ppt_path"] = highlighted_ppt_path.read_bytes()

            st.success("Validation completed!")

    if "csv_path" in st.session_state:
        st.download_button("Download Validation Report (CSV)", st.session_state["csv_path"],
                           file_name="validation_report.csv")

    if "ppt_path" in st.session_state:
        st.download_button("Download Highlighted PPT", st.session_state["ppt_path"],
                           file_name="highlighted_presentation.pptx")

if __name__ == "__main__":
    main()






# #########################################################################################

# # import streamlit as st
# # import tempfile
# # from pathlib import Path
# # from pptx import Presentation
# # from spellchecker import SpellChecker
# # import language_tool_python
# # import csv
# # import re
# # import string

# # # LanguageTool API initialization
# # def initialize_language_tool():
# #     try:
# #         return language_tool_python.LanguageToolPublicAPI('en-US')  # Use Public API mode
# #     except Exception as e:
# #         st.error(f"LanguageTool initialization failed: {e}")
# #         return None

# # grammar_tool = initialize_language_tool()

# # # Function to detect grammar issues
# # def validate_grammar(input_ppt):
# #     presentation = Presentation(input_ppt)
# #     grammar_issues = []

# #     for slide_index, slide in enumerate(presentation.slides, start=1):
# #         for shape in slide.shapes:
# #             if shape.has_text_frame:
# #                 for paragraph in shape.text_frame.paragraphs:
# #                     for run in paragraph.runs:
# #                         text = run.text.strip()
# #                         if text:
# #                             # Grammar Check using LanguageTool
# #                             if grammar_tool:
# #                                 matches = grammar_tool.check(text)
# #                                 if matches:
# #                                     corrected = language_tool_python.utils.correct(text, matches)
# #                                     if corrected != text:  # Only log if correction is made
# #                                         grammar_issues.append({
# #                                             'slide': slide_index,
# #                                             'issue': 'Grammar Error',
# #                                             'text': text,
# #                                             'corrected': corrected
# #                                         })
# #     return grammar_issues

# # # Function to detect and correct misspellings
# # def detect_misspellings(text):
# #     spell = SpellChecker()
# #     words = text.split()
# #     misspellings = {}

# #     for word in words:
# #         # Remove punctuation from the word for checking
# #         clean_word = word.strip(string.punctuation)
        
# #         # Check if the word is misspelled
# #         if clean_word and clean_word.lower() not in spell:
# #             correction = spell.correction(clean_word)
# #             if correction:  # Only suggest if there's a valid correction
# #                 misspellings[word] = correction

# #     return misspellings

# # # Function to validate spelling
# # def validate_spelling(input_ppt):
# #     presentation = Presentation(input_ppt)
# #     spelling_issues = []

# #     for slide_index, slide in enumerate(presentation.slides, start=1):
# #         for shape in slide.shapes:
# #             if shape.has_text_frame:
# #                 for paragraph in shape.text_frame.paragraphs:
# #                     for run in paragraph.runs:
# #                         if run.text.strip():
# #                             misspellings = detect_misspellings(run.text)
# #                             for original_word, correction in misspellings.items():
# #                                 spelling_issues.append({
# #                                     'slide': slide_index,
# #                                     'issue': 'Misspelling',
# #                                     'text': f"Original: {original_word}",
# #                                     'corrected': f"Suggestion: {correction}"
# #                                 })

# #     return spelling_issues

# # # Function to validate fonts in a presentation
# # def validate_fonts(input_ppt, default_font):
# #     presentation = Presentation(input_ppt)
# #     issues = []

# #     for slide_index, slide in enumerate(presentation.slides, start=1):
# #         for shape in slide.shapes:
# #             if shape.has_text_frame:
# #                 for paragraph in shape.text_frame.paragraphs:
# #                     for run in paragraph.runs:
# #                         if run.text.strip():  # Skip empty text
# #                             # Check for inconsistent fonts
# #                             if run.font.name != default_font:
# #                                 issues.append({
# #                                     'slide': slide_index,
# #                                     'issue': 'Inconsistent Font',
# #                                     'text': run.text,
# #                                     'corrected': f"Detected: {run.font.name}, Expected: {default_font}"
# #                                 })
# #     return issues

# # # Function to detect punctuation issues
# # def validate_punctuation(input_ppt):
# #     presentation = Presentation(input_ppt)
# #     punctuation_issues = []

# #     # Define patterns for punctuation problems
# #     excessive_punctuation_pattern = r"([!?.:,;]{2,})"  # Two or more punctuation marks
# #     repeated_word_pattern = r"\b(\w+)\s+\1\b"  # Repeated words (e.g., "the the")

# #     for slide_index, slide in enumerate(presentation.slides, start=1):
# #         for shape in slide.shapes:
# #             if shape.has_text_frame:
# #                 for paragraph in shape.text_frame.paragraphs:
# #                     for run in paragraph.runs:
# #                         if run.text.strip():
# #                             text = run.text

# #                             # Check excessive punctuation
# #                             match = re.search(excessive_punctuation_pattern, text)
# #                             if match:
# #                                 punctuation_marks = match.group(1)  # Extract punctuation
# #                                 punctuation_issues.append({
# #                                     'slide': slide_index,
# #                                     'issue': 'Punctuation Marks',
# #                                     'text': text,
# #                                     'corrected': f"Excessive punctuation marks detected ({punctuation_marks})"
# #                                 })

# #                             # Check repeated words
# #                             if re.search(repeated_word_pattern, text, flags=re.IGNORECASE):
# #                                 punctuation_issues.append({
# #                                     'slide': slide_index,
# #                                     'issue': 'Punctuation Marks',
# #                                     'text': text,
# #                                     'corrected': "Repeated words detected"
# #                                 })

# #     return punctuation_issues

# # # Function to save issues to CSV
# # def save_to_csv(issues, output_csv):
# #     with open(output_csv, mode='w', newline='', encoding='utf-8') as file:
# #         writer = csv.DictWriter(file, fieldnames=['slide', 'issue', 'text', 'corrected'])
# #         writer.writeheader()
# #         writer.writerows(issues)

# # # Main Streamlit app
# # def main():
# #     # CSS to hide Streamlit footer and profile menu
# #     hide_streamlit_style = """
# #     <style>
# #     footer {visibility: hidden;}
# #     [title~="View analytics"] {visibility: hidden;}
# #     </style>
# #     """
# #     st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# #     st.title("PPT Validator")

# #     uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["pptx"])

# #     font_options = ["Arial", "Calibri", "Times New Roman", "Verdana", "Helvetica", "EYInterstate"]
# #     default_font = st.selectbox("Select the default font for validation", font_options)

# #     if uploaded_file and st.button("Run Validation"):
# #         with tempfile.TemporaryDirectory() as tmpdir:
# #             # Save uploaded file temporarily
# #             temp_ppt_path = Path(tmpdir) / "uploaded_ppt.pptx"
# #             with open(temp_ppt_path, "wb") as f:
# #                 f.write(uploaded_file.getbuffer())

# #             # Output path
# #             csv_output_path = Path(tmpdir) / "validation_report.csv"

# #             # Run validations
# #             font_issues = validate_fonts(temp_ppt_path, default_font)
# #             punctuation_issues = validate_punctuation(temp_ppt_path)
# #             spelling_issues = validate_spelling(temp_ppt_path)
# #             grammar_issues = validate_grammar(temp_ppt_path)

# #             # Combine issues and save to CSV
# #             combined_issues = font_issues + punctuation_issues + spelling_issues + grammar_issues
# #             save_to_csv(combined_issues, csv_output_path)

# #             # Display success and download link
# #             st.success("Validation completed!")
# #             st.download_button("Download Validation Report (CSV)", csv_output_path.read_bytes(),
# #                                file_name="validation_report.csv")

# # if __name__ == "__main__":
# #     main()

# import streamlit as st
# import tempfile
# from pathlib import Path
# from pptx import Presentation
# from spellchecker import SpellChecker
# import language_tool_python
# import csv
# import re
# import string
# from pptx.dml.color import RGBColor

# # LanguageTool API initialization
# def initialize_language_tool():
#     try:
#         return language_tool_python.LanguageToolPublicAPI('en-US')
#     except Exception as e:
#         st.error(f"LanguageTool initialization failed: {e}")
#         return None

# grammar_tool = initialize_language_tool()

# # Function to highlight issues in a PPT
# def highlight_ppt(input_ppt, output_ppt, issues):
#     presentation = Presentation(input_ppt)
#     for issue in issues:
#         slide_index = issue['slide'] - 1  # Slide index starts at 0
#         slide = presentation.slides[slide_index]
#         for shape in slide.shapes:
#             if shape.has_text_frame:
#                 for paragraph in shape.text_frame.paragraphs:
#                     for run in paragraph.runs:
#                         if issue['text'] in run.text:
#                             run.font.color.rgb = RGBColor(255, 255, 0)  # Highlight text in yellow
#     presentation.save(output_ppt)

# # Validation functions
# def validate_grammar(input_ppt, progress_callback):
#     presentation = Presentation(input_ppt)
#     grammar_issues = []
#     total_slides = len(presentation.slides)

#     for slide_index, slide in enumerate(presentation.slides, start=1):
#         for shape in slide.shapes:
#             if shape.has_text_frame:
#                 for paragraph in shape.text_frame.paragraphs:
#                     for run in paragraph.runs:
#                         text = run.text.strip()
#                         if text and grammar_tool:
#                             matches = grammar_tool.check(text)
#                             if matches:
#                                 corrected = language_tool_python.utils.correct(text, matches)
#                                 if corrected != text:
#                                     grammar_issues.append({
#                                         'slide': slide_index,
#                                         'issue': 'Grammar Error',
#                                         'text': text,
#                                         'corrected': corrected
#                                     })
#         progress_callback(slide_index, total_slides, "Grammar Validation")
#     return grammar_issues

# def validate_spelling(input_ppt, progress_callback):
#     presentation = Presentation(input_ppt)
#     spelling_issues = []
#     spell = SpellChecker()
#     total_slides = len(presentation.slides)

#     for slide_index, slide in enumerate(presentation.slides, start=1):
#         for shape in slide.shapes:
#             if shape.has_text_frame:
#                 for paragraph in shape.text_frame.paragraphs:
#                     for run in paragraph.runs:
#                         if run.text.strip():
#                             words = run.text.split()
#                             for word in words:
#                                 clean_word = word.strip(string.punctuation)
#                                 if clean_word and clean_word.lower() not in spell:
#                                     correction = spell.correction(clean_word)
#                                     if correction:
#                                         spelling_issues.append({
#                                             'slide': slide_index,
#                                             'issue': 'Misspelling',
#                                             'text': f"Original: {word}",
#                                             'corrected': f"Suggestion: {correction}"
#                                         })
#         progress_callback(slide_index, total_slides, "Spelling Validation")
#     return spelling_issues

# def validate_fonts(input_ppt, default_font, progress_callback):
#     presentation = Presentation(input_ppt)
#     font_issues = []
#     total_slides = len(presentation.slides)

#     for slide_index, slide in enumerate(presentation.slides, start=1):
#         for shape in slide.shapes:
#             if shape.has_text_frame:
#                 for paragraph in shape.text_frame.paragraphs:
#                     for run in paragraph.runs:
#                         if run.text.strip() and run.font.name != default_font:
#                             font_issues.append({
#                                 'slide': slide_index,
#                                 'issue': 'Inconsistent Font',
#                                 'text': run.text,
#                                 'corrected': f"Detected: {run.font.name}, Expected: {default_font}"
#                             })
#         progress_callback(slide_index, total_slides, "Font Validation")
#     return font_issues

# def validate_punctuation(input_ppt, progress_callback):
#     presentation = Presentation(input_ppt)
#     punctuation_issues = []
#     total_slides = len(presentation.slides)

#     excessive_punctuation_pattern = r"([!?.:,;]{2,})"
#     repeated_word_pattern = r"\b(\w+)\s+\1\b"

#     for slide_index, slide in enumerate(presentation.slides, start=1):
#         for shape in slide.shapes:
#             if shape.has_text_frame:
#                 for paragraph in shape.text_frame.paragraphs:
#                     for run in paragraph.runs:
#                         text = run.text.strip()
#                         if re.search(excessive_punctuation_pattern, text):
#                             punctuation_issues.append({
#                                 'slide': slide_index,
#                                 'issue': 'Punctuation Error',
#                                 'text': text,
#                                 'corrected': "Excessive punctuation marks detected"
#                             })
#                         if re.search(repeated_word_pattern, text, flags=re.IGNORECASE):
#                             punctuation_issues.append({
#                                 'slide': slide_index,
#                                 'issue': 'Punctuation Error',
#                                 'text': text,
#                                 'corrected': "Repeated words detected"
#                             })
#         progress_callback(slide_index, total_slides, "Punctuation Validation")
#     return punctuation_issues

# # Save issues to CSV
# def save_to_csv(issues, output_csv):
#     with open(output_csv, mode='w', newline='', encoding='utf-8') as file:
#         writer = csv.DictWriter(file, fieldnames=['slide', 'issue', 'text', 'corrected'])
#         writer.writeheader()
#         writer.writerows(issues)

# # Password protection
# PREDEFINED_PASSWORD = "securepassword123"

# def password_protection():
#     if "authenticated" not in st.session_state:
#         st.session_state.authenticated = False

#     if not st.session_state.authenticated:
#         with st.form("password_form", clear_on_submit=True):
#             password_input = st.text_input("Enter Password", type="password")
#             submitted = st.form_submit_button("Submit")
#             if submitted:
#                 if password_input == PREDEFINED_PASSWORD:
#                     st.session_state.authenticated = True
#                     st.success("Access Granted! Click 'Run Validation' to proceed.")
#                 else:
#                     st.error("Incorrect Password")
#         return False
#     return True

# # Main function
# def main():
#     if not password_protection():
#         return  # Stop execution if not authenticated

#     st.title("PPT Validator")
#     uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["pptx"])
#     font_options = ["Arial", "Calibri", "Times New Roman", "Verdana", "Helvetica", "EYInterstate"]
#     default_font = st.selectbox("Select the default font for validation", font_options)

#     if uploaded_file and st.button("Run Validation"):
#         with tempfile.TemporaryDirectory() as tmpdir:
#             temp_ppt_path = Path(tmpdir) / "uploaded_ppt.pptx"
#             with open(temp_ppt_path, "wb") as f:
#                 f.write(uploaded_file.getbuffer())

#             csv_output_path = Path(tmpdir) / "validation_report.csv"
#             highlighted_ppt_path = Path(tmpdir) / "highlighted_presentation.pptx"
#             progress_bar = st.progress(0)
#             progress_text = st.empty()

#             def update_progress(current, total, task_name):
#                 percentage = int((current / total) * 100)
#                 progress_bar.progress(percentage / 100)
#                 progress_text.text(f"{task_name}: {percentage}%")

#             grammar_issues = validate_grammar(temp_ppt_path, update_progress)
#             spelling_issues = validate_spelling(temp_ppt_path, update_progress)
#             punctuation_issues = validate_punctuation(temp_ppt_path, update_progress)
#             font_issues = validate_fonts(temp_ppt_path, default_font, update_progress)

#             combined_issues = grammar_issues + spelling_issues + punctuation_issues + font_issues
#             save_to_csv(combined_issues, csv_output_path)
#             highlight_ppt(temp_ppt_path, highlighted_ppt_path, combined_issues)

#             st.success("Validation completed!")
#             st.download_button("Download Validation Report (CSV)", csv_output_path.read_bytes(),
#                                file_name="validation_report.csv")
#             st.download_button("Download Highlighted PPT", highlighted_ppt_path.read_bytes(),
#                                file_name="highlighted_presentation.pptx")

# if __name__ == "__main__":
#     main()


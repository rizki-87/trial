# Updated validate_combined function
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
                            # Grammar or Spelling Check
                            if grammar_tool:
                                matches = grammar_tool.check(text)
                                if matches:
                                    corrected = language_tool_python.utils.correct(text, matches)
                                    if corrected != text:  # Only log if correction is made
                                        combined_issues.append({
                                            'slide': slide_index,
                                            'issue': 'Grammar or Spelling Error',
                                            'text': text,
                                            'corrected': corrected  # Do not include punctuation corrections
                                        })

                            # Punctuation Check (Excessive Punctuation)
                            excessive_punctuation_pattern = r"([!?.:,;]{2,})"
                            match = re.search(excessive_punctuation_pattern, text)
                            if match:
                                punctuation_marks = match.group(1)
                                combined_issues.append({
                                    'slide': slide_index,
                                    'issue': 'Punctuation Issue',
                                    'text': text,
                                    'corrected': f"Excessive punctuation marks detected ({punctuation_marks})"
                                })

    return combined_issues

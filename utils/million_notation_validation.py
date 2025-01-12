import re
import logging

def validate_million_notations(slide, slide_index, notation='m'):
    issues = []
    
    # Determine the regex pattern based on the selected notation
    if notation.lower() == 'm':
        pattern = r'[\€\$]?\s*\d+\s?m\b'  # Allow currency symbols before the number
    elif notation.lower() == 'mn':
        pattern = r'[\€\$]?\s*\d+\s?Mn\b'  # Allow currency symbols before the number
    else:
        pattern = r'[\€\$]?\s*\d+\s?M\b'  # Default to 'M' with currency symbols

    notation_set = set()
    all_matches = []
    logging.debug(f"Slide {slide_index}: Checking shapes for million notations")

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue  # Skip shapes that do not have text frames
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                matches = re.findall(pattern, run.text, re.IGNORECASE)  # Find matches based on the pattern
                all_matches.extend(matches)  # Collect all matches
                for match in matches:
                    notation_set.add(match.strip())  # Add the match to the notation set
    
    # Check for consistency of notation only if any notation was found
    if notation_set:
        if len(notation_set) > 1:
            for match in all_matches:
                issues.append({
                    'slide': slide_index,
                    'issue': 'Inconsistent Million Notations',
                    'text': match,
                    'details': f'Found inconsistent million notations: [using {", ".join(notation_set)}]'
                })
        else:
            # If only one notation is found, no need to log an issue
            logging.debug(f"Slide {slide_index}: Consistent notation found: {notation_set}")
    else:
        # If no notation was found, no need to log an issue
        logging.debug(f"Slide {slide_index}: No million notation found.")

    return issues

import re
import logging

def validate_million_notations(slide, slide_index, notation='m'):
    issues = []
    
    # Determine the regex pattern based on the selected notation
    if notation.lower() == 'm':
        # Allow currency symbols, optional spaces, and flexible number formats for 'm'
        pattern = r'[\€\$]?\s*\d{1,3}(?:\.\d{3})*(?:,\d+)?\s?[mM]\b'  
    elif notation.lower() == 'mn':
        # Allow currency symbols, optional spaces, and flexible number formats for 'Mn'
        pattern = r'[\€\$]?\s*\d{1,3}(?:\.\d{3})*(?:,\d+)?\s?Mn\b'  
    else:
        # Default to 'M' with currency symbols and flexible number formats
        pattern = r'[\€\$]?\s*\d{1,3}(?:\.\d{3})*(?:,\d+)?\s?M\b'  

    notation_set = set()  # Set to store unique notations found
    all_matches = []  # List to store all matches found
    logging.debug(f"Slide {slide_index}: Checking shapes for million notations")  # Log the slide being checked

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue  # Skip shapes that do not have text frames
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                # Find matches based on the pattern
                matches = re.findall(pattern, run.text, re.IGNORECASE)  
                all_matches.extend(matches)  # Collect all matches found
                for match in matches:
                    notation_set.add(match.strip())  # Add the match to the notation set
    
    # Check for consistency of notation only if any notation was found
    if notation_set:
        if len(notation_set) > 1:
            for match in all_matches:
                issues.append({
                    'slide': slide_index,
                    'issue': 'Inconsistent Million Notations',  # Issue type
                    'text': match,  # The text that caused the issue
                    'details': f'Found inconsistent million notations: [using {", ".join(notation_set)}]'  # Details of the issue
                })
        else:
            # If only one notation is found, no need to log an issue
            logging.debug(f"Slide {slide_index}: Consistent notation found: {notation_set}")
    else:
        # If no notation was found, no need to log an issue
        logging.debug(f"Slide {slide_index}: No million notation found.")

    return issues  # Return the list of issues found

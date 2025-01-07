from pptx import Presentation  
from pptx.dml.color import RGBColor  
import csv  
import logging  
  
def highlight_ppt(input_ppt, output_ppt, issues):  
    """Highlights issues in a PowerPoint presentation, ignoring likely background images."""  
    try:  
        presentation = Presentation(input_ppt)  
        for issue in issues:  
            if isinstance(issue, dict):  
                slide_index = issue['slide'] - 1  
                if 0 <= slide_index < len(presentation.slides):  
                    slide = presentation.slides[slide_index]  
                    for shape in slide.shapes:  
                        if shape.has_text_frame:  
                            for paragraph in shape.text_frame.paragraphs:  
                                for run in paragraph.runs:  
                                    if issue['text'] in run.text:  
                                        run.font.color.rgb = RGBColor(255, 255, 0)  
                        elif is_likely_background_image(shape, slide):  
                            continue  # Ignore likely background images  
  
        presentation.save(output_ppt)  
    except Exception as e:  
        logging.error(f"Error highlighting PPT: {e}")  
  
  
def is_likely_background_image(shape, slide):  
    """Checks if a shape is likely a background image based on its position and size."""  
    slide_width = slide.slide_width  
    slide_height = slide.slide_height  
    # Adjust thresholds as needed  
    position_threshold = 10  # Pixels  
    size_threshold = 0.9  # Percentage of slide size  
  
    if shape.shape_type == 13: # MSL_PICTURE.  Check if this is correct for your PPTX files.  
        if (shape.left <= position_threshold and shape.top <= position_threshold and  
                shape.width >= slide_width * size_threshold and shape.height >= slide_height * size_threshold):  
            return True  
    return False  
  
  
def save_to_csv(issues, output_csv):  
    """Saves validation issues to a CSV file."""  
    try:  
        with open(output_csv, mode='w', newline='', encoding='utf-8') as file:  
            fieldnames = ['slide', 'issue', 'text', 'corrected', 'details']  
            writer = csv.DictWriter(file, fieldnames=fieldnames)  
            writer.writeheader()  
            for issue in issues:  
                if isinstance(issue, dict):  
                    row = {k: issue.get(k, '') for k in fieldnames}  
                    writer.writerow(row)  
    except Exception as e:  
        logging.error(f"Error saving to CSV: {e}")  
  

from pptx import Presentation
from pptx.dml.color import RGBColor

def highlight_ppt(input_ppt, output_ppt, issues):
    presentation = Presentation(input_ppt)
    for issue in issues:
        if isinstance(issue, dict):
            slide_index = issue['slide'] - 1
            slide = presentation.slides[slide_index]
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if issue['text'] in run.text:
                                run.font.color.rgb = RGBColor(255, 255, 0)

    # Akses slide_width dan slide_height dari objek Presentation
    slide_width = presentation.slide_width
    slide_height = presentation.slide_height

    # Simpan presentasi yang telah disorot
    presentation.save(output_ppt)

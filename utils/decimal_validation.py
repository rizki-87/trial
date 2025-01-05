# utils/decimal_validation.py

import re
import logging

# Simpan pola regex dalam variabel
decimal_pattern = re.compile(r'\b\d+[\.,]\d+\b')

def validate_decimal_consistency(slide, slide_index, decimal_places):
    issues = []
    
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    text = run.text
                    # Cari semua angka desimal dengan titik atau koma sebagai pemisah desimal
                    matches = decimal_pattern.findall(text)
                    logging.debug(f"Slide {slide_index}: Found matches: {matches}")
                    for match in matches:
                        # Ganti koma dengan titik untuk konsistensi
                        match = match.replace(',', '.')
                        # Hitung jumlah digit setelah titik
                        decimal_part = match.split('.')[-1]
                        if len(decimal_part) != decimal_places:
                            issues.append({
                                'slide': slide_index,
                                'issue': 'Inconsistent Decimal Points',
                                'text': match,
                                'details': f'Expected {decimal_places} decimal place(s), found {len(decimal_part)} in "{match}".'
                            })
                            logging.debug(f"Slide {slide_index}: Inconsistent decimal points found in "{match}". Expected {decimal_places}, found {len(decimal_part)}.")
    
    return issues

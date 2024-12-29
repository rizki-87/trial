# utils/decimal_validation.py

import re
import logging

def validate_decimal_consistency(slide, slide_index, reference_decimal_points):
    issues = []
    decimal_points = set()
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    text = run.text
                    # Cari semua angka desimal dengan titik atau koma sebagai pemisah desimal
                    matches = re.findall(r'\b\d+[\.,]\d+\b', text)
                    for match in matches:
                        # Ganti koma dengan titik untuk konsistensi
                        match = match.replace(',', '.')
                        # Hitung jumlah digit setelah titik
                        decimal_part = match.split('.')[-1]
                        decimal_points.add(len(decimal_part))
    
    # Jika ada lebih dari satu jumlah digit setelah titik, tambahkan issue
    if len(decimal_points) > 1:
        issues.append({
            'slide': slide_index,
            'issue': 'Inconsistent Decimal Points',
            'text': '',
            'details': f'Found inconsistent decimal points: {list(decimal_points)}'
        })
    
    # Jika reference_decimal_points belum ada, set sebagai referensi
    if reference_decimal_points is None and decimal_points:
        reference_decimal_points = decimal_points.copy()
        logging.debug(f"Set reference_decimal_points to {reference_decimal_points} on slide {slide_index}")
    
    # Jika reference_decimal_points sudah ada, periksa konsistensi dengan referensi
    if reference_decimal_points is not None:
        if decimal_points != reference_decimal_points:
            # Pastikan reference_decimal_points tidak kosong sebelum melakukan pop()
            if reference_decimal_points:
                ref_point = reference_decimal_points.pop()
                reference_decimal_points.add(ref_point)  # Kembalikan ke set
                issues.append({
                    'slide': slide_index,
                    'issue': 'Inconsistent Decimal Points',
                    'text': '',
                    'details': f'Decimal points on slide {slide_index} are inconsistent. Expected {ref_point} decimal places, found {list(decimal_points)}.'
                })
            else:
                logging.error(f"reference_decimal_points is empty on slide {slide_index}")
    
    return issues, reference_decimal_points

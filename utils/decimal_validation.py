import re
import logging

# Simpan pola regex dalam variabel
decimal_pattern = re.compile(r'\b\d+[\.,]\d+\b')

def validate_decimal_consistency(slide, slide_index, reference_decimal_points):
    issues = []
    decimal_points = set()
    
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
                        decimal_points.add(len(decimal_part))
                        logging.debug(f"Slide {slide_index}: Match: {match}, Decimal part: {decimal_part}, Decimal points: {decimal_points}")
    
    # Jika slide tidak memiliki angka desimal, kembalikan issues dan reference_decimal_points tanpa perubahan
    if not decimal_points:
        return issues, reference_decimal_points
    
    # Jika reference_decimal_points belum ada, set sebagai referensi
    if reference_decimal_points is None:
        reference_decimal_points = next(iter(decimal_points))
        logging.debug(f"Set reference_decimal_points to {reference_decimal_points} on slide {slide_index}")
    
    # Jika reference_decimal_points sudah ada, periksa konsistensi dengan referensi
    else:
        for decimal_point in decimal_points:
            if decimal_point != reference_decimal_points:
                issues.append({
                    'slide': slide_index,
                    'issue': 'Inconsistent Decimal Points',
                    'text': '',
                    'details': f'Decimal points on slide {slide_index} are inconsistent. Expected {reference_decimal_points} decimal places, found {decimal_point}.'
                })
                logging.debug(f"Slide {slide_index}: Inconsistent decimal points found. Expected {reference_decimal_points}, found {decimal_point}")
    
    return issues, reference_decimal_points

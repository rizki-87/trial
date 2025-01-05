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
                    for match in matches:
                        # Ganti koma dengan titik untuk konsistensi
                        match = match.replace(',', '.')
                        # Hitung jumlah digit setelah titik
                        decimal_part = match.split('.')[-1]
                        decimal_points.add(len(decimal_part))
    
    # Jika slide tidak memiliki angka desimal, kembalikan issues dan reference_decimal_points tanpa perubahan
    if not decimal_points:
        return issues, reference_decimal_points
    
    # Jika reference_decimal_points belum ada, set sebagai referensi
    if reference_decimal_points is None:
        reference_decimal_points = decimal_points.copy()
        logging.debug(f"Set reference_decimal_points to {reference_decimal_points} on slide {slide_index}")
    
    # Jika reference_decimal_points sudah ada, periksa konsistensi dengan referensi
    else:
        ref_point = next(iter(reference_decimal_points))  # Ambil nilai referensi
        for decimal_point in decimal_points:
            if decimal_point != ref_point:
                issues.append({
                    'slide': slide_index,
                    'issue': 'Inconsistent Decimal Points',
                    'text': '',
                    'details': f'Decimal points on slide {slide_index} are inconsistent. Expected {ref_point} decimal places, found {decimal_point}.'
                })
    
    return issues, reference_decimal_points

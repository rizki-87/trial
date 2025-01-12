import re  
import logging  
  
def validate_million_notations(slide, slide_index, notation='m'):  
    issues = []  
      
    # Tentukan pola regex berdasarkan notasi yang dipilih  
    if notation.lower() == 'm':  
        pattern = r'[\€\$]?\s*\d{1,3}(?:\.\d{3})*(?:,\d+)?\s?[mM]\b'  # M atau m  
    elif notation.lower() == 'mn':  
        pattern = r'[\€\$]?\s*\d{1,3}(?:\.\d{3})*(?:,\d+)?\s?Mn\b'  # Mn  
    else:  
        pattern = r'[\€\$]?\s*\d{1,3}(?:\.\d{3})*(?:,\d+)?\s?M\b'  # M  
  
    notation_set = set()  # Set untuk menyimpan notasi unik yang ditemukan  
    all_matches = []  # List untuk menyimpan semua match yang ditemukan  
    logging.debug(f"Slide {slide_index}: Checking shapes for million notations")  # Log slide yang diperiksa  
  
    for shape in slide.shapes:  
        if not shape.has_text_frame:  
            continue  # Lewati bentuk yang tidak memiliki frame teks  
        for paragraph in shape.text_frame.paragraphs:  
            for run in paragraph.runs:  
                logging.debug(f"Checking text: {run.text}")  # Log teks yang sedang diperiksa  
                matches = re.findall(pattern, run.text, re.IGNORECASE)    
                logging.debug(f"Matches found: {matches}")  # Log hasil match  
                all_matches.extend(matches)  # Kumpulkan semua match yang ditemukan  
                for match in matches:  
                    notation_set.add(match.strip())  # Tambahkan match ke set notasi  
  
    # Log semua notasi yang ditemukan  
    logging.debug(f"Slide {slide_index}: Found notations: {notation_set}")  
  
    # Periksa konsistensi notasi hanya jika ada notasi yang ditemukan  
    if notation_set:  
        # Cek apakah ada notasi yang berbeda dari yang dipilih  
        if any((n.lower() != notation.lower() for n in notation_set)):  
            for match in all_matches:  
                issues.append({  
                    'slide': slide_index,  
                    'issue': 'Inconsistent Million Notations',  # Jenis masalah  
                    'text': match,  # Teks yang menyebabkan masalah  
                    'details': f'Found inconsistent million notations: [using {", ".join(notation_set)}]'  # Detail masalah  
                })  
        else:  
            # Jika hanya satu notasi ditemukan, tidak perlu mencatat masalah  
            logging.debug(f"Slide {slide_index}: Consistent notation found: {notation_set}")  
    else:  
        # Jika tidak ada notasi ditemukan, tidak perlu mencatat masalah  
        logging.debug(f"Slide {slide_index}: No million notation found.")  
  
    return issues  # Kembalikan daftar masalah yang ditemukan  

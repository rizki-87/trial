# import re  
# import pandas as pd  
# import logging  
  
# def validate_million_notations_with_pandas(df, selected_notation):  
#     issues = []  
      
#     # Tentukan pola regex berdasarkan notasi yang dipilih  
#     if selected_notation.lower() == 'm':  
#         pattern = r'[\€\$]?\s*\d+(?:\.\d+)?\s*[mM]\b'  # Mencari 'm' atau 'M'  
#     elif selected_notation.lower() == 'mn':  
#         pattern = r'[\€\$]?\s*\d+(?:\.\d+)?\s?Mn\b'  # Mencari 'Mn'  
#     else:  
#         pattern = r'[\€\$]?\s*\d+(?:\.\d+)?\s?[mM]\b'  # Default ke 'm' atau 'M'  
  
#     for index, row in df.iterrows():  
#         text = row['text']  
#         logging.debug(f"Checking text: {text}")  # Log teks yang sedang diperiksa  
          
#         # Cek apakah ada notasi juta dalam teks  
#         found_million_notation = re.search(pattern, text)  
          
#         if found_million_notation:  
#             # Jika notasi ditemukan, periksa apakah sesuai dengan notasi yang dipilih  
#             if selected_notation.lower() == 'mn' and 'm' in text.lower():  
#                 issues.append({  
#                     'slide': index + 1,  
#                     'issue': 'Found "m" notation but "Mn" was expected',  
#                     'text': text  
#                 })  
#             elif selected_notation.lower() == 'm' and 'Mn' in text:  
#                 issues.append({  
#                     'slide': index + 1,  
#                     'issue': 'Found "Mn" notation but "m" was expected',  
#                     'text': text  
#                 })  
#         else:  
#             logging.debug(f"No valid million notation found in: {text}")  # Log jika tidak ditemukan  
  
#     return issues  

import spacy  
import pandas as pd  
  
# Muat model bahasa Inggris  
nlp = spacy.load("en_core_web_sm")  
  
def validate_million_notations_with_spacy(df, selected_notation):  
    issues = []  
      
    for index, row in df.iterrows():  
        text = row['text']  
        doc = nlp(text)  
          
        # Cek untuk notasi juta  
        found_notation = False  
          
        for token in doc:  
            # Cek apakah token adalah angka dan diikuti oleh notasi yang sesuai  
            if token.like_num:  # Token adalah angka  
                next_token = token.nbor()  # Ambil token berikutnya  
                if selected_notation.lower() == 'm' and next_token.text.lower() in ['m']:  
                    found_notation = True  
                elif selected_notation.lower() == 'mn' and next_token.text.lower() in ['mn']:  
                    found_notation = True  
          
        # Jika notasi ditemukan, periksa apakah sesuai dengan notasi yang dipilih  
        if found_notation:  
            # Jika notasi ditemukan, periksa apakah sesuai dengan notasi yang dipilih  
            if selected_notation.lower() == 'mn' and 'm' in text.lower():  
                issues.append({  
                    'slide': index + 1,  
                    'issue': 'Found "m" notation but "Mn" was expected',  
                    'text': text  
                })  
            elif selected_notation.lower() == 'm' and 'Mn' in text:  
                issues.append({  
                    'slide': index + 1,  
                    'issue': 'Found "Mn" notation but "m" was expected',  
                    'text': text  
                })  
        else:  
            # Jika tidak ditemukan, Anda bisa menambahkan logika lain jika perlu  
            pass  
  
    return issues  

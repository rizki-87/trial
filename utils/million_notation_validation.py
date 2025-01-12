import re  
import pandas as pd  
import logging  
  
def validate_million_notations_with_pandas(df, selected_notation):  
    issues = []  
      
    # Tentukan pola regex berdasarkan notasi yang dipilih  
    if selected_notation.lower() == 'm':  
        pattern = r'[\€\$]?\s*\d+(?:\.\d+)?\s*[mM]\b'  # Mencari 'm' atau 'M'  
    elif selected_notation.lower() == 'mn':  
        pattern = r'[\€\$]?\s*\d+(?:\.\d+)?\s?Mn\b'  # Mencari 'Mn'  
    else:  
        pattern = r'[\€\$]?\s*\d+(?:\.\d+)?\s?[mM]\b'  # Default ke 'm' atau 'M'  
  
    for index, row in df.iterrows():  
        text = row['text']  
        logging.debug(f"Checking text: {text}")  # Log teks yang sedang diperiksa  
          
        # Cek apakah ada notasi juta dalam teks  
        found_million_notation = re.search(pattern, text)  
          
        if found_million_notation:  
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
            logging.debug(f"No valid million notation found in: {text}")  # Log jika tidak ditemukan  
  
    return issues  

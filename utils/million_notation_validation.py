import re  
import logging  
  
def validate_million_notations_with_pandas(df, notation='m'):  
    issues = []  
      
    # Tentukan pola notasi juta  
    if notation.lower() == 'm':  
        pattern = r'[\€\$]?\s*\d{1,3}(?:\.\d{3})*(?:,\d+)?\s?[mM]\b'  
    elif notation.lower() == 'mn':  
        pattern = r'[\€\$]?\s*\d{1,3}(?:\.\d{3})*(?:,\d+)?\s?Mn\b'  
    else:  
        pattern = r'[\€\$]?\s*\d{1,3}(?:\.\d{3})*(?:,\d+)?\s?M\b'  
  
    for index, row in df.iterrows():  
        if re.search(pattern, row['text']):  
            issues.append({  
                'slide': index + 1,  # Indeks slide  
                'issue': 'Found Million Notation',  
                'text': row['text']  
            })  
  
    return issues  

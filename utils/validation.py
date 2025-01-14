import re        
import logging        
import string        
import pandas as pd  # Pastikan pandas diimpor    
from utils.spelling_validation import validate_spelling_slide, validate_spelling_in_text        
  
def validate_tables(slide, slide_index):  # Hapus selected_notation dari parameter  
    issues = []        
    for shape in slide.shapes:        
        if shape.has_table:        
            table = shape.table        
            for row in table.rows:        
                for cell in row.cells:        
                    # Validasi teks di dalam sel        
                    text = cell.text.strip()        
                    if text:  # Jika ada teks        
                        issues.extend(validate_spelling_in_text(text, slide_index))        
            
    # Validasi notasi juta menggunakan DataFrame      
    # Mengambil teks dari slide untuk validasi      
    slide_text = []      
    for shape in slide.shapes:      
        if shape.has_text_frame:      
            for paragraph in shape.text_frame.paragraphs:      
                for run in paragraph.runs:      
                    slide_text.append(run.text)      
    df = pd.DataFrame(slide_text, columns=['text'])  # Membuat DataFrame dari teks slide      
      
    issues.extend(validate_million_notations(df, slide_index))  # Memanggil fungsi baru tanpa selected_notation        
            
    return issues        
  
def validate_charts(slide, slide_index):  # Hapus selected_notation dari parameter  
    issues = []        
    for shape in slide.shapes:        
        if shape.has_chart:        
            chart = shape.chart        
            # Validasi data di dalam chart        
            for series in chart.series:        
                for point in series.points:        
                    label = point.data_label.text.strip()        
                    if label:        
                        issues.extend(validate_spelling_in_text(label, slide_index))        
            # Jika chart memiliki data yang ditampilkan dalam tabel, validasi juga        
            if chart.has_data_table:        
                for row in chart.data_table.rows:        
                    for cell in row.cells:        
                        text = cell.text.strip()        
                        if text:        
                            issues.extend(validate_spelling_in_text(text, slide_index))        
            
    # Validasi notasi juta menggunakan DataFrame      
    # Mengambil teks dari slide untuk validasi      
    slide_text = []      
    for shape in slide.shapes:      
        if shape.has_text_frame:      
            for paragraph in shape.text_frame.paragraphs:      
                for run in paragraph.runs:      
                    slide_text.append(run.text)      
    df = pd.DataFrame(slide_text, columns=['text'])  # Membuat DataFrame dari teks slide      
      
    issues.extend(validate_million_notations(df, slide_index))  # Memanggil fungsi baru tanpa selected_notation        
            
    return issues        

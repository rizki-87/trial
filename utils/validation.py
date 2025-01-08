from pptx import Presentation
import re
import logging

def validate_tables(slide, slide_index):
    issues = []
    for shape in slide.shapes:
        if shape.has_table:
            table = shape.table
            for row in table.rows:
                for cell in row.cells:
                    # Validasi teks di dalam sel
                    text = cell.text.strip()
                    if text:  # Jika ada teks
                        # Tambahkan logika validasi sesuai kebutuhan
                        issues.extend(validate_spelling_in_text(text, slide_index))
    return issues

def validate_charts(slide, slide_index):
    issues = []
    for shape in slide.shapes:
        if shape.has_chart:
            chart = shape.chart
            # Validasi data di dalam chart
            for series in chart.series:
                for point in series.points:
                    # Misalnya, validasi label data
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
    return issues

def validate_spelling_in_text(text, slide_index):
    issues = []
    words = re.findall(r"\b[\w+]+\b", text)
    for word in words:
        clean_word = word.strip(string.punctuation)
        if clean_word.lower() not in spell:
            correction = spell.correction(clean_word)
            if correction and correction != clean_word:
                issues.append({
                    'slide': slide_index,
                    'issue': 'Misspelling in Table/Chart',
                    'text': word,
                    'corrected': correction
                })
    return issues
def validate_spelling_in_text(text, slide_index):
    issues = []
    words = re.findall(r"\b[\w+]+\b", text)
    for word in words:
        clean_word = word.strip(string.punctuation)
        if clean_word.lower() not in spell:
            correction = spell.correction(clean_word)
            if correction and correction != clean_word:
                issues.append({
                    'slide': slide_index,
                    'issue': 'Misspelling',
                    'text': word,
                    'corrected': correction
                })
    return issues

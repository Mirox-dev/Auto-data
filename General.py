import os
import hashlib
from datetime import datetime
from docx import Document

def get_file_hash(file_path):
    """Выбирает алгоритм в зависимости от размера файла и возвращает тип и хэш."""
    size_bytes = os.path.getsize(file_path)

    if size_bytes < 1 * 1024 * 1024:  # <1 Мб
        algorithm = 'md5'
    elif size_bytes < 100 * 1024 * 1024:  # 1–100 Мб
        algorithm = 'sha1'
    else:  # >100 Мб
        algorithm = 'sha256'

    hash_func = hashlib.new(algorithm)
    with open(file_path, 'rb') as file:
        while chunk := file.read(8192):
            hash_func.update(chunk)
    return algorithm.upper(), hash_func.hexdigest()

def get_creation_date(file_path):
    timestamp = os.path.getctime(file_path)
    return datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S')

def format_file_size(size_bytes):
    """Форматирует размер в Кб или Мб"""
    kb = size_bytes / 1024
    mb = kb / 1024
    gb = mb / 1024
    if gb >= 1:
        return f'{gb:.2f} Гб'
    elif mb >= 1:
        return f"{mb:.2f} Мб"
    else:
        return f"{kb:.2f} Кб"

def create_word_report(directory_path):
    doc = Document()
    folder_name = os.path.basename(os.path.normpath(directory_path))
    doc.add_heading(f'Отчет о файлах в папке:\n{folder_name}', level=1)
    doc.add_paragraph()

    for filename in os.listdir(directory_path):
        full_path = os.path.join(directory_path, filename)

        if os.path.isfile(full_path):
            try:
                creation_date = get_creation_date(full_path)
                size_formatted = format_file_size(os.path.getsize(full_path))
                hash_type, file_hash = get_file_hash(full_path)

                doc.add_paragraph(filename)
                doc.add_paragraph(f"Дата создания: {creation_date}")
                doc.add_paragraph(f"Размер: {size_formatted}")
                doc.add_paragraph(f"Хэш ({hash_type}): {file_hash}")
                doc.add_paragraph("")  # пустая строка для разделения файлов
            except Exception as e:
                doc.add_paragraph(filename)
                doc.add_paragraph(f"Ошибка при обработке файла: {e}")
                doc.add_paragraph("")

    output_path = os.path.join(directory_path, "report.docx")
    doc.save(output_path)
    print(f"Отчет создан: {output_path}")
if __name__ == "__main__":
    folder = input("Введите путь к папке: ").strip('"')
    create_word_report(folder)

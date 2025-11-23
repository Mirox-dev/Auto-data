import os
import hashlib
from datetime import datetime
from docx import Document

def get_file_hash(file_path, algorithm='sha256'):
    """Вычисляет hash-сумму файла."""
    hash_func = hashlib.new(algorithm)

    try:
        with open(file_path, 'rb') as file:
            while True:
                chunk = file.read(8192)
                if not chunk:
                    break
                hash_func.update(chunk)
        return hash_func.hexdigest()
    except Exception as e:
        return "Ошибка хеширования: " + str(e)

def get_creation_date(file_path):
    """Получает дату создания файла (Windows)."""
    try:
        timestamp = os.path.getctime(file_path)
        return datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return "Не удалось получить дату"

def create_word_report(directory_path, output_file="report.docx"):
    doc = Document()
    doc.add_heading("Отчет о файлах в директории:", level=1)
    doc.add_paragraph(directory_path)

    doc.add_paragraph("-" * 50)

    try:
        files = os.listdir(directory_path)
    except Exception as e:
        print("Ошибка доступа к папке:", e)
        return

    for filename in files:
        full_path = os.path.join(directory_path, filename)

        if os.path.isfile(full_path):
            try:
                creation_date = get_creation_date(full_path)
                size = os.path.getsize(full_path)
                file_hash = get_file_hash(full_path)

                line = "{} — {}, {} байт, {}".format(
                    filename, creation_date, size, file_hash
                )

                doc.add_paragraph(line)

            except Exception as e:
                doc.add_paragraph("{} — ОШИБКА: {}".format(filename, e))

    try:
        doc.save(output_file)
        print("Отчет создан:", output_file)
    except Exception as e:
        print("Ошибка при сохранении файла:", e)

if __name__ == "__main__":
    folder = input("Введите путь к папке: ").strip('"')
    create_word_report(folder)

import PyPDF2
from datetime import datetime
import sys
import os

def extract_pdf_modification_date(pdf_path):
    """
    Извлекает дату последнего изменения содержимого PDF-файла (ModDate)
    из внутренних метаданных документа.
    
    Args:
        pdf_path (str): Путь к PDF-файлу
        
    Returns:
        dict: Словарь с информацией о датах или None в случае ошибки
    """
    try:
        # Проверяем существует ли файл
        if not os.path.exists(pdf_path):
            return {"error": f"Файл не найден: {pdf_path}"}
        
        # Открываем PDF-файл
        with open(pdf_path, 'rb') as file:
            # Создаем объект для чтения PDF
            pdf_reader = PyPDF2.PdfReader(file)
            
            # Получаем метаданные документа
            metadata = pdf_reader.metadata
            
            if metadata is None:
                return {"error": "Метаданные не найдены в PDF-файле"}
            
            # Результат
            result = {
                "filename": os.path.basename(pdf_path),
                "file_path": pdf_path,
                "file_system_info": {
                    "file_system_modified": datetime.fromtimestamp(
                        os.path.getmtime(pdf_path)
                    ).strftime("%Y-%m-%d %H:%M:%S"),
                    "file_system_created": datetime.fromtimestamp(
                        os.path.getctime(pdf_path)
                    ).strftime("%Y-%m-%d %H:%M:%S")
                },
                "pdf_internal_info": {}
            }
            
            # Извлекаем стандартные поля метаданных PDF
            pdf_fields = {
                '/ModDate': 'pdf_internal_modified',
                '/CreationDate': 'pdf_internal_created',
                '/Title': 'title',
                '/Author': 'author',
                '/Subject': 'subject',
                '/Producer': 'producer',
                '/Creator': 'creator'
            }
            
            for pdf_field, result_field in pdf_fields.items():
                if pdf_field in metadata:
                    value = metadata[pdf_field]
                    # Пробуем преобразовать дату в читаемый формат
                    if pdf_field in ['/ModDate', '/CreationDate'] and value:
                        try:
                            # PDF даты обычно в формате: D:20240101120000+02'00'
                            date_str = str(value)
                            if date_str.startswith('D:'):
                                date_str = date_str[2:]
                            
                            # Простой парсинг даты (YYYYMMDDHHMMSS)
                            if len(date_str) >= 14:
                                formatted_date = f"{date_str[0:4]}-{date_str[4:6]}-{date_str[6:8]} {date_str[8:10]}:{date_str[10:12]}:{date_str[12:14]}"
                                result["pdf_internal_info"][result_field] = {
                                    "raw": str(value),
                                    "formatted": formatted_date
                                }
                            else:
                                result["pdf_internal_info"][result_field] = str(value)
                        except:
                            result["pdf_internal_info"][result_field] = str(value)
                    else:
                        result["pdf_internal_info"][result_field] = str(value)
            
            return result
            
    except Exception as e:
        return {"error": f"Ошибка при обработке PDF: {str(e)}"}

def main():
    """
    Основная функция программы
    """
    print("=" * 60)
    print("ИЗВЛЕЧЕНИЕ ВНУТРЕННЕЙ ДАТЫ ИЗМЕНЕНИЯ PDF")
    print("=" * 60)
    
    # Проверяем аргументы командной строки
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    else:
        # Если аргумент не передан, запрашиваем путь
        pdf_path = input("Введите путь к PDF-файлу: ").strip()
    
    if not pdf_path:
        print("Ошибка: Путь к файлу не указан")
        return
    
    # Извлекаем информацию
    print(f"\nАнализ файла: {pdf_path}")
    print("-" * 60)
    
    result = extract_pdf_modification_date(pdf_path)
    
    if "error" in result:
        print(f"ОШИБКА: {result['error']}")
        return
    
    # Выводим информацию о файловой системе
    print("\n📁 ИНФОРМАЦИЯ ИЗ ФАЙЛОВОЙ СИСТЕМЫ WINDOWS:")
    print(f"   Последнее изменение (файловая система): {result['file_system_info']['file_system_modified']}")
    print(f"   Дата создания (файловая система): {result['file_system_info']['file_system_created']}")
    
    # Выводим внутренние метаданные PDF
    print("\n📄 ВНУТРЕННИЕ МЕТАДАННЫЕ PDF (КАК В PDF-XCHANGE):")
    
    pdf_info = result['pdf_internal_info']
    
    if 'pdf_internal_modified' in pdf_info:
        mod_data = pdf_info['pdf_internal_modified']
        if isinstance(mod_data, dict):
            print(f"   📅 ПОСЛЕДНЯЯ ДАТА ИЗМЕНЕНИЯ СОДЕРЖИМОГО (ModDate): {mod_data['formatted']}")
            print(f"      (сырое значение: {mod_data['raw']})")
        else:
            print(f"   📅 ПОСЛЕДНЯЯ ДАТА ИЗМЕНЕНИЯ СОДЕРЖИМОГО (ModDate): {mod_data}")
    else:
        print("   ⚠️ Поле ModDate не найдено в метаданных PDF")
    
    if 'pdf_internal_created' in pdf_info:
        create_data = pdf_info['pdf_internal_created']
        if isinstance(create_data, dict):
            print(f"   📅 ДАТА СОЗДАНИЯ ДОКУМЕНТА (CreationDate): {create_data['formatted']}")
    
    # Другая информация
    print("\n📋 ДОПОЛНИТЕЛЬНАЯ ИНФОРМАЦИЯ О ДОКУМЕНТЕ:")
    for key, value in pdf_info.items():
        if key not in ['pdf_internal_modified', 'pdf_internal_created']:
            if isinstance(value, dict):
                print(f"   {key}: {value.get('formatted', value)}")
            else:
                print(f"   {key}: {value}")
    
    # Сравнение дат
    print("\n🔍 АНАЛИЗ:")
    fs_mod = result['file_system_info']['file_system_modified']
    
    if 'pdf_internal_modified' in pdf_info:
        pdf_mod_data = pdf_info['pdf_internal_modified']
        pdf_mod = pdf_mod_data['formatted'] if isinstance(pdf_mod_data, dict) else pdf_mod_data
        
        print(f"   Дата в файловой системе: {fs_mod}")
        print(f"   Внутренняя дата PDF:      {pdf_mod}")
        
        if fs_mod[:10] != pdf_mod[:10] if isinstance(pdf_mod, str) else False:
            print("\n   ⚠️ Даты РАЗЛИЧАЮТСЯ! Внутренняя дата изменения")
            print("      содержимого не совпадает с датой файла в проводнике.")
            print("      PDF-XChange Editor показывает внутреннюю дату (ModDate).")
        else:
            print("\n   ✅ Даты совпадают (по крайней мере по день).")
    else:
        print("   ⚠️ Невозможно сравнить даты: отсутствует ModDate")
    
    print("\n" + "=" * 60)

if __name__ == "__main__":
    # Проверяем наличие необходимых библиотек
    try:
        import PyPDF2
    except ImportError:
        print("Ошибка: Необходимо установить библиотеку PyPDF2")
        print("Установите её командой: pip install PyPDF2")
        sys.exit(1)
    
    main()
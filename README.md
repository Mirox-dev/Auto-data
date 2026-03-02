# Auto-data — Генератор ИУЛ

> **⚠️ Внимание:** С 1 марта согласно постановлению ([doc_iul.pdf](doc_iul.pdf)) ИУЛ **больше не принимается** в экспертизе.

*[Русский](#ru) | [English](#en)*

---

<a name="ru"></a>

## Описание

GUI-приложение на Python для автоматической генерации **Информационно-удостоверяющих листов (ИУЛ)** в формате `.docx`. Вычисляет контрольные суммы файлов (MD5 или CRC32), собирает метаданные и формирует структурированный документ в соответствии со стандартным бланком ИУЛ.

## Возможности

- Обработка одного файла или всей папки
- Выбор алгоритма хэширования: **MD5** или **CRC32**
- Настраиваемое количество строк для подписей работников
- Генерация документа `.docx` с таблицей: имя файла, дата изменения, размер, контрольная сумма
- Графический интерфейс на Tkinter

## Требования

- Python **3.7+**
- Библиотека `python-docx`

## Установка и запуск

```bash
# Установить зависимости
pip install -r requirements.txt

# Запустить приложение
python Авто_ИУЛ_1.1.py
```

## Использование

1. Нажмите **«Выбрать файл»** или **«Выбрать папку»**
2. Выберите тип контрольной суммы (MD5 / CRC32)
3. Введите имя выходного файла (без `.docx`)
4. Введите количество работников
5. Нажмите **«Создать УЛ»** — файл сохранится в той же папке, что и исходный

## Автор

Голуб Егор Евгеньевич — regooogolub@gmail.com

---

<a name="en"></a>

## Description

A Python GUI application for automatic generation of **Information and Certification Sheets (ICS)** in `.docx` format. Computes file checksums (MD5 or CRC32), collects metadata, and produces a structured document following the standard ICS template.

## Features

- Process a single file or an entire folder
- Choose hash algorithm: **MD5** or **CRC32**
- Configurable number of signature rows for personnel
- Generates a `.docx` table with: file name, modification date, file size, checksum
- Graphical interface built with Tkinter

## Requirements

- Python **3.7+**
- `python-docx` library

## Installation & Run

```bash
# Install dependencies
pip install -r requirements.txt

# Launch the application
python Авто_ИУЛ_1.1.py
```

## Usage

1. Click **"Select File"** or **"Select Folder"**
2. Choose the checksum type (MD5 / CRC32)
3. Enter the output file name (without `.docx`)
4. Enter the number of workers
5. Click **"Create ICS"** — the file is saved in the same directory as the source

## Author

Egor Golub — regooogolub@gmail.com

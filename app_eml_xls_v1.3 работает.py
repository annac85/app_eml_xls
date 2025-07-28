import os
import pandas as pd
import webbrowser
from tkinter import Tk, filedialog, messagebox, Button
from tkinter import ttk
from email import policy
from email.parser import BytesParser
import re

# Глобальные переменные для путей к папкам
input_folder = ""
output_folder = ""

# Функция для извлечения данных из текста
def extract_data_from_email(email_text):
    data = {}
    try:
        # 
        webinar_match = re.search(r'вебинар\s*["«»]*(.*?)(?=\n|$)', email_text, re.IGNORECASE)
        fio_match = re.search(r'ФИО обучающегося[:\s]*(.*?)(?=\n|$)', email_text, re.IGNORECASE)
        org_match = re.search(r'Название организации[:\s]*(.*?)(?=\n|$)', email_text, re.IGNORECASE)
        position_match = re.search(r'Должность[:\s]*(.*?)(?=\n|$)', email_text, re.IGNORECASE)
        phone_match = re.search(r'Контактный телефон[:\s]*(.*?)(?=\n|$)', email_text, re.IGNORECASE)
        email_match = re.search(r'Электронная почта слушателя[:\s]*(.*?)(?=\n|$)', email_text, re.IGNORECASE)
        
        # Доп переменные
        program_name_match = re.search(r'Название программы[:\s]*(.*?)(?=\n|$)', email_text, re.IGNORECASE)
        qualification_match = re.search(r'Образование, квалификация в соответствии с дипломом[:\s]*(.*?)(?=\n|$)', email_text, re.IGNORECASE)
        org_email_match = re.search(r'Электронная почта организации[:\s]*(.*?)(?=\n|$)', email_text, re.IGNORECASE)
        org_phone_match = re.search(r'Телефон организации[:\s]*(.*?)(?=\n|$)', email_text, re.IGNORECASE)

        # Заполняем данные, если они найдены
        data['Название вебинара'] = webinar_match.group(1).strip('«»') if webinar_match else 'нет данных'
        data['ФИО'] = fio_match.group(1).strip() if fio_match else 'нет данных'
        data['Название организации'] = org_match.group(1).strip() if org_match else 'нет данных'
        data['Должность'] = position_match.group(1).strip() if position_match else 'нет данных'
        data['Контактный телефон'] = phone_match.group(1).strip() if phone_match else 'нет данных'
        data['Электронная почта слушателя'] = email_match.group(1).strip() if email_match else 'нет данных'
        
        # Новые данные
        data['Название программы'] = program_name_match.group(1).strip() if program_name_match else 'нет данных'
        data['Образование, квалификация'] = qualification_match.group(1).strip() if qualification_match else 'нет данных'
        data['Электронная почта организации'] = org_email_match.group(1).strip() if org_email_match else 'нет данных'
        data['Телефон организации'] = org_phone_match.group(1).strip() if org_phone_match else 'нет данных'

    except AttributeError:
        return None
    return data

# Функция для чтения .eml файлов
def process_eml_file(eml_path):
    try:
        with open(eml_path, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)
            email_text = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
        return extract_data_from_email(email_text)
    except Exception as e:
        print(f"Ошибка при чтении файла {eml_path}: {e}")
        return None

# Функция для обхода папок и сбора данных
def process_directory(directory):
    collected_data = []
    for root, _, files in os.walk(directory):
        print(f"Обрабатывается папка: {root}")
        for file in files:
            if file.endswith('.eml'):
                eml_path = os.path.join(root, file)
                print(f"Обрабатывается файл: {eml_path}")
                email_data = process_eml_file(eml_path)
                if not email_data:
                    email_data = {
                        'Название вебинара': 'нет данных',
                        'ФИО': 'нет данных',
                        'Название организации': 'нет данных',
                        'Должность': 'нет данных',
                        'Контактный телефон': 'нет данных',
                        'Электронная почта слушателя': 'нет данных',
                        'Название программы': 'нет данных',
                        'Образование, квалификация': 'нет данных',
                        'Электронная почта организации': 'нет данных',
                        'Телефон организации': 'нет данных'
                    }
                    print(f"Нет данных в файле: {eml_path}")
                collected_data.append(email_data)
    return collected_data

# Функции для открытия ссылок
def open_website(event):
    webbrowser.open("https://web-domsolnca.ru")

def open_donation(event):
    webbrowser.open("https://yoomoney.ru/to/41001596412502")

# Функции для выбора папок и сохранения файла
def select_input_folder():
    global input_folder
    folder_path = filedialog.askdirectory()
    if folder_path:
        input_folder_label.config(text=f"Папка с .eml файлами: {folder_path}")
        input_folder = folder_path

def select_output_folder():
    global output_folder
    folder_path = filedialog.askdirectory()
    if folder_path:
        output_folder_label.config(text=f"Папка для сохранения: {folder_path}")
        output_folder = folder_path

def save_excel():
    try:
        if not (input_folder and output_folder):
            raise Exception("Выберите папку с .eml файлами и папку для сохранения Excel файла!")
        
        # Сбор данных
        data = process_directory(input_folder)
        
        if not data:
            raise ValueError("Нет данных для сохранения")
        
        file_name = filedialog.asksaveasfilename(
            initialdir=output_folder,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Сохранить файл как"
        )
        if not file_name:
            raise Exception("Не указан файл для сохранения!")
        
        df = pd.DataFrame(data)
        
        # Попробуйте использовать openpyxl для сохранения Excel файла
        df.to_excel(file_name, index=False, engine='openpyxl')
        messagebox.showinfo("Успех", f"Данные успешно сохранены в {file_name}")
    
    except Exception as e:
        messagebox.showerror("Ошибка", str(e))

# Инициализация главного окна
root = Tk()
root.title("Обработка .eml файлов и сохранение в Excel")
root.geometry("500x400")  # Увеличен размер окна
root.configure(bg='#e1f0fa')  # Светло-голубой фон

# Основной фрейм
frame = ttk.Frame(root, padding="20", relief='flat')
frame.pack(fill='both', expand=True)

# Кнопка для выбора папки с .eml файлами
input_folder_button = Button(frame, text="Выбрать папку с .eml файлами", command=select_input_folder, 
                             font=('Roboto', 16), bg='#3498db', fg='white', activebackground='#ADD8E6', borderwidth=0)
input_folder_button.pack(pady=10, fill='x')

# Метка для отображения пути к выбранной папке
input_folder_label = ttk.Label(frame, text="Папка не выбрана")
input_folder_label.pack(pady=5)

# Кнопка для выбора папки для сохранения Excel файла
output_folder_button = Button(frame, text="Выбрать папку для сохранения", command=select_output_folder, 
                              font=('Roboto', 14), bg='#3498db', fg='white', activebackground='#ADD8E6', borderwidth=0)
output_folder_button.pack(pady=10, fill='x')

# Метка для отображения пути к выбранной папке
output_folder_label = ttk.Label(frame, text="Папка не выбрана")
output_folder_label.pack(pady=5)

# Кнопка для генерации и сохранения Excel файла
save_button = Button(frame, text="Сохранить данные в Excel", command=save_excel, 
                     font=('Roboto', 14), bg='#3498db', fg='white', activebackground='#ADD8E6', borderwidth=0)
save_button.pack(pady=20, fill='x')

# Автор и ссылки
author_label = ttk.Label(frame, text="Автор: Анна Черкасова", foreground="blue", cursor="hand2")
author_label.pack(pady=5)
author_label.bind("<Button-1>", open_website)

donation_label = ttk.Label(frame, text="Автору на кофе", foreground="blue", cursor="hand2")
donation_label.pack(pady=5)
donation_label.bind("<Button-1>", open_donation)

# Запуск основного цикла
root.mainloop()

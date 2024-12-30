# -*- coding: utf-8 -*-

import re
import os
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import threading

class ConverterGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Конвертер VCF - Excel ©УОТЗ")
        self.root.geometry("400x200")
        
        # Центруємо вікно
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - 400) // 2
        y = (screen_height - 200) // 2
        self.root.geometry(f"400x200+{x}+{y}")
        
        # Створюємо та розміщуємо елементи інтерфейсу
        self.setup_ui()
        
    def setup_ui(self):
        # Основний контейнер
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Мітка поточного файлу
        self.file_label = ttk.Label(main_frame, text="Готовий до обробки файлів")
        self.file_label.grid(row=0, column=0, pady=10, sticky=tk.W)
        
        # Прогрес файлів
        self.files_progress = ttk.Progressbar(main_frame, length=360, mode='determinate')
        self.files_progress.grid(row=1, column=0, pady=5)
        
        # Мітка прогресу обробки контактів
        self.contacts_label = ttk.Label(main_frame, text="")
        self.contacts_label.grid(row=2, column=0, pady=5, sticky=tk.W)
        
        # Прогрес контактів
        self.contacts_progress = ttk.Progressbar(main_frame, length=360, mode='determinate')
        self.contacts_progress.grid(row=3, column=0, pady=5)
        
        # Кнопка старту
        self.start_button = ttk.Button(main_frame, text="Почати конвертацію", command=self.start_conversion)
        self.start_button.grid(row=4, column=0, pady=20)

    def clean_name(self, name):
        """Очищення імені від усіх символів крім букв та пробілів"""
        cleaned = re.sub(r'[^a-zA-Zа-яА-ЯёЁіІїЇєЄ\s]', '', name)
        cleaned = re.sub(r'\s+', ' ', cleaned)
        return cleaned.strip()

    def format_phone(self, phone):
        """Форматування телефону у формат 380XXXXXXXXX"""
        digits = re.sub(r'\D', '', phone)
        
        if len(digits) >= 9:
            if digits.startswith('380'):
                digits = digits[3:]
            elif digits.startswith('80'):
                digits = digits[2:]
            elif digits.startswith('0'):
                digits = digits[1:]
                
            digits = digits[-9:]
            
            if len(digits) == 9:
                return f"380{digits}"
        
        return ''

    def parse_vcf_contacts(self, filename):
        """Парсинг VCF файлу"""
        contacts = {}
        current_contact = {}
        
        try:
            with open(filename, 'r', encoding='utf-8') as file:
                for line in file:
                    line = line.strip()
                    
                    if line == "BEGIN:VCARD":
                        current_contact = {}
                        continue
                        
                    elif line == "END:VCARD":
                        if 'FN' in current_contact:
                            contacts[current_contact['FN']] = current_contact
                        continue
                    
                    if ':' in line:
                        field, value = line.split(':', 1)
                        base_field = field.split(';')[0]
                        current_contact[base_field] = value
        
        except Exception as e:
            messagebox.showerror("Помилка", f"Помилка читання файлу {filename}: {str(e)}")
            return {}
            
        return contacts

    def process_files(self):
        """Основний процес конвертації"""
        try:
            # Створюємо папку exports якщо її немає
            export_dir = 'exports'
            if not os.path.exists(export_dir):
                os.makedirs(export_dir)
            
            # Отримуємо список VCF файлів
            vcf_files = [f for f in os.listdir() if f.lower().endswith('.vcf')]
            
            if not vcf_files:
                messagebox.showinfo("Інформація", "VCF файли не знайдено у поточній директорії")
                self.root.quit()
                return
            
            # Налаштовуємо прогрес для файлів
            self.files_progress['maximum'] = len(vcf_files)
            
            # Обробляємо кожен файл
            for i, vcf_file in enumerate(vcf_files, 1):
                self.file_label['text'] = f"Обробка файлу: {vcf_file}"
                contacts = self.parse_vcf_contacts(vcf_file)
                
                if contacts:
                    # Налаштовуємо прогрес для контактів
                    self.contacts_progress['maximum'] = len(contacts)
                    
                    # Підготовка даних для Excel
                    contacts_list = []
                    for j, (name, details) in enumerate(contacts.items(), 1):
                        phone = details.get('TEL', '')
                        if isinstance(phone, list):
                            phone = phone[0]
                        
                        contact_data = {
                            'Ім\'я': self.clean_name(name),
                            'Телефон': self.format_phone(phone)
                        }
                        
                        if contact_data['Телефон']:
                            contacts_list.append(contact_data)
                        
                        # Оновлюємо прогрес контактів
                        self.contacts_progress['value'] = j
                        self.contacts_label['text'] = f"Оброблено контактів: {j} з {len(contacts)}"
                        self.root.update()
                    
                    # Зберігаємо в Excel
                    if contacts_list:
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        base_name = os.path.splitext(vcf_file)[0]
                        excel_filename = os.path.join(export_dir, f"{base_name}_{timestamp}.xlsx")
                        
                        df = pd.DataFrame(contacts_list)
                        df.to_excel(excel_filename, index=False, engine='openpyxl')
                
                # Оновлюємо прогрес файлів
                self.files_progress['value'] = i
                self.root.update()
            
            messagebox.showinfo("Готово", "Конвертацію успішно завершено!")
            self.root.quit()
            
        except Exception as e:
            messagebox.showerror("Помилка", f"Виникла помилка: {str(e)}")
            self.root.quit()

    def start_conversion(self):
        """Запуск конвертації в окремому потоці"""
        self.start_button['state'] = 'disabled'
        threading.Thread(target=self.process_files, daemon=True).start()

    def run(self):
        """Запуск програми"""
        self.root.mainloop()

if __name__ == "__main__":
    app = ConverterGUI()
    app.run()
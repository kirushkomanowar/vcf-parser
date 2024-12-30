# -*- coding: utf-8 -*-
import re
import os
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import threading
import quopri

class ConverterGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Конвертер VCF у Excel")
        self.root.geometry("400x200")
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - 400) // 2
        y = (screen_height - 200) // 2
        self.root.geometry(f"400x200+{x}+{y}")
        
        self.setup_ui()
        
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.file_label = ttk.Label(main_frame, text="Готовий до обробки файлів")
        self.file_label.grid(row=0, column=0, pady=10, sticky=tk.W)
        
        self.files_progress = ttk.Progressbar(main_frame, length=360, mode='determinate')
        self.files_progress.grid(row=1, column=0, pady=5)
        
        self.contacts_label = ttk.Label(main_frame, text="")
        self.contacts_label.grid(row=2, column=0, pady=5, sticky=tk.W)
        
        self.contacts_progress = ttk.Progressbar(main_frame, length=360, mode='determinate')
        self.contacts_progress.grid(row=3, column=0, pady=5)
        
        self.start_button = ttk.Button(main_frame, text="Почати конвертацію", command=self.start_conversion)
        self.start_button.grid(row=4, column=0, pady=20)

    def decode_quoted_printable(self, text):
        try:
            # Убираем переносы строк и пробелы в конце
            text = text.replace('=\r\n', '').replace('=\n', '').strip()
            # Декодируем из quoted-printable в байты, затем в UTF-8
            decoded_bytes = quopri.decodestring(text)
            return decoded_bytes.decode('utf-8')
        except Exception as e:
            print(f"Ошибка декодирования: {str(e)}")
            return text

    def parse_vcf_line(self, line):
        try:
            if ':' not in line:
                return None, None

            field_part, value = line.split(':', 1)
            field_parts = field_part.split(';')
            base_field = field_parts[0]
            
            # Проверяем параметры кодировки
            if 'QUOTED-PRINTABLE' in field_part.upper():
                decoded_value = self.decode_quoted_printable(value)
                return base_field, decoded_value
            return base_field, value

        except Exception as e:
            print(f"Ошибка парсинга строки: {str(e)}")
            return None, None

    def parse_vcf_contacts(self, filename):
        contacts = {}
        current_contact = {}
        
        try:
            with open(filename, 'r', encoding='utf-8', errors='replace') as file:
                lines = file.readlines()
                i = 0
                while i < len(lines):
                    line = lines[i].strip()
                    
                    if line == "BEGIN:VCARD":
                        current_contact = {}
                        phones = []
                    elif line == "END:VCARD":
                        if 'FN' in current_contact:
                            if phones:
                                current_contact['TEL'] = phones
                            contacts[current_contact['FN']] = current_contact
                    elif line.startswith('TEL'):
                        phone_number = line.split(':')[-1].strip()
                        phones.append(phone_number)
                    else:
                        field, value = self.parse_vcf_line(line) or (None, None)
                        if field and field != 'TEL':
                            current_contact[field] = value
                    i += 1
                    
            print(f"Прочитано контактов: {len(contacts)}")
                
        except Exception as e:
            messagebox.showerror("Помилка", f"Помилка читання файлу {filename}: {str(e)}")
            return {}
            
        return contacts

    def clean_name(self, name):
        try:
            cleaned = re.sub(r'[^\w\s\-А-Яа-яЁёІіЇїЄєҐґ]', '', name)
            cleaned = re.sub(r'\s+', ' ', cleaned)
            return cleaned.strip()
        except Exception as e:
            print(f"Ошибка очистки имени: {str(e)}")
            return name

    def format_phone(self, phone):
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

    def process_files(self):
        try:
            export_dir = 'exports'
            if not os.path.exists(export_dir):
                os.makedirs(export_dir)
            
            vcf_files = [f for f in os.listdir() if f.lower().endswith('.vcf')]
            
            if not vcf_files:
                messagebox.showinfo("Інформація", "VCF файли не знайдено у поточній директорії")
                self.root.quit()
                return
            
            self.files_progress['maximum'] = len(vcf_files)
            
            for i, vcf_file in enumerate(vcf_files, 1):
                self.file_label['text'] = f"Обробка файлу: {vcf_file}"
                contacts = self.parse_vcf_contacts(vcf_file)
                
                if contacts:
                    contacts_list = []
                    total_phones = sum(len(details.get('TEL', [])) for details in contacts.values())
                    self.contacts_progress['maximum'] = total_phones
                    processed_phones = 0
                    
                    for name, details in contacts.items():
                        clean_contact_name = self.clean_name(name)
                        phones = details.get('TEL', [])
                        
                        for phone_number in phones:
                            formatted_phone = self.format_phone(phone_number)
                            if formatted_phone:
                                contact_data = {
                                    'Ім\'я': clean_contact_name,
                                    'Телефон': formatted_phone
                                }
                                contacts_list.append(contact_data)
                            
                            processed_phones += 1
                            self.contacts_progress['value'] = processed_phones
                            self.contacts_label['text'] = f"Оброблено номерів: {processed_phones} з {total_phones}"
                            self.root.update()
                    
                    if contacts_list:
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        base_name = os.path.splitext(vcf_file)[0]
                        excel_filename = os.path.join(export_dir, f"{base_name}_{timestamp}.xlsx")
                        
                        df = pd.DataFrame(contacts_list)
                        df.to_excel(excel_filename, index=False, engine='openpyxl')
                
                self.files_progress['value'] = i
                self.root.update()
            
            messagebox.showinfo("Готово", "Конвертацію успішно завершено!")
            self.root.quit()
            
        except Exception as e:
            messagebox.showerror("Помилка", f"Виникла помилка: {str(e)}")
            self.root.quit()

    def start_conversion(self):
        self.start_button['state'] = 'disabled'
        threading.Thread(target=self.process_files, daemon=True).start()

    def run(self):
        self.root.mainloop()

def remove_duplicates(contacts):
    """
    Remove duplicate contacts based on name and phone number.
    Returns list of unique contacts.
    """
    seen = set()
    unique_contacts = []
    
    for contact in contacts:
        # Create tuple of name and phone for comparison
        identity = (contact.name, contact.phone)
        if identity not in seen:
            seen.add(identity)
            unique_contacts.append(contact)
    
    return unique_contacts

def process_vcf(input_file):
    contacts = parse_vcf(input_file)
    # Add this line after parsing
    contacts = remove_duplicates(contacts)
    return contacts


def main():
    if os.name == 'nt':
        os.system('chcp 65001')
    app = ConverterGUI()
    app.run()

if __name__ == "__main__":
    main()
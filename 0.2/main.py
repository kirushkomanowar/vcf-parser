import os
from cx_Freeze import setup, Executable

# Зависимости, которые нужно включить
build_options = {
    'packages': ['pandas', 'openpyxl'],
    'excludes': [],
    'include_files': []
}

# Создаем исполняемый файл
setup(
    name="VCF to Excel Converter",
    version="1.0",
    description="Конвертер VCF файлов в Excel",
    options={'build_exe': build_options},
    executables=[
        Executable(
            "vcf_converter.py",  # Имя вашего основного Python файла
            base=None,  # 'Win32GUI' для приложения без консоли
            target_name="VCF_Converter.exe"
        )
    ]
)
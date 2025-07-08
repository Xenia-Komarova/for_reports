# for_reports
A program for generating a report on representative expenses

Программа для генерации отчётов по представительским расходам.

Для использования необходимо установить 

python -m pip install python-docx

Для создания .exe

pip install pyinstaller
pyinstaller --onefile --noconsole --add-data "template_prikaz.docx;." --add-data "template_smeta.docx;." --add-data "template_otchet.docx;." index.py


Для обновления:
pyinstaller --onefile --noconsole --add-data "template_prikaz.docx;." --add-data "template_smeta.docx;." --add-data "template_otchet.docx;." index.py

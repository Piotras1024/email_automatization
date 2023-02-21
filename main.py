from RAW_docx import RAWDocx
from RAW_xlsx import RAWXlsx
import os

print(os.getcwd())


##[0] 0 oznacza pierwszego pracownika, [0][0] = im. naz. 1 prac, [0][1] text do 1 prac., [0][2] messe in outlook



# template_docx = RAWDocx('template.docx')
# template_docx.change_word_in_pattern('Name', 'Zmiana1')
# template_docx.change_word_in_pattern('{tekst1}', 'Zmiana2')
# template_docx.save_xml('document')
# template_docx.zip_RAW_objbect_docx()

xlsx = RAWXlsx("input.xlsm")
xlsx.fill_tab_with_workers()
number_of_workers = xlsx.number_of_workers()
excel_1kolumna_1pracownik = xlsx.workers_tab

#tworzy z każdego wiersza excel pracownika i wrzuca dyplom do Diploma
for worker in range(number_of_workers):
    name = excel_1kolumna_1pracownik[worker][0]
    template_docx = RAWDocx('template.docx')
    template_docx.change_word_in_pattern('Name', name)
    template_docx.change_word_in_pattern('{tekst1}', str(excel_1kolumna_1pracownik[worker][1]))
    template_docx.save_xml('document')
    template_docx.zip_RAW_objbect_docx(f"Diplomas/{name}.docx")

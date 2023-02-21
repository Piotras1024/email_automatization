import shutil
from bs4 import BeautifulSoup
import shutil


class RAWDocx:
    def __init__(self, doc_pattern_path):
        self.doc_pattern_path = doc_pattern_path
        self.root_dir = 'Docx_in_zip'   # gdzie idzie new_doc
        self.xmls_objects: dict[str: BeautifulSoup] = {}
        self.xmls_paths: dict[str: str] = {}
        self.unzip_RAW_object_docx()
        self.read_xml("document", "word/document.xml")
        self.all_wt = self.get_xml_object("document").find_all("w:t")

    #funkcja wypakowuje plik.docx i wkłada go do folderu diploma.
    def unzip_RAW_object_docx(self):
        shutil.unpack_archive(self.doc_pattern_path, self.root_dir, "zip", )

    #self.root_dir wskazuje gdzie ma być nowo utworzony dyplom. self.new_doc - to jego nazwa bez formatu.
    # def zip_RAW_objbect_docx(self):
    #     shutil.make_archive(self.new_doc, 'zip', self.root_dir)
    #     shutil.move(f'{self.new_doc}.zip', self.root_dir + f"{self.new_doc}.docx")
    def zip_RAW_objbect_docx(self, new_filename):
        # try:
        #     os.mkdir('Diplomas')
        # except FileExistsError:
        #     pass
        # try:
        #     os.remove(f'Diplomas/{excel_1kolumna_1pracownik[index][0]}{self.new_doc}.docx')
        # except FileNotFoundError:
        #     pass

        shutil.make_archive(new_filename, 'zip', self.root_dir)
        shutil.move(f"{new_filename}.zip", f"{new_filename}")
        shutil.rmtree('Docx_in_zip')



    #self.xmls_name = 'Document'
    def read_xml(self, xml_name, xml_fn):
        self.xmls_paths[xml_name] = xml_fn
        with open(f"{self.root_dir}/{xml_fn}", "r") as xml_file:
            xml_str = xml_file.read()
        self.xmls_objects[xml_name] = BeautifulSoup(xml_str, "xml")

    def save_xml(self, xml_name):
        with open(f"{self.root_dir}/{self.get_xml_path(xml_name)}", "w") as xml_file:
            xml_file.write(str(self.get_xml_object(xml_name)))

    def get_xml_path(self, xml_name) -> str:
        return self.xmls_paths[xml_name]

    def get_xml_object(self, xml_name) -> BeautifulSoup:
        return self.xmls_objects[xml_name]

    def change_word_in_pattern(self, old_word, new_word):
        for element in filter(lambda wt: old_word in wt.text, self.all_wt):
            element.string.replace_with(element.text.replace(old_word, new_word))



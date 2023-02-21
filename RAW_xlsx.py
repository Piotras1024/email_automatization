## potrzebuje klase która po plik.xlsx.make_list_from("Name and Surname")
## i tworzy liste z teho co pod spodobem.
import pandas as pd
import math



class RAWXlsx:
    def __init__(self, fn, ):
        self.fn = fn
        self.workers = pd.read_excel(f'{fn}')
        self.changed_fn = f"changed_{fn}"
        self.extract_dir = self.fn.replace(".", "_")
        self.workers_tab = []
        self.outlook = []

    def fill_tab_with_workers(self):
        for _, record in self.workers.iterrows():
            text = record['Text']
            name_surname = record['Name and Surname']
            message_in_outlook = record['Message in outlook']
            self.workers_tab.append([name_surname, text])
            self.outlook.append(message_in_outlook)

    def number_of_workers(self):
        i = 0
        for worker in self.workers_tab:
            try:
                if math.isnan(worker[0]):
                    return i
            except TypeError:
                pass
            i += 1


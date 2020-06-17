"""A simple class based xlsx serialization system

Functions
---------
This is where you can specify any functions and what they return.

Limitations
-----------
Due to the limitations of XlsxWritter (see: https://xlsxwriter.readthedocs.io/introduction.html ) the following are not possible:
- Deserialization of existing files is not supported
- Modifying/updating existing files is not supported

Examples
--------
Store some animal instances in a spreadsheet called 'animals.xlsx'

```
from ezexcel import Spreadsheet

class Animal():
    def __init__(self, name:str, conservation_status:str):
        self.name = name
        self.conservation_status = conservation_status

leopard_gecko = Animal('Leopard Gecko', 'Least Concern')

philippine_eagle = Animal('Philippine Eagle', 'Threatened')

with Spreadsheet('animals.xlsx', Animal) as output_sheet:
    output_sheet.store(leopard_gecko, philippine_eagle)
```
"""

from xlsxwriter import Workbook

class Spreadsheet():
    def __init__(self:spreadsheet, file_name:str, class_identifier:object):
        if not file_name.endswith(".xlsx"):
            file_name += ".xlsx"
        self.file_name = file_name
        self.class_identifier = class_identifier
        self.workbook = None
        self.worksheet = None

    def store(self, *instances):
        """Takes an instance of the specified class to store"""
        for current_instance in instances:
            if (type(current_instance) == list) or (type(current_instance) == tuple):
                #TODO
                pass
            elif not isinstance(instance, self.class_identifier):
                raise ValueError(f"Provided instance is not of type {self.class_identifier}")


    def __enter__(self):
        self.workbook = Workbook(self.file_name)
        self.worksheet = self.workbook.add_worksheet()
        return self

    def __exit__(self, type, value, traceback):
        self.workbook.close()
        return True
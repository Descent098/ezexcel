"""A simple class based xlsx serialization system

Limitations
-----------
Currently only classes with 51 or less attributes are supported

Examples
--------
### Store some animal instances in a spreadsheet called 'animals.xlsx'
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

### Store a list of instances into a spreadsheet called 'users.xlsx'
```
from ezexcel import Spreadsheet

import random
import string
from dataclasses import dataclass

@dataclass
class User():
    Name:str
    Age:int
    Weight:int
    Family: list # Note that Iterables will be flattened to a string with newline seperators

instances = []
ranstring = lambda: ''.join(random.choices(string.ascii_uppercase, k=10)) # Generates a random 10 character string
for i in range(1000):
    instances.append(User(ranstring(), random.randint(12,100), random.randint(75,400), [ranstring(), ranstring(), ranstring()]))

with Spreadsheet('users.xlsx', User) as output_sheet:
    output_sheet.store(instances)
```
"""
import logging                               # Used to log data for debugging
import datetime                              # Used to validate type assertions for datetime instances
from typing import Union, Iterable           # Used for type hinting and type assertions

import colored                               # Colours terminal output for emphasis
from openpyxl import Workbook                # Used to open and operate with xlsx files
from openpyxl.styles import Font, Alignment  # Used to pretty output to files


class Spreadsheet():
    """A class that takes in instances of objects and serializes them to xlsx files

    Parameters
    ----------
    file_name : (str)
        The name of the .xlsx file that will be saved out (extension can be included or excluded)

    class_identifier : (object)
        The class object for instances you want to store, see example(s) for details

    Raises
    ------
    ValueError

        In two cases:

            1. If instances provided to Spreadsheet.store() do not match type used to construct Spreadsheet instance
            2. If class provided has more than 51 attributes (see limitations section of docs for details)

    Examples
    --------
    #### Store some animal instances in a spreadsheet called 'animals.xlsx'
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
    def __init__(self, file_name:str, class_identifier:object):
        self.file_name = file_name
        self.class_identifier = class_identifier
        self.workbook = None
        self.worksheet = None

        # Make sure filename has .xlsx extension
        if not file_name.endswith(".xlsx"):
            logging.debug(f"added .xlsx to {file_name}")
            file_name += ".xlsx"

        # Get all attributes of class defined in __init__
        self.class_attributes = class_identifier.__init__.__code__.co_varnames[1::] # Skip the self
        if len(self.class_attributes) > 51:
            raise ValueError(f"Provided class {class_identifier.__name__} has more than 51 attributes")


    def __enter__(self):
        """Entrypoint for the context manager

        Returns
        -------
        Spreadsheet
            Reference to self
        """
        self.workbook = Workbook()
        self.worksheet = self.workbook.active
        self.worksheet.page_setup.fitToWidth = 1
        return self


    def __exit__(self, exc_type, exc_value, traceback):
        """Exitpoint for the context manager

        Returns
        -------
        bool
            True if the context manager ran into no issues saving files
        """
        if exc_type is None and exc_value is None:
            try:
                self.workbook.save(self.file_name)
                print(f"{self.file_name} successfully saved")
            except PermissionError:
                input(f"{colored.fg(1)}File {self.file_name} is currently open{colored.fg(15)}\nPlease close it and hit enter to save file: ")
                self.workbook.save(self.file_name)
            return True

        else:
            print(f"{colored.fg(1)}Ran into exception {exc_type}, with value {exc_value} here is traceback{colored.fg(15)}")
            return False


    def _add_row_to_spreadsheet(self, data:list, row:int, style:Font = False):
        """Take in some data, and an int for a row and add that data to that row

        Parameters
        ----------
        data : list
            The data you want to store in the row

        row : int
            The index of the row to store the data to

        style : Font, optional
            If you want to supply custom formatting for the row, by default False
        """
        # The value that will be converted using chr() for column identifiers i.e. A1 B1 etc.
        column_identifier = 65  # Initialize to ord() value of 'A'

        for value in data:
            if column_identifier == 91:  # Roll over to Ax column identifiers from x column identifiers
                label = f"AA{row}"
            elif column_identifier > 91:  # If beyond Z in column identifiers
                label = f"A{chr(column_identifier-26)}{row}"
            else:  # If before or at Z in column identifiers
                label = f"{chr(column_identifier)}{row}"
            logging.debug(f"{value} will be written to {label}")

            # Apply styles if specified
            if style:
                self.worksheet[label].font = style

            # Add value to worksheet
            if isinstance(value, Iterable) and not type(value) in [str, int, float]:
                # If value is an Iterable that's not a str, int or float then flatten it to a str
                flattened_value = ""
                for sub_value in value: 
                    flattened_value += f"{str(sub_value)}\n"
                self.worksheet[label] = flattened_value
            elif type(value) not in [str, int, float, datetime.datetime]:
                # Value is not a str, int, float or datetime object (all can be natively serialized)
                self.worksheet[label] = str(value)
            else: # If value is a string, int, float or datetime object
                self.worksheet[label] = value

            # Apply wrap text formatting to all rows that aren't the heading
            if not row == 1: 
                self.worksheet[label].alignment = Alignment(wrapText=True)

            # Increment the column identifiers variable to move to next column letter
            column_identifier += 1


    def _get_values_from_instance(self, instance:object) -> list:
        """Get's the instance's attribute values

        Parameters
        ----------
        instance : object
            The instance to pull the attribute values from

        Returns
        -------
        list
            The values for the attributes from the instance
        """
        logging.debug(f"Attributes are {self.class_attributes}")
        values = [] # All the values of the attributes in order
        for attribute in self.class_attributes:
            logging.debug(f"Looking for attribute {attribute} found value {instance.__dict__[attribute]}")
            values.append(instance.__dict__[attribute]) 
        return values


    def store(self, *instances:Union[object, Iterable[object]]):
        """Takes in instance(s) of the specified class to store

        Parameters
        ----------
        instances : (Iterable[object] or arbitrary number of isntances)
            The instances with the data you want to store

        Raises
        ------
        ValueError
            If an instance is not the correct type
        
        Examples
        --------
        #### Store some animal instances in a spreadsheet called 'animals.xlsx'
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
        print(f"Beginning to store {self.class_identifier.__name__} instances to {self.file_name}")
        current_row = 1  # The current row that the iteration is at

        # Add heading with the list of class attributes to A1
        self._add_row_to_spreadsheet(self.class_attributes, current_row, Font(bold=True, size=14))
        current_row += 1  # Increment row to start with row right after heading
        logging.debug(f"Instances are {instances}")

        # Check if instance provided is a class of correct type, or an Iterable
        for current_instance in instances:
            logging.debug(f"Instance is {str(current_instance)}")
            if isinstance(current_instance, Iterable):  # If argument is an Iterable (i.e. list, tuple etc.)
                for sub_instance in current_instance:
                    if not isinstance(sub_instance, self.class_identifier):  # Validate sub-instance is correct type
                        raise ValueError(f"Provided instance: {sub_instance} is not of type {self.class_identifier}")
                    else:
                        self._add_row_to_spreadsheet(self._get_values_from_instance(sub_instance), current_row)
                        current_row += 1
            elif not isinstance(current_instance, self.class_identifier):  # If argument is not correct type
                raise ValueError(f"Provided instance: {current_instance} is not of type {self.class_identifier}")
            
            else:  # If argument is a single class instance of the correct type
                logging.debug(f"Adding values from {str(current_instance)}: {self._get_values_from_instance(current_instance)}")
                self._add_row_to_spreadsheet(self._get_values_from_instance(current_instance), current_row)
                current_row += 1

![ezspreadsheet logo](https://raw.githubusercontent.com/Descent098/ezspreadsheet/master/.github/logo.png)

# EZ Spreadsheet

*A simple API to store/load python objects to/from spreadsheets*

## Table of contents
- [Goals](#goals)
- [Features](#features)
  - [Object Serialization](#object-serialization)
  - [Object Deserialization](#object-deserialization)
  - [Flexibility](#flexibility)
  - [Readability](#readability)
- [Installation](#installation)
- [Quick-start](#quick-start)
- [Additional Documentation](#additional-documentation)

## Goals
This project has a few goals:
1. Make OO projects easier to serialize
2. Make interfacing with spreadsheet files simple
3. Use as minimal syntax as possible to complete the above

## Features

Below are a some of the features in ezspreadsheet.

### Object serialization

The api lets you directly serialize object instances into spreadsheets:

```python
from ezspreadsheet import Spreadsheet

class Animal():
    def __init__(self, name:str, conservation_status:str):
        self.name = name
        self.conservation_status = conservation_status

leopard_gecko = Animal('Leopard Gecko', 'Least Concern')

with Spreadsheet('animals.xlsx', Animal) as output_sheet: # to use .csv just change file extension
    output_sheet.store(leopard_gecko)
```

### Object deserialization

The api lets you deserialize objects from spreadsheets back into instances (Note if you don't provide a class to build back into a dynamic [namedtuple](https://docs.python.org/3/library/collections.html#collections.namedtuple) is generated):

```python
from ezspreadsheet import Spreadsheet

class Animal():
    def __init__(self, name:str, conservation_status:str):
        self.name = name
        self.conservation_status = conservation_status

leopard_gecko = Animal('Leopard Gecko', 'Least Concern')

philippine_eagle = Animal('Philippine Eagle', 'Threatened')

# Store Data
with Spreadsheet('animals.xlsx', Animal) as output_sheet: # to use .csv just change file extension
    output_sheet.store(leopard_gecko, philippine_eagle)

# Retrieve data using the same class construtor
with Spreadsheet('animals.xlsx', Animal) as input_sheet: # to use .csv just change file extension
    animals, instances = input_sheet.load("animals")

print(instances) # prints: [<__main__.Animal object at 0x0000011BAB89A3A0>, <__main__.Animal object at 0x0000011BAD4289A0>]
# Note the class constructor is the same
print(Animal == animals) # Prints: True

# Retrieve namedtuple classes when no class constructor is available
with Spreadsheet('animals.xlsx') as input_sheet: # to use .csv just change file extension
    animals, instances = input_sheet.load("animals")

print(animals) # Prints: <class 'ezspreadsheet.animals'>
print(instances) # Prints: [animals(name='Leopard Gecko', conservation_status='Least Concern'), animals(name='Philippine Eagle', conservation_status='Threatened')]

# Note the class constructor is now different
print(Animal == animals) # Prints: False
```

### Flexibility

There is syntactic flexibility to allow an arbitrary number of instance arguments, or simple Iterables (like lists and tuples):

```python
from ezspreadsheet import Spreadsheet

class Animal():
    def __init__(self, name:str, conservation_status:str):
        self.name = name
        self.conservation_status = conservation_status

leopard_gecko = Animal('Leopard Gecko', 'Least Concern')

philippine_eagle = Animal('Philippine Eagle', 'Threatened')

# Direct instances
with Spreadsheet('animals.xlsx', Animal) as output_sheet: # to use .csv just change file extension
    output_sheet.store(leopard_gecko, philippine_eagle)

# Iterables
instances = []
instances.append(leopard_gecko)
instances.append(philippine_eagle)

with Spreadsheet('animals.xlsx', Animal) as output_sheet: # to use .csv just change file extension
    output_sheet.store(instances)
```

### Readability

You can specify a ```readable``` variable in the ```Spreadsheet.store()``` method to allow Iterable instance attributes to be written in a readable format (note they will be deserialized as strings):

```python
from ezspreadsheet import Spreadsheet
from dataclasses import dataclass

@dataclass
class User():
    Name:str
    Age:int
    Weight:int
    Family: list # Note that Iterables will be flattened to a string with newline seperators

jd = User("John Doe", 20, 75, ["Abby", "Mike", "Janice"])

# Store Data as readable
with Spreadsheet('users.xlsx', User) as output_sheet: # to use .csv just change file extension
    output_sheet.store(jd, readable=True)

# Retrieve namedtuple classes when no class constructor is available
with Spreadsheet('users.xlsx') as input_sheet: # to use .csv just change file extension
    users, instances = input_sheet.load("users")

print(instances[0].Family) # Prints: - Abby\n- Mike\n- Janice

# Store Data as not readable
with Spreadsheet('users.xlsx', User) as output_sheet: # to use .csv just change file extension
    output_sheet.store(jd)

# Retrieve namedtuple classes when no class constructor is available
with Spreadsheet('users.xlsx') as input_sheet: # to use .csv just change file extension
    users, instances = input_sheet.load("users")

print(instances[0].Family) # Prints: ['Abby', 'Mike', 'Janice']
```

## Installation

### From PyPi

1. Run ```pip install ezspreadsheet``` or ```sudo pip3 install ezspreadsheet```

### From source

1. Clone this repo: (https://github.com/Descent098/ezspreadsheet)
2. Run ```pip install .``` or ```sudo pip3 install .```in the root directory


## Quick-start

**Note that syntax does not change between ```.xlsx``` and ```.csv``` files, you can replace ```.xlsx``` with ```.csv``` and below examples will work**

### Store some animal instances in a spreadsheet called 'animals.xlsx', then read back the data
```python
from ezspreadsheet import Spreadsheet

class Animal():
    def __init__(self, name:str, conservation_status:str):
        self.name = name
        self.conservation_status = conservation_status

leopard_gecko = Animal('Leopard Gecko', 'Least Concern')

philippine_eagle = Animal('Philippine Eagle', 'Threatened')

# Store
with Spreadsheet('animals.xlsx', Animal) as output_sheet: # to use .csv just change file extension
    output_sheet.store(leopard_gecko, philippine_eagle)

# Load
with Spreadsheet('animals.xlsx', Animal) as input_sheet: # to use .csv just change file extension
    _, instances = output_sheet.load("animals")

print(instances) # prints: [<__main__.Animal object at 0x0000011BAB89A3A0>, <__main__.Animal object at 0x0000011BAD4289A0>]
```

### Store a list of instances into a spreadsheet called 'users.xlsx'
```python
from ezspreadsheet import Spreadsheet

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

with Spreadsheet('users.xlsx', User) as output_sheet: # to use .csv just change file extension
    output_sheet.store(instances)
```

## Differences between xlsx and csv

Please note there are some differnces between ```.xlsx``` files and ```.csv``` files:
1. ```.xlsx``` files are significantly faster because they isntantiate cell objects and are C-level optimized instead of just doing plain text generation. If you need to save thousands of objects, I would recommend using ```.xlsx``` files.
2. When passing ```readable=True``` to ```Spreadsheet.store()``` the formatting for ```.xlsx``` files allows for newlines, so iterables are broken by newlines ```\n```. CSV readers use newlines to read csv files (even when told not to), so they are broken by tabs ```\t``` in ```.csv``` files instead.

## Additional Documentation

Additional documentation can be found at https://kieranwood.ca/ezspreadsheet

For details on how contributing to the project, please see [CONTRIBUTING.md](https://github.com/Descent098/ezspreadsheet/blob/master/CONTRIBUTING.md), for details on upcoming changes see [our roadmap](https://github.com/Descent098/ezspreadsheet/projects).

For most recent changes see [CHANGELOG.md](https://github.com/Descent098/ezspreadsheet/blob/master/CHANGELOG.md).

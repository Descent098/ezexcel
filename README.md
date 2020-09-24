![execel logo](https://github.com/Descent098/sdu/blob/master/.github/logo.png)

# EZ Excel

*A simple class based xlsx serialization system*

## Table of contents
- [Goals](#goals)
- [Installation](#installation)
- [Quick-start](#quick-start)
- [Additional Documentation](#additional-documentation)

## Goals
This project has a few goals:
1. Make OO projects easier to serialize to xlsx
2. Make use of xlsx serialization and deserialization (coming soon) easier
3. Write the simplest to use api for xlsx serialization

## Installation

### From PyPi

1. Run ```pip install ezexcel``` or ```sudo pip3 install ezexcel```

### From source

1. Clone this repo: (https://github.com/Descent098/ez-excel)
2. Run ```pip install .``` or ```sudo pip3 install .```in the root directory


## Quick-start

### Store some animal instances in a spreadsheet called 'animals.xlsx'
```python
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
```python
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

## Additional Documentation

Additional documentation can be found at https://kieranwood.ca/ezexcel

For details on how contributing to the project, please see [CONTRIBUTING.md](https://github.com/Descent098/ezexcel/blob/master/CONTRIBUTING.md), for details on upcoming changes see [our roadmap](https://github.com/Descent098/ezexcel/projects).

For most recent changes see [CHANGELOG.md](https://github.com/Descent098/ezexcel/blob/master/CHANGELOG.md).

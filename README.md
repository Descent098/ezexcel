# EZ Excel

*A simple class based xlsx serialization system*

## Quick-start

*Include how people can get started using your project in the shortest time possible*

### Installation

#### From PyPi

1. Run ```pip install ezexcel``` or ```sudo pip3 install ezexcel```

#### From source

1. Clone this repo: (https://github.com/Descent098/ez-excel)
2. Run ```pip install .``` or ```sudo pip3 install .```in the root directory

#### Usage

*Include how to use your package as an API (if that's what you're going for)*

## Additional Documentation

Additional documentation can be found at https://kieranwood.ca/ezexcel

## Development-Contribution guide

### Installing development dependencies

There are a few dependencies you will need to use this package fully, they are specified in the extras require parameter in setup.py but you can install them manually:

```
nox   	# Used to run automated processes
pytest 	# Used to run the test code in the tests directory
```

Just go through and run ```pip install <name>``` or ```sudo pip3 install <name>```. These dependencies will help you to automate documentation creation, testing, and build + distribution (through PyPi) automation.

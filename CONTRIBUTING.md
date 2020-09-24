## Development-Contribution guide

### Installing development dependencies

There are a few dependencies you will need to use this package fully, they are specified in the extras require parameter in setup.py but you can install them manually:

```
nox   	# Used to run automated processes
pytest 	# Used to run the test code in the tests directory
```

Just go through and run ```pip install <name>``` or ```sudo pip3 install <name>```. These dependencies will help you to automate documentation creation, testing, and build + distribution (through PyPi) automation.

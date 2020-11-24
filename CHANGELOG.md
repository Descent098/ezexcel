# Changelog

## 0.2.1; November 24th 2020

- Updated description to match name change

## 0.2.0; November 24th 2020

- Added deserialization methods for xlsx files
- Added serialization and deserialization methods for CSV files
- Changed name from ezexcel to ezspreadsheet
- Split Xlxs processing to internal class and converted Spreadsheet class to dispatching class

## 0.1.1; September 25th 2020

Fixed logo loading on PyPi

### Bug fixes

- Fixed loading issue with logo on PyPi

## 0.1.0; September 25th

Initial release of EZ Excel

### Features

- Ability to provide a class to instantiate a Spreadsheet
- Ability to pass instances in an iterable of class to Spreadsheet to be serialized
- Automatically flatten Iterable attributes within instances to endline delimited strings
- Added testing suite for all existing functionality

### Documentation improvements

- Wrote all existing documentation :)

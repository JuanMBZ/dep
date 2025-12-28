# dep
Simple python script to convert excel files to dep.

## Requirements
* Install Python 3.+ in https://www.python.org/downloads/
* Then run the command below in the terminal after installing python:
```
pip install requirements.txt
```
OR
```
python -m pip install requirements.txt
```

## Usage
Open terminal in same directory as the downloaded scripts then call the script by typing:
`dep "excel_file"`  
* excel_file = path to excel file. (ctl-shift-c in windows to get path)  
More detailed info below:
```
usage: dep [-h] [-s [SHEET ...]] [-c COLUMN] [-z ZEROES] [-u] excel_file

positional arguments:
  excel_file            Location of excel book

options:
  -h, --help            show this help message and exit
  -s [SHEET ...], --sheet [SHEET ...]
                        Name of sheets to be converted to .dep. Default will check for all sheets in excel file.
  -c COLUMN, --column COLUMN
                        Set column number of the whole part of EMP_PAY
  -z ZEROES, --zeroes ZEROES
                        Add inputted number of zeroes to the left of EMP_ACCTNO
  -u, --sheetname       Sets output file names to their correspinding sheet name. Default uses the excel file's name.
```
* Square brackets are optional and can be left out

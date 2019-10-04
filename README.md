# xlsx2csv

It's a python3 script to convert XLSX file into CSV files by using [xlrd](https://pypi.org/project/xlrd/) library

## Install

1. install python 3.7+
2. install xlrd library `pip install xlrd`
3. git clone `https://github.com/peepeepopapapeepeepo/xlsx2csv.git`

## Usage

``` bash
xlsx2csv.py <xlsx> <sheet> <csv>
```

- \<xlsx\> = xlsx file
- \<sheet\> = sheet in that xlsx file. (default=Sheet1)
- \<csv\> = output csv file. (default=same as xlsx file except extension is csv)
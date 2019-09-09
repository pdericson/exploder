[![Build Status](https://travis-ci.org/pdericson/exploder.svg?branch=master)](https://travis-ci.org/pdericson/exploder)

# exploder

Exploder.

Usage:

```
$ ./exploder.py ---help
usage: exploder.py [-h] --worksheet1 WORKSHEET --worksheet2 WORKSHEET
                   --columns COLUMNS
                   path

Exploder.

positional arguments:
  path                  the workbook path

optional arguments:
  -h, --help            show this help message and exit
  --worksheet1 WORKSHEET
                        worksheet 1
  --worksheet2 WORKSHEET
                        worksheet 2
  --columns COLUMNS     the columns to explode
```

e.g.

```
./exploder.py --worksheet1 Sheet1 --worksheet2 Sheet2 --column '1,3' test1.xlsx
```

## development

```
./exploder_test.py
rm -rf vendor
pip3 install --target vendor openpyxl
```

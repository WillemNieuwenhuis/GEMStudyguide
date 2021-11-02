# Extract GEM study guide topics

## Purpose
The study guide information is stored in Excel workbooks.
This app extracts this information and turns it into
a Microsoft Word document (.docx). Optionally it can be 
saved as a PDF as well.

## Installation
A python installation is required, with the following (root) packages, the dependencies should be installed automatically:
 - pandas
 - openpyxl
 - python-docx
 - docx2pdf

The included **_requirement.yaml_** file can be used to create a conda environment.

The python-docx and docx2pdf modules need to be installed with pip.
The doc2pdf package could not be installed in a conda environment, so it was installed with the *--user* option

## Syntax
```
usage: gemguidemain.py [-h] [-d] [-p] [-v] source output

Generate GEM study guide

positional arguments:
  source         Input Excel workbook
  output         Output document name

optional arguments:
  -h, --help     show this help message and exit
  -d, --docx     Generate DOCX output (default)
  -p, --pdf      Generate PDF output
  -v, --version  show program's version number and exit
```

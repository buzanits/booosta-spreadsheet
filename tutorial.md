# Booosta Spreadsheet module - Tutorial

## Abstract

This tutorial covers the spreadsheet module of the Booosta PHP framework. If you are new to this framework, we strongly
recommend, that you first read the [general tutorial of Booosta](https://github.com/buzanits/booosta-installer/blob/master/tutorial/tutorial.md).

## Purpose

The purpose of this module is to provide functionality to read spreadsheet files and import the read data to PHP data structures. It also provides functions to write data to spreadsheet files. Currently this module supports the formats XLSX (Microsoft Excel) and ODS (Open Document Spreadsheet).

This module uses PhpOffice (https://github.com/PHPOffice) for access to xlsx and ods files.

## Installation

This module can be installed with

```
composer require booosta/spreadsheet
```

This also loads addtional dependent modules.

## Usage

### Read spreadsheet files

```
$sheet = $this->makeInstance('spreadsheet', 'myfile.xlsx');

# get_indexed_data($convert_to_utf8 = false, $use_header = true)

# get an array with the data indexed with the content of the first line
$data = $sheet->get_indexed_data();

# addtionally convert data to UTF-8
$data = $sheet->get_indexed_data(true);

# index the data with 0, 1, 2... instead of first line (first line is also interpreted as data)
$data = $sheet->get_indexed_data(false, false);


# get_mapped_data($mapping, $convert_to_utf8 = false, $use_header = true)

# in this example the header of the spreadsheet says 'firstname', but the result is indexed with 'First Name'
$mapping = ['firstname' => 'First Name', 'lastname' => 'Last Name'];
$data = $sheet->get_mapped_data($mapping);

# get an array with the header (first line of spreadsheet)
$data = $sheet->get_header();
```

### Write spreadsheet files

```
# provide the data in a two dimensional array
$data = [['name', 'phone', 'email'], ['Alice', '0303033', 'alice@example.com']];
$sheet = $this->makeInstance('spreadsheet');
$sheet->set_data($data);

# write teh loaded data into a spreadsheet file
$sheet->save('mydata.xlsx');
```

### Setting various flags

```
# get an additional array field with a hyperlink from the spreadsheet
# presumed there is a column "message" in the spreadsheet, that holds text and a hyperlink
# $data will have a field "message" and a field "message_hyperlink" holding only the hyperlink
$sheet->extract_hyperlinks();
$data = $sheet->get_indexed_data();

# automatically resize the column withs
$sheet->set_autoresize();
$sheet->save('mydata.xlsx');
```

### Functions from PhpSpreadsheet

You can use any function that is provided by PhpSpreadsheet on a `\PhpOffice\PhpSpreadsheet\Spreadsheet` object. See the [documentation](https://github.com/PHPOffice/PhpSpreadsheet/blob/master/docs/index.md) of PhpSpreadsheet.

```
$active_sheet = $sheet->getActiveSheet();
```

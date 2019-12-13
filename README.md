# PHP Excel encoder

[![Build Status](https://travis-ci.org/Ang3/php-excel-encoder.svg?branch=master)](https://travis-ci.org/Ang3/php-excel-encoder) [![Latest Stable Version](https://poser.pugx.org/ang3/php-excel-encoder/v/stable)](https://packagist.org/packages/ang3/php-excel-encoder) [![Latest Unstable Version](https://poser.pugx.org/ang3/php-excel-encoder/v/unstable)](https://packagist.org/packages/ang3/php-excel-encoder) [![Total Downloads](https://poser.pugx.org/ang3/php-excel-encoder/downloads)](https://packagist.org/packages/ang3/php-excel-encoder)

Encode and decode xls/xlsx files and more thanks to the component [phpoffice/phpspreadsheet](https://phpspreadsheet.readthedocs.io/en/latest/).

Also, this component uses the component ```symfony/serializer``` - Please read [documentation](https://symfony.com/doc/current/components/serializer.html) for more informations about serializer usage.

## Installation

You can install the component in 2 different ways:

- Install it via Composer
```
$ composer require ang3/php-excel-encoder
```

- Use the official Git repository (https://github.com/Ang3/php-excel-encoder).

Then, require the vendor/autoload.php file to enable the autoloading mechanism provided by Composer. 
Otherwise, your application won't be able to find the classes of this component.

## Usage

### Create encoder

```php
<?php

require_once 'vendor/autoload.php';

use Ang3\Component\Serializer\Encoder\ExcelEncoder;

// Create the encoder with default context
$encoder = new ExcelEncoder($defaultContext = []);
```

**Context parameters:**
- ```ExcelEncoder::NB_HEADERS_ROW_KEY``` (boolean): Count of header rows [default: ```1```]
- ```ExcelEncoder::FLATTENED_HEADERS_SEPARATOR_KEY``` (string): separator for flattened entries key [default: ```.```]
- ```ExcelEncoder::HEADERS_IN_BOLD_KEY``` (boolean): put headers in bold (encoding only) [default: ```true```]
- ```ExcelEncoder::HEADERS_HORIZONTAL_ALIGNMENT_KEY``` (string): put headers in bold (encoding only: ```left```, ```center``` or ```right```) [default: ```center```]
- ```ExcelEncoder::COLUMNS_AUTOSIZE_KEY``` (boolean): column autosize feature (encoding only) [default: ```true```]
- ```ExcelEncoder::COLUMNS_MAXSIZE_KEY``` (integer): column maxsize feature (encoding only) [default: ```50```]
- ```ExcelEncoder::CSV_DELIMITER_KEY``` (string): columns separator (CSV decoding only) [default: ```;```]
- ```ExcelEncoder::CSV_ENCLOSURE_KEY``` (string): Cell values enclosure (CSV decoding only) [default: ```"```]

### Encoding

**Accepted formats:**
- ```ExcelEncoder::XLS```
- ```ExcelEncoder::XLSX```

```php
<?php
// Create the encoder...

// Test data
$data = [
  // Array by sheet
  'My sheet' => [
    // Array by rows
    [
      'bool' => false,
      'int' => 1,
      'float' => 1.618,
      'string' => 'Hello',
      'object' => new DateTime('2000-01-01 13:37:00'),
      'array' => [
        'bool' => true,
        'int' => 3,
        'float' => 3.14,
        'string' => 'World',
        'object' => new DateTime('2000-01-01 13:37:00'),
        'array' => [
          'again'
        ]
      ],
    ],
  ]
];

// Encode data with specific format
$xls = $encoder->encode($data, ExcelEncoder::XLSX);

// Put the content in a file with format extension for example
file_put_contents('my_excel_file.xlsx', $xls);
```

### Decoding

**Accepted formats:**
- ```ExcelEncoder::XLS```
- ```ExcelEncoder::XLSX```
- ```ExcelEncoder::SPREADSHEET``` *The component ```phpspreadsheet``` will try to resolve format automatically*

Please read [phpspreadsheet documentation](https://phpspreadsheet.readthedocs.io/en/latest/) to know wich file format can be read.

**Accepted data format in 2019:**
- Open Document Format/OASIS (.ods)
- Office Open XML (.xlsx) Excel 2007 and above
- BIFF 8 (.xls) Excel 97 and above
- BIFF 5 (.xls) Excel 95
- SpreadsheetML (.xml) Excel 2003
- Gnumeric
- HTML
- SYLK
- CSV

**/!\ When the format does not support multi-sheets (CSV for example), the only one sheet is called "Worksheet".**

```php
<?php
// Create the encoder...

// Decode data with no specific format
$data = $encoder->decode('my_excel_file.xlsx', ExcelEncoder::SPREADSHEET);

var_dump($data);

// Output:
// 
// array(1) {
//  ["Sheet_0"]=> array(1) {
//    [0]=> array(15) {
//      ["bool"] => int(0)
//      ["int"] => int(1)
//      ["float"] => float(1.618)
//      ["string"] => string(5) "Hello"
//      ["object.date"] => string(26) "2000-01-01 13:37:00.000000"
//      ["object.timezone_type"] => int(3)
//      ["object.timezone"] => string(13) "Europe/Berlin"
//      ["array.bool"] => int(1)
//      ["array.int"] => int(3)
//      ["array.float"] => float(3.14)
//      ["array.string"] => string(5) "World"
//      ["array.object.date"] => string(26) "2000-01-01 13:37:00.000000"
//      ["array.object.timezone_type"] => int(3)
//      ["array.object.timezone"] => string(13) "Europe/Berlin"
//      ["array.array.0"] => string(5) "again"
//    }
//  }
// }

```

## Run tests

```$ git clone https://github.com/Ang3/php-excel-encoder.git```

```$ composer install```

```$ vendor/bin/simple-phpunit```

That's it!
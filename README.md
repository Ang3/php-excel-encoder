# PHP Excel encoder

[![Build Status](https://travis-ci.org/Ang3/php-excel-encoder.svg?branch=master)](https://travis-ci.org/Ang3/php-excel-encoder) [![Latest Stable Version](https://poser.pugx.org/ang3/php-excel-encoder/v/stable)](https://packagist.org/packages/ang3/php-excel-encoder) [![Latest Unstable Version](https://poser.pugx.org/ang3/php-excel-encoder/v/unstable)](https://packagist.org/packages/ang3/php-excel-encoder) [![Total Downloads](https://poser.pugx.org/ang3/php-excel-encoder/downloads)](https://packagist.org/packages/ang3/php-excel-encoder)

Encode and decode xls/xlsx files and more thanks to the component ```phpoffice/phpspreadsheet```. This component uses the component ```symfony/serializer``` - Please read [documentation](https://symfony.com/doc/current/components/serializer.html) for more informations about serializer usage.

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

Context parameters:
- ```ExcelEncoder::NB_HEADERS_ROW_KEY```: Count of header rows [default: ```1```]
- ```ExcelEncoder::FLATTENED_HEADERS_SEPARATOR_KEY```: separator for flattened entries key [default: ```.```]
- ```ExcelEncoder::HEADERS_IN_BOLD_KEY```: put headers in bold (encoding only: boolean) [default: ```true```]
- ```ExcelEncoder::HEADERS_HORIZONTAL_ALIGNMENT_KEY```: put headers in bold (encoding only: ```left```, ```center``` or ```right```) [default: ```center```]
- ```ExcelEncoder::COLUMNS_AUTOSIZE_KEY```: column autosize feature (encoding only: boolean) [default: ```true```]
- ```ExcelEncoder::COLUMNS_MAXSIZE_KEY```: column maxsize feature (encoding only: integer) [default: ```50```]

### Encoding

Accepted formats:
- ```ExcelEncoder::XLS```: ```xls```
- ```ExcelEncoder::XLSX```: ```xlsx```

```php
<?php

require_once 'vendor/autoload.php';

use Ang3\Component\Serializer\Encoder\ExcelEncoder;

// Create the encoder
$encoder = new ExcelEncoder;

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
//

// Put the content in a file with format extension
file_put_contents('my_excel_file.xlsx', $xls);
```

### Decoding

Accepted formats:
- ```ExcelEncoder::XLS```: ```xls```
- ```ExcelEncoder::XLSX```: ```xlsx```
- ```ExcelEncoder::WORKSHEET```: ```worksheet``` (the component will try to resolve format from ```phpspreadsheet```)

Please read [phpspreadsheet documentation](https://phpspreadsheet.readthedocs.io/en/latest/) to know wich file format can be read.

Accepted data format in 2019:
- Open Document Format/OASIS (.ods)
- Office Open XML (.xlsx) Excel 2007 and above
- BIFF 8 (.xls) Excel 97 and above
- BIFF 5 (.xls) Excel 95
- SpreadsheetML (.xml) Excel 2003
- Gnumeric
- HTML
- SYLK
- CSV

```php
<?php
// ...

// Create the encoder
$encoder = new ExcelEncoder;

// Decode data with no specific format
$data = $encoder->decode('my_excel_file.xlsx', ExcelEncoder::WORKSHEET);

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

That's it!
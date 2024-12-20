PHP Excel encoder
=================

[![Code Quality](https://github.com/assoconnect/php-excel-encoder/actions/workflows/build.yml/badge.svg)](https://github.com/assoconnect/php-excel-encoder/actions/workflows/build.yml)
[![Latest Stable Version](https://poser.pugx.org/assoconnect/php-excel-encoder/v/stable)](https://packagist.org/packages/assoconnect/php-excel-encoder)

Encode and decode xls/xlsx files and more thanks to the component 
[phpoffice/phpspreadsheet](https://phpspreadsheet.readthedocs.io/en/latest/).

Also, this component uses the component Serializer from package 
[```symfony/serializer```](https://packagist.org/packages/symfony/serializer) - Please read the 
[documentation](https://symfony.com/doc/current/components/serializer.html) for more information about serializer usage.

Summary
-------

- [Installation](#installation)
- [Upgrade](#upgrade)
- [Usage](#usage)
- [Symfony integration](#symfony-integration)
- [Run tests](#run-tests)


Installation
------------

Open a command console and execute the following command to download the latest stable version of this bundle:

```shell
composer require assoconnect/php-excel-encoder
```

If you install this component outside of a Symfony application, you must require the vendor/autoload.php 
file in your code to enable the class autoloading mechanism provided by Composer. 
Read [this article](https://symfony.com/doc/current/components/using_components.html) for more details.

Upgrade
-------

To upgrade from 1.x to 2.x or from 2.x to 3.x, please read the file [UPGRADE-2.0.md](/UPGRADE-2.0.md).

Usage
-----

### Create encoder

```php
<?php

require_once 'vendor/autoload.php';

use assoconnect\Serializer\Encoder\ExcelEncoder;

// Create the encoder with default context
$encoder = new ExcelEncoder($defaultContext = []);
```

**Context parameters:**
- ```ExcelEncoder::AS_COLLECTION_KEY``` (boolean): Load data as collection [default: ```true```]
- ```ExcelEncoder::FLATTENED_HEADERS_SEPARATOR_KEY``` (string): separator for flattened entries key [default: ```.```]
- ```ExcelEncoder::HEADERS_IN_BOLD_KEY``` (boolean): put headers in bold (encoding only) [default: ```true```]
- ```ExcelEncoder::HEADERS_HORIZONTAL_ALIGNMENT_KEY``` (string): put headers in bold (encoding only: ```left```, ```center``` or ```right```) [default: ```center```]
- ```ExcelEncoder::COLUMNS_AUTOSIZE_KEY``` (boolean): column autosize feature (encoding only) [default: ```true```]
- ```ExcelEncoder::COLUMNS_MAXSIZE_KEY``` (integer): column maxsize feature (encoding only) [default: ```50```]

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

```php
// Create the encoder...

// Decode data with no specific format
$data = $encoder->decode('my_excel_file.xlsx', ExcelEncoder::XLS);

var_dump($data);

// Output:
// 
// array(1) {
//  ["Sheet_0"] => array(1) { // Name of the sheet
//    [0] => array(15) { // First row
//      ["bool"] => int(0) // First cell
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

#### Headers

By default, the encoder loads data as collection: the header name is the key of each data. 
Indeed, the default value of context parameter ```ExcelEncoder::AS_COLLECTION_KEY``` is ```true```. 
To disable this feature, pass the context parameter as ```false```.

```php
$data = $encoder->decode('my_excel_file.xlsx', ExcelEncoder::XLS, [
    ExcelEncoder::AS_COLLECTION_KEY => false
]);
```

```
// Output:
// 
// array(1) {
//  ["Sheet_0"] => array(1) { // Name of the sheet
//    [0] => array(15) { // Headers row
//      0 => "bool" // First cell
//      1 => "int"
//      2 => "float"
//      3 => "string"
//      4 => "object.date"
//      5 => "object.timezone_type"
//      6 => "object.timezone"
//      7 => "array.bool"
//      8 => "array.int"
//      9 => "array.float"
//      10 => "array.string"
//      11 => "array.object.date"
//      12 => "array.object.timezone_type"
//      13 => "array.object.timezone"
//      14 => "array.array.0"
//    },
//    [1] => array(15) { // First data row
//      [0] => int(0) // First cell
//      [1] => int(1)
//      [2] => float(1.618)
//      [3] => string(5) "Hello"
//      [4] => string(26) "2000-01-01 13:37:00.000000"
//      [5] => int(3)
//      [6] => string(13) "Europe/Berlin"
//      [7] => int(1)
//      [8] => int(3)
//      [9] => float(3.14)
//      [10] => string(5) "World"
//      [11] => string(26) "2000-01-01 13:37:00.000000"
//      [12] => int(3)
//      [13] => string(13) "Europe/Berlin"
//      [14] => string(5) "again"
//    }
//  }
// }
```

Symfony integration
-------------------

The encoder implements the interfaces of the Serializer component. To enable this encoder, you could just 
configure it as service by adding the contents below into the file `config/services.yaml`:

```yaml
# config/services.yaml
services:
  # ...
  AssoConnect\Serializer\Encoder\ExcelEncoder: ~
```

Run tests
---------

```$ git clone https://github.com/assoconnect/php-excel-encoder.git```

```$ composer install```

```$ vendor/bin/simple-phpunit```

That's it!
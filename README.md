# PHP PDF component

[![Build Status](https://travis-ci.org/Ang3/php-pdf.svg?branch=master)](https://travis-ci.org/Ang3/php-pdf) [![Latest Stable Version](https://poser.pugx.org/ang3/php-pdf/v/stable)](https://packagist.org/packages/ang3/php-pdf) [![Latest Unstable Version](https://poser.pugx.org/ang3/php-pdf/v/unstable)](https://packagist.org/packages/ang3/php-pdf) [![Total Downloads](https://poser.pugx.org/ang3/php-pdf/downloads)](https://packagist.org/packages/ang3/php-pdf)

This component helps you to generate and merge PDF's.

**Tested platforms**

- Linux Ubuntu 18.04.3 LTS

## Requirements

### Google chrome

Install Google Chrome on your server.

#### On Ubuntu

Source: https://doc.ubuntu-fr.org/google_chrome

```console
$ sudo apt-get install google-chrome-stable
```

### pdfunite

#### On Ubuntu

See http://manpages.ubuntu.com/manpages/bionic/man1/pdfunite.1.html

## Installation

You can install the component in 2 different ways:

- Install it via Composer (ang3/php-pdf on Packagist)
- Use the official Git repository (https://github.com/Ang3/php-pdf).

Then, require the vendor/autoload.php file to enable the autoloading mechanism provided by Composer. 
Otherwise, your application won't be able to find the classes of this component.

## Usage

### Create the factory

```php
<?php

require_once 'vendor/autoload.php';

use Ang3\Component\Pdf\PdfFactory;

// Parameters default values
$parameters = [
	'chrome_path' => '/usr/bin/google-chrome-stable',
	'pdfunite_path' => '/usr/bin/pdfunite',
];

// Créate the factory with optional parameters
$factory = new PdfFactory($parameters = []);
```

### Generating PDF files

The parameter ```chrome_path``` is used to generate the PDF.

```php
<?php

// Create the factory
// ...

/**
 * You can pass an output to the factory
 * 
 * @var Symfony\Component\Console\Output\OutputInterface|null
 */
$output = null;

// Generate the PDF from content as temporary file
$file = $factory->createFromContent('Hello world!'); // temp file
// Of generate the PDF with specific target filename (optional output)
$file = $factory->createFromContent('Hello world!', '<target_file>', $output);

// Generate the PDF from URL as temporary file
$file = $factory->createFromUrl('<content_url>'); // temp file
// Of generate the PDF with specific target filename (optional output)
$file = $factory->createFromUrl('<content_url>', '<target_file>', $output);
```

Both methods return the generated filename.

### Merging PDF files

The parameter ```pdfunite_path``` is used to merge PDF files.

```php
<?php

// Create the factory
// ...

/**
 * You can pass an output to the factory
 * 
 * @var Symfony\Component\Console\Output\OutputInterface|null
 */
$output = null;

// Create and get the merged PDF filename
$file = $factory->merge('<target_file>', [
	'<pdf_1>',
	'<pdf_2>',
	// ...
], $output);
```

The method returns the merged filename.
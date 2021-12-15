CHANGELOG
=========

2.0
---

* PHP 8 support
* Symfony v6.x support
* Removed CSV support (prefer using the Symfony encoder instead).
* Deleted format ```ExcelEncoder::SPREADSHEET```
* Added context parameter ```ExcelEncoder::AS_COLLECION_KEY```
* Decoding without merging headers row (as no collection)
* Removed context parameter ```ExcelEncoder::NB_HEADER_ROWS_KEY```
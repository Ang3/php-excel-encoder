UPGRADE FROM 1.x to 2.x
=======================

Binary compatibility break (BC break).

**No CSV support anymore**

Support
-------

- Do not use this encoder to encode/decode CSV data.
- Replace the encoding format to ```ExcelEncoder::XLS``` or ```ExcelEncoder::XLSX```

Context
-------

- Delete usage of context parameter ```ExcelEncoder::NB_HEADER_ROWS_KEY``` (int).
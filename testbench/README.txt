README
------

run_test_and_diff.sh

Just a quick tool for diffing spreadsheets, from a baseline openoffice/libreoffice spreadsheet.
Requires xmllint and meld as command line tools in linux.  The idea is you can manipulate the
xlsx spreadsheet and then see what the resulting xml is, and diff it with your test.xlsx

```sudo apt-get install xmllint libxml2-utils```

Unit tests

The PHPUnit tests live in ../tests (run them from the project root with `composer test`).

# openpyxl 2.5.0 date-handling-test

## Summary

MS Excel dates are a horrible, buggy and unpredictable mess, and
nobody who records time or date in any capacity should use
it. However, word has not gotten around yet.

The developers of the great [`openpyxl` module](https://bitbucket.org/openpyxl/openpyxl) recently introduced an improved handling of dates. A couple of test however show that this is no silver bullet. Based on the results below maybe the best advice is to not under any circumstances use a numerical data-type, including "Date" types in Excel and rather express time related stuff as strings. You can't calculate with dates then, you say? If you want to calculate anything you should not have used Excel in the first place! [:grumpy:].

## Testdata

Problems in the past have revolved around times and dates that get somewhere expressed as `0`, around Excel's counterfactual belief that 1900 was a leap year, and around the earliest allowed date (somewhere around 1900-01-01). We use a couple of test-dates and times that seem relevant in this regard, see `datetest.py` and images below. 

## Setup

+ MS Excel 2010 on Windows 7 Enterprise   
Note that differing versions of anything above, let alone using Excel for Mac, will likely change all outcomes.
+ I could not be bothered to play with different language setting of OS and Office. It is both "Swiss German", I believe.
+ openpyxl 2.5.0

## Usage

Run `./datetest.py` and follow instructions.

## Tests

### 1. Roundtrip

`openpyxl is used to write a), a set of dates, b), a set of times and c), a set of datetimes to an xlsx file. Then `openpyxl is used to read these data again from the file and write them back, into additional columns. Note: The columns starting with "string_" are string representations of the respective dates to serve as a reference as what was the true value should be. This is done both with parameter `iso_dates=True` for the initial creation of the workbook, and without. The resulting files are `testdates.roundtrip_isodates.xlsx` and `testdates_roundtrip.xlsx`, respectively.

![**Roundtrip with `iso_dates=True`**: looks OK!](./img/roundtrip_isodates.png)

![**Roundtrip without `iso_dates`**: re-written date sports Jan 0th, has funny formatting and midnight doesn't work.Note that `1900-01-01`has changed to `1900-01-00`even in the column `written_date`, which had correct value when the file was written the first time.](./img/roundtrip.png)

### 2. Excel Interrupt

Just like "**Roundtrip**" except that after the initial file is written, the user has opened it with Excel and saved it under a different name. Without applying any changes! The resulting files are `testdates_isodates_xlinterrupt.xlsx` and `testdates_xlinterrupt.xlsx`, respectively.

![**Excel Interrupt with `iso_dates=True`**: Same botched results as above.](./img/xlinterrupt_isodates.png)

![**Excel Interrupt without `iso_dates`**: Same botched results as above](./img/xlinterrupt.png)

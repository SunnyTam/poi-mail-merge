[![Build Status](https://travis-ci.org/centic9/poi-mail-merge.svg)](https://travis-ci.org/centic9/poi-mail-merge) [![Gradle Status](https://gradleupdate.appspot.com/centic9/poi-mail-merge/status.svg?branch=master)](https://gradleupdate.appspot.com/centic9/poi-mail-merge/status)

This is a small application which allows to repeatedely replace markers in a Microsoft Word document with items taken from a CSV/Microsoft Excel file. 

I started this project as I was quite disappointed with the functionality that LibreOffice offers, I especially wanted something that is repeatable/automateable
and does not produce spurious strange results and also does not need re-configuration each time the mail-merge is (re-)run.

## How it works

All you need is a Word-Document in Excel >= 2003 format (.doc) which acts as template and an Excel .xls/.xlsx or CSV file which contains one row for each time the template should be populated.

The word-document can contain template-markers (enclosed in ${...}) for things that should be replaced, e.g. "${first-name} ${last-name}".

The first sheet of the Excel/CSV file is read as a header-row which is used to match the template-names used in the Word-template.

The result is a single merged Word-document which contains a copy of the template for each line in the Excel file.

## Use it

### Grab and compile it

    git clone git://github.com/centic9/poi-mail-merge
    cd poi-mail-merge
    ./gradlew installDist

### Run it

    ./run.sh <word-template> <excel/csv-file> <output-file>

### Sample files

There are some sample files in the directory `samples`, you can run these as follows

    ./gradlew installDist
	build\install\poi-mail-merge\bin\poi-mail-merge.bat samples\Template.docx samples\Lines.xlsx build\Result.docx

on Unix you can use the following steps

    ./gradlew installDist
	./run.sh samples/Template.docx samples/Lines.xlsx build/Result.docx
	
## Tips

### Convert to PDF

You can use the too ```unoconv``` to further convert the resulting docx, e.g. to PDF:

    unoconv -vvv --timeout=10 --doctype=document --output=result.pdf result.docx

## Known issues

### Only one CSV format supported

Currently only CSV files which use comma as delimiter and double-quotes for quoting text are supported. Other formats required code-changes, but should be easy to to by adjusting the CSFFormat definition (it uses add as it uses http://commons.apache.org/proper/commons-csv/ for CSV handling).

### Word-Formatting can confuse the replacement

If there are multiple formattings applied to a strings that holds a template-pattern, the resulting XML-representation of the document might be split into multiple XML-Tags and thus might prevent the replacement from happening. 

A workaround is to use the formatting tool in LibreOffice/OpenOfficeto ensure that the replacement tags have only one formatting applied to them. 

See #6 for possible improvements.
## Change it

### Create Eclipse project files

    ./gradlew eclipse

### Build it and run tests

    cd poi-mail-merge
    ./gradlew check jacocoTestReport

#### Licensing

* poi-mail-merge is licensed under the [BSD 2-Clause License].
* A few pieces are imported from other sources, the source-files contain the necessary license pieces/references.

[BSD 2-Clause License]: http://www.opensource.org/licenses/bsd-license.php

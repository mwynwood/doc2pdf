# doc2pdf
Compiles Word Documents and PDF files into one big PDF file.

## About
I created doc2pdf so that I could quickly and easily create "Assessment Packages". These "Assessment Packages" are typically made up of a number of documents, so having them all together in one PDF is a lot more convient than having them in multimple seperate files, all in different formats.

With doc2pdf, it is possible to merge any combination of .doc, .docx, and .pdf files together to create one PDF file. You also have the option to add a Cover Page to the merged document, and save your settings to recall and edit later.

## Screenshot
<img src="https://github.com/mwynwood/doc2pdf/blob/master/screenshot.png">

## Save File Format
doc2pdf allows you to save your settings, so you can easily load them up later.

These settings are saved in plain text files with the extension ".doc2pdf"

Each line in the file represents one setting. Here is the format:
```
Include Cover Page. Value will be: True|False
Delete PDFs after Merge. Value will be: True|False
Cover Page Line 1. Value will be a string.
Cover Page Line 2. Value will be a string.
Cover Page Line 3. Value will be a string.
Cover Page Line 4. Value will be a string.
Logo. Value will be a string with the path of the image.
Spare spot for a future setting.
Spare spot for a future setting.
Spare spot for a future setting.
Spare spot for a future setting.
Spare spot for a future setting.
Spare spot for a future setting.
Spare spot for a future setting.
Spare spot for a future setting.
Spare spot for a future setting.
Spare spot for a future setting.
The remaining lines contain the path to each .doc, .docx, or .pdf file
```
## Installing
doc2pdf compiles into one little EXE file.

As long as you've got the .NET Framework and Microsoft Office installed, it should run.

You can download an EXE here: https://github.com/mwynwood/doc2pdf/blob/master/doc2pdf.exe

## Built With
* [PDFSharp](http://www.pdfsharp.net/) - Open Source .NET library that easily creates and processes PDFs
* [Microsoft Word](https://www.office.com/) - doc2pdf uses "Microsoft Word 2016" to convert Word Documents to PDF.

## History
This program is based on a little PowerShell script I wrote to do the same thing.

You can see the PowerShell version here: https://github.com/mwynwood/doc2pdf-PS

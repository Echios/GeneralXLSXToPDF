# Disclaimer:
Software provided as it is, with all bugs and features!
I take NO responsibility for lost data/corrupted files
Advice: make a backup prior to run this program! (General a good idea for Data)


# How To:
copy .exe file from the folder "dist" to target folder. double klick will start a command window and execute the programm


## Function:

- checks for ./export folder in its current folder, if not: makes it
- searches for .xlsx files in its current folder including subfolder, excluding currently opened files (fslename starts with ~)
- searches for "Laborjournal" in Cell(1,2) (resembles B1 and merged cells) of all sheets of the workbook
- saves Sheet as PDF in ./export with pdfname= filename+sheetname

Report any bugs and requests on github "https://github.com/Echios/GeneralXLSXToPDF" or via Mail (adress will be added later)

# Testing environment:
    Win10 Home x64
    Python 3.7.9 (anaconda)
    Excel (Office 365, Version2009 Build 16.0.13231.20110, 32bit)
used compiler for exe: pyinstaller (https://www.pyinstaller.org/)

# To Do
- [ ] add pdf merger part (merge all pdfs into one)

 
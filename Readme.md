Disclaimer:
Software provided as it is, with all bugs and features!
I take NO responsibility for lost data/corrupted files
Advice: make a backup prior to run this program! (General a good idea for Data)


How To:
copy .exe file from the folder "dist" to target folder. double klick will start a command window and execute the programm


Function:
-checks for ./export folder in its current folder, if not: makes it
-searches for .xlsx files in its current folder including subfolder, excluding currently opened files (filename starts with ~)
-searches for "Laborjournal" in Cell(1,2)[resembles B1 and merged cells] of all sheets of the workbook
saves Sheet as PDF in ./export with pdfname= filename+sheetname

Report any bugs and requests to "winterle@gmx.de"




used compiler for exe: pyinstaller
https://www.pyinstaller.org/
@echo off
set "sourceFolder=C:\Users\AL\OneDrive\Documents\CM"
set "destinationFolder=C:\Users\AL\OneDrive\Documents\ICT"

:: Convert DOCX to PDF using Microsoft Word
for %%F in ("%sourceFolder%\*.docx") do (
    "C:\Program Files\Microsoft Office\root\Office16\winword.exe" /MFilePrintDefault /MFileExit /q "%%~fF"
)

echo Conversion complete.
pause

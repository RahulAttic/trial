@echo off
setlocal enabledelayedexpansion

rem Create a temporary folder for extracted pages
mkdir temp

rem Loop through all PDF files in the directory
for %%f in (*.pdf) do (
    pdftk "%%f" cat 2 output "temp\page2_%%~nf.pdf"
)

rem Merge all the second pages into one PDF
pdftk temp\*.pdf cat output merged.pdf

rem Clean up
rmdir /s /q temp

echo Merged second pages into merged.pdf

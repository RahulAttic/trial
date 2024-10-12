Dim excelApp, workbook, sheet
Dim excelFile, pdfFileName, folderPath, file
Dim fso, folder

' Path to the folder containing Excel files (change this to your actual folder path)
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set folderPath =  objFSO.GetFolder(".")

'folderPath = "C:\Users\RAHUL KUMAR\Desktop\Bank Document\Yashpal ji\07-08-2024\NBC\Excel" ' Change this to your folder path

' Create a FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

' Create an instance of Excel
Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False

' Get the folder containing Excel files
Set folder = fso.GetFolder(folderPath)

' Loop through each file in the folder
For Each file In folder.Files
    ' Check if the file is an Excel file
    If LCase(fso.GetExtensionName(file.Name)) = "xlsx" Or _
       LCase(fso.GetExtensionName(file.Name)) = "xls" Then
       
        ' Open the Excel workbook
        Set workbook = excelApp.Workbooks.Open(file.Path)

        ' Loop through each sheet in the workbook
        For Each sheet In workbook.Sheets
            ' Create the PDF filename
            pdfFileName = fso.GetParentFolderName(file.Path) & "\" & _
                          fso.GetBaseName(file.Name) & "_" & sheet.Name & ".pdf"
            
            ' Export the sheet as PDF
            sheet.ExportAsFixedFormat 0, pdfFileName
        Next

        ' Close the workbook
        workbook.Close False
    End If
Next

' Quit Excel
excelApp.Quit

' Clean up
Set sheet = Nothing
Set workbook = Nothing
Set excelApp = Nothing
Set fso = Nothing
Set folder = Nothing

WScript.Echo "Conversion completed for all Excel files."

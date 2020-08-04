Sub OpenAndImportTxtFile()
'Open a txt to import all of its data

    Dim wkbTXTFile As Workbook
    Dim wksImport As Worksheet

    'Set sheet to import data to
    Set wksImport = Worksheets("Name of sheet") 'MODIFY THIS

    'Open txt file to import data from
    Set wkbTXTFile = Workbooks.Open("Path of txt file") 'MODIFY THIS

    'Copy data
    wkbTXTFile.Sheets(1).Cells.Copy wksImport.Cells

    'Close txt file
    wkbTXTFile.Close savechanges:=False
End Sub

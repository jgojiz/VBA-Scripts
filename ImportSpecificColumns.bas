Attribute VB_Name = "modImportSpecificColumns"
Option Explicit

Sub CopySpecificColumns()
'Copy specific columns from data source sheet to destination sheet

    Dim wksSource As Worksheet, wksDestination As Worksheet, wksCriteria As Worksheet
    Dim lastRow As Long
    Dim arrColumns As Variant
    
    'Get source and destination sheet
    Set wksSource = shData
    Set wksCriteria = shCriteria
    Set wksDestination = ThisWorkbook.Sheets.Add(After:=Worksheets(ThisWorkbook.Worksheets.Count))
    
    'Get last row in source sheet
    lastRow = wksSource.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Get column numbers to copy, col1 are columns from source, col2 are columns from destination
    arrColumns = wksCriteria.Range("A7:B12").Value
    
    'Loop through each column to copy from source to destination
    Dim i As Integer
    For i = LBound(arrColumns) To UBound(arrColumns)
        wksSource.Columns(arrColumns(i, 1)).Copy wksDestination.Columns(arrColumns(i, 2))
    Next i
    
End Sub

Sub ImportSpecificColumns()
'Open a file to copy specific columns to a sheet

    Dim wksSource As Worksheet, wksDestination As Worksheet, wksCriteria As Worksheet
    Dim wkb As Workbook ''if source sheet is in current workbook delete this line
    Dim lastRow As Long
    Dim arrColumns As Variant
    
    'set sheets in thisworkbook
    Set wksCriteria = shCriteria
    Set wksDestination = ThisWorkbook.Sheets.Add(After:=Worksheets(ThisWorkbook.Worksheets.Count))
    
    'Set source
    Set wkb = Workbooks.Open(ThisWorkbook.Path & "\Sample Tutelas.txt") 'if source sheet is in current workbook delete previous line
    Set wksSource = wkb.Worksheets(1)
    
    'Get last row in source sheet
    lastRow = wksSource.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Get column numbers to copy, col1 are columns from source, col2 are columns from destination
    arrColumns = wksCriteria.Range("A7:B12").Value
    
    'Loop through each column to copy from source to destination
    Dim i As Integer
    For i = LBound(arrColumns) To UBound(arrColumns)
        wksSource.Columns(arrColumns(i, 1)).Copy wksDestination.Columns(arrColumns(i, 2))
    Next i
    
    wkb.Close savechanges:=False 'if source sheet is in current workbook delete this line
End Sub









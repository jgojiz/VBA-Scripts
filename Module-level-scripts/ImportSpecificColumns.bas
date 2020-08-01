Sub CopySpecificColumns()
'Copy specific columns from data source sheet in thisworkbook to destination sheet

    Dim wksSource As Worksheet, wksDestination As Worksheet, wksCriteria As Worksheet
    Dim lastRow As Long
    Dim arrColumns As Variant
    
    'Get source and destination sheet
    Set wksSource = Worksheets("Name of source sheet") 'MODIFY THIS
    Set wksCriteria = Worksheets("Name of sheet where column numbers range is stored") 'MODIFY THIS
    Set wksDestination = ThisWorkbook.Sheets.Add(After:=Worksheets(ThisWorkbook.Worksheets.Count))
    
    'Get last row in source sheet
    lastRow = wksSource.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Get column numbers to copy, col1 are columns from source, col2 are columns from destination
    arrColumns = wksCriteria.Range("range in criteria sheet").Value 'MODIFY THIS
    
    'Loop through each column to copy from source to destination
    Dim i As Integer
    For i = LBound(arrColumns) To UBound(arrColumns)
        wksSource.Columns(arrColumns(i, 1)).Copy wksDestination.Columns(arrColumns(i, 2))
    Next i
    
End Sub

Sub ImportSpecificColumnsFromFile()
'Open a file to copy specific columns to a sheet
'diff from previous sub: add line 41 to open source file and add line 56 to close it

    Dim wksSource As Worksheet, wksDestination As Worksheet, wksCriteria As Worksheet
    Dim wkb As Workbook ''if source sheet is in current workbook delete this line
    Dim lastRow As Long
    Dim arrColumns As Variant
    
    'set sheets in thisworkbook
    Set wksCriteria = Worksheets("Name of sheet where column numbers range is stored") 'MODIFY THIS
    Set wksDestination = ThisWorkbook.Sheets.Add(After:=Worksheets(ThisWorkbook.Worksheets.Count))
    
    'Set source
    Set wkb = Workbooks.Open("Path of source data file") 'MODIFY THIS
    Set wksSource = wkb.Worksheets(1)
    
    'Get last row in source sheet
    lastRow = wksSource.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Get column numbers to copy, col1 are columns from source, col2 are columns from destination
    arrColumns = wksCriteria.Range("range in criteria sheet").Value 'MODIFY THIS
    
    'Loop through each column to copy from source to destination
    Dim i As Integer
    For i = LBound(arrColumns) To UBound(arrColumns)
        wksSource.Columns(arrColumns(i, 1)).Copy wksDestination.Columns(arrColumns(i, 2))
    Next i
    
    wkb.Close savechanges:=False
End Sub



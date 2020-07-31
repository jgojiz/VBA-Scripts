Sub SortColumns()
'Rearrange columns in a table
    
    'Prevent switching windows and alerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim wksData As Worksheet, wksCriteria As Worksheet, wksTemp As Worksheet
    Dim arrColumns As Variant
    
    'Set thisworkbook sheets
    Set wksData = Worksheets("NameOfDataSheet") 'MODIFY THIS
    Set wksCriteria = Worksheets("Name of sheet where column numbers range is stored") 'MODIFY THIS
    
    'Get column numbers to sort. Col1 original order, Col2 new order
    arrColumns = wksCriteria.Range("range from criteria sheet").Value 'MODIFY THIS
    
    'Temporary wks to put reorder columns
    Set wksTemp = Worksheets.Add(After:=Worksheets(ThisWorkbook.Worksheets.Count))
    
    'Loop through each column to CUT data to temporary sheet
    'This cleans the data sheet
    Dim i As Integer
    For i = LBound(arrColumns) To UBound(arrColumns)
        wksData.Columns(arrColumns(i, 1)).Cut wksTemp.Columns(arrColumns(i, 2))
    Next i
    
    'Copy reordered columns from temp sheet to data sheet
    wksTemp.Cells.Copy wksData.Range("A1")
    
    'delete temporary sheet
    wksTemp.Delete
    
    'Activate data sheet in Excel
    wksData.Activate
    
    'Restore Excel properties
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

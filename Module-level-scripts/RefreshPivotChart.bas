'Put one of these inside the worksheet object code that contains the data table for your pivot chart
'Remember that the data table should be formated as Excel Table

Private Sub Worksheet_Change (ByVal Target As Range)
'This refreshes all pivot caches and queries in the workbook
    ThisWorkbook.RefresAll
End Sub

Private Sub Worksheet_Change (ByVal Target As Range)
'This refreshes all pivot caches
    Dim pvtCache As PivotCache 
    For Each pvtCache In ThisWorkbook.PivotCaches
        pvtCache.Refresh
    Next pc
End Sub

Private Sub Worksheet_Change (ByVal Target As Range)
'This refreshes a specific pivot cache
    Worksheets("Name of sheet").PivotTables("Name of Pivot Table").PivotCache.Refresh 'MODIFY THIS
End Sub


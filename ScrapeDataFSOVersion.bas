'Requires checking Microsoft Scripting Runtime reference
Sub SourceDataOnMultipleFiles()
'Loop through files in folder and open Excel files (source) to copy its data in this wkb (destination)
'Your destination sheet should have the appropriate headers

    Dim wkbSource As Workbook
    Dim wksDestination As Worksheet
    Dim lngLastRowDest As Long, lngLastRowSource As Long
    Dim fso As New FileSystemObject
    Dim fo As Folder
    
    'Speed up your code
    Application.ScreenUpdating = False
    
    'Destination sheet for all data gathered
    Set wksDestination = ThisWorkbook.Worksheets("Name of sheet of destination") 'MODIFIY THIS
    
    'Define path of folder with files to Source data from
    Set fo = fso.GetFolder("PATH") 'MODIFY THIS
    
    'Loop through files in folder
    Dim f As File
    Dim wks As Worksheet
    For Each f In fo.Files
        'Identify Excel file extention
        If fso.GetExtensionName(f.Name) = "xlsx" Then
            'Open Excel file to Source
            Set wkbSource = Workbooks.Open(f.Path)
            
            'Loop through each sheet in the Source file
            For Each wks In wkbSource.Worksheets
                'update last row in destination sheet
                lngLastRowDest = wksDestination.Cells(Rows.Count, 1).End(xlUp).Row + 1
                
                'Copy data from source sheet (without headers) to destination sheet
                wks.Range("A1").CurrentRegion.Offset(1).Copy wksDestination.Cells(lngLastRowDest, 1)
            Next wks
            'Close Source file
            wkbSource.Close
        End If
    Next
    'Restore settings
    Application.ScreenUpdating = True
End Sub







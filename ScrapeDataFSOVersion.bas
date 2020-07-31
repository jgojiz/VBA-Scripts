'TODO: Improve script, there are 3 loops!
Sub ScrapeDataOnMultipleFiles()
'Loop through all files in a folder to open Excel files to copy values and join them in current sheet

    Dim scrapeWkb As Workbook
    Dim destinationWks As Worksheet
    Dim fso As New FileSystemObject
    Dim fo As Folder
    Dim lastRowDestination As Long, lastRowScrape As Long
    
    'Save RAM by preventing window switching
    Application.ScreenUpdating = False
    
    'Destination sheet for all data gathered
    Set destinationWks = ThisWorkbook.Worksheets("NameOfSheetWithData")
    
    'Define path of folder with files to scrape data from
    Set fo = fso.GetFolder("PATH")
    
    'Loop through files in folder
    Dim f As File
    Dim x As Long
    Dim wks As Worksheet
    For Each f In fo.Files
        'Identify Excel file extention
        If fso.GetExtensionName(f.Name) = "xlsx" Then
            'update last row in destination sheet
            lastRowDestination = destinationWks.Cells(Rows.Count, 1).End(xlUp).Row + 1
            
            'Open Excel file to scrape
            Set scrapeWkb = Workbooks.Open(f.Path)
            
            'Loop through each sheet in the scrape file
            For Each wks In scrapeWkb.Worksheets
                'update last row in scrape sheet
                lastRowScrape = wks.Cells(Rows.Count, 1).End(xlUp).Row
                
                'Loop through each row
                For x = 2 To lastRowScrape
                    'Assign values from scrape row to destination row
                    With destinationWks
                        .Cells(lastRowDestination, 1).Value = wks.Cells(x, 1).Value
                        .Cells(lastRowDestination, 2).Value = wks.Cells(x, 2).Value
                        .Cells(lastRowDestination, 3).Value = wks.Cells(x, 3).Value
                        .Cells(lastRowDestination, 4).Value = wks.Cells(x, 4).Value
                        .Cells(lastRowDestination, 5).Value = wks.Cells(x, 5).Value
                        .Cells(lastRowDestination, 6).Value = wks.Cells(x, 6).Value
                        .Cells(lastRowDestination, 7).Value = wks.Cells(x, 7).Value
                    End With
                    'update last row in destination sheet
                    lastRowDestination = lastRowDestination + 1
                Next x
            Next
            'Close scrape file
            scrapeWkb.Close
        End If
    Next
    Application.ScreenUpdating = True
End Sub









'TODO: Improve script, there are 3 loops!
Sub ScrapeDataOnMultipleFiles()
'Loop through all files in a folder to open Excel files to copy values and join them in current sheet

    Dim scrape_workbook As Workbook
    Dim destination_sheet As Worksheet, scrape_sheet As Worksheet
    Dim fso As New FileSystemObject
    Dim fo As Folder
    Dim last_row_destination As Long, last_row_scrape As Long
    
    'Save RAM by preventing window switching
    Application.ScreenUpdating = False
    
    'Destination sheet for all data gathered
    Set destination_sheet = ThisWorkbook.Worksheets("NameOfSheetWithData")
    
    'Define path of folder with files to scrape data from
    Set fo = fso.GetFolder("PATH")
    
    'Loop through files in folder
    Dim f As File
    Dim x As Long
    For Each f In fo.Files
        'Identify Excel file extention
        If fso.GetExtensionName(f.Name) = "xlsx" Then
            'update last row in destination sheet
            last_row_destination = destination_sheet.Cells(Rows.Count, 1).End(xlUp).Row + 1
            
            'Open Excel file to scrape
            Set scrape_workbook = Workbooks.Open(f.Path)
            
            'Loop through each sheet in the scrape file
            For Each scrape_sheet In scrape_workbook.Worksheets
                'update last row in scrape sheet
                last_row_scrape = scrape_sheet.Cells(Rows.Count, 1).End(xlUp).Row
                
                'Loop through each row
                For x = 2 To last_row_scrape
                    'Assign values from scrape row to destination row
                    With destination_sheet
                        .Cells(last_row_destination, 1).Value = scrape_sheet.Cells(x, 1).Value
                        .Cells(last_row_destination, 2).Value = scrape_sheet.Cells(x, 2).Value
                        .Cells(last_row_destination, 3).Value = scrape_sheet.Cells(x, 3).Value
                        .Cells(last_row_destination, 4).Value = scrape_sheet.Cells(x, 4).Value
                        .Cells(last_row_destination, 5).Value = scrape_sheet.Cells(x, 5).Value
                        .Cells(last_row_destination, 6).Value = scrape_sheet.Cells(x, 6).Value
                        .Cells(last_row_destination, 7).Value = scrape_sheet.Cells(x, 7).Value
                    End With
                    'update last row in destination sheet
                    last_row_destination = last_row_destination + 1
                Next x
            Next
            'Close scrape file
            scrape_workbook.Close
        End If
    Next
    Application.ScreenUpdating = True
End Sub









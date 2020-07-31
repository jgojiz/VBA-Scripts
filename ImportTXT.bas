Sub OpenAndImportTxtFile()
    Dim importWkb As Workbook, txtFile As Workbook
    Dim importWks As Worksheet

    Set importWkb = ThisWorkbook
    Set importWks = importWkb.Sheets("sheet name") 'Sheet where you want to import

    Set txtFile = Workbooks.Open("FilePath") 'The txt file is opened as a xlsx file well formatted

    txtFile.Sheets(1).Cells.Copy importWks.Cells

    txtFile.Close savechanges:=False
End Sub
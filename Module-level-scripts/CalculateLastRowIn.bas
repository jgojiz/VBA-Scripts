' Make your code cleaner with this function
Function CalculateLastRowIn(wks as Worksheet) as Long
    CalculateLastRowIn = wks.Cells(Rows.Count, 1).End(xlUp).Row
End Function

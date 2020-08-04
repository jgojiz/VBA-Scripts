' Passing worksheet code name
Function CalculateLastRowIn(wks as Worksheet) as Long
    CalculateLastRowIn = wks.Cells(Rows.Count, 1).End(xlUp).Row
End Function

' Passing worksheet name
Function CalculateLastRowIn(wksName as String) as Long
    CalculateLastRowIn = Worksheets(wksName).Cells(Rows.Count, 1).End(xlUp).Row
End Function

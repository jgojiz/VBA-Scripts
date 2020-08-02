' This function provides similar behaviour as XLOOKUP function, only available in Microsoft 365
' https://support.microsoft.com/en-us/office/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929
'
' Put this function in a module and save the workbook as .xlam, for Add-in. Then add it in
' Data -> Excel Add-ins -> Browse the .xlam file -> Click Ok
'
' After calling the function in Excel, press CTRL -> SHIFT -> A (Windows) to show the arguments
Function XLOOKUP(lookup_value As Variant, lookup_array As Range, return_array As Range)
    'Recalculate function whenever calculation occurs in any cells on the worksheet
    Application.Volatile True
    XLOOKUP = WorksheetFunction.Index(return_array, _
                                        WorksheetFunction.Match(lookup_value, _
                                                                    lookup_array, 0))
End Function
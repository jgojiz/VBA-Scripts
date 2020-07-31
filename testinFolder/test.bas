'just testing nested folders in git
Sub testinvbaextension()
    '// add declarations
    On Error GoTo catchError
    Dim objvarName As Object
    For Each wks In thisworkbook.worksheets 
        'blabla
    Next wks

    Function functionName() As functionType
        '// add declarations
        On Error GoTo catchError
    exitFunction:
        Exit Function
    catchError:
        '// add error handling
        GoTo exitFunction
    End Function

    If expression Then
        
    End If

    If expression Then
        
    ElseIf expression Then
        
    Else
        
    End If

    Select Case testexpression
        Case expressionlist-n
            statements-n
        Case Else
            elsestatements
    End Select

exitSub:
    Exit Sub
catchError:
    '// add error handling
    GoTo exitSub
End Sub
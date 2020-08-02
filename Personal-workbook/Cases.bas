' Put this subs in a module inside your PERSONAL.XLSB macro workbook
' So they're available eache time you open up Excel

Private rngCell As Range

Private Sub UPPERCASE()
   For Each rngCell In Selection
      rngCell.Value = UCase(rngCell.Value)
   Next
End Sub

Private Sub lowercase()
   For Each rngCell In Selection
      rngCell.Value = LCase(rngCell.Value)
   Next
End Sub

Private Sub ProperCase()
   For Each rngCell In Selection
      ' There is not a Proper function in VBA, 
      ' so you need to use this form:
      rngCell.Value = Application.Proper(rngCell.Value)
   Next
End Sub

'Sentence case for accentuated words (like in spanish)
Private Sub SentenceCaseAccentuated()
    Dim s, ch As String
    Dim start As Boolean
    Dim i As Long
    
    For Each rngCell In Selection.Cells
        s = rngCell.Value
        start = True
        
        If Not IsEmpty(s) Then
            For i = 1 To Len(s)
                'Extract the i-th char in s
                ch = Mid(s, i, 1)
                
                Select Case ch
                    Case "."
                        start = True
                    Case "?"
                        start = True
                    Case "a" To "z"
                        If start Then ch = UCase(ch): start = False
                    Case "A" To "Z"
                        If start Then start = False Else ch = LCase(ch)
                    Case "Á"
                        If start Then start = False Else ch = "á"
                    Case "É"
                        If start Then start = False Else ch = "é"
                    Case "Í"
                        If start Then start = False Else ch = "í"
                    Case "Ó"
                        If start Then start = False Else ch = "ó"
                    Case "Ú"
                        If start Then start = False Else ch = "ú"
                    Case "Ñ"
                        If start Then start = False Else ch = "ñ"
                End Select
                ' Replaces the i-th char in s
                Mid(s, i, 1) = ch
            Next
            rngCell.Value = s
        End If
    Next
End Sub




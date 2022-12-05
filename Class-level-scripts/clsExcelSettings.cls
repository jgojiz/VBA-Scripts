' This class provides the means to manipulate application settings
' to speed up code execution in VBA.
' Adapted from https://excelmacromastery.com/

' General usage:
'   Dim settings as New clsExcelSettings ( class name )
'   settings.TurnOff
'   
'   ------YOUR CODE-----
'
'   settings.TurnOn
'
' If you have specific settings you want to save later:
'   Dim settings as New clsExcelSettings
'   settings.Backup
'   settings.Turnof
'   
'   ------YOUR CODE-----
'
'   settings.Restore

Private calculation As XlCalculation
Private displayStatus As Boolean
Private enableEvents As Boolean
Private screenUpdating As Boolean
Private displayalerts As Boolean

' Purpose: Backup Current settings
Public Sub Backup()
    calculation = Application.calculation
    displayStatus = Application.DisplayStatusBar
    enableEvents = Application.enableEvents
    screenUpdating = Application.screenUpdating
    displayalerts = Application.displayalerts
End Sub

' Purpose: Restore Current settings
Public Sub Restore()
    Application.calculation = calculation
    Application.DisplayStatusBar = displayStatus
    Application.enableEvents = enableEvents
    Application.screenUpdating = screenUpdating
End Sub

' Purpose: Turn off automatic calculations, events and screen updating
Public Sub TurnOff()
    Application.calculation = xlCalculationManual
    Application.DisplayStatusBar = False
    Application.enableEvents = False
    Application.screenUpdating = False
    Application.displayalerts = False
End Sub

' Purpose: Turn on automatic calculations, events and screen updating
Public Sub TurnOn()
    Application.calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    Application.enableEvents = True
    Application.screenUpdating = True
    Application.displayalerts = True
End Sub





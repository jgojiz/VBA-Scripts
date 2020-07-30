Attribute VB_Name = "modUnir"
Option Explicit

Sub Consolidacion_Oficinas()
    
    'Declarar variables
    Dim wb As Workbook, wsThis As Worksheet, ws As Worksheet
    Dim fso As New FileSystemObject
    Dim fo As Folder
    Dim r As Long, lr As Long 'variables fila o row
    
    ' no cambiar de ventana
    Application.ScreenUpdating = False
    
    'Definir la ruta donde estan los archivos
    Set wsThis = ThisWorkbook.Worksheets("Datos")
    Set fo = fso.GetFolder("D:\Google Drive\VBA\Excel Labs\Loop through files to scrape data\Files")
    
    Dim f As File, x As Long
    For Each f In fo.Files
        
        'Identificar las extensiones de Excel xlsx
        If fso.GetExtensionName(f.Name) = "xlsx" Then
            
            r = wsThis.Cells(Rows.Count, 1).End(xlUp).Row + 1 'actualiza la ultima fila del consolidado
            
            Debug.Print r
            
            'Abrir el archivo
            Set wb = Workbooks.Open(f.Path)
            
            'Copiar y pegar la informacion que necesito
            For Each ws In wb.Worksheets
                lr = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
                For x = 2 To lr
                    wsThis.Cells(r, 1).Value = ws.Cells(x, 1).Value
                    wsThis.Cells(r, 2).Value = ws.Cells(x, 2).Value
                    wsThis.Cells(r, 3).Value = ws.Cells(x, 3).Value
                    wsThis.Cells(r, 4).Value = ws.Cells(x, 4).Value
                    wsThis.Cells(r, 5).Value = ws.Cells(x, 5).Value
                    wsThis.Cells(r, 6).Value = ws.Cells(x, 6).Value
                    wsThis.Cells(r, 7).Value = ws.Cells(x, 7).Value
                    
                    r = r + 1
                Next x
        
            Next
            
            wb.Close
            'Debug.Print f.Name
    
        End If
    Next
    
    Application.ScreenUpdating = True
End Sub









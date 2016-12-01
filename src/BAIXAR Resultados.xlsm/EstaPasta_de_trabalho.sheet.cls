Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    Call ExportaCodigo
End Sub

Private Sub Workbook_Open()
'
'    Dim wMenu As Worksheet
'
'    Set wMenu = Sheets("Menu")
'
'    wMenu.Activate
'
'    wMenu.Range("M6").Activate
    
End Sub
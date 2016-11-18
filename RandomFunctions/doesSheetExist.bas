Function doesSheetExist( _
                        ByVal SheetName As String, _
                        Optional ByRef wkbk As Workbook) As Boolean
    'Variables
    Dim sht As Worksheet    'Generic worksheet
    If wkbk Is Nothing Then
        Set wkbk = ActiveWorkbook
    End If
    SheetName = LCase(SheetName)
    
    doesSheetExist = False  'Defaults to not existing
    
    For Each sht In wkbk.Sheets
        If LCase(sht.Name) = SheetName Then
            doesSheetExist = True
            Exit For
        End If
    Next sht
    
End Function

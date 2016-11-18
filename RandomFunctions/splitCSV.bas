Function splitCSV( _
                    ByVal toSplit As String, _
                    Optional ByVal trimValue As Boolean = True) As Variant
    'Variables
    Dim counter As Long         'Generic Counter
    Dim strSplit() As String    'Array to hold the split values
    
    strSplit = Split(toSplit, ",")  'Split by comma (CSV)
    
    If trimValue = True Then
        For counter = 0 To UBound(strSplit)
            strSplit(counter) = Trim(strSplit(counter)) 'Trim all the spaces from front and end of values
        Next counter
    End If
    
    splitCSV = strSplit 'Set array for return
    
End Function

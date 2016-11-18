Sub removeColumn( _
                 ByVal stringHeaderToRemoveColumn As String, _
                 Optional ByRef shtWorking As Worksheet, _
                 Optional ByVal OnlyRemoveIfRowOne As Boolean = False)

    If shtWorking Is Nothing Then _
        Set shtWorking = ActiveSheet

    Dim rngFound As Range
    
    Set rngFound = findMe(stringHeaderToRemoveColumn, shtWorking.UsedRange)
    
    If rngFound Is Nothing Then
        'Do nothing. Nothing was found.
    Else
        If rngFound.Column > 0 Then
            If OnlyRemoveIfRowOne = True And _
               rngFound.Row = 1 Then _
                    rngFound.EntireColumn.Delete
            Else
                rngFound.EntireColumn.Delete
            End If
        End If
    End If

End Sub

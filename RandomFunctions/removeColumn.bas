Sub removeColumn( _
                 ByVal strHeader As String, _
                 Optional ByVal boolRemoveAllInstances As Boolean = False, _
                 Optional ByRef shtWorking As Worksheet, _
                 Optional ByVal headerRow As Long = 1)
    '23 Nov 2016
    'Removing all instances will remove all that is found in the given row. _
    '   Care should be taken when setting to true, and headerRow is 0. You may _
    '   remove much more than anticipated.
    'If headerRow is 0, then anywhere it is will be searched and removed. _
    '   Otherwise, only the "header row" given will be searched and removed.
    
    'Set ShtWorking
    If shtWorking Is Nothing Then _
        Set shtWorking = ActiveSheet
    
    Dim rngFound As Range   'Range we find and then remove that column
    Dim rngSearch As Range  'Range to search to remove the column(s)
    
    'Set the range to search, depending on where we want to remove things
    If headerRow > 0 Then _
        Set rngSearch = shtWorking.Rows(headerRow) Else _
        Set rngSearch = shtWorking.UsedRange
    
    'Try to find at least one instance
    Set rngFound = findMe(strHeader, rngSearch)
    
    If rngFound Is Nothing Then
        'Do nothing. Nothing was found.
    Else
        If boolRemoveAllInstances = True Then
            'Removing all instances
            Do While Not rngFound Is Nothing
                If rngFound.Column > 0 Then _
                    rngFound.EntireColumn.Delete
                Set rngFound = findMe(strHeader, rngSearch)
            Loop
        Else
            'Only remove one instance
            If rngFound.Column > 0 Then _
                rngFound.EntireColumn.Delete
        End If
    End If
End Sub

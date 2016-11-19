Sub sortListbox( _
                ByRef ourList As MSForms.ListBox, _
                Optional ByVal columnIndexToSort As Long = 0, _
                Optional ByVal Ascending As Boolean = True)
    Dim counterOuter As Long
    Dim counterInner As Long
    Dim counterColumns As Long
    Dim numColumns As Long
        numColumns = ourList.ColumnCount
    Dim temp As Variant
        
    With ourList
        'Bubble Sort
        '   Taking each value and comparing it to the next value, _
        '   bubbling the values we want to the top
        For counterOuter = 0 To .ListCount - 2
            For counterInner = 0 To .ListCount - 2
                If Ascending = True Then
                    If LCase(.List(counterInner, columnIndexToSort)) > LCase(.List(counterInner + 1, columnIndexToSort)) Then
                        For counterColumns = 0 To numColumns - 1
                            temp = .List(counterInner, counterColumns)
                            .List(counterInner, counterColumns) = .List(counterInner + 1, counterColumns)
                            .List(counterInner + 1, counterColumns) = temp
                        Next counterColumns
                    End If
                Else
                    If LCase(.List(counterInner, columnIndexToSort)) < LCase(.List(counterInner + 1, columnIndexToSort)) Then
                        For counterColumns = 0 To numColumns - 1
                            temp = .List(counterInner, counterColumns)
                            .List(counterInner, counterColumns) = .List(counterInner + 1, counterColumns)
                            .List(counterInner + 1, counterColumns) = temp
                        Next counterColumns
                    End If
                End If
            Next counterInner
        Next counterOuter
    End With
End Sub

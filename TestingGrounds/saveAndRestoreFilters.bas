Option Explicit

Sub WhatFilters()
    Dim iFilt As Integer
    Dim i, j As Integer
    Dim numFilters As Integer
    Dim crit1 As Variant

    If Not ActiveSheet.AutoFilterMode Then
        Debug.Print "Please enable AutoFilter for the active worksheet"
        Exit Sub
    End If

    numFilters = ActiveSheet.AutoFilter.Filters.Count
    Debug.Print "Sheet(" & ActiveSheet.Name & ") has " & numFilters & " filters."

    For i = 1 To numFilters
        If ActiveSheet.AutoFilter.Filters.Item(i).On Then
            crit1 = ActiveSheet.AutoFilter.Filters.Item(i).Criteria1
            If IsArray(crit1) Then
                '--- multiple criteria are selected in this column
                For j = 1 To UBound(crit1)
                    Debug.Print "crit1(" & i & ") is '" & crit1(j) & "'"
                Next j
            Else
                '--- only a single criteria is selected in this column
                Debug.Print "crit1(" & i & ") is '" & crit1 & "'"
            End If
        End If
    Next i
End Sub


Function saveFilters(ByRef rngFiltered As Range) As String
    'Ideas from http://stackoverflow.com/questions/29393006/loop-through-filter-criteria
    If rngFiltered.Parent.AutoFilterMode = False Then
        Exit Sub
    End If
    Dim colCounter As Long  'Generic column counter

End Function

Function restoreFilters(ByRef rngFiltered As Range, _
                        ByVal filtersSaved As String) As Boolean
    'True if filters restored; False if not
    If rngFiltered.Parent.AutoFilterMode = False Then
        Exit Sub
    End If
    Dim colCounter As Long  'Generic column counter
    
    
End Function


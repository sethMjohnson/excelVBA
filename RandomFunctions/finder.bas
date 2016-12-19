Option Explicit

Function findMe(ByVal strSearchTerm As String, _
                Optional ByRef rngToSearch As Range, _
                Optional ByVal findAll As Boolean = False, _
                Optional ByRef rngAfter As Range, _
                Optional ByVal turnOffAutoFilter As Boolean) As Range
    '19 Dec 2016
    'Returns the range that the string was found.
    ' If no range is passed, will search the entire active sheet.
    ' If nothing is found, will return Nothing
    ' If the value we are looking for is in a filtered range _
    '    (header or otherwise), it might not be found.
    
    If rngToSearch Is Nothing Then _
        Set rngToSearch = ActiveSheet.UsedRange
        
    If rngAfter Is Nothing Then _
        Set rngAfter = rngToSearch.Cells(rngToSearch.Cells.Count)
        
    'Make sure rngAfter is in the searcheable range
    If rngAfter.Parent.Name <> rngToSearch.Parent.Name Then
        Set rngAfter = rngToSearch.Cells(rngToSearch.Cells.Count)
    Else
        If Application.Intersect(rngAfter, rngToSearch) Is Nothing Then _
            Set rngAfter = rngToSearch.Cells(rngToSearch.Cells.Count)
    End If
        
    'If the range to search is passed as a range, but doesn't have anything in it, _
    '   an error will occur. So if you try to count rows or columns and there actually _
    '   isn't anything there, it will exit the function.
    On Error Resume Next
    If rngToSearch.Rows.Count = 0 Or rngToSearch.Columns.Count = 0 Then _
        Exit Function
    On Error GoTo 0
        
    If turnOffAutoFilter = True And _
       rngToSearch.Parent.AutoFilterMode = True Then
       rngToSearch.Parent.AutoFilterMode = False
    End If
    
    Set findMe = rngToSearch.Find(What:=strSearchTerm, _
                                    After:=rngAfter, _
                                    LookIn:=xlValues, _
                                    LookAt:=xlWhole, _
                                    SearchOrder:=xlByRows, _
                                    SearchDirection:=xlNext, _
                                    MatchCase:=False)
    'Either we found or we didn't!
    'You'll need to error catch to make sure that it's not Nothing
End Function

Option Explicit

Function findMe(ByVal strSearchTerm As String, _
                Optional ByRef rngToSearch As Range, _
                Optional ByVal turnOffAutoFilter As Boolean) As Range
    '23 Nov 2016
    'Returns the range that the string was found.
    ' If no range is passed, will search the entire active sheet.
    ' If nothing is found, will return Nothing
    ' If the value we are looking for is in a filtered range _
    '    (header or otherwise), it might not be found.
    
    If rngToSearch Is Nothing Then _
        Set rngToSearch = ActiveSheet.UsedRange
    
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
                                    After:=rngToSearch.Cells(rngToSearch.Cells.Count), _
                                    LookIn:=xlValues, _
                                    LookAt:=xlWhole, _
                                    SearchOrder:=xlByRows, _
                                    SearchDirection:=xlNext, _
                                    MatchCase:=False)
    'Either we found or we didn't!
    'You'll need to error catch to make sure that it's not Nothing
End Function

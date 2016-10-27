Option Explicit


Function findMe(ByVal strSearchTerm As String, _
                Optional ByRef rngToSearch As Range, _
                Optional ByVal turnOffAutoFilter As Boolean) As Range
    'Returns the range that the string was found.
    ' If no range is passed, will search the entire active sheet.
    ' If nothing is found, will return Nothing
    ' If the value we are looking for is in a filtered range _
    '    (header or otherwise), it might not be found.
    
    If rngToSearch Is Nothing Then
        Set rngToSearch = ActiveSheet.UsedRange
    End If
    
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

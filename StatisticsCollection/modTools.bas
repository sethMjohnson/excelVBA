Attribute VB_Name = "modTools"
Option Explicit

Function lastRow(Optional columnNumber As Long = 1, _
                 Optional wksht As Worksheet) As Long
    If wksht Is Nothing Then Set wksht = ActiveSheet
    lastRow = wksht.Cells(wksht.Rows.Count, columnNumber).End(xlUp).Row
End Function

Function lastColumn(Optional rowNumber As Long = 1, _
                    Optional wksht As Worksheet)
    If wksht Is Nothing Then Set wksht = ActiveSheet
    lastColumn = wksht.Cells(rowNumber, wksht.Columns.Count).End(xlToLeft).Column
End Function

Function getAllKeyValueDict( _
        ByRef ourDict As Dictionary) As String
'    'See what's in the Dictionary
'    Call getAllKeyValueDict(theDict)
    Dim ourKey As Variant
    
    For Each ourKey In ourDict.Keys
        getAllKeyValueDict = getAllKeyValueDict & vbCrLf & " KEY: " & ourKey & " || ITEM: " & ourDict(ourKey)
    Next ourKey

End Function

Function numCertainCharacterInString( _
        ByVal ourString As String, _
        ByVal findCharacter As String) As Long
    numCertainCharacterInString = Len(ourString) - Len(Replace(ourString, findCharacter, ""))
End Function

Function colNumToLetter(ByVal number As Long) As String
    Dim strArr() As String
    strArr = Split(Cells(1, number).Address, "$")
    If UBound(strArr) > 0 Then _
        colNumToLetter = strArr(1)
End Function

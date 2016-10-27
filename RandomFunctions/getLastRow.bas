Option Explicit

Function lastRow(Optional columnNumber As Long = 1, _
                 Optional wksht As Worksheet) As Long
    If wksht Is Nothing Then Set wksht = ActiveSheet
    lastRow = wksht.Cells(wksht.Rows.Count, columnNumber).End(xlUp).Row
End Function

Function nextBlankRow(ByRef rngLook As Range) As Long
    nextBlankRow = 0
    Do While rngLook.Offset(nextBlankRow, 0).Value <> ""
        nextBlankRow = nextBlankRow + 1
    Loop
    nextBlankRow = rngLook.Row + nextBlankRow
End Function

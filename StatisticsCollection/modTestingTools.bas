Attribute VB_Name = "modTestingTools"
Option Explicit

Sub insertRandomData()

    Dim rangeToInsert As Range
        Set rangeToInsert = Selection
    Dim cell As Range   'Generic Range to loop through
    Dim lowValue As Long
        lowValue = 1
    Dim highValue As Long
        highValue = 200
        
    For Each cell In rangeToInsert
        cell.Value = (highValue - lowValue + 1) * Rnd + lowValue
        cell.Value = Format(cell.Value, "0.00")
    Next cell

End Sub

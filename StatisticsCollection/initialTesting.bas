Attribute VB_Name = "modStuff"
Option Explicit
Const EXCEL_MAX_ROWS = 1048576

'13 Feb 2017

Sub main()

    'Variable Declaration
    Dim statSheet As Worksheet  'Sheet to place data on
    Dim dataSheet As Worksheet  'Sheet working with
    Dim theDict As Scripting.Dictionary 'Dictionary to hold all our IDs
    Dim colWorkingOn As Long    'Current column getting data from
    
    Dim rowCounter As Long  'Generic counter for rows
    Dim colCounter As Long  'Generic counter for columns

    'Set some Variables
    Set dataSheet = ActiveSheet                 'Set the sheet we are working with
    Set statSheet = ActiveWorkbook.Sheets.Add   'Create New Sheet for our analyses
        dataSheet.Activate 'Move back to current sheet
    Set theDict = New Scripting.Dictionary      'Create Dictionary
    
    'Populate the Dictionary
    Call getSampleIDsAndAddresses(theDict, dataSheet)
    
    'Set IDs and Number of them
    Call setIDs(theDict, statSheet)
    Call numSamples(theDict, statSheet)
    
    'Run which process we want
    For colCounter = 2 To lastColumn(, dataSheet)
        Call wkshtFormulaSamples(theDict, statSheet, dataSheet, colCounter, "AVERAGE")
        Call wkshtFormulaSamples(theDict, statSheet, dataSheet, colCounter, "STDEV.S")
    Next colCounter
    
    'AutoFit the stats
    statSheet.Columns.AutoFit
    
End Sub
Sub wkshtFormulaSamples( _
        ByRef ourDict As Dictionary, _
        ByRef shtPlacing As Worksheet, _
        ByRef shtFrom As Worksheet, _
        ByVal columnFrom As Long, _
        ByVal strFormula As String)
    Dim ourKey As Variant   'Loop through all keys
    Dim rowCounter As Long  'Generic row counter
    
    If shtPlacing.Cells(1, lastColumn(, shtPlacing)).Value = "" Then
        shtPlacing.Cells(1, lastColumn(, shtPlacing)).Value = strFormula & ":" & shtFrom.Cells(1, columnFrom).Value
    Else
        shtPlacing.Cells(1, lastColumn(, shtPlacing) + 1).Value = strFormula & ":" & shtFrom.Cells(1, columnFrom).Value
    End If
    rowCounter = 2  'Under header
    
    For Each ourKey In ourDict.Keys
        shtPlacing.Cells(rowCounter, lastColumn(, shtPlacing)).Value = "=" & strFormula & "(" & Replace(ourDict(ourKey), "A", colNumToLetter(columnFrom)) & ")"
        rowCounter = rowCounter + 1
    Next ourKey
End Sub

Sub numSamples(ByRef ourDict As Dictionary, _
                ByRef shtPlacing As Worksheet)
    Dim ourKey As Variant   'Loop through all keys
    Dim rowCounter As Long  'Generic row counter
    
    If shtPlacing.Cells(1, lastColumn(, shtPlacing)).Value = "" Then
        shtPlacing.Cells(1, lastColumn(, shtPlacing)).Value = "Num IDs"
    Else
        shtPlacing.Cells(1, lastColumn(, shtPlacing) + 1).Value = "Num IDs"
    End If
    rowCounter = 2  'Under header
    
    For Each ourKey In ourDict.Keys
        shtPlacing.Cells(rowCounter, lastColumn(, shtPlacing)).Value = numCertainCharacterInString(ourDict(ourKey), ",") + 1
        rowCounter = rowCounter + 1
    Next ourKey
End Sub

Sub setIDs(ByRef ourDict As Dictionary, _
            ByRef shtPlacing As Worksheet)
    Dim ourKey As Variant   'Loop through all keys
    Dim rowCounter As Long  'Generic row counter
    
    If shtPlacing.Cells(1, lastColumn).Value = "" Then
        shtPlacing.Cells(1, lastColumn(, shtPlacing)).Value = "ID"
    Else
        shtPlacing.Cells(1, lastColumn(, shtPlacing) + 1).Value = "ID"
    End If
    rowCounter = 2  'Under header
    
    For Each ourKey In ourDict.Keys
        shtPlacing.Cells(rowCounter, lastColumn(, shtPlacing)).Value = ourKey
        rowCounter = rowCounter + 1
    Next ourKey
End Sub

Sub getSampleIDsAndAddresses( _
        ByRef ourDict As Dictionary, _
        ByRef ourSheet As Worksheet)
    
    ' Chr(39) is the apostrophe "'" character.
    '   It's needed for the sheets
    
    'Variable Declaration
    Dim optionNum As Long   'Sample IDs based on selection. Options:
                            ' 1) If one cell, use whole column, and check for header
                            ' 1) If whole column, use all samples, and check for header
                            ' 2) If >1 cell, then on selection, and check for header
    Dim counter As Long 'Generic counter
    Dim hasHeader As Boolean    'If data has a header or not
    Dim rngForSamples As Range  'The range we'll loop through to get the Keys and Items for our Dictionary
    
    'Set some variables
    If Selection.Rows.Count = 1 Or _
       Selection.Rows.Count = EXCEL_MAX_ROWS Then
        optionNum = 1
        hasHeader = True
    Else
        optionNum = 2
        hasHeader = False
    End If
    
    'Get the range for IDs
    Select Case optionNum
        Case 1
            Set rngForSamples = ourSheet.Range(Cells(1, 1).Address & ":" & _
                                Cells(lastRow(1, ourSheet), 1).Address)
        Case 2
            Set rngForSamples = ourSheet.Range(Cells(Selection.Row, 1).Address & ":" & _
                                Cells(Selection.Row + Selection.Rows.Count - 1, 1).Address)
        Case Else
            'Error
    End Select
    
    'Get all IDs as keys
    If hasHeader = True Then
        For counter = 2 To rngForSamples.Rows.Count
            If ourDict.Exists(rngForSamples.Cells(counter, 1).Value) = True Then
                ourDict.Item(rngForSamples.Cells(counter, 1).Value) = _
                            ourDict.Item(rngForSamples.Cells(counter, 1).Value) & "," & Chr(39) & ourSheet.Name & Chr(39) & "!" & rngForSamples.Cells(counter, 1).Address
            Else
                ourDict.Add rngForSamples.Cells(counter, 1).Value, _
                            Chr(39) & ourSheet.Name & Chr(39) & "!" & rngForSamples.Cells(counter, 1).Address
            End If
        Next counter
        
    Else
        For counter = 1 To rngForSamples.Rows.Count
            If ourDict.Exists(rngForSamples.Cells(counter, 1).Value) = True Then
                ourDict.Item(rngForSamples.Cells(counter, 1).Value) = _
                            ourDict.Item(rngForSamples.Cells(counter, 1).Value) & "," & Chr(39) & ourSheet.Name & Chr(39) & "!" & rngForSamples.Cells(counter, 1).Address
            Else
                ourDict.Add rngForSamples.Cells(counter, 1).Value, _
                            Chr(39) & ourSheet.Name & Chr(39) & "!" & rngForSamples.Cells(counter, 1).Address
            End If
        Next counter
    End If

End Sub

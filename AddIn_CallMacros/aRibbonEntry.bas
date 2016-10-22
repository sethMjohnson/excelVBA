Attribute VB_Name = "aRibbonEntry"
Option Explicit
Sub callMeTheMacro(chosenOption As IRibbonControl)
'Reset our Public Variables, from HNS
    Call setPublicVariablesFromHNS
    
'Initialize for network or not, every time a button is pressed
    Call initializeMe
    
Dim macroWantingToRun As classMacroInfo: Set macroWantingToRun = New classMacroInfo
Dim strPassedTags As String: strPassedTags = chosenOption.Tag
Dim splitUpArray() As String

'Split the string (tag in the customui) up with the delimiter of "|"
'EX: tag="ExcelTools|Change Enter Direction.xlsm|Main"
splitUpArray() = Split(strPassedTags, "|")

'Set each of the individual things for the class
macroWantingToRun.fileLocation = PUBstrMacroPath & splitUpArray(0)
macroWantingToRun.fileName = splitUpArray(1)
macroWantingToRun.moduleName = splitUpArray(2)

'Keep current workbook and open macro, then switch back to current
On Error GoTo errorNoActiveWorkbook
Dim strCurrentWorkbook As String
strCurrentWorkbook = ActiveWorkbook.Name
GoTo errorNoActiveWorkbookEND
errorNoActiveWorkbook:
Workbooks.Add
strCurrentWorkbook = ActiveWorkbook.Name
errorNoActiveWorkbookEND:

If macroWantingToRun.fileName = "folder" Then
'Open up our folder
    Shell "explorer """ & macroWantingToRun.fileLocation, vbNormalFocus

Else
'Open up the Macro
    Workbooks.Open (macroWantingToRun.fileLocation & "\" & macroWantingToRun.fileName)
    If strCurrentWorkbook <> "Book1" Then
        Workbooks(strCurrentWorkbook).Activate
    End If
End If

If macroWantingToRun.moduleName = "noMethod" Then
    Application.DisplayAlerts = True
    Exit Sub
Else
    Application.Run ("'" & macroWantingToRun.fileName & "'!" & macroWantingToRun.moduleName)
    'Workbooks(macroWantingToRun.fileName).Close
    Application.DisplayAlerts = True
End If

End Sub
Function Eval(Ref As String)
    Application.Volatile
    Eval = Evaluate(Ref)
End Function

Function seeIfSheetExistsSample(sheetName As String) As Boolean

    Dim wsSheet As Worksheet
    On Error Resume Next
    Set wsSheet = Sheets(sheetName)
    On Error GoTo 0
    If Not wsSheet Is Nothing Then
        seeIfSheetExists = True
    Else
        seeIfSheetExists = False
    End If

End Function


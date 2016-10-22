Attribute VB_Name = "ribbonTools"
Option Explicit

Sub CallbackGetEnabled(control As IRibbonControl, _
                       ByRef enabled)
'getEnabled="CallbackGetEnabled" tag for whatever you want in the ribbon.

End Sub
Sub CallbackGetLabel(control As IRibbonControl, _
                     ByRef label)
'getLabel tag
    Select Case control.ID
    Case "bttnUpdateAddin"
        If AddInUpdateNeeded = True Then
            label = "Update Available."
        Else
            label = "No Update Needed."
        End If
    
  End Select

End Sub

Function AddInUpdateNeeded() As Boolean
'Variable Declaration
Dim strFromPath As String
Dim strToPath As String
Dim strAddinName As String
Dim dateFromPath As Date
Dim dateToPath As Date

On Error GoTo ErrorCaught 'Won't update if there is an error. Keeps things simple
'Some potential errors: _
    File does not exist on network _
    File directory does not exist on computer

    Call setPublicVariablesFromHNS

    'Variable Sets
    strFromPath = PUBstrNetworkMacroPath & "\Toolbar\customUI.xlam"
    strToPath = CStr(Environ("APPDATA") & "\Microsoft\AddIns\customUI.xlam")
    'TO: "%APPDATA%\Microsoft\AddIns\"
    strAddinName = "customUI"
    dateFromPath = FileDateTime(strFromPath)
    dateToPath = FileDateTime(strToPath)

    'Check if both files were modified the same date
        'http://www.techonthenet.com/excel/formulas/datediff.php
        ' s is for second
    If DateDiff("s", dateFromPath, dateToPath) = 0 Then
        'Modified the same date/ time. No update needed.
        AddInUpdateNeeded = False
    Else
        'They are different.
        AddInUpdateNeeded = True
    End If

ErrorCaught:

On Error GoTo 0

End Function

Sub changeDirectionOfEnter(control As IRibbonControl)

    If Application.MoveAfterReturnDirection = xlToRight Then
        Application.MoveAfterReturnDirection = xlDown
        toastNotification ("Pressing enter will move DOWN.")
    Else
        Application.MoveAfterReturnDirection = xlToRight
        toastNotification ("Pressing enter will move RIGHT.")
    End If
    
End Sub

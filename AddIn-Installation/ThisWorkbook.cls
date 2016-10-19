VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Const addinAutoDisabled = "HSNautoDisabledAddinForUpgrade"
'Hidden Name Space

'When the workbook opens, it will check if this is an addin that's already installed.
'  If it's already installed, the getvisible won't work properly, so we will disable the addin being installed.
'  If we didn't upgrade the addin, we will re-enable the addin we disabled at workbook open.
'  To persist whether we disabled the addin or not, we will use the Hidden Name Space to store a boolean value.

Private Sub Workbook_Open()
    Dim addInName As String
    addInName = Replace(ThisWorkbook.Name, UPGRADE_TEXT, "")
    addInName = Replace(addInName, ".xlam", "")

    'Format (String of Variable Name, Value, Overwrite)
    Call AddHiddenName(addinAutoDisabled, False, True)

    If InStr(ThisWorkbook.Name, UPGRADE_TEXT) > 0 Then
        If Excel.Application.AddIns(addInName).Installed = True Then
            Excel.Application.AddIns(addInName).Installed = False
            Call AddHiddenName(addinAutoDisabled, True, True)
        End If
    End If
End Sub


Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If IsNull(GetHiddenNameValue(addinAutoDisabled)) = False And _
       GetHiddenNameValue(addinAutoDisabled) = True Then
        Dim addInName As String
        addInName = Replace(ThisWorkbook.Name, UPGRADE_TEXT, "")
        addInName = Replace(addInName, ".xlam", "")
        Excel.Application.AddIns(addInName).Installed = True
    End If
End Sub

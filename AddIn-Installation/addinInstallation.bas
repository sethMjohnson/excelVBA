Attribute VB_Name = "addinInstallation"
Option Explicit
Global Const UPGRADE_TEXT = "_upgrade"  'Text at the end of the file name that signifies this is an upgraded addin
Const ADDIN_EXT = ".xlam"               'AddIn Extension text


Sub VisibilityOfAddInInstallation(control As IRibbonControl, ByRef visible)
    If thisAddInInstalled = True Then
        visible = False
    Else
        visible = True
    End If
End Sub
Sub getLabelbttnInstallThisAddIn(control As IRibbonControl, ByRef labelValue)
    If thisAddInUpgrade = True Then
        labelValue = "Upgrade This AddIn"
    Else
        labelValue = "Install This AddIn"
    End If
End Sub

Function thisAddInUpgrade() As Boolean
    If InStr(1, ThisWorkbook.Name, UPGRADE_TEXT) > 0 Then
        thisAddInUpgrade = True
    Else
        thisAddInUpgrade = False
    End If
End Function
Function thisAddInInstalled() As Boolean
    'Checks and returns if this AddIn is installed or not
    If ThisWorkbook.Path & "\" = (Environ$("APPDATA")) & "\Microsoft\AddIns\" Then
        thisAddInInstalled = True
    Else
        thisAddInInstalled = False
    End If
End Function

Sub installThisAddin(buttonClicked As IRibbonControl)
    'Check if this is an AddIn first
    If Right(ThisWorkbook.Name, Len(ADDIN_EXT)) <> ADDIN_EXT Then
        MsgBox "This is not an addin for me to install."
        Exit Sub
    End If
    
    'Ask user if they actually meant to install this
    Dim upgradeVersion As Boolean   'To know the words to upgrade or initial install
        upgradeVersion = thisAddInUpgrade
    If upgradeVersion = True Then
        If MsgBox("Do you wish to upgrade to [" & ThisWorkbook.Name & "]?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    Else
        If MsgBox("Do you wish to install [" & ThisWorkbook.Name & "]?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    'Variable Declaration and Sets
    Dim thisAddInLocation As String         'Location on this computer of this current AddIn
        thisAddInLocation = ThisWorkbook.Path
    Dim officeAddinsFolderLink As String    'Location to install AddIn
        officeAddinsFolderLink = (Environ$("APPDATA")) & "\Microsoft\AddIns\"
    Dim addInName As String                 'AddIn Name to install
        addInName = Replace(ThisWorkbook.Name, ADDIN_EXT, "")
    Dim strFromPath As String               'Full from path to copy
        strFromPath = thisAddInLocation & "\" & addInName & ADDIN_EXT
    If upgradeVersion = True Then
        addInName = Replace(addInName, UPGRADE_TEXT, "")
        Excel.Application.AddIns(addInName).Installed = False
    End If
    Dim strToPath As String                 'Full to path to copy
        strToPath = officeAddinsFolderLink & addInName & ADDIN_EXT
    Dim objExcel As Object                  'Excel Object to install AddIn in the background
    Dim objAddin As Object                  'AddIn Object to install in the background
    Dim fso As Object                       'File Scripting Object to copy over AddIn
        Set fso = CreateObject("Scripting.FileSystemObject")

    If upgradeVersion = True Then
        addInName = Replace(addInName, UPGRADE_TEXT, "")
        Excel.Application.AddIns(addInName).Installed = False
    End If


On Error GoTo ErrorsHappens
    
    'Copy over file (From, To, Overwrite)
     fso.CopyFile strFromPath, strToPath, True

    'Adding the Add-In straight with excel. _
        Create an Excel Object, then an Addin Object, then set it to true. _
        The file was already copied over.
    'Need an excel object or we may not be able to add an addin (automation error)
    Set objExcel = CreateObject("Excel.Application")
    
    objExcel.Workbooks.Add
    If upgradeVersion = True Then
        objExcel.Application.AddIns(addInName).Installed = True
    Else
        Set objAddin = objExcel.AddIns.Add(officeAddinsFolderLink & addInName & ADDIN_EXT, True)
        objAddin.Installed = True
    End If
        
    objExcel.Quit

    'Clean up Objects
    Set objExcel = Nothing
    Set objAddin = Nothing
    Set fso = Nothing

    MsgBox ("Installation Successful. You may need to restart Excel to see changes." & vbCrLf & vbCrLf & _
            "If the AddIn does not show, you may need to enable it in Excel Options: " & vbCrLf & _
            "   File > Options > Add-Ins > Manage: [Excel Add-Ins] [GO...]")
    
On Error GoTo 0
    Exit Sub

ErrorsHappens:
On Error GoTo 0
    MsgBox "Something went wrong with the installation. AddIn not installed."

End Sub

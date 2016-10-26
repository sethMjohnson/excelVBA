Attribute VB_Name = "iniForms"
Option Explicit
Function FormINI(ByRef formUsing As UserForm, _
                 ByVal strINIpath As String, _
                 ByVal actionINI As iniAction)
    Dim ctrl As Control 'Generic control to loop through
    Dim strINIValue As String   'Holding string to make sure all values exist,
                                ' as we wouldn't want to write over a value if there is an error.
    
    For Each ctrl In formUsing.Controls
        If actionINI = iniWrite Then
            'Write out all the things for our form
            Select Case TypeName(ctrl)
                Case "CheckBox"
                    Call ManageINI(actionINI, ctrl.Name, "Name", strINIpath, ctrl.Name)
                    Call ManageINI(actionINI, ctrl.Name, "Caption", strINIpath, ctrl.Caption)
                    Call ManageINI(actionINI, ctrl.Name, "Value", strINIpath, ctrl.Value)
                    Call ManageINI(actionINI, ctrl.Name, "Control Tip Text", strINIpath, ctrl.ControlTipText)
                    
                Case "Label"
                    Call ManageINI(actionINI, ctrl.Name, "Name", strINIpath, ctrl.Name)
                    Call ManageINI(actionINI, ctrl.Name, "Caption", strINIpath, ctrl.Caption)
                    
                Case "TextBox"
                    Call ManageINI(actionINI, ctrl.Name, "Name", strINIpath, ctrl.Name)
                    Call ManageINI(actionINI, ctrl.Name, "Value", strINIpath, ctrl.Value)
                    
                Case "ComboBox"
                
                Case "CommandButton"
                    
                Case Else
                    MsgBox TypeName(ctrl) & " :: Not found in Select."
            End Select
        Else
            'Reading, so read all the things into the controls
            Select Case TypeName(ctrl)
                Case "CheckBox"
                    strINIValue = ManageINI(actionINI, ctrl.Name, "Caption", strINIpath)
                        If strINIValue <> "DNE" Then ctrl.Caption = strINIValue
                    strINIValue = ManageINI(actionINI, ctrl.Name, "Value", strINIpath)
                        If strINIValue <> "DNE" Then ctrl.Value = strINIValue
                    strINIValue = ManageINI(actionINI, ctrl.Name, "Control Tip Text", strINIpath)
                        If strINIValue <> "DNE" Then ctrl.ControlTipText = strINIValue
                    
                Case "Label"
                    strINIValue = ManageINI(actionINI, ctrl.Name, "Caption", strINIpath)
                        If strINIValue <> "DNE" Then ctrl.Caption = strINIValue
                    
                Case "TextBox"
                    strINIValue = ManageINI(actionINI, ctrl.Name, "Value", strINIpath)
                        If strINIValue <> "DNE" Then ctrl.Value = strINIValue
                    
                Case "ComboBox"
                
                Case "CommandButton"
                    
                Case Else
                    MsgBox TypeName(ctrl) & " :: Not found in Select."
            End Select
        End If
    Next ctrl
End Function



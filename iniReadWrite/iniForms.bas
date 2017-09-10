Attribute VB_Name = "iniForms"
Option Explicit

Function FormINI(ByRef formUsing As UserForm, _
                 ByVal actionINI As iniAction, _
                 Optional ByVal strINIpath As String)
    ' Updated 21 Nov 2016
    Dim ctrl As control 'Generic control to loop through
    Dim strINIValue As String   'Holding string to make sure all values exist,
                                ' as we wouldn't want to write over a value if there is an error.
    Dim strINIValueInner As String 'Inner loop inner
    Dim boolAlreadyDeleted As Boolean: boolAlreadyDeleted = False   'If we already deleted then don't delete again
    'Section / Key / Path / Value for Key
    
    'No path specified, use the workbooks path
    If strINIpath = "" Then _
        strINIpath = ThisWorkbook.Path & "\" & ThisWorkbook.Name & ".ini"
    
    'For ListBoxes
    Dim counterIndex As Long 'Generic counter to loop index
    Dim counterCol As Long  'Generic counter to loop through columns
    Dim numCol As Long  'Column number for listbox
    
    For Each ctrl In formUsing.Controls
        If actionINI = iniWrite Then
            If boolAlreadyDeleted = False Then
                'delete ini first
                'http://stackoverflow.com/questions/67835/deleting-a-file-in-vba
                If (Dir(strINIpath) <> "") = True Then 'See above
                  ' First remove readonly attribute, if set
                  SetAttr strINIpath, vbNormal
                  ' Then delete the file
                  Kill strINIpath
                  boolAlreadyDeleted = True
                End If
            End If
            
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
                    Call ManageINI(actionINI, ctrl.Name, "Value", strINIpath, ctrl.Value)
                
                Case "CommandButton"
                
                Case "OptionButton"
                    Call ManageINI(actionINI, ctrl.Name, "Name", strINIpath, ctrl.Name)
                    Call ManageINI(actionINI, ctrl.Name, "Value", strINIpath, ctrl.Value)
                
                Case "Frame"
                
                Case "MultiPage"
                
                Case "ListBox"
                                        
                    'Action / Section / Key / Path / Value for Key
                    
                    For counterIndex = 0 To ctrl.ListCount - 1
                        For counterCol = 0 To ctrl.ColumnCount - 1
                            Call ManageINI(actionINI, _
                                           ctrl.Name, _
                                           counterIndex & "," & counterCol, _
                                           strINIpath, _
                                           ctrl.List(counterIndex, counterCol))
                        Next counterCol
                    Next counterIndex
                    
                Case Else
                    MsgBox TypeName(ctrl) & " :" & ctrl.Name & ": Not found in routine to save the form's options."
            End Select
        Else
            'Reading, so read all the things into the controls
            Select Case TypeName(ctrl)
                Case "CheckBox"
                    strINIValue = ManageINI(actionINI, ctrl.Name, "Caption", strINIpath)
                        If strINIValue <> c_KEY_DOES_NOT_EXIST Then ctrl.Caption = strINIValue
                    strINIValue = ManageINI(actionINI, ctrl.Name, "Value", strINIpath)
                        If strINIValue <> c_KEY_DOES_NOT_EXIST Then ctrl.Value = strINIValue
                    strINIValue = ManageINI(actionINI, ctrl.Name, "Control Tip Text", strINIpath)
                        If strINIValue <> c_KEY_DOES_NOT_EXIST Then ctrl.ControlTipText = strINIValue
                    
                Case "Label"
                    strINIValue = ManageINI(actionINI, ctrl.Name, "Caption", strINIpath)
                        If strINIValue <> c_KEY_DOES_NOT_EXIST Then ctrl.Caption = strINIValue
                    
                Case "TextBox"
                    strINIValue = ManageINI(actionINI, ctrl.Name, "Value", strINIpath)
                        If strINIValue <> c_KEY_DOES_NOT_EXIST Then ctrl.Value = strINIValue
                    
                Case "ComboBox"
                    strINIValue = ManageINI(actionINI, ctrl.Name, "Value", strINIpath)
                        If strINIValue <> c_KEY_DOES_NOT_EXIST Then ctrl.Value = strINIValue
                
                Case "CommandButton"
                
                Case "OptionButton"
                    strINIValue = ManageINI(actionINI, ctrl.Name, "Value", strINIpath)
                        If strINIValue <> c_KEY_DOES_NOT_EXIST Then ctrl.Value = strINIValue
                
                Case "Frame"
                
                Case "MultiPage"
                
                Case "ListBox"
                                                            
                    'Action / Section / Key / Path / Value for Key
                    
                    counterIndex = 0    'Reset list index
                    Do
                        counterCol = 0  'Reset column index
                        strINIValue = ManageINI(actionINI, _
                                                ctrl.Name, _
                                                counterIndex & "," & counterCol, _
                                                strINIpath)
                        If strINIValue <> c_KEY_DOES_NOT_EXIST Then
                            'We have a row, so add one so we can add column data
                            ctrl.AddItem
                        End If
                        Do
                            strINIValueInner = ManageINI(actionINI, _
                                                        ctrl.Name, _
                                                        counterIndex & "," & counterCol, _
                                                        strINIpath)
                            If strINIValueInner <> c_KEY_DOES_NOT_EXIST Then
                                If strINIValueInner <> "" Then
                                    'Add into list
                                    ctrl.List(counterIndex, counterCol) = strINIValueInner
                                End If
                            End If

                            counterCol = counterCol + 1 'Increment Indes of Columns
                            
                        Loop While strINIValueInner <> c_KEY_DOES_NOT_EXIST

                        counterIndex = counterIndex + 1 'Increment Index of List
                    Loop While strINIValue <> c_KEY_DOES_NOT_EXIST
                    
                Case Else
                    MsgBox TypeName(ctrl) & " :" & ctrl.Name & ": Not found in routine to load the form's options."
            End Select
        End If
                
    Next ctrl
End Function

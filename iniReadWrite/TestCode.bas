'The following code can be put in a userform to load on form initialization.
' You could save the form values when the form is terminated, or use a button like below.

Private Sub UserForm_Initialize()
    'Restore Options
    Call FormINI(Me, ThisWorkbook.Path & "\" & ThisWorkbook.Name & ".ini", iniRead)
End Sub

Private Sub bttnSaveOptions_Click()
    'Save Options
    Call FormINI(Me, ThisWorkbook.Path & "\" & ThisWorkbook.Name & ".ini", iniWrite)
End Sub

Sub removeSelectedItemListbox( _
                              ByRef ourList As MSForms.ListBox)
    'Call this routine without parentheses
    ' Ex: removeSelectedItemListbox Me.TheList
    ' OR, just use the Call keyword
    ' Ex: Call removeSelectedItemListbox(Me.TheList)
    Dim counter As Long
    Dim userChoice As Long
    
    Do While counter <= ourList.ListCount - 1
        If ourList.Selected(counter) = True Then
            'This is the one selected
            userChoice = MsgBox("Remove [" & ourList.List(counter, 0) & "] from the list?", _
                                vbYesNoCancel, _
                                "Remove this from List")
            Select Case userChoice
                Case vbYes:
                    'Delete it!
                    ourList.RemoveItem (counter)
                    counter = counter - 1
                    
                Case vbNo:
                    'Do nothing
                    
                Case vbCancel:
                    'Exit out of removing things
                    Exit Do
          
                Case Else
                    'NOTHING
            End Select
        End If
        
        counter = counter + 1
    Loop
    
End Sub

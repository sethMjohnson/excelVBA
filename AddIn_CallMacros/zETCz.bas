Attribute VB_Name = "zETCz"
Option Explicit

Sub openDevTools()
    devTools.Show
End Sub

Sub initializeMe()


    
    'Set some of our Public Variables
    PUBstrUserName = Application.UserName
    PUBstrNetworkMacroPath = ""
    PUBstrLocalMacroPath = "C:\[MACRO-Local]\"
    setLocalOrNetwork 'Find if we are local or networked
    
    If PUBstrUserName Like "*" Then
        'Set the Special Developer Key
        Application.OnKey "^+d", "openDevTools"
    End If
    
    Call setHNSFromPublicVariables
    
End Sub

Attribute VB_Name = "modDirInsideAnother"
Option Explicit

Public Function DirectoryInsideAnother(ByVal directoryOne As String, ByVal directoryTwo As String) As Integer
    'Return +1 if One is Inside of Two
    'Return  0 if they are not inside each other
    'Return -1 if Two is Inside of One; Also if they are the same directory.
    Dim splitDirOne() As String
        splitDirOne = Split(directoryOne, "\")
    Dim splitDirTwo() As String
        splitDirTwo = Split(directoryTwo, "\")
    Dim uBoundOne As Long
        uBoundOne = UBound(splitDirOne)
    Dim uBoundTwo As Long
        uBoundTwo = UBound(splitDirTwo)
    Dim counter As Long
    
'    If directoryOne = directoryTwo Then
'        'They are the same
'        DirectoryInsideAnother = -1    'This will be the same as below
'        Exit Function
'    End If
     
    If uBoundOne > uBoundTwo Then
        For counter = 0 To uBoundTwo
            If splitDirTwo(counter) = splitDirOne(counter) Then
                'Two is currently inside one
                DirectoryInsideAnother = 1
            Else
                'Not part as directories fork
                DirectoryInsideAnother = 0
                Exit For
            End If
        Next counter
    Else
        For counter = 0 To uBoundOne
            If splitDirOne(counter) = splitDirTwo(counter) Then
                'One is currently inside two
                DirectoryInsideAnother = -1
            Else
                'Not part as directories fork
                DirectoryInsideAnother = 0
                Exit For
            End If
        Next counter
    End If
    
End Function

Sub test_FOR_DirectoryInsideAnother()
    Dim strOne As String
    Dim strTwo As String
    
    strOne = "C:\a\b\c"
    strTwo = "C:\a\b"
    Debug.Print DirectoryInsideAnother(strOne, strTwo)
    strOne = "C:\a\b"
    strTwo = "C:\a\b\c"
    Debug.Print DirectoryInsideAnother(strOne, strTwo)
    strOne = "C:\a\b"
    strTwo = "C:\a\c"
    Debug.Print DirectoryInsideAnother(strOne, strTwo)
    strOne = "C:\a\b"
    strTwo = "C:\a\b"
    Debug.Print DirectoryInsideAnother(strOne, strTwo)

End Sub


Attribute VB_Name = "hiddenNameSpace"
Option Explicit
Option Compare Text
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modHiddenNames
' By Chip Pearson, www.cpearson.com , chip@cpearson.com
'
' This module expands on "HiddenNameSpace" as documented by Laurent Longre at
' "www.cpearson.com/excel/hidden.htm". This code is my own, but concept comes from
' Laurent Longre.
' Note that these names persist as long as the application is running, even if the
' workbook that create then names (or any other workbook) is closed.
'
' This module contains the following procedures for working with names in the hidden name
' space in Excel.
'           HiddenNameExists
'               Returns True or False indicating whether then specified name exists.
'           IsValidName
'               Returns True or False indicating whether the specified name is valid.
'           AddHiddenName
'               Adds a hidden name and value. Optionally overwrites the name if it
'               already exists.
'           DeleteHiddenName
'               Deletes a name in the hidden name space. Ignores the condition if the
'               name does not exist. This function does not return a value.
'           GetHiddenNameValue
'               Returns the value of a hidden if that name exists. Returns the value of the
'               name if it exists, or NULL if it does not exist.
'
' To change the value of an existing name, first call DeleteHiddenName to remove the name,
' then call AddHiddenName to add the name with the next value.
'
' There is no way to enumerate the existing names. You must know the name in order to
' access or delete it.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Const C_ILLEGAL_CHARS = " /-:;!@#$%^&*()+=,<>"

Public Function IsValidName(HiddenName As String) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsValidName
' This function returns True if HiddenName is a valid name, i.e., it
' is not an empty string and does not contain any character in the
' C_ILLEGAL_CHARS constant.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim C As String
Dim NameNdx As Long
Dim CharNdx As Long
If Trim(HiddenName) = vbNullString Then
    IsValidName = False
    Exit Function
End If
''''''''''''''''''''''''''''''''''''
' Test each character in HiddenName
' against each character in
' C_ILLEGALCHARS. If a match is
' found, get out and return False.
'''''''''''''''''''''''''''''''''''
For NameNdx = 1 To Len(HiddenName)
    For CharNdx = 1 To Len(C_ILLEGAL_CHARS)
        If StrComp(Mid(HiddenName, NameNdx, 1), Mid(C_ILLEGAL_CHARS, CharNdx, 1), vbBinaryCompare) = 0 Then
            '''''''''''''''''''''''''''''
            ' Once one invalid character
            ' is found, there is no
            ' need to continue. Get out
            ' with a result of False.
            '''''''''''''''''''''''''''''
            IsValidName = False
            Exit Function
        End If
    Next CharNdx
Next NameNdx

''''''''''''''''''''''''''''''
' If we made out of the loop,
' the name is valid.
''''''''''''''''''''''''''''''
IsValidName = True

End Function


Public Function HiddenNameExists(HiddenName As String) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' HiddenNameExists
' This function returns True if the hidden name HiddenName
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim V As Variant
On Error Resume Next

''''''''''''''''''''''''''''''''''''''''
' Ensure the name is valid
''''''''''''''''''''''''''''''''''''''''
If IsValidName(HiddenName) = False Then
    HiddenNameExists = False
    Exit Function
End If


V = Application.ExecuteExcel4Macro(HiddenName)
On Error GoTo 0
If IsError(V) = False Then
    ''''''''''''''''''''''''''''''
    ' No error. Name exists.
    ''''''''''''''''''''''''''''''
    HiddenNameExists = True
Else
    ''''''''''''''''''''''''''''''
    ' Error. Name does not exists.
    ''''''''''''''''''''''''''''''
    HiddenNameExists = False
End If

End Function

Public Function AddHiddenName(HiddenName As String, NameValue As Variant, _
    Optional OverWriteExisting As Boolean = False) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' AddHiddenName
' This adds the hidden name HiddenName with a value NameValue to Excel's
' hidden name space. If OverWriteExisting is omitted or False, the function
' will not overwrite the existing name and will return False if the name
' already exists. If OverWriteExisting is True, the original name is
' deleted and replace with the values passed to this function, and the
' function will return True.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim V As Variant
Dim Res As Variant

''''''''''''''''''''''''''''''''''''''''
' Ensure the name is valid
''''''''''''''''''''''''''''''''''''''''
If IsValidName(HiddenName) = False Then
    AddHiddenName = False
    Exit Function
End If


'''''''''''''''''''''''''''''''''
' If V is an object, an array,
' or a user-defined type, then
' return False.
''''''''''''''''''''''''''''''''
If VarType(V) >= vbArray Then
    AddHiddenName = False
    Exit Function
End If
If (VarType(V) = vbUserDefinedType) Or (VarType(V) = vbObject) Then
    AddHiddenName = False
    Exit Function
End If


'''''''''''''''''''''''''''''''''
' Test to see if the name exists.
'''''''''''''''''''''''''''''''''
On Error Resume Next
V = Application.ExecuteExcel4Macro(HiddenName)
On Error GoTo 0
If IsError(V) = False Then
    '''''''''''''''''''''''''''''
    ' Error. Name Exists. If
    ' OverWriteExisting is False,
    ' exit with False. Otherwise
    ' delete the name.
    '''''''''''''''''''''''''''''
    If OverWriteExisting = False Then
        AddHiddenName = False
        Exit Function
    Else
        DeleteHiddenName HiddenName:=HiddenName
    End If
End If
V = Application.ExecuteExcel4Macro("SET.NAME(" & Chr(34) & HiddenName & Chr(34) & "," & Chr(34) & NameValue & Chr(34) & ")")
If IsError(V) = True Then
    AddHiddenName = False
Else
    AddHiddenName = True
End If

End Function

Public Sub DeleteHiddenName(HiddenName As String)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DeleteHiddenName
' This deletes an name from Excel's hidden name space. It ignores the
' condition that the name does not exist. The procedure does not return
' an result.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
Application.ExecuteExcel4Macro ("SET.NAME(" & Chr(34) & HiddenName & Chr(34) & ")")

End Sub

Public Function GetHiddenNameValue(HiddenName As String) As Variant
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetHiddenNameValue
' This function returns the value of HiddenName. If the name does
' not exist, the function returns NULL. Otherwise, it returns the
' value of HiddenName. Note that the value returned by this function
' is always a string value. You'll have to convert it to another
' data type is desired.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim V As Variant

''''''''''''''''''''''''''''''''''''''''
' Ensure the name is valid
''''''''''''''''''''''''''''''''''''''''
If IsValidName(HiddenName) = False Then
    GetHiddenNameValue = Null
    Exit Function
End If

If HiddenNameExists(HiddenName:=HiddenName) = False Then
    GetHiddenNameValue = Null
    Exit Function
End If
On Error Resume Next
V = Application.ExecuteExcel4Macro(HiddenName)
On Error GoTo 0
If IsError(V) = True Then
    GetHiddenNameValue = Null
    Exit Function
End If

GetHiddenNameValue = V
    

End Function


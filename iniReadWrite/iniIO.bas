Attribute VB_Name = "iniIO"
Option Explicit

Global Const c_KEY_DOES_NOT_EXIST = "DNE: Key Does Not Exist"
'*******************************************************************************
' Declaration for Reading and Writing to an INI file.
'*******************************************************************************

'++++++++++++++++++++++++++++++++++++++++++++++++++++
' API Functions for Reading and Writing to INI File
'++++++++++++++++++++++++++++++++++++++++++++++++++++

#If Win64 Then
   'System Functions, see comment below:
   'All of these are the same as in the "Else" portion, except with "PtrSafe" included,
   '  for 64-bit systems. Below this section are the explanations for what these do.
    ' Declare for reading INI files.
    Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" _
        Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                          ByVal lpKeyName As Any, _
                                          ByVal lpDefault As String, _
                                          ByVal lpReturnedString As String, _
                                          ByVal nSize As Long, _
                                          ByVal lpFileName As String) As Long
                                          
    ' Declare for writing INI files.
    Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" _
        Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                            ByVal lpKeyName As Any, _
                                            ByVal lpString As Any, _
                                            ByVal lpFileName As String) As Long
#Else
    Private Declare Function GetPrivateProfileString Lib "kernel32" _
        Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                          ByVal lpKeyName As Any, _
                                          ByVal lpDefault As String, _
                                          ByVal lpReturnedString As String, _
                                          ByVal nSize As Long, _
                                          ByVal lpFileName As String) As Long
                                          
    ' Declare for writing INI files.
    Private Declare Function WritePrivateProfileString Lib "kernel32" _
        Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                            ByVal lpKeyName As Any, _
                                            ByVal lpString As Any, _
                                            ByVal lpFileName As String) As Long
#End If

'++++++++++++++++++++++++++++++++++++++++++++++++++++
' Enumeration for ManageINI funtion
'++++++++++++++++++++++++++++++++++++++++++++++++++++

Enum iniAction
    iniRead = 1
    iniWrite = 2
End Enum
'*******************************************************************************
' End INI file declaratin Section.
'*******************************************************************************

Function ManageINI(actionINI As iniAction, _
                    iniSection As String, _
                    iniSectionKey As String, _
                    iniFilePath As String, _
                    Optional iniKeyValue As String) As String
'*******************************************************************************
' Description:  This reads an INI file section/key combination and
'               returns the read value as a string.
'
' Author:       Scott Lyerly (Edited by Seth Johnson, for personal readability)
' Contact:      scott_lyerly@tjx.com, or scott.c.lyerly@gmail.com
' Obtained:     https://scottlyerly.wordpress.com/2013/12/03/excel-geeking-using-vba-with-configuration-files/
'
' Notes:        Requires "Private Declare Function GetPrivateProfileString" and
'               "WritePrivateProfileString" to be added in the declarations
'               at the top of the module.
'
' Name:                 Date:           Init:   Modification:
' ManageINI             26-Nov-2013     SCL     Original development
' Function Declares     09-Sep-2017     SMJ     Added to allow 64-bit
'
' Arguments:    actionINI       The action to take in the function, reading or writing to
'                               to the INI file. Uses the enumeration iniAction in the
'                               declarations section.
'               iniSection      The section of the INI file to search
'               iniSectionKey   The key of the INI from which to retrieve a value
'               iniFilePath     The name and directory location of the INI file
'               iniKeyValue     The value to be written to the INI file (if writing - optional)
'
' Returns:      string      The return string is one of three things:
'                           1) The value being sought from the INI file.
'                           2) The value being written to the INI file (should match
'                              the iniKeyValue parameter).
'                           3) The word "Error". This can be changed to whatever makes
'                              the most sense to the programmer using it.
'*******************************************************************************

On Error GoTo Err_ManageSectionEntry

    ' Variable declarations.
    Dim sRetBuf         As String
    Dim iLenBuf         As Integer
    Dim sFileName       As String
    Dim sReturnValue    As String
    Dim lRetVal         As Long
    
    ' Based on the actionINI parameter, take action.
    If actionINI = iniRead Then  ' If reading from the INI file.

        ' Set the return buffer to by 256 spaces. This should be enough to
        ' hold the value being returned from the INI file, but if not,
        ' increase the value.
        sRetBuf = Space(256)

        ' Get the size of the return buffer.
        iLenBuf = Len(sRetBuf)

        ' Read the INI Section/Key value into the return variable.
        sReturnValue = GetPrivateProfileString(iniSection, _
                                               iniSectionKey, _
                                               "", _
                                               sRetBuf, _
                                               iLenBuf, _
                                               iniFilePath)

        ' Trim the excess garbage that comes through with the variable.
        sReturnValue = Trim(Left(sRetBuf, sReturnValue))

        ' If we get a value returned, pass it back as the argument.
        ' Else pass "False".
        If Len(sReturnValue) > 0 Then
            ManageINI = sReturnValue
        Else
            ManageINI = c_KEY_DOES_NOT_EXIST 'Does not exist
        End If
    ElseIf actionINI = iniWrite Then ' If writing to the INI file.

        ' Check to see if a value was passed in the iniKeyValue parameter.
        If Len(iniKeyValue) = 0 Then
            ManageINI = "Error"

        Else
            
            ' Write to the INI file and capture the value returned
            ' in the API function.
            lRetVal = WritePrivateProfileString(iniSection, _
                                               iniSectionKey, _
                                               iniKeyValue, _
                                               iniFilePath)

            ' Check to see if we had an error wrting to the INI file.
            If lRetVal = 0 Then ManageINI = "Error"

        End If
End If
    
Exit_Clean:
    Exit Function
    
Err_ManageSectionEntry:
    MsgBox Err.Number & ": " & Err.Description
    Resume Exit_Clean

End Function


Attribute VB_Name = "zToolBoxz"
Option Explicit
'----------------------------------------------------------------------------
' This ToolBox needs to be referenced to be used within other files:
'   Tools>Refereces>"This Addin Project Name" (ATools)
' Also, need to make sure that a reference to "Microsoft Scripting Runtime" _
'   is used as well.
' Best of luck!
'----------------------------------------------------------------------------
'
' Blank Header as follows, for each different Sub/Function:
'
'----------------------------------------------------------------------------
' About   : What does this Sub/Function do?
' Requires: What other Subs/Functions do I need for this?
' Passed  : FileType - Why pass this variable?
' Optional: FileType - Why pass this variable? (In the order that they are listed)
'           FileType - Why pass this variable?
' Returns : What data type and why does this return this value? (If Function)
'----------------------------------------------------------------------------
'    Private Sub Default()
'    'Set Constants
'
'    'Variable Declaration
'
'    'Variable Initialization
'
'    'Validation
'
'    End Sub

'Public Variables, for ease of access between things
    'These won't stay in memory long it seems, so use the HiddenNameSpace, with these _
     strings as the HiddenNameSpace name
Public PUBstrLogTxt           As String     'Text to pass to printLog()
Public PUBstrLogPath          As String     'Path to save Logs

Public PUBstrUserName         As String     'User Name of this system

Public PUBstrLocalMacroPath   As String     'Local Path where macros are stored
Public PUBstrNetworkMacroPath As String     'Network Path where macros are stored
Public PUBstrMacroPath        As String     'Path we use when invoking macros (gets set to Local or Network)

Public PUBboolNetworkAvailable As Boolean   'True or False if our Network Path is available


'----------------------------------------------------------------------------
' About   : Prints a string to a logfile and debug window (if available)
' Requires: boolProjectModelAccess
'           fsoDir
' Passed  : String - To print to file
' Optional: String - Name of file to print to
'           String - Name of file path to write to file
'----------------------------------------------------------------------------
    Public Sub printLog(ByVal inputString As String, _
                 Optional ByVal filePath As String = "C:\temp\", _
                 Optional ByVal fileName As String = "LOGFILE", _
                 Optional ByVal printHeader As Boolean = True)
    'Set Constants
        Const ForAppending = 8
        
    'Variable Declaration
        Dim fso     As Object 'File System Object
        Dim txtFile As Object 'Text File Object
        
    'Variable Initialization
        Set fso = CreateObject("Scripting.FileSystemObject")
        
    'Validation
        If Right(filePath, 1) <> "\" Then filePath = filePath & "\"
        
    'See if File Exists, and if not, make it
        On Error GoTo ErrorHandler
        If Dir(filePath & fileName & ".txt") <> "" Then
            'Already exists, So Open Txt
            Set txtFile = fso.OpenTextFile(filePath & fileName & ".TXT", ForAppending)
        Else
            'DNE, So make it
            fsoDir (filePath)
            
            'Make the Text File
            Set txtFile = fso.CreateTextFile(filePath & fileName & ".TXT", False)
        End If
        
    'Print to Debug if possible
        If boolProjectModelAccess Then
            If printHeader = True Then
            Debug.Print vbCrLf & _
                        "----------------------------------------"
            End If
            Debug.Print inputString
        End If
        
    'Print to File
        If printHeader = True Then
            txtFile.WriteLine ("----------------------------------------" & vbCrLf & _
                               "Log File : " & Now())
        End If
        txtFile.WriteLine (inputString)
        
    'Close File
        txtFile.Close
            
        On Error GoTo 0
    Exit Sub
       
       
       
ErrorHandler:
    If boolProjectModelAccess = True Then
            Select Case Err.Number
                Case 57
                    PUBstrLogTxt = "Error 57. Cannot create Path : " & filePath
                    Call printLog(PUBstrLogTxt)
                    
                Case 52
                    PUBstrLogTxt = "Error 52. Bad File name or Number."
                    Call printLog(PUBstrLogTxt)
                    
                Case Else
                    PUBstrLogTxt = "ERROR" & vbCrLf & _
                        Err.Number & " : " & Err.Description
                    Call printLog(PUBstrLogTxt)
            End Select
    End If
    
        On Error GoTo 0
    Exit Sub
        
    End Sub



'----------------------------------------------------------------------------
' About   : Check for Project Model Access
' Requires: NONE
' Passed  : NONE
' Optional: NONE
' Returns : Bool - True or False depending on Project Model Access
'----------------------------------------------------------------------------
Public Function boolProjectModelAccess() As Boolean
'Declare Variables
    Dim VBProject As Object ' as VBProject
'Try to set and check if we have access
    On Error Resume Next
        Set VBProject = ActiveWorkbook.VBProject
        If Err.Number <> 0 Then
            boolProjectModelAccess = False 'No access
        Else
            boolProjectModelAccess = True 'Yes access
        End If
    On Error GoTo 0
End Function



'----------------------------------------------------------------------------
' About   : Creates directory if it Does Not Exist (DNE)
' Requires: NONE
' Passed  : String - Name of file path for directory check
' Optional: NONE
' Returns : NONE
'----------------------------------------------------------------------------
Public Sub fsoDir(ByVal filePath As String)
'Variable Declaration
    Dim fso     As Object 'File System Object
    Dim strSplit() As String 'Generic Split String Array
    Dim strMkPath  As String 'Path that we need to make
    Dim counter As Long 'Generic Counter
    Dim isUNC As Boolean 'Check to see if we have a "\\" at the beginning, which would much things up
    
'Variable Initialization
    Set fso = CreateObject("Scripting.FileSystemObject")
    isUNC = False
    
'Validation
    If Right(filePath, 1) <> "\" Then filePath = filePath & "\"
    
'See if File Exists, and if not, make it
    If fso.FolderExists(filePath) = True Then
        'Already exists
    Else
        'DNE, So make it
        strSplit = Split(filePath, "\") 'Split the filePath to check if all directories exist
        If strSplit(0) = "" Then
            isUNC = True
        End If
        If isUNC = True Then
            'Need to do a little workaround
            strMkPath = ""
        Else
            strMkPath = strSplit(0) 'Set the drive letter
        End If
        
        For counter = 1 To UBound(strSplit) - 1 'One less, or we will try to make the final "\" and error
            If isUNC = True And (counter = 1 Or counter = 2 Or counter = 3) Then
                'Don't try to make directories: first  (+ "\") _
                                                second (+ "SERVER\") _
                                                third  (+ "VOLUME\") _
                        with the fourth spitting it out to the file _
                        in the Else portion
                strMkPath = strMkPath & "\" & strSplit(counter)
            Else
                'Build the path we are trying to create for this iteration
                strMkPath = strMkPath & "\" & strSplit(counter)
                If fso.FolderExists(strMkPath) = False Then
                    'This part of the path DNE, So Make it
                    MkDir (strMkPath)
                End If
                'Loop to check the next part of the path
            End If
        Next
        strMkPath = strMkPath & "\" 'Create the inside path part of our folder
    End If
    
End Sub

'----------------------------------------------------------------------------
' About   : Sees if the Network folder exists, and if yes sets network,
'           and if no, will set to Local Drive
' Requires: NONE
' Passed  : NONE
' Optional: NONE
'----------------------------------------------------------------------------
    Public Sub setLocalOrNetwork()
    'Variable Declaration
        Dim fso As Object
        
    'Variable Initialization
        Set fso = CreateObject("Scripting.FileSystemObject")
        
    'Check if Network Folder is available
        If fso.FolderExists(PUBstrNetworkMacroPath) = False Then
            'Network is not available, use Local
            PUBboolNetworkAvailable = False
            PUBstrMacroPath = PUBstrLocalMacroPath
        Else
            'Network is available
            PUBboolNetworkAvailable = True
            PUBstrMacroPath = PUBstrNetworkMacroPath
        End If
        
    End Sub


'----------------------------------------------------------------------------
' About   : Sets the public variables, taking the values from the HiddenNameSpace
' Requires: GetHiddenNameValue
' Passed  : NONE
' Optional: NONE
'----------------------------------------------------------------------------
    Public Sub setPublicVariablesFromHNS()
        
    'Get all our Public variables set, again, with the values stored in HiddenNameSpace
        PUBstrLogTxt = GetHiddenNameValue("PUBstrLogTxt")
        PUBstrLogPath = GetHiddenNameValue("PUBstrLogPath")
    
        PUBstrUserName = GetHiddenNameValue("PUBstrUserName")
        
        PUBstrLocalMacroPath = GetHiddenNameValue("PUBstrLocalMacroPath")
        PUBstrNetworkMacroPath = GetHiddenNameValue("PUBstrNetworkMacroPath")
        PUBstrMacroPath = GetHiddenNameValue("PUBstrMacroPath")

        PUBboolNetworkAvailable = GetHiddenNameValue("PUBboolNetworkAvailable")
                
    End Sub

'----------------------------------------------------------------------------
' About   : Sets the HiddenNameSpaces, from the Public Variables
' Requires: AddHiddenName
' Passed  : NONE
' Optional: NONE
'----------------------------------------------------------------------------
    Public Sub setHNSFromPublicVariables()
        'Make sure the Log Path is up-to-date
        PUBstrLogPath = PUBstrMacroPath & "[Logs]\"
        
        'Setup Me HiddenNameSpaces, with all the Public Variables
        'Format (String of Variable Name, Value of the Variable)
        Call AddHiddenName("PUBstrLogTxt", PUBstrLogTxt, True)
        Call AddHiddenName("PUBstrLogPath", PUBstrLogPath, True)
        
        Call AddHiddenName("PUBstrUserName", PUBstrUserName, True)
        
        Call AddHiddenName("PUBstrLocalMacroPath", PUBstrLocalMacroPath, True)
        Call AddHiddenName("PUBstrNetworkMacroPath", PUBstrNetworkMacroPath, True)
        Call AddHiddenName("PUBstrMacroPath", PUBstrMacroPath, True)
        
        Call AddHiddenName("PUBboolNetworkAvailable", PUBboolNetworkAvailable, True)
                
    End Sub



'----------------------------------------------------------------------------
' About   : Backs up a folder from one location to another
' Requires: NONE
' Passed  : String - Location of the folder to backup
'           String - Location of the folder to save the source to
' Optional: String - Folder(s) to exclude from the backup
'           String - Path to save a log file to
'----------------------------------------------------------------------------
Public Sub roboBackup(ByVal SourcePath As String, _
                      ByVal DestinationPath As String, _
                      Optional ByVal ExcludedFolders As String, _
                      Optional ByVal LogPath As String)
'See if we are on a Mac or Windows. RoboCopy only works on Windows.
    If Application.OperatingSystem Like "Windows*NT*" Then
        'Continue
    Else
        MsgBox "OS needs to be Windows for RoboCopy to work. Cannot Continue."
        Exit Sub
    End If

'Variable Declaration
    Dim WShell As Object    'WScripting shell
    Dim command As String ' Command to send to shell

'    'Removes the last '\' if it exists, From and To Paths
'    If Right(SourcePath, 1) = "\" Then SourcePath = Left(SourcePath, Len(SourcePath) - 1)
'    If Right(DestinationPath, 1) = "\" Then DestinationPath = Left(DestinationPath, Len(DestinationPath) - 1)
    
    'Information about RoboCopy.exe
        'http://ss64.com/nt/robocopy.html
        'http://social.technet.microsoft.com/wiki/contents/articles/1073.robocopy-and-a-few-examples.aspx
        '/COPYALL : Copy ALL file info (equivalent to /COPY:DATSOU).
            '/COPY:copyflag[s] : What to COPY (default is /COPY:DAT) _
                      (copyflags : D=Data, A=Attributes, T=Timestamps _
                       S=Security=NTFS ACLs, O=Owner info, U=aUditing info).
        '/S : Copy Subfolders.
        '/E : Copy Subfolders, including Empty Subfolders
        '/B : Copy files in Backup mode.
            'https://social.technet.microsoft.com/Forums/scriptcenter/en-US/899e3b9c-2576-4160-9c76-de6d0c8c4fc6/question-about-how-the-robocopy-b-switch-works _
             [/B (backup mode) will allow Robocopy to override file and folder permission settings (ACLs).]
        '/MIR : MIRror a directory tree - equivalent to /PURGE plus all subfolders (/E)
        '/XF file [file]... : eXclude Files matching given names/paths/wildcards.
        '/XD dirs [dirs]... : eXclude Directories matching given names/paths.
                'XF and XD can be used in combination  e.g. _
                 ROBOCOPY c:\source d:\dest /XF *.doc *.xls /XD c:\unwanted /S
        '/XO : eXclude Older - if destination file exists and is the same date _
                 or newer than the source - don’t bother to overwrite it.
        '/LOG:file : Output status to LOG file (overwrite existing log).

    Set WShell = CreateObject("WScript.Shell")
    
    'Set Command with Switches. /COPYALL /B /MIR can error out without sufficient rights
    command = "robocopy.exe " & SourcePath & " " & _
                    DestinationPath & _
                    " /E" & _
                    " /XO" & _
                    " /XD " & ExcludedFolders
    
    If LogPath = "" Then
        'Don't Log
        Call WShell.Run(command, , True)
    Else
        'Logging
        Call WShell.Run(command & _
                        " /LOG+:" & LogPath & "RoboBackup.txt", , True)
    End If
        
    Set WShell = Nothing
    
End Sub

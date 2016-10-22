Attribute VB_Name = "importModule"
Option Explicit
'http://www.cpearson.com/excel/vbe.aspx
'http://www.rondebruin.nl/win/s9/win002.htm
'https://support.microsoft.com/en-us/kb/219905

Sub imported()
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent

    Set VBProj = ActiveWorkbook.VBProject
    
    VBProj.VBComponents.Import "C:\temp\importedModule.txt"
    Set VBComp = VBProj.VBComponents("importedModule")

    callImportedFunction
    
    VBProj.VBComponents.Remove VBComp

End Sub

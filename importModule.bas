Attribute VB_Name = "importModule"
Option Explicit

Sub imported()
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent

    Set VBProj = ActiveWorkbook.VBProject
    
    VBProj.VBComponents.Import "C:\temp\importedModule.txt"
    Set VBComp = VBProj.VBComponents("importedModule")

    callImportedFunction
    
    VBProj.VBComponents.Remove VBComp

End Sub

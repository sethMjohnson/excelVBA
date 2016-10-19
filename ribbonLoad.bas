Attribute VB_Name = "ribbonLoad"
Option Explicit
'This stuff will allow you to edit the labels, etc of anything from the ribbon even after it loads
'https://msdn.microsoft.com/en-us/library/aa338202(v=office.12)#OfficeCustomizingRibbonUIforDevelopers_Dynamically

'Global variables
Dim customRibbon As IRibbonUI

Sub customRibbonLoad(custRibbon As IRibbonUI)
    Set customRibbon = custRibbon
End Sub

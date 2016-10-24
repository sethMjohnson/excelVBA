Attribute VB_Name = "modCollections"
Option Explicit

Public Function CollectionContains(ByRef collection As Variant, ByVal key As Variant) As Boolean
'http://stackoverflow.com/questions/137845/determining-whether-an-object-is-a-member-of-a-collection-in-vba
    Dim obj As Variant
    On Error GoTo err
        CollectionContains = True
        obj = collection(key)
        Exit Function
err:
    
        CollectionContains = False
End Function

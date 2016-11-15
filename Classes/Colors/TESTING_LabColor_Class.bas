Attribute VB_Name = "TESTING_LabColor_Class"
Option Explicit

Sub testerColorer()

    Dim mycolor As LabColor
        Set mycolor = New LabColor
    Dim mycolor2 As LabColor
        Set mycolor2 = New LabColor
        
    Dim strDiff As String

    mycolor.LongRGBval = Cells(1, 1).Interior.color
    mycolor2.LongRGBval = Cells(1, 2).Interior.color

    mycolor.setColors
    mycolor2.setColors
    
    Debug.Print mycolor.allColorValues
    Debug.Print mycolor2.allColorValues
    strDiff = "Color Difference : " & Format(mycolor.colorDifferenceLab(mycolor, mycolor2), "#0.000")
    Debug.Print strDiff
    Cells(1, 3).Value = strDiff

End Sub
Sub testerColorer22()

    Dim mycolor As LabColor
        Set mycolor = New LabColor
    Dim mycolor2 As LabColor
        Set mycolor2 = New LabColor
        
    Dim strDiff As String

    Dim selection As Long
        selection = 13
        
    Select Case selection
        Case Is = 1
            'similar colors, but in 76 standard have high difference
            mycolor.LongRGBval = RGB(255, 25, 137)
            mycolor2.LongRGBval = RGB(255, 25, 131)
        Case Is = 2
            'slightly less similar
            mycolor.LongRGBval = RGB(31, 146, 255)
            mycolor2.LongRGBval = RGB(31, 140, 255)
        Case Is = 3
            'good difference
            mycolor.LongRGBval = RGB(146, 146, 31)
            mycolor2.LongRGBval = RGB(131, 131, 31)
        Case Else
            'radically different
            mycolor.LongRGBval = RGB(255, 255, 255)
            mycolor2.LongRGBval = RGB(0, 0, 0)
    End Select

    mycolor.setColors
    mycolor2.setColors
    
'    Debug.Print mycolor.allColorValues
'    Debug.Print mycolor2.allColorValues
    Debug.Print vbCrLf
    strDiff = "Color Difference CIE76: " & Format(mycolor.colorDifferenceLab(mycolor, mycolor2), "#0.000")
    Debug.Print strDiff
    strDiff = "Color Difference CIE94: " & Format(mycolor.colorDifferenceLab(mycolor, mycolor2, False, False), "#0.000")
    Debug.Print strDiff

End Sub
Sub colorTesting()

    Dim mycolor1 As LabColor
    Dim mycolor2 As LabColor
    Dim mycolor3 As LabColor
    Dim mycolor4 As LabColor
    Dim mycolor5 As LabColor
    Dim mycolor6 As LabColor
    Dim response1 As String
    Dim response2 As String
    Dim response3 As String
    Dim response4 As String
    Dim response5 As String
    Dim response6 As String
    
    'Everything is set off of mycolor1.
    ' Need to set off mycolor1, or it's time-consuming to get the XYZ and Lab _
    '   values to set the rest. You'll need more than 4 decimal points out _
    '   to make sure that all the responses match.
        
        Set mycolor1 = New LabColor
    mycolor1.HEX_Color = "#F02B3F"
    response1 = mycolor1.allColorValues
        
        Set mycolor2 = New LabColor
    mycolor2.LongRGBval = mycolor1.LongRGBval
    response2 = mycolor2.allColorValues
    
        Set mycolor3 = New LabColor
    mycolor3.LongRGBval = RGB(mycolor1.RGB_R, mycolor1.RGB_G, mycolor1.RGB_B)
    response3 = mycolor3.allColorValues
    
        Set mycolor4 = New LabColor
    mycolor4.Linear_X = mycolor1.Linear_X
    mycolor4.Linear_Y = mycolor1.Linear_Y
    mycolor4.Linear_Z = mycolor1.Linear_Z
    response4 = mycolor4.allColorValues
    
        Set mycolor5 = New LabColor
    mycolor5.Lab_L = mycolor1.Lab_L
    mycolor5.Lab_A = mycolor1.Lab_A
    mycolor5.Lab_B = mycolor1.Lab_B
    response5 = mycolor5.allColorValues
    
        Set mycolor6 = New LabColor
    response6 = mycolor6.allColorValues
        
    If response1 = response2 And _
       response2 = response3 And _
       response3 = response4 And _
       response4 = response5 Then
        Debug.Print response1
        Debug.Print "All responses matched for 1-5."
    Else
        Debug.Print vbCrLf & "From Hex" & response1
        Debug.Print vbCrLf & "From Long" & response2
        Debug.Print vbCrLf & "From RGB" & response3
        Debug.Print vbCrLf & "From XYZ" & response4
        Debug.Print vbCrLf & "From Lab" & response5
    End If
    
    Debug.Print "Empty color, 6: " & response6

End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LabColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Notes:
'The equation for sRGB to Linear is:
'     Condition           |        Value
' 0<=S<=0.04045           | L = S/12.92
' 0.04045<S<=1            | L = ((S+0.055)/1.055)^2.4
'
'All equations are taken as:
'   Observer. = 2°, Illuminant = D65
'
'Lab color is used to approximate human vision.
'   The closer the values are together, the closer to what it looks like.
'
'The equations used are standard equations for conversions.
'   See the links below for more information.
'
'Links:
'   https://en.wikipedia.org/wiki/Lab_color_space
'       \General knowledge/ equations
'   http://www.easyrgb.com/?X=MATH
'       \Equations
'   http://entropymine.com/imageworsener/srgbformula/
'   https://en.wikipedia.org/wiki/Color_difference
'   https://www.researchgate.net/publication/236023905_Color_difference_Delta_E_-_A_survey
'       \Has a scale for difference
'   http://www.brucelindbloom.com/index.html
'       \Equations
'
'Validation with:
'   http://colormine.org/color-converter
'   http://colormine.org/delta-e-calculator
'   http://colormine.org/delta-e-calculator/cie94
'
'ToDos:
' Write in validation of values for everything
' Work in difference ranges. "Just Noticeable Difference" is ~2.3, but need to find a scale.
' Make the doubles in the differences as decimal? The values don't validate exactly to the validation site.
'   Have to declare as variant, then typecast as cDec()
'
'ERRORS:
'
'Fixed:
' There is an off-by-one error with RGB values.
'   Helped by: http://www.mrexcel.com/forum/members/stephencrump.html
'   Forum Link: http://www.mrexcel.com/forum/excel-questions/975682-conversion-hex-rgb-xyz-lab-off-one-error.html
'   Also inspired to look into the built-in functions so I wouldn't have to use my own for Hex values
' Upper and lower limits are broken? When using Hex code of 00FF00, eg
'   00 was the cause, and it was because it only saw as 0, meaning the string wasn't 7 characters long, so only the
'   default value was remaining.


Const DEFAULT_EMPTY_VAL = -1

    'Class Variables
    'RGB (Red, Blue, Green)
    Private pRed As Long
    Private pGreen As Long
    Private pBlue As Long
    'Long
    Private pRGB_Long As Long
    'XYZ (linear X, Y, Z)
    Private pLinX As Double
    Private pLinY As Double
    Private pLinZ As Double
    'Lab (Light, A color, B color)
    Private pLabLight As Double
    Private pLabColorA As Double
    Private pLabColorB As Double
    Private pLabColorFull As Double
    'Hex
    Private pHexColor As String
    
    'Initilization of values
    Private Sub Class_Initialize()
        'If the values have been set, they'll be positive
        pRed = DEFAULT_EMPTY_VAL
        pGreen = DEFAULT_EMPTY_VAL
        pBlue = DEFAULT_EMPTY_VAL
        
        pRGB_Long = DEFAULT_EMPTY_VAL
        
        pLinX = DEFAULT_EMPTY_VAL
        pLinY = DEFAULT_EMPTY_VAL
        pLinZ = DEFAULT_EMPTY_VAL
        
        pLabLight = DEFAULT_EMPTY_VAL
        pLabColorA = DEFAULT_EMPTY_VAL
        pLabColorB = DEFAULT_EMPTY_VAL
        
        pHexColor = DEFAULT_EMPTY_VAL
    End Sub

    'Class Properties
'RED
    Public Property Get RGB_R() As Long
        RGB_R = pRed
    End Property
    Public Property Let RGB_R(Value As Long)
        If Value >= 0 And Value < 256 Then _
            pRed = Value
    End Property
    
'GREEN
    Public Property Get RGB_G() As Long
        RGB_G = pGreen
    End Property
    Public Property Let RGB_G(Value As Long)
        If Value >= 0 And Value < 256 Then _
            pGreen = Value
    End Property
    
'BLUE
    Public Property Get RGB_B() As Long
        RGB_B = pBlue
    End Property
    Public Property Let RGB_B(Value As Long)
        If Value >= 0 And Value < 256 Then _
            pBlue = Value
    End Property
    
'LONG
    Public Property Get LongRGBval() As Long
        LongRGBval = pRGB_Long
    End Property
    Public Property Let LongRGBval(Value As Long)
        If Abs(Value) <= 16777215 Then
            pRGB_Long = Value
        End If
    End Property
    
'X
    Public Property Get Linear_X() As Double
        Linear_X = pLinX
    End Property
    Public Property Let Linear_X(Value As Double)
        pLinX = Value
    End Property

'Y
    Public Property Get Linear_Y() As Double
        Linear_Y = pLinY
    End Property
    Public Property Let Linear_Y(Value As Double)
        pLinY = Value
    End Property

'Z
    Public Property Get Linear_Z() As Double
        Linear_Z = pLinZ
    End Property
    Public Property Let Linear_Z(Value As Double)
        pLinZ = Value
    End Property

    
'LAB_ Light
    Public Property Get Lab_L() As Double
        Lab_L = pLabLight
    End Property
    Public Property Let Lab_L(Value As Double)
        pLabLight = Value
    End Property


'LAB_ color dimension A
    Public Property Get Lab_A() As Double
        Lab_A = pLabColorA
    End Property
    Public Property Let Lab_A(Value As Double)
        pLabColorA = Value
    End Property


'LAB_ color dimension B
    Public Property Get Lab_B() As Double
        Lab_B = pLabColorB
    End Property
    Public Property Let Lab_B(Value As Double)
        pLabColorB = Value
    End Property
    
'Hex color
    Public Property Get HEX_Color() As String
        HEX_Color = pHexColor
    End Property
    Public Property Let HEX_Color(Value As String)
        If Len(Value) = 7 And Left(Value, 1) = "#" Then
            pHexColor = Value
        End If
    End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'     Sets of the different colors from all the other color spaces          '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub setRGBFromLong()
    'For division here we use \ instead of / because of integer division
    'https://msdn.microsoft.com/en-us/library/b6ex274z.aspx
    RGB_R = LongRGBval Mod 256
    RGB_G = (LongRGBval \ 256) Mod 256
    RGB_B = (LongRGBval \ 256 \ 256) Mod 256
End Sub

Private Sub setLongFromRGB()
    If RGB_R <> DEFAULT_EMPTY_VAL And _
       RGB_G <> DEFAULT_EMPTY_VAL And _
       RGB_B <> DEFAULT_EMPTY_VAL Then
        LongRGBval = RGB(RGB_R, RGB_G, RGB_B)
    End If
End Sub

Private Sub setLABFromRGB()
    'First need to set to XYZ, then to Lab (Hunter) color
    setXYZFromRGB
    setLabFromXYZ
End Sub

Private Sub setXYZFromRGB()
    Dim partR As Double
    Dim partG As Double
    Dim partB As Double
    
    'Conversion to XYZ
    partR = RGB_R / 255 'Get in Nominal Range (0 to 1)
    partG = RGB_G / 255
    partB = RGB_B / 255
    
    If partR > 0.04045 Then
        partR = ((partR + 0.055) / 1.055) ^ 2.4
    Else
        partR = partR / 12.92
    End If
    
    If partG > 0.04045 Then
        partG = ((partG + 0.055) / 1.055) ^ 2.4
    Else
        partG = partG / 12.92
    End If
    
    If partB > 0.04045 Then
        partB = ((partB + 0.055) / 1.055) ^ 2.4
    Else
        partB = partB / 12.92
    End If
    
    partR = partR * 100
    partG = partG * 100
    partB = partB * 100
    
    'Observer. = 2°, Illuminant = D65
    Linear_X = (partR * 0.4124) + (partG * 0.3576) + (partB * 0.1805)
    Linear_Y = (partR * 0.2126) + (partG * 0.7152) + (partB * 0.0722)
    Linear_Z = (partR * 0.0193) + (partG * 0.1192) + (partB * 0.9505)
    
End Sub

Private Sub setLabFromXYZ()
    'Conversion to Hunter Lab
    If Linear_Y = 0 Then
        Lab_L = 0
        Lab_A = 0
        Lab_B = 0
    Else
        Lab_L = 10 * Sqr(Linear_Y)
        Lab_A = 17.5 * (((1.02 * Linear_X) - Linear_Y) / Sqr(Linear_Y))
        Lab_B = 7 * ((Linear_Y - (0.847 * Linear_Z)) / Sqr(Linear_Y))
    End If
End Sub

Private Sub setRGBFromLab()
    'Lab -> XYZ -> RGB
    setXYZFromLab
    setRGBFromXYZ
End Sub

Private Sub setXYZFromLab()
    'Lab -> XYZ
    Dim partX As Double
    Dim partY As Double
    Dim partZ As Double
    
    partY = Lab_L / 10
    partX = Lab_A / 17.5 * Lab_L / 10
    partZ = Lab_B / 7 * Lab_L / 10
    
    Linear_Y = partY ^ 2
    Linear_X = (partX + Linear_Y) / 1.02
    Linear_Z = -(partZ - Linear_Y) / 0.847
End Sub

Private Sub setRGBFromXYZ()
    'XYZ -> RGB
    Dim partX As Double
    Dim partY As Double
    Dim partZ As Double
    
    Dim partR As Double
    Dim partG As Double
    Dim partB As Double
    
    partX = Linear_X / 100
    partY = Linear_Y / 100
    partZ = Linear_Z / 100
    
    partR = (partX * 3.2406) + (partY * -1.5372) + (partZ * -0.4986)
    partG = (partX * -0.9689) + (partY * 1.8758) + (partZ * 0.0415)
    partB = (partX * 0.0557) + (partY * -0.204) + (partZ * 1.057)
    
    If partR > 0.0031308 Then
        partR = 1.055 * (partR ^ (1 / 2.4)) - 0.055
    Else
        partR = 12.92 * partR
    End If
    
    If partG > 0.0031308 Then
        partG = 1.055 * (partG ^ (1 / 2.4)) - 0.055
    Else
        partG = 12.92 * partG
    End If
    
    If partB > 0.0031308 Then
        partB = 1.055 * (partB ^ (1 / 2.4)) - 0.055
    Else
        partB = 12.92 * partB
    End If
    
    RGB_R = partR * 255
    RGB_G = partG * 255
    RGB_B = partB * 255
End Sub

Private Sub setHexFromRGB()
    Dim strHex As String
    'Hex() is a built-in function of excel
    'Format using @ displays a character or a space. Then we'll replace the spaces with 0s.
    strHex = "#" & Format(hex(RGB_R), "@@") & Format(hex(RGB_G), "@@") & Format(hex(RGB_B), "@@")
    strHex = Replace(strHex, " ", "0")
    HEX_Color = strHex
End Sub

Private Sub setRGBFromHex()
    'Anything in VB with &H in it will say it's a Hex value
    RGB_R = CLng("&H" & Mid(HEX_Color, 2, 2))
    RGB_G = CLng("&H" & Mid(HEX_Color, 4, 2))
    RGB_B = CLng("&H" & Mid(HEX_Color, 6, 2))
End Sub

Public Function setColors() As Boolean
    Select Case True
        Case RGB_R <> DEFAULT_EMPTY_VAL And RGB_G <> DEFAULT_EMPTY_VAL And RGB_B <> DEFAULT_EMPTY_VAL
            setLongFromRGB
            setXYZFromRGB
            setLabFromXYZ
            setHexFromRGB
            setColors = True
        Case LongRGBval <> DEFAULT_EMPTY_VAL
            setRGBFromLong
            setXYZFromRGB
            setLabFromXYZ
            setHexFromRGB
            setColors = True
        Case Lab_L <> DEFAULT_EMPTY_VAL And Lab_A <> DEFAULT_EMPTY_VAL And Lab_B <> DEFAULT_EMPTY_VAL
            setXYZFromLab
            setRGBFromXYZ
            setLongFromRGB
            setHexFromRGB
            setColors = True
        Case Linear_X <> DEFAULT_EMPTY_VAL And Linear_Y <> DEFAULT_EMPTY_VAL And Linear_Z <> DEFAULT_EMPTY_VAL
            setRGBFromXYZ
            setLongFromRGB
            setLabFromXYZ
            setHexFromRGB
            setColors = True
        Case CStr(HEX_Color) <> CStr(DEFAULT_EMPTY_VAL)
            setRGBFromHex
            setLongFromRGB
            setXYZFromRGB
            setLabFromXYZ
            setColors = True
        Case Else
            setColors = False
    End Select
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Print table of all the different colors with their spaces           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function allColorValues(Optional ByVal tableStyle As Boolean = True) As String

    If setColors = True Then
        If tableStyle = True Then
            allColorValues = vbCrLf & "________________________________________________________________" & vbCrLf & _
                             "___RGB__________" & vbTab & vbTab & "___XYZ__________" & vbTab & vbTab & "___Lab__________|" & vbCrLf & _
                             "R:" & vbTab & Format(RGB_R, "000") & vbTab & vbTab & vbTab & vbTab & vbTab & "X:" & vbTab & Format(Linear_X, "000.0000") & vbTab & vbTab & vbTab & "L:" & vbTab & Format(Lab_L, "000.0000") & vbTab & "|" & vbCrLf & _
                             "G:" & vbTab & Format(RGB_G, "000") & vbTab & vbTab & vbTab & vbTab & vbTab & "Y:" & vbTab & Format(Linear_Y, "000.0000") & vbTab & vbTab & vbTab & "A:" & vbTab & Format(Lab_A, "000.0000") & vbTab & "|" & vbCrLf & _
                             "B:" & vbTab & Format(RGB_B, "000") & vbTab & vbTab & vbTab & vbTab & vbTab & "Z:" & vbTab & Format(Linear_Z, "000.0000") & vbTab & vbTab & vbTab & "B:" & vbTab & Format(Lab_B, "000.0000") & vbTab & "|" & vbCrLf & _
                             vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "|" & vbCrLf & _
                             "___Long_________" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "|" & vbCrLf & _
                             "Long:" & vbTab & Format(LongRGBval, "00000000") & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "|" & vbCrLf & _
                             vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "|" & vbCrLf & _
                             "___Hexadecimal__" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "|" & vbCrLf & _
                             "Hex:" & vbTab & HEX_Color & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "|" & vbCrLf & _
                             "________________________________________________________________|"

        
        Else
            allColorValues = vbCrLf & "________________________________________________________________" & vbCrLf & _
                             "___RGB___" & vbCrLf & _
                             "Red:" & vbTab & RGB_R & vbCrLf & _
                             "Green:" & vbTab & RGB_G & vbCrLf & _
                             "Blue:" & vbTab & RGB_B & vbCrLf & _
                             vbCrLf & _
                             "___Long___" & vbCrLf & _
                             "Long:" & vbTab & LongRGBval & vbCrLf & _
                             "___Hex___" & vbCrLf & _
                             "Hex Color:" & vbTab & HEX_Color & vbCrLf & _
                             vbCrLf & _
                             "___XYZ___" & vbCrLf & _
                             "Linear X:" & vbTab & Format(Linear_X, "0.0000") & vbCrLf & _
                             "Linear Y:" & vbTab & Format(Linear_Y, "0.0000") & vbCrLf & _
                             "Linear Z:" & vbTab & Format(Linear_Z, "0.0000") & vbCrLf & _
                             vbCrLf & _
                             "___Lab___" & vbCrLf & _
                             "Light:" & vbTab & vbTab & Format(Lab_L, "0.0000") & vbCrLf & _
                             "Color A:" & vbTab & Format(Lab_A, "0.0000") & vbCrLf & _
                             "Color B:" & vbTab & Format(Lab_B, "0.0000")
        End If
    Else
        allColorValues = "Cannot set the colors from the information provided."
    End If

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'     Functions that work on the color spaces and their differences         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function colorDifferenceLab( _
    ByRef colorOne As LabColor, _
    ByRef colorTwo As LabColor, _
    Optional ByVal CIE76 As Boolean = True, _
    Optional ByVal graphics As Boolean = True) As Double
    
    'CIE76 by default
    If CIE76 = True Then
        'Using CIE76 formula
        colorDifferenceLab = Sqr( _
                                ((colorTwo.Lab_L - colorOne.Lab_L) ^ 2) + _
                                ((colorTwo.Lab_A - colorOne.Lab_A) ^ 2) + _
                                ((colorTwo.Lab_B - colorOne.Lab_B) ^ 2) _
                                )
    Else
        'Use the CIE94 formula
        'd for Delta, or change (difference)
        'http://www.brucelindbloom.com/index.html?Eqn_DeltaE_CIE94.html
        
        'Variables
        Dim dL As Double
        Dim C1 As Double
        Dim C2 As Double
        Dim dC As Double
        Dim dH As Double
        Dim da As Double
        Dim db As Double
        Dim SL As Double
        Dim SC As Double
        Dim SH As Double
        
        'Constants
        Dim K1 As Double
        Dim K2 As Double
        Dim kL As Double
        Dim kC As Double: kC = 1#
        Dim kH As Double: kH = 1#
        
        'Set Constants
        If graphics = True Then
            'Use the constants for graphics
            K1 = 0.045
            K2 = 0.015
            kL = 1#
        Else
            'Use the constants for textiles
            K1 = 0.048
            K2 = 0.014
            kL = 2#
        End If
        
        'Set Variables
        dL = colorOne.Lab_L - colorTwo.Lab_L
        C1 = Sqr((colorOne.Lab_A ^ 2) + (colorOne.Lab_B ^ 2))
        C2 = Sqr((colorTwo.Lab_A ^ 2) + (colorTwo.Lab_B ^ 2))
        dC = C1 - C2
        da = colorOne.Lab_A - colorTwo.Lab_A
        db = colorOne.Lab_B - colorTwo.Lab_B
        'dH = Sqr((da ^ 2) + (db ^ 2) - (dC ^ 2)) 'Limited precision might cause this to be slightly negative.
            'To combat this, we will not take the square root, but square it. We will compensate for the rest of equation.
            ' Bruce Lindbloom site shows this.
        dH = (da ^ 2) + (db ^ 2) - (dC ^ 2)
        SL = 1#
        SC = 1# + (K1 * C1)
        SH = 1# + (K2 * C1)
        'Compensating for possible negative dH by using the equation below this one.
'        colorDifferenceLab = _
'            Sqr( _
'                ((dL) / (kL * SL)) ^ 2 + _
'                ((dC) / (kC * SC)) ^ 2 + _
'                ((dH) / (kH * SH)) ^ 2 _
'                )
        colorDifferenceLab = _
            Sqr( _
                ((dL / (kL * SL)) ^ 2) + _
                ((dC / (kC * SC)) ^ 2) + _
                (dH / ((kH * SH) ^ 2)) _
                )
    End If
End Function

'dE is Delta E, or change
'   0 < dE < 1 - observer does not notice the difference,
'   1 < dE < 2 - only experienced observer can notice the difference
'   2 < dE < 3.5 - unexperienced observer also notices the difference,
' 3.5 < dE < 5 - clear difference in color is noticed,
'   5 < dE - observer notices two different colors.
'
'From:
'   https://www.researchgate.net/publication/236023905_Color_difference_Delta_E_-_A_survey

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStopWatch 
   Caption         =   "StopWatch"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "frmStopWatch.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const captionStart = "Start"
Const captionStop = "Stop"
Const captionTimer = "Time elapsed in seconds."

Public startTime As Double
Public endTime As Double

Private Sub bttnStopWatch_Click()
    If bttnStopWatch.Caption = captionStart Then
        startTime = Timer
        bttnStopWatch.Caption = captionStop
    Else
        endTime = Timer
        If IsNumeric(lblSeconds.Caption) Then
            lblSeconds.Caption = lblSeconds.Caption + (endTime - startTime)
        Else
            lblSeconds.Caption = endTime - startTime
        End If
        bttnStopWatch.Caption = captionStart
    End If
End Sub

Private Sub lblSeconds_Click()
    If MsgBox("Reset Stop Watch?", vbYesNo) = vbYes Then
        lblSeconds.Caption = captionTimer
    End If
End Sub

Private Sub UserForm_Initialize()
    bttnStopWatch.Caption = captionStart
    lblSeconds.Caption = captionTimer
End Sub

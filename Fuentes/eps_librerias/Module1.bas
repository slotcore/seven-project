Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Declare Sub InitCommonControls Lib "Comctl32" ()

Sub main()
    InitCommonControls
    Call SetErrorMode(2)
    Form1.Show
End Sub

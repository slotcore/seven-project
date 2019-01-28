VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'****************************************************
'* Title       : Stack                              *
'* Programmer  : Deri Dayandri                      *
'* Date        : (I can't remember when...)         *
'* Description : to hold stack when create postfix  *
'*               form and when calculate the result *
'****************************************************

Private vStack() As Variant
Private vCurrentPos As Integer

Private Const ADDITIONALSIZE = 5


Friend Property Get StackSize() As Integer
  StackSize = vCurrentPos
End Property

Friend Sub Push(Value As Variant)
Dim NewSize As Integer

  vCurrentPos = vCurrentPos + 1
  If vCurrentPos > UBound(vStack) Then
    NewSize = UBound(vStack) + ADDITIONALSIZE
    ReDim Preserve vStack(NewSize)
  End If
  vStack(vCurrentPos - 1) = Value
End Sub

Friend Function Pop() As Variant
  If vCurrentPos > 0 Then
    Pop = vStack(vCurrentPos - 1)
    vCurrentPos = vCurrentPos - 1
  End If
End Function

Friend Function GetPopValue() As Variant
  If vCurrentPos > 0 Then
    GetPopValue = vStack(vCurrentPos - 1)
  End If
End Function

Private Sub Class_Initialize()
  vCurrentPos = 0
  ReDim vStack(0)
End Sub

Private Sub Class_Terminate()
  Erase vStack
End Sub

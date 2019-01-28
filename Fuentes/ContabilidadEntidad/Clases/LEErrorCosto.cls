VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LEErrorCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mLErrorCosto As New Collection
Private mLastIndex As Long

Private Sub Class_Initialize()
    LastIndex = 1
End Sub

Public Property Get LastIndex() As Long
    LastIndex = mLastIndex
End Property
Public Property Let LastIndex(ByVal NewValue As Long)
    mLastIndex = NewValue
End Property

Public Sub Add(mErrorCosto As EErrorCosto)
    mErrorCosto.Index = LastIndex
    mLErrorCosto.Add mErrorCosto
    LastIndex = LastIndex + 1
End Sub

Public Function Count() As Long
    Count = mLErrorCosto.Count
End Function

Public Sub Delete(ByVal Index As Variant)
    Dim mIndice As Long
    mLErrorCosto.Remove Index
    ' Se actualizan los indices de la coleccion
    For mIndice = 1 To mLErrorCosto.Count
        mLErrorCosto(mIndice).Index = mIndice
    Next
End Sub

Public Function Item(ByVal Index As Variant) As EErrorCosto
Attribute Item.VB_UserMemId = 0
    Set Item = mLErrorCosto.Item(Index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mLErrorCosto.[_NewEnum]
End Function

Public Sub Fetch()
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LETransferenciaDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mTransferenciaDetalles As New Collection

Public Sub Add(mTransferenciaDetalle As ETransferenciaDet)
    mTransferenciaDetalles.Add mTransferenciaDetalle
End Sub

Public Function Count() As Long
    Count = mTransferenciaDetalles.Count
End Function

Public Sub Delete(ByVal Index As Variant)
    mTransferenciaDetalles.Remove Index
End Sub

Public Function Item(ByVal Index As Variant) As ETransferenciaDet
Attribute Item.VB_UserMemId = 0
    Set Item = mTransferenciaDetalles.Item(Index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mTransferenciaDetalles.[_NewEnum]
End Function

Public Function Fetch(Opcion As TransferenciaDetalleEnum, _
                        Optional IdTransferencia As String, _
                        Optional Codigo As String, _
                        Optional Descripcion As String, _
                        Optional FechaCreacionRegistroInicio As Date, _
                        Optional FechaCreacionRegistroFin As Date, _
                        Optional FechaUltimaActualizacionInicio As Date, _
                        Optional FechaUltimaActualizacionFin As Date) As ETransferenciaDet
    
End Function







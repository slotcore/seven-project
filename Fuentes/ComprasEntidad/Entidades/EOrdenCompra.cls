VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EOrdenCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mOrdenCompra As Integer
Private mCodigo As String
Private mFechaCreacion As Date
Private mFechaInventario As Date
Private mNumeroSerie As String
Private mNumeroDocumento As String
Private mIdResponsable As Integer
Private mResponsable As String
'Private mMercaderias As EMercaderiaCollection

Public Property Get OrdenCompra() As Integer
    OrdenCompra = mOrdenCompra
End Property
Public Property Let OrdenCompra(ByVal NewValue As Integer)
    mOrdenCompra = NewValue
End Property

Public Property Get Codigo() As String
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal NewValue As String)
    mCodigo = NewValue
End Property

Public Property Get FechaCreacion() As Date
    FechaCreacion = mFechaCreacion
End Property
Public Property Let FechaCreacion(ByVal NewValue As Date)
    mFechaCreacion = NewValue
End Property

Public Property Get FechaInventario() As Date
    FechaInventario = mFechaInventario
End Property
Public Property Let FechaInventario(ByVal NewValue As Date)
    mFechaInventario = NewValue
End Property

Public Property Get NumeroSerie() As String
    NumeroSerie = mNumeroSerie
End Property
Public Property Let NumeroSerie(ByVal NewValue As String)
    mNumeroSerie = NewValue
End Property

Public Property Get NumeroDocumento() As String
    NumeroDocumento = mNumeroDocumento
End Property
Public Property Let NumeroDocumento(ByVal NewValue As String)
    mNumeroDocumento = NewValue
End Property

Public Property Get IdResponsable() As Integer
    IdResponsable = mIdResponsable
End Property
Public Property Let IdResponsable(ByVal NewValue As Integer)
    mIdResponsable = NewValue
End Property

Public Property Get Responsable() As String
    Responsable = mResponsable
End Property
Public Property Let Responsable(ByVal NewValue As String)
    mResponsable = NewValue
End Property

'Public Property Get Mercaderias() As EMercaderiaCollection
'    Mercaderias = mMercaderias
'End Property
'Public Property Set Mercaderias(ByVal NewValue As EMercaderiaCollection)
'    Set mMercaderias = NewValue
'End Property

Public Function Fetch(OrdenCompra As String) As EOrdenCompra

End Function

Public Sub SetEntity(Record As ADODB.Recordset)
'    Dim F As New SistemaLogica.Funciones
'    Dim M As New SistemaEntidad.EMercaderiaCollection
'
'    Record.MoveFirst
'    Me.OrdenCompra = F.NuloString(Record("OrdenCompra"))
'    Me.Codigo = F.NuloString(Record("Codigo"))
'    Me.FechaCreacion = F.ConvertirNumeroAFecha(F.NuloNumeric(Record("FechaCreacion")))
'    Me.FechaInventario = F.ConvertirNumeroAFecha(F.NuloNumeric(Record("FechaInventario")))
'    Me.NumeroSerie = F.NuloString(Record("NumeroSerie"))
'    Me.NumeroDocumento = F.NuloString(Record("NumeroDocumento"))
'    Me.IdResponsable = F.NuloString(Record("IdResponsable"))
'    Me.Responsable = F.NuloString(Record("Responsable"))
'    Me.Mercaderias = M.Fetch(InventarioE, Me.OrdenCompra)
    Set Record = Nothing
End Sub






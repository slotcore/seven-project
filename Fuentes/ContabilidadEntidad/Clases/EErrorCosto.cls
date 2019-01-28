VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EErrorCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigoError As String
Private mOrigen As String
Private mDetalleError As String
Private mSolucionError As String

Private F As New SistemaLogica.Funciones
Private mIndex As Long
Private mConexion As New ADODB.Connection
Private mCalled As Boolean
Private mLoadChild As Boolean

Public Property Get Index() As Long
    Index = mIndex
End Property
Public Property Let Index(ByVal NewValue As Long)
    mIndex = NewValue
End Property

Private Sub Class_Initialize()
    mCalled = False
    mLoadChild = True
End Sub

Private Sub Class_Terminate()
    'CloseConnection
End Sub

Public Property Get Conexion() As ADODB.Connection
    Set Conexion = mConexion
End Property
Public Property Set Conexion(ByVal NewValue As ADODB.Connection)
    Set mConexion = NewValue
End Property

Public Property Get Called() As Boolean
    Called = mCalled
End Property
Public Property Let Called(ByVal NewValue As Boolean)
    mCalled = NewValue
End Property

Public Property Get LoadChild() As Boolean
    LoadChild = mLoadChild
End Property
Public Property Let LoadChild(ByVal NewValue As Boolean)
    mLoadChild = NewValue
End Property

Public Property Get CodigoError() As String
    CodigoError = mCodigoError
End Property
Public Property Let CodigoError(ByVal NewValue As String)
    mCodigoError = NewValue
End Property

Public Property Get Origen() As String
    Origen = mOrigen
End Property
Public Property Let Origen(ByVal NewValue As String)
    mOrigen = NewValue
End Property

Public Property Get DetalleError() As String
    DetalleError = mDetalleError
End Property
Public Property Let DetalleError(ByVal NewValue As String)
    mDetalleError = NewValue
End Property

Public Property Get SolucionError() As String
    SolucionError = mSolucionError
End Property
Public Property Let SolucionError(ByVal NewValue As String)
    mSolucionError = NewValue
End Property

Public Sub Fetch(IdErrorCosto As Long)
End Sub

Public Sub SetEntity(Record As ADODB.Recordset)
End Sub

Public Function Save(IdUsuario As Long, MachineName As String) As Boolean
End Function

Private Sub Insert(IdUsuario As Long, MachineName As String)
End Sub

Private Sub Update(IdUsuario As Long, MachineName As String)
End Sub

Public Sub Delete(IdUsuario As Long, MachineName As String)
End Sub



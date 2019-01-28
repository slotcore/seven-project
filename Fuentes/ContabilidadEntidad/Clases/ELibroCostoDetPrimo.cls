VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ELibroCostoPrimo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdLibroCostoDet As Long
Private mIdTipoItem As Long
Private mTipoItem As String
Private mIdItem As Long
Private mItem As String
Private mCodigoItem As String
Private mCantidad As Double
Private mCostoPromedio As Double

Private mAnhoTrabajo As Long
Private F As New SistemaLogica.Funciones
Private mConexion As New ADODB.Connection
Private mCalled As Boolean

Private Sub Class_Initialize()
    mCalled = False
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

Public Property Get IdLibroCostoDet() As Long
    IdLibroCostoDet = mIdLibroCostoDet
End Property
Public Property Let IdLibroCostoDet(ByVal NewValue As Long)
    mIdLibroCostoDet = NewValue
End Property

Public Property Get IdTipoItem() As Long
    IdTipoItem = mIdTipoItem
End Property
Public Property Let IdTipoItem(ByVal NewValue As Long)
    mIdTipoItem = NewValue
End Property

Public Property Get TipoItem() As String
    TipoItem = mTipoItem
End Property
Public Property Let TipoItem(ByVal NewValue As String)
    mTipoItem = NewValue
End Property

Public Property Get Cantidad() As Double
    Cantidad = mCantidad
End Property
Public Property Let Cantidad(ByVal NewValue As Double)
    mCantidad = NewValue
End Property

Public Property Get CostoPromedio() As Double
    CostoPromedio = mCostoPromedio
End Property
Public Property Let CostoPromedio(ByVal NewValue As Double)
    mCostoPromedio = NewValue
End Property

Public Property Get CodigoItem() As String
    CodigoItem = mCodigoItem
End Property
Public Property Let CodigoItem(ByVal NewValue As String)
    mCodigoItem = NewValue
End Property

Public Property Get IdItem() As Long
    IdItem = mIdItem
End Property
Public Property Let IdItem(ByVal NewValue As Long)
    mIdItem = NewValue
End Property

Public Property Get Item() As String
    Item = mItem
End Property
Public Property Let Item(ByVal NewValue As String)
    mItem = NewValue
End Property

Public Sub Fetch(IdLibroCostoDet As String)

End Sub

Public Sub SetEntity(Record As ADODB.Recordset)
End Sub

Public Sub Save(IdUsuario As Long, MachineName As String)
On Error GoTo BLoqueError
    If (F.NuloNumeric(Me.IdLibroCostoDet) = 0) Then
        Insert IdUsuario, MachineName
    Else
        Update IdUsuario, MachineName
    End If
    Exit Sub
    
BLoqueError:
    Err.Raise &HFFFFFF01, , "Error al Grabar :" + Trim(Err.Description)
End Sub

Private Sub Insert(IdUsuario As Long, MachineName As String)
    Dim xId As Long
    Dim dataBase As New SistemaData.EDataBase
    Dim FData As New SistemaData.Funciones
    Dim F As New SistemaLogica.Funciones
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    xId = FData.GetCodeTable("alm_tomaLibroCosto", "idtomaLibroCosto", mConexion)
    Me.IdLibroCostoDet = xId
                
    dataBase.CommandText = "INSERT INTO alm_tomaLibroCosto(idtomaLibroCosto, idalm, idresponsable, idtipoLibroCosto, " _
                & "idestadoLibroCosto, numser, numdoc, nombre, descripcion, fchinv, fchvig, Item, ano) " _
            + vbCr + "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    
    dataBase.AddParameter "@idtomaLibroCosto", adInteger, adParamInput, xId
    dataBase.AddParameter "@idalm", adInteger, adParamInput, F.NuloNumeric(Me.IdTipoItem)
    dataBase.AddParameter "@numser", adVarChar, adParamInput, F.NuloString(Me.CodigoItem)
    dataBase.AddParameter "@fchinv", adDouble, adParamInput, Me.Cantidad
    dataBase.AddParameter "@fchvig", adDouble, adParamInput, Me.CostoPromedio
    dataBase.AddParameter "@Item", adInteger, adParamInput, F.NuloNumeric(Me.IdItem)
    dataBase.Execute
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "Insert - LibroCosto :" + Trim(Err.Description)
End Sub

Private Sub Update(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    dataBase.CommandText = "UPDouble alm_tomaLibroCosto SET idalm = ?, idresponsable = ?, idtipoLibroCosto = ?, " _
                & "idestadoLibroCosto = ?, numser = ?, numdoc = ?, nombre = ?, descripcion = ?, " _
                & "fchinv = ?, fchvig = ?, Item = ?, ano = ? " _
            + vbCr + "WHERE (alm_tomaLibroCosto.idtomaLibroCosto = ?)"
    
    dataBase.AddParameter "@idalm", adInteger, adParamInput, F.NuloNumeric(Me.IdTipoItem)
    dataBase.AddParameter "@numser", adVarChar, adParamInput, F.NuloString(Me.CodigoItem)
    dataBase.AddParameter "@fchinv", adDouble, adParamInput, Me.Cantidad
    dataBase.AddParameter "@fchvig", adDouble, adParamInput, Me.CostoPromedio
    dataBase.AddParameter "@Item", adInteger, adParamInput, F.NuloNumeric(Me.IdItem)
    dataBase.AddParameter "@idtomaLibroCosto", adInteger, adParamInput, Me.IdLibroCostoDet
    dataBase.Execute
          
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "UpDouble - LibroCosto :" + Trim(Err.Description)
    Exit Sub
End Sub

Public Sub Delete(IdUsuario As Long, MachineName As String)
    Dim xId As Double
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
        
    dataBase.ClearParameter
    dataBase.CommandText = "DELETE FROM alm_tomaLibroCosto " _
            + vbCr + "WHERE alm_tomaLibroCosto.idtomaLibroCosto = ?"
    
    dataBase.AddParameter "@idtomaLibroCosto", adInteger, adParamInput, Me.IdLibroCostoDet
    dataBase.Execute
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el LibroCosto por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub

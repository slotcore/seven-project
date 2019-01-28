VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdAlmacen As Long
Private mCodigo As String
Private mDescripcion As String
Private mGlosa As String
Private mIdMetodoValorizacion As Long
Private mMetodoValorizacion As String
Private mIdTipoAlmacen As Long
Private mTipoAlmacen As String

Private F As New SistemaLogica.Funciones
Private mConexion As New ADODB.Connection
Private mCalled As Boolean
Private mLoadChild As Boolean

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

Public Property Get IdAlmacen() As Long
    IdAlmacen = mIdAlmacen
End Property
Public Property Let IdAlmacen(ByVal NewValue As Long)
    mIdAlmacen = NewValue
End Property

Public Property Get Codigo() As String
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal NewValue As String)
    mCodigo = NewValue
End Property

Public Property Get Descripcion() As String
    Descripcion = mDescripcion
End Property
Public Property Let Descripcion(ByVal NewValue As String)
    mDescripcion = NewValue
End Property

Public Property Get Glosa() As String
    Glosa = mGlosa
End Property
Public Property Let Glosa(ByVal NewValue As String)
    mGlosa = NewValue
End Property

Public Property Get IdMetodoValorizacion() As Long
    IdMetodoValorizacion = mIdMetodoValorizacion
End Property
Public Property Let IdMetodoValorizacion(ByVal NewValue As Long)
    mIdMetodoValorizacion = NewValue
End Property

Public Property Get MetodoValorizacion() As String
    MetodoValorizacion = mMetodoValorizacion
End Property
Public Property Let MetodoValorizacion(ByVal NewValue As String)
    mMetodoValorizacion = NewValue
End Property

Public Property Get IdTipoAlmacen() As Long
    IdTipoAlmacen = mIdTipoAlmacen
End Property
Public Property Let IdTipoAlmacen(ByVal NewValue As Long)
    mIdTipoAlmacen = NewValue
End Property

Public Property Get TipoAlmacen() As String
    TipoAlmacen = mTipoAlmacen
End Property
Public Property Let TipoAlmacen(ByVal NewValue As String)
    mTipoAlmacen = NewValue
End Property

Public Sub Fetch(IdAlmacen As Long)
    Dim dataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    dataBase.CommandText = "SELECT alm_almacenes.id, alm_almacenes.codigo, alm_almacenes.descripcion, alm_almacenes.obs, alm_almacenes.idtipalm, alm_almacenes.idmetval, mae_tipoalmacen.codigo AS codtipalm, mae_tipoalmacen.descripcion AS destipalm, mae_metodoval.abrev AS codmetval, mae_metodoval.descripcion AS desmetval " _
        + vbCr + "FROM (alm_almacenes LEFT JOIN mae_tipoalmacen ON alm_almacenes.idtipalm = mae_tipoalmacen.id) LEFT JOIN mae_metodoval ON alm_almacenes.idmetval = mae_metodoval.id " _
        + vbCr + "WHERE ((alm_almacenes.idMetodoValorizacionajeauto)=" & IdAlmacen & ")"
    
    Set mRecord = dataBase.GetRecordset
    If mRecord.RecordCount > 0 Then Me.SetEntity mRecord
    Exit Sub
    
BLoqueError:
    Err.Raise Err.Number, "[Fetch] " & Err.Source, Err.Description
End Sub

Public Sub SetEntity(Record As ADODB.Recordset)
    Me.IdAlmacen = F.NuloNumeric(Record("id"))
    Me.Codigo = F.NuloString(Record("codigo"))
    Me.Descripcion = F.NuloString(Record("descripcion"))
    Me.Glosa = F.NuloString(Record("obs"))
    Me.IdMetodoValorizacion = F.NuloNumeric(Record("idmetval"))
    Me.MetodoValorizacion = F.NuloString(Record("desmetval"))
    Me.IdTipoAlmacen = F.NuloNumeric(Record("idtipalm"))
    Me.TipoAlmacen = F.NuloString(Record("destipalm"))
End Sub

Public Function Save(IdUsuario As Long, MachineName As String) As Boolean
On Error GoTo BLoqueError
    If (F.NuloNumeric(Me.IdAlmacen) = 0) Then
        Insert IdUsuario, MachineName
    Else
        Update IdUsuario, MachineName
    End If
    Save = True
    Exit Function
    
BLoqueError:
    Save = False
    MsgBox Err.Description, vbInformation + vbOKOnly + vbDefaultButton1, "[Save] " & Err.Source
End Function

Private Sub Insert(IdUsuario As Long, MachineName As String)
    Dim xId As Double
    Dim dataBase As New SistemaData.EDataBase
    Dim FData As New SistemaData.Funciones
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    xId = FData.GetCodeTable("alm_almacenes", "id", mConexion)
    
    Me.IdAlmacen = xId
                
    dataBase.CommandText = "INSERT INTO alm_almacenes(id, codigo, descripcion, obs, idtipalm, idmetval, fechacreacion, horacreacion, usuariocreacion, placacreacion) " _
            + vbCr + "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    
    dataBase.AddParameter "@id", adInteger, adParamInput, xId
    dataBase.AddParameter "@codigo", adVarChar, adParamInput, F.NuloString(Me.Codigo)
    dataBase.AddParameter "@descripcion", adVarChar, adParamInput, F.NuloString(Me.Descripcion)
    dataBase.AddParameter "@obs", adVarChar, adParamInput, F.NuloString(Me.Glosa)
    dataBase.AddParameter "@idtipalm", adInteger, adParamInput, F.NuloNumeric(Me.IdTipoAlmacen)
    dataBase.AddParameter "@idmetval", adInteger, adParamInput, F.NuloNumeric(Me.IdMetodoValorizacion)
    '************************************
    ' Parametros de auditoria
    dataBase.AddParameter "@fechacreacion", adInteger, adParamInput, F.ConvertirFechaANumero(Date)
    dataBase.AddParameter "@horacreacion", adInteger, adParamInput, F.ConvertirHoraANumero(Time)
    dataBase.AddParameter "@usuariocreacion", adVarChar, adParamInput, IdUsuario
    dataBase.AddParameter "@placacreacion", adVarChar, adParamInput, MachineName
    '************************************
    dataBase.Execute
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise Err.Number, "[Insert] " & Err.Source, Err.Description
End Sub

Private Sub Update(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    dataBase.CommandText = "UPDATE alm_almacenes SET codigo = ?, descripcion = ?, obs = ?, idtipalm = ?, idmetval = ?, fechamodificacion = ?, horamodificacion = ?, usuariomodificacion = ?, placamodificacion = ? " _
            + vbCr + "WHERE (alm_almacenes.id = ?)"
    
    dataBase.AddParameter "@codigo", adVarChar, adParamInput, F.NuloString(Me.Codigo)
    dataBase.AddParameter "@descripcion", adVarChar, adParamInput, F.NuloString(Me.Descripcion)
    dataBase.AddParameter "@obs", adVarChar, adParamInput, F.NuloString(Me.Glosa)
    dataBase.AddParameter "@idtipalm", adInteger, adParamInput, F.NuloNumeric(Me.IdTipoAlmacen)
    dataBase.AddParameter "@idmetval", adInteger, adParamInput, F.NuloNumeric(Me.IdMetodoValorizacion)
    '************************************
    ' Parametros de auditoria
    dataBase.AddParameter "@fechamodificacion", adInteger, adParamInput, F.ConvertirFechaANumero(Date)
    dataBase.AddParameter "@horamodificacion", adInteger, adParamInput, F.ConvertirHoraANumero(Time)
    dataBase.AddParameter "@usuariomodificacion", adVarChar, adParamInput, IdUsuario
    dataBase.AddParameter "@placamodificacion", adVarChar, adParamInput, MachineName
    '************************************
    dataBase.AddParameter "@id", adInteger, adParamInput, F.NuloNumeric(Me.IdAlmacen)
    dataBase.Execute
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise Err.Number, "[Update] " & Err.Source, Err.Description
    Exit Sub
End Sub

Public Sub Delete(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    dataBase.ClearParameter
    dataBase.CommandText = "DELETE FROM alm_almacenes " _
            + vbCr + "WHERE alm_almacenes.id = ?"
    dataBase.AddParameter "@id", adInteger, adParamInput, Me.IdAlmacen
    dataBase.Execute
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    MsgBox Err.Description, vbInformation + vbOKOnly + vbDefaultButton1, "[Delete] " & Err.Source
    Exit Sub
End Sub



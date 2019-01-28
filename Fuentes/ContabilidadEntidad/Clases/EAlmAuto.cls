VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EAlmAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdAlmacenajeAutomatico As Long
Private mIdAlmacen As Long
Private mAlmacen As String
Private mIdItem As Long
Private mItem As String
Private mGlosa As String

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

Public Property Get IdAlmacenajeAutomatico() As Long
    IdAlmacenajeAutomatico = mIdAlmacenajeAutomatico
End Property
Public Property Let IdAlmacenajeAutomatico(ByVal NewValue As Long)
    mIdAlmacenajeAutomatico = NewValue
End Property

Public Property Get IdAlmacen() As Long
    IdAlmacen = mIdAlmacen
End Property
Public Property Let IdAlmacen(ByVal NewValue As Long)
    mIdAlmacen = NewValue
End Property

Public Property Get Almacen() As String
    Almacen = mAlmacen
End Property
Public Property Let Almacen(ByVal NewValue As String)
    mAlmacen = NewValue
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

Public Property Get Glosa() As String
    Glosa = mGlosa
End Property
Public Property Let Glosa(ByVal NewValue As String)
    mGlosa = NewValue
End Property

Public Sub Fetch(IdAlmacenajeAutomatico As Long, Optional IdItem As Long)
    Dim cWhere As String
    Dim dataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    cWhere = ""
    If IdItem > 0 Then
        cWhere = "WHERE ((alm_almacenajeauto.iditem)=" & IdItem & ")"
    Else
        cWhere = "WHERE ((alm_almacenajeauto.idalmacenajeauto)=" & IdAlmacenajeAutomatico & ")"
    End If
    dataBase.CommandText = "SELECT alm_almacenajeauto.idalmacenajeauto AS id, alm_almacenajeauto.idalm, alm_almacenes.descripcion AS almacen, alm_almacenajeauto.iditem, alm_inventario.codpro, alm_inventario.descripcion AS item, alm_almacenajeauto.glosa " _
        + vbCr + "FROM (alm_almacenajeauto INNER JOIN alm_almacenes ON alm_almacenajeauto.idalm = alm_almacenes.id) INNER JOIN alm_inventario ON alm_almacenajeauto.iditem = alm_inventario.id " _
        + vbCr + cWhere
    
    Set mRecord = dataBase.GetRecordset
    If mRecord.RecordCount > 0 Then Me.SetEntity mRecord
    Exit Sub
    
BLoqueError:
    Err.Raise Err.Number, "[Fetch] " & Err.Source, Err.Description
End Sub

Public Sub SetEntity(Record As ADODB.Recordset)
    Me.IdAlmacenajeAutomatico = F.NuloNumeric(Record("idalmacenajeauto"))
    Me.IdAlmacen = F.NuloNumeric(Record("idalm"))
    Me.Almacen = F.NuloString(Record("almacen"))
    Me.IdItem = F.NuloNumeric(Record("IdItem"))
    Me.Item = F.NuloString(Record("item"))
    Me.Glosa = F.NuloString(Record("glosa"))
End Sub

Public Function Save(IdUsuario As Long, MachineName As String) As Boolean
On Error GoTo BLoqueError
    If (F.NuloNumeric(Me.IdAlmacenajeAutomatico) = 0) Then
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
    xId = FData.GetCodeTable("alm_almacenajeauto", "idalmacenajeauto", mConexion)
    
    Me.IdAlmacenajeAutomatico = xId
                
    dataBase.CommandText = "INSERT INTO alm_almacenajeauto(idalmacenajeauto, idalm, iditem, glosa, fechacreacion, horacreacion, usuariocreacion, placacreacion) " _
            + vbCr + "VALUES(?, ?, ?, ?, ?, ?, ?, ?)"
    
    dataBase.AddParameter "@idalmacenajeauto", adInteger, adParamInput, xId
    dataBase.AddParameter "@idalm", adInteger, adParamInput, F.NuloNumeric(Me.IdAlmacen)
    dataBase.AddParameter "@iditem", adInteger, adParamInput, F.NuloNumeric(Me.IdItem)
    dataBase.AddParameter "@glosa", adVarChar, adParamInput, F.NuloString(Me.Glosa)
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
    
    dataBase.CommandText = "UPDATE alm_almacenajeauto SET idalm = ?, iditem = ?, glosa = ?, fechamodificacion = ?, horamodificacion = ?, usuariomodificacion = ?, placamodificacion = ? " _
            + vbCr + "WHERE (alm_almacenajeauto.idalmacenajeauto = ?)"
    
    dataBase.AddParameter "@idalm", adInteger, adParamInput, F.NuloNumeric(Me.IdAlmacen)
    dataBase.AddParameter "@iditem", adInteger, adParamInput, F.NuloNumeric(Me.IdItem)
    dataBase.AddParameter "@glosa", adVarChar, adParamInput, F.NuloString(Me.Glosa)
    '************************************
    ' Parametros de auditoria
    dataBase.AddParameter "@fechamodificacion", adInteger, adParamInput, F.ConvertirFechaANumero(Date)
    dataBase.AddParameter "@horamodificacion", adInteger, adParamInput, F.ConvertirHoraANumero(Time)
    dataBase.AddParameter "@usuariomodificacion", adVarChar, adParamInput, IdUsuario
    dataBase.AddParameter "@placamodificacion", adVarChar, adParamInput, MachineName
    '************************************
    dataBase.AddParameter "@idalmacenajeauto", adInteger, adParamInput, F.NuloNumeric(Me.IdAlmacenajeAutomatico)
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
    dataBase.CommandText = "DELETE FROM alm_almacenajeauto " _
            + vbCr + "WHERE alm_almacenajeauto.idalmacenajeauto = ?"
    dataBase.AddParameter "@idalmacenajeauto", adInteger, adParamInput, Me.IdAlmacenajeAutomatico
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

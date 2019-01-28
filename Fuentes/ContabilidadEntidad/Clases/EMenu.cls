VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdMenu As Long
Private mIdModulo As Long
Private mModulo As String
Private mCodigoModulo As String
Private mIdCategoria As Long
Private mCodigoCategoria As String
Private mCategoria As String
Private mCodigoOrden As String
Private mDescripcion As String
Private mNombreControl As String

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

Public Property Get IdMenu() As Long
    IdMenu = mIdMenu
End Property
Public Property Let IdMenu(ByVal NewValue As Long)
    mIdMenu = NewValue
End Property

Public Property Get IdModulo() As Long
    IdModulo = mIdModulo
End Property
Public Property Let IdModulo(ByVal NewValue As Long)
    mIdModulo = NewValue
End Property

Public Property Get Modulo() As String
    Modulo = mModulo
End Property
Public Property Let Modulo(ByVal NewValue As String)
    mModulo = NewValue
End Property

Public Property Get CodigoModulo() As String
    CodigoModulo = mCodigoModulo
End Property
Public Property Let CodigoModulo(ByVal NewValue As String)
    mCodigoModulo = NewValue
End Property

Public Property Get IdCategoria() As Long
    IdCategoria = mIdCategoria
End Property
Public Property Let IdCategoria(ByVal NewValue As Long)
    mIdCategoria = NewValue
End Property

Public Property Get Categoria() As String
    Categoria = mCategoria
End Property
Public Property Let Categoria(ByVal NewValue As String)
    mCategoria = NewValue
End Property

Public Property Get CodigoCategoria() As String
    CodigoCategoria = mCodigoCategoria
End Property
Public Property Let CodigoCategoria(ByVal NewValue As String)
    mCodigoCategoria = NewValue
End Property

Public Property Get CodigoOrden() As String
    CodigoOrden = mCodigoOrden
End Property
Public Property Let CodigoOrden(ByVal NewValue As String)
    mCodigoOrden = NewValue
End Property

Public Property Get Descripcion() As String
    Descripcion = mDescripcion
End Property
Public Property Let Descripcion(ByVal NewValue As String)
    mDescripcion = NewValue
End Property

Public Property Get NombreControl() As String
    NombreControl = mNombreControl
End Property
Public Property Let NombreControl(ByVal NewValue As String)
    mNombreControl = NewValue
End Property

Public Sub Fetch(IdMenu As Long)
    Dim dataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    
    dataBase.CommandText = "SELECT mae_menu.id, mae_menu.idmodulo, mae_modulo.codigo AS codmodulo, mae_modulo.descripcion AS modulo, mae_menu.categoria, mae_menucategoria.codigo As codcategoria, mae_menucategoria.descripcion AS categoria, mae_menu.codord, mae_menu.descripcion, mae_menu.nomcon " _
        + vbCr + "FROM (mae_menu INNER JOIN mae_modulo ON mae_menu.idmodulo = mae_modulo.idmodulo) INNER JOIN mae_menucategoria ON mae_menu.categoria = mae_menucategoria.idmenucategoria " _
        + vbCr + "WHERE ((mae_menu.id)=" & IdMenu & ")"
    
    Set mRecord = dataBase.GetRecordset
    If mRecord.RecordCount > 0 Then Me.SetEntity mRecord
    Exit Sub
    
BLoqueError:
    Err.Raise Err.Number, "[Fetch] " & Err.Source, Err.Description
End Sub

Public Sub SetEntity(Record As ADODB.Recordset)
    Me.IdMenu = F.NuloNumeric(Record("id"))
    Me.IdModulo = F.NuloNumeric(Record("idmodulo"))
    Me.CodigoModulo = F.NuloString(Record("codmodulo"))
    Me.Modulo = F.NuloString(Record("modulo"))
    Me.IdCategoria = F.NuloNumeric(Record("idcategoria"))
    Me.CodigoCategoria = F.NuloString(Record("codcategoria"))
    Me.Categoria = F.NuloString(Record("categoria"))
    Me.CodigoOrden = F.NuloString(Record("codord"))
    Me.Descripcion = F.NuloString(Record("descripcion"))
    Me.NombreControl = F.NuloString(Record("nomcon"))
End Sub

Public Function Save(IdUsuario As Long, MachineName As String) As Boolean
On Error GoTo BLoqueError
    If (F.NuloNumeric(Me.IdMenu) = 0) Then
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
    xId = FData.GetCodeTable("mae_menu", "id", mConexion)
    
    Me.IdMenu = xId
                
    dataBase.CommandText = "INSERT INTO mae_menu(id, idmodulo, categoria, codord, descripcion, nomcon, fechacreacion, horacreacion, usuariocreacion, placacreacion) " _
            + vbCr + "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    
    dataBase.AddParameter "@id", adInteger, adParamInput, xId
    dataBase.AddParameter "@idmodulo", adInteger, adParamInput, F.NuloNumeric(Me.IdModulo)
    dataBase.AddParameter "@categoria", adInteger, adParamInput, F.NuloNumeric(Me.IdCategoria)
    dataBase.AddParameter "@codord", adVarChar, adParamInput, F.NuloString(Me.CodigoOrden)
    dataBase.AddParameter "@descripcion", adVarChar, adParamInput, F.NuloString(Me.Descripcion)
    dataBase.AddParameter "@nomcon", adVarChar, adParamInput, F.NuloString(Me.NombreControl)
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
    
    dataBase.CommandText = "UPDATE mae_menu SET idmodulo = ?, categoria = ?, codord = ?, descripcion = ?, nomcon = ?, fechamodificacion = ?, horamodificacion = ?, usuariomodificacion = ?, placamodificacion = ? " _
            + vbCr + "WHERE (mae_menu.id = ?)"
    
    dataBase.AddParameter "@idmodulo", adInteger, adParamInput, F.NuloNumeric(Me.IdModulo)
    dataBase.AddParameter "@categoria", adInteger, adParamInput, F.NuloNumeric(Me.IdCategoria)
    dataBase.AddParameter "@codord", adVarChar, adParamInput, F.NuloString(Me.CodigoOrden)
    dataBase.AddParameter "@descripcion", adVarChar, adParamInput, F.NuloString(Me.Descripcion)
    dataBase.AddParameter "@nomcon", adVarChar, adParamInput, F.NuloString(Me.NombreControl)
    '************************************
    ' Parametros de auditoria
    dataBase.AddParameter "@fechamodificacion", adInteger, adParamInput, F.ConvertirFechaANumero(Date)
    dataBase.AddParameter "@horamodificacion", adInteger, adParamInput, F.ConvertirHoraANumero(Time)
    dataBase.AddParameter "@usuariomodificacion", adVarChar, adParamInput, IdUsuario
    dataBase.AddParameter "@placamodificacion", adVarChar, adParamInput, MachineName
    '************************************
    dataBase.AddParameter "@id", adInteger, adParamInput, F.NuloNumeric(Me.IdMenu)
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
    dataBase.CommandText = "DELETE FROM mae_menu " _
            + vbCr + "WHERE mae_menu.id = ?"
    dataBase.AddParameter "@id", adInteger, adParamInput, Me.IdMenu
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

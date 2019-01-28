VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ELibroCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdLibroCosto As Long
Private mDescripcion As String
Private mIdMes As Long
Private mMes As String
Private mIdMetodoVal As Long
Private mCodigoMetodoVal As String
Private mMetodoVal As String
Private mIdAplicaDistribucion As Long
Private mIdTipoDistribucion As Long
Private mIdCampoDistribucion As Long
Private mCostoMOD As Double
Private mCostoCIF As Double
Private mIdAlmacen As Long
Private mCodigoAlmacen As String
Private mAlmacen As String
Private mLLibroCostoDet As New LELibroCostoDet

Private F As New SistemaLogica.Funciones
Private mIndex As Long
Private mConexion As New ADODB.Connection
Private mCalled As Boolean
Private mIsRecursive As Boolean
Private mLoadChild As Boolean

Public Property Get Index() As Long
    Index = mIndex
End Property
Public Property Let Index(ByVal NewValue As Long)
    mIndex = NewValue
End Property

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

Public Property Get IsRecursive() As Boolean
    IsRecursive = mIsRecursive
End Property
Public Property Let IsRecursive(ByVal NewValue As Boolean)
    mIsRecursive = NewValue
End Property

Public Property Get LoadChild() As Boolean
    LoadChild = mLoadChild
End Property
Public Property Let LoadChild(ByVal NewValue As Boolean)
    mLoadChild = NewValue
End Property

Public Property Get IdLibroCosto() As Long
    IdLibroCosto = mIdLibroCosto
End Property
Public Property Let IdLibroCosto(ByVal NewValue As Long)
    mIdLibroCosto = NewValue
End Property

Public Property Get Descripcion() As String
    Descripcion = mDescripcion
End Property
Public Property Let Descripcion(ByVal NewValue As String)
    mDescripcion = NewValue
End Property

Public Property Get IdMes() As Long
    IdMes = mIdMes
End Property
Public Property Let IdMes(ByVal NewValue As Long)
    mIdMes = NewValue
End Property

Public Property Get Mes() As String
    Mes = mMes
End Property
Public Property Let Mes(ByVal NewValue As String)
    mMes = NewValue
End Property

Public Property Get IdMetodoVal() As Long
    IdMetodoVal = mIdMetodoVal
End Property
Public Property Let IdMetodoVal(ByVal NewValue As Long)
    mIdMetodoVal = NewValue
End Property

Public Property Get CodigoMetodoVal() As String
    CodigoMetodoVal = mCodigoMetodoVal
End Property
Public Property Let CodigoMetodoVal(ByVal NewValue As String)
    mCodigoMetodoVal = NewValue
End Property

Public Property Get MetodoVal() As String
    MetodoVal = mMetodoVal
End Property
Public Property Let MetodoVal(ByVal NewValue As String)
    mMetodoVal = NewValue
End Property

Public Property Get IdAplicaDistribucion() As Long
    IdAplicaDistribucion = mIdAplicaDistribucion
End Property
Public Property Let IdAplicaDistribucion(ByVal NewValue As Long)
    mIdAplicaDistribucion = NewValue
End Property

Public Property Get IdTipoDistribucion() As Long
    IdTipoDistribucion = mIdTipoDistribucion
End Property
Public Property Let IdTipoDistribucion(ByVal NewValue As Long)
    mIdTipoDistribucion = NewValue
End Property

Public Property Get IdCampoDistribucion() As Long
    IdCampoDistribucion = mIdCampoDistribucion
End Property
Public Property Let IdCampoDistribucion(ByVal NewValue As Long)
    mIdCampoDistribucion = NewValue
End Property

Public Property Get CostoMOD() As Double
    CostoMOD = mCostoMOD
End Property
Public Property Let CostoMOD(ByVal NewValue As Double)
    mCostoMOD = NewValue
End Property

Public Property Get CostoCIF() As Double
    CostoCIF = mCostoCIF
End Property
Public Property Let CostoCIF(ByVal NewValue As Double)
    mCostoCIF = NewValue
End Property

Public Property Get IdAlmacen() As Long
    IdAlmacen = mIdAlmacen
End Property
Public Property Let IdAlmacen(ByVal NewValue As Long)
    mIdAlmacen = NewValue
End Property

Public Property Get CodigoAlmacen() As String
    CodigoAlmacen = mCodigoAlmacen
End Property
Public Property Let CodigoAlmacen(ByVal NewValue As String)
    mCodigoAlmacen = NewValue
End Property

Public Property Get Almacen() As String
    Almacen = mAlmacen
End Property
Public Property Let Almacen(ByVal NewValue As String)
    mAlmacen = NewValue
End Property

Public Property Get LLibroCostoDet() As LELibroCostoDet
    Set LLibroCostoDet = mLLibroCostoDet
End Property
Public Property Set LLibroCostoDet(ByVal NewValue As LELibroCostoDet)
    Set mLLibroCostoDet = NewValue
End Property

Private Sub Class_Initialize()
    mCalled = False
    mIsRecursive = False
    mIndex = 0
End Sub

Public Sub Fetch(IdLibroCosto As Long)
    Dim mDataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BLoqueError
    If Not mConexion Is Nothing Then Set mDataBase.Connection = mConexion
    mDataBase.CommandText = "SELECT con_librocosto.id, con_librocosto.descripcion, con_librocosto.idmes, con_librocosto.idmetodoval, mae_configval.abrev AS codmetval, mae_configval.descripcion AS desmetval, con_librocosto.idalm, alm_almacenes.codigo AS codalm, alm_almacenes.descripcion AS desalm, con_librocosto.costomod, con_librocosto.costocif " _
            + vbCr + "FROM (con_librocosto INNER JOIN mae_configval ON con_librocosto.idmetodoval = mae_configval.id) INNER JOIN alm_almacenes ON con_librocosto.idalm = alm_almacenes.id " _
            + vbCr + "WHERE (((con_librocosto.id)=" & IdLibroCosto & "))"
    Set mRecord = mDataBase.GetRecordset
    If mRecord.RecordCount > 0 Then Me.SetEntity mRecord
    Exit Sub
    
BLoqueError:
    Err.Raise Err.Number, "[Fetch] " & Err.Source, Err.Description
End Sub

Public Sub SetEntity(Record As ADODB.Recordset)
On Error GoTo BLoqueError
    Me.IdLibroCosto = F.NuloNumeric(Record("id"))
    Me.Descripcion = F.NuloString(Record("descripcion"))
    Me.IdMes = F.NuloNumeric(Record("idmes"))
    Me.IdMetodoVal = F.NuloNumeric(Record("idmetodoval"))
    Me.CodigoMetodoVal = F.NuloString(Record("codmetval"))
    Me.MetodoVal = F.NuloString(Record("desmetval"))
    Me.IdAlmacen = F.NuloNumeric(Record("idalm"))
    Me.CodigoAlmacen = F.NuloString(Record("codalm"))
    Me.Almacen = F.NuloString(Record("desalm"))
    Me.CostoMOD = F.NuloNumeric(Record("costomod"))
    Me.CostoCIF = F.NuloNumeric(Record("costocif"))
    If mLoadChild Then
        Set Me.LLibroCostoDet.Conexion = mConexion
        Me.LLibroCostoDet.Fetch Me.IdLibroCosto
    End If
    Exit Sub
    
BLoqueError:
    Err.Raise Err.Number, "[SetEntity] " & Err.Source, Err.Description
End Sub

Public Function Save(IdUsuario As Long, MachineName As String) As Boolean
On Error GoTo BLoqueError
    If (F.NuloNumeric(Me.IdLibroCosto) = 0) Then
        Insert IdUsuario, MachineName
    Else
        Update IdUsuario, MachineName
    End If
    Save = True
    Exit Function
    
BLoqueError:
    Save = False
    F.MostrarMensajeError Err.Description, "Grabar", "[Save] " & Err.Source
End Function

Private Sub Insert(IdUsuario As Long, MachineName As String)
    Dim xId As Double
    Dim dataBase As New SistemaData.EDataBase
    Dim FData As New SistemaData.Funciones
    Dim F As New SistemaLogica.Funciones
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    xId = FData.GetCodeTable("con_librocosto", "id", mConexion)
    Me.IdLibroCosto = xId
    dataBase.CommandText = "INSERT INTO con_librocosto(id, descripcion, idmes, idmetodoval, " _
                & "idalm, costomod, costocif) " _
            + vbCr + "VALUES(?, ?, ?, ?, ?, ?, ?)"
    
    dataBase.AddParameter "@id", adInteger, adParamInput, xId
    dataBase.AddParameter "@descripcion", adVarChar, adParamInput, F.NuloString(Me.Descripcion)
    dataBase.AddParameter "@idmes", adInteger, adParamInput, F.NuloNumeric(Me.IdMes)
    dataBase.AddParameter "@idmetodoval", adInteger, adParamInput, F.NuloNumeric(Me.IdMetodoVal)
    dataBase.AddParameter "@idalm", adInteger, adParamInput, F.NuloNumeric(Me.IdAlmacen)
    dataBase.AddParameter "@costomod", adDouble, adParamInput, F.NuloNumeric(Me.CostoMOD)
    dataBase.AddParameter "@costocif", adDouble, adParamInput, F.NuloNumeric(Me.CostoCIF)
    dataBase.Execute
    
    Dim CostoDet As ELibroCostoDet
    Dim IdCostoDet As Long
    IdCostoDet = F.NuloNumeric(FData.GetCodeTable("con_librocostodet", "id", mConexion))
    For Each CostoDet In Me.LLibroCostoDet
        If Not mConexion Is Nothing Then
            Set CostoDet.Conexion = mConexion
            CostoDet.Called = True
        End If
        CostoDet.MarkNew = True
        CostoDet.IdLibroCostoDet = IdCostoDet
        CostoDet.IdLibroCosto = xId
        If Not CostoDet.Save(IdUsuario, MachineName) Then Err.Raise &HFFFFFF01, , "Error al intentar grabar el LibroCosto"
        IdCostoDet = IdCostoDet + 1
    Next
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , Err.Description
End Sub

Private Sub Update(IdUsuario As Long, MachineName As String)
    Dim xId As Double
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    xId = Me.IdLibroCosto
    
    dataBase.CommandText = "UPDATE con_librocosto SET descripcion = ?, idmes = ?, idmetodoval = ?, " _
                & "idalm = ?, costomod = ?, costocif = ? " _
            + vbCr + "WHERE (con_librocosto.id = ?)"
        
    dataBase.AddParameter "@descripcion", adVarChar, adParamInput, F.NuloString(Me.Descripcion)
    dataBase.AddParameter "@idmes", adInteger, adParamInput, F.NuloNumeric(Me.IdMes)
    dataBase.AddParameter "@idmetodoval", adInteger, adParamInput, F.NuloNumeric(Me.IdMetodoVal)
    dataBase.AddParameter "@idalm", adInteger, adParamInput, F.NuloNumeric(Me.IdAlmacen)
    dataBase.AddParameter "@costomod", adDouble, adParamInput, F.NuloNumeric(Me.CostoMOD)
    dataBase.AddParameter "@costocif", adDouble, adParamInput, F.NuloNumeric(Me.CostoCIF)
    dataBase.AddParameter "@id", adInteger, adParamInput, xId
    dataBase.Execute
    
    ' Borramos Hijos
    DeleteChild IdUsuario, MachineName
    
    ' Grabamos Hijos
    Dim CostoDet As ELibroCostoDet
    For Each CostoDet In Me.LLibroCostoDet
        If Not mConexion Is Nothing Then
            Set CostoDet.Conexion = mConexion
        End If
        CostoDet.Called = True
        CostoDet.IdLibroCosto = Me.IdLibroCosto
        CostoDet.Save IdUsuario, MachineName
    Next
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el LibroCosto por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub

Public Sub Delete(IdUsuario As Long, MachineName As String)
    Dim xId As Double
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    ' Eliminamos a los Hijos
    DeleteChild IdUsuario, MachineName
    
    dataBase.ClearParameter
    dataBase.CommandText = "DELETE FROM con_librocosto " _
            + vbCr + "WHERE con_librocosto.id = ?"
    
    dataBase.AddParameter "@id", adInteger, adParamInput, Me.IdLibroCosto
    dataBase.Execute
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el Transferencia por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub

Private Sub DeleteChild(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    
    dataBase.ClearParameter
    dataBase.CommandText = "DELETE FROM con_librocostodet " _
            + vbCr + "WHERE con_librocostodet.idlibro = ?"

    dataBase.AddParameter "@idlibro", adInteger, adParamInput, Me.IdLibroCosto
    dataBase.Execute
    
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el Transferencia por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub




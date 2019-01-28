VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EKardexDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdKardexDetalle As Long
Private mIdKardex As Long
Private mIdAlmacen As Long
Private mUltimaFecha As Date
Private mCantidad As Double
Private mCostoUnitario As Double
Private mCostoUnitarioPromedio As Double
Private mLKardexDetalleMov As New LEKardexDetMov

Private F As New SistemaLogica.Funciones
Private mConexion As New ADODB.Connection
Private mCalled As Boolean
Private mLoadChild As Boolean
Private mUpdateParent As Boolean

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

Public Property Get UpdateParent() As Boolean
    UpdateParent = mUpdateParent
End Property
Public Property Let UpdateParent(ByVal NewValue As Boolean)
    mUpdateParent = NewValue
End Property

Public Property Get IdKardexDetalle() As Long
    IdKardexDetalle = mIdKardexDetalle
End Property
Public Property Let IdKardexDetalle(ByVal NewValue As Long)
    mIdKardexDetalle = NewValue
End Property

Public Property Get IdKardex() As Long
    IdKardex = mIdKardex
End Property
Public Property Let IdKardex(ByVal NewValue As Long)
    mIdKardex = NewValue
End Property

Public Property Get IdAlmacen() As Long
    IdAlmacen = mIdAlmacen
End Property
Public Property Let IdAlmacen(ByVal NewValue As Long)
    mIdAlmacen = NewValue
End Property

Public Property Get UltimaFecha() As Date
    UltimaFecha = mUltimaFecha
End Property
Public Property Let UltimaFecha(ByVal NewValue As Date)
    mUltimaFecha = NewValue
End Property

Public Property Get Cantidad() As Double
    Cantidad = mCantidad
End Property
Public Property Let Cantidad(ByVal NewValue As Double)
    mCantidad = NewValue
End Property

Public Property Get CostoUnitario() As Double
    CostoUnitario = mCostoUnitario
End Property
Public Property Let CostoUnitario(ByVal NewValue As Double)
    mCostoUnitario = NewValue
End Property

Public Property Get CostoUnitarioPromedio() As Double
    CostoUnitarioPromedio = mCostoUnitarioPromedio
End Property
Public Property Let CostoUnitarioPromedio(ByVal NewValue As Double)
    mCostoUnitarioPromedio = NewValue
End Property

Public Property Get LKardexDetalleMov() As LEKardexDetMov
    Set LKardexDetalleMov = mLKardexDetalleMov
End Property
Public Property Set LKardexDetalleMov(ByVal NewValue As LEKardexDetMov)
    Set mLKardexDetalleMov = NewValue
End Property

Public Sub Fetch(IdKardexDetalle As Long, _
                        Optional IdKardex As Long = 0, _
                        Optional IdAlmacen As Long = 0)
    Dim cWhere As String
    Dim dataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    cWhere = ""
    If IdKardexDetalle > 0 Then
        cWhere = "WHERE ((alm_kardexdet.idkardexdet)=" & IdKardexDetalle & ")"
    End If
    If IdKardex > 0 Then
        If cWhere = "" Then cWhere = "WHERE" Else cWhere = cWhere & " AND"
        cWhere = cWhere & " ((alm_kardexdet.idkardex)=" & IdKardex & ")"
    End If
    If IdAlmacen > 0 Then
        If cWhere = "" Then cWhere = "WHERE" Else cWhere = cWhere & " AND"
        cWhere = cWhere & " ((alm_kardexdet.idalm)=" & IdAlmacen & ")"
    End If
    dataBase.CommandText = "SELECT alm_kardexdet.idkardexdet, alm_kardexdet.idkardex, alm_kardexdet.idalm, alm_kardexdet.fecha, alm_kardexdet.cantidad, alm_kardexdet.costounitario, alm_kardexdet.costounitariopromedio " _
            + vbCr + "FROM alm_kardexdet " _
            + vbCr + cWhere
    
    Set mRecord = dataBase.GetRecordset
    If mRecord.RecordCount > 0 Then Me.SetEntity mRecord
    Exit Sub
    
BLoqueError:
    Err.Raise Err.Number, "[Fetch] " & Err.Source, Err.Description
End Sub

Public Sub SetEntity(Record As ADODB.Recordset)
    Me.IdKardexDetalle = F.NuloNumeric(Record("idkardexdet"))
    Me.IdKardex = F.NuloNumeric(Record("idkardex"))
    Me.IdAlmacen = F.NuloNumeric(Record("idalm"))
    Me.UltimaFecha = Record("fecha")
    Me.Cantidad = F.NuloNumeric(Record("cantidad"))
    Me.CostoUnitario = F.NuloNumeric(Record("costounitario"))
    Me.CostoUnitarioPromedio = F.NuloNumeric(Record("costounitariopromedio"))
    If mLoadChild Then
        Set Me.LKardexDetalleMov.Conexion = mConexion
        Me.LKardexDetalleMov.Fetch Me.IdKardexDetalle
    End If
End Sub

Public Function Save(IdUsuario As Long, MachineName As String) As Boolean
On Error GoTo BLoqueError
    If (F.NuloNumeric(Me.IdKardexDetalle) = 0) Then
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
    xId = FData.GetCodeTable("alm_kardexdet", "idkardexdet", mConexion)
    
    Me.IdKardexDetalle = xId
                
    dataBase.CommandText = "INSERT INTO alm_kardexdet(idkardexdet, idkardex, idalm, fecha, cantidad, costounitario, " _
                & "costounitariopromedio, fechacreacion, horacreacion, usuariocreacion, placacreacion) " _
            + vbCr + "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    
    dataBase.AddParameter "@idkardexdet", adInteger, adParamInput, xId
    dataBase.AddParameter "@idkardex", adInteger, adParamInput, F.NuloNumeric(Me.IdKardex)
    dataBase.AddParameter "@idalm", adInteger, adParamInput, F.NuloNumeric(Me.IdAlmacen)
    dataBase.AddParameter "@fecha", adDate, adParamInput, Me.UltimaFecha
    dataBase.AddParameter "@cantidad", adDouble, adParamInput, F.NuloNumeric(Me.Cantidad)
    dataBase.AddParameter "@costounitario", adDouble, adParamInput, F.NuloNumeric(Me.CostoUnitario)
    dataBase.AddParameter "@costounitariopromedio", adDouble, adParamInput, F.NuloNumeric(Me.CostoUnitarioPromedio)
    '************************************
    ' Parametros de auditoria
    dataBase.AddParameter "@fechacreacion", adInteger, adParamInput, F.ConvertirFechaANumero(Date)
    dataBase.AddParameter "@horacreacion", adInteger, adParamInput, F.ConvertirHoraANumero(Time)
    dataBase.AddParameter "@usuariocreacion", adVarChar, adParamInput, IdUsuario
    dataBase.AddParameter "@placacreacion", adVarChar, adParamInput, MachineName
    '************************************
    dataBase.Execute
    
    Dim KardexDetalleMov As EKardexDetMov
    For Each KardexDetalleMov In Me.LKardexDetalleMov
        If Not mConexion Is Nothing Then
            Set KardexDetalleMov.Conexion = mConexion
            KardexDetalleMov.Called = True
        End If
        KardexDetalleMov.IdKardexDetalle = xId
        If Not KardexDetalleMov.Save(IdUsuario, MachineName) Then Err.Raise &HFFFFFF01, "[Kardex Detalle - Insert] " & Err.Source, "Error al intentar grabar detalle"
        Set KardexDetalleMov = Nothing
    Next
    
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
    
    dataBase.CommandText = "UPDATE alm_kardexdet SET idkardex = ?, idalm = ?, fecha = ?, cantidad = ?, costounitario = ?, " _
                & "costounitariopromedio = ?, fechamodificacion = ?, horamodificacion = ?, usuariomodificacion = ?, placamodificacion = ? " _
            + vbCr + "WHERE (alm_kardexdet.idkardexdet = ?)"
    
    dataBase.AddParameter "@idkardex", adInteger, adParamInput, F.NuloNumeric(Me.IdKardex)
    dataBase.AddParameter "@idalm", adInteger, adParamInput, F.NuloNumeric(Me.IdAlmacen)
    dataBase.AddParameter "@fecha", adDate, adParamInput, Me.UltimaFecha
    dataBase.AddParameter "@cantidad", adDouble, adParamInput, F.NuloNumeric(Me.Cantidad)
    dataBase.AddParameter "@costounitario", adDouble, adParamInput, F.NuloNumeric(Me.CostoUnitario)
    dataBase.AddParameter "@costounitariopromedio", adDouble, adParamInput, F.NuloNumeric(Me.CostoUnitarioPromedio)
    '************************************
    ' Parametros de auditoria
    dataBase.AddParameter "@fechamodificacion", adInteger, adParamInput, F.ConvertirFechaANumero(Date)
    dataBase.AddParameter "@horamodificacion", adInteger, adParamInput, F.ConvertirHoraANumero(Time)
    dataBase.AddParameter "@usuariomodificacion", adVarChar, adParamInput, IdUsuario
    dataBase.AddParameter "@placamodificacion", adVarChar, adParamInput, MachineName
    '************************************
    dataBase.AddParameter "@idkardexdet", adInteger, adParamInput, F.NuloNumeric(Me.IdKardexDetalle)
    dataBase.Execute
    
    Dim KardexDetalleMov As EKardexDetMov
    For Each KardexDetalleMov In Me.LKardexDetalleMov
        If Not mConexion Is Nothing Then
            KardexDetalleMov.Conexion = mConexion
            KardexDetalleMov.Called = True
        End If
        If Not KardexDetalleMov.Save(IdUsuario, MachineName) Then Err.Raise &HFFFFFF01, "[Kardex Detalle - Insert] " & Err.Source, "Error al intentar grabar detalle"
        Set KardexDetalleMov = Nothing
    Next
    
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
    ' Eliminamos a los Hijos
    DeleteChild IdUsuario, MachineName
    
    dataBase.ClearParameter
    dataBase.CommandText = "DELETE FROM alm_kardexdet " _
            + vbCr + "WHERE alm_kardexdet.idkardexdet = ?"
    dataBase.AddParameter "@idkardexdet", adInteger, adParamInput, Me.IdKardexDetalle
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

Private Sub DeleteChild(IdUsuario As Long, MachineName As String)
On Error GoTo BLoqueError

    Dim KardexDetalleMov As EKardexDetMov
    For Each KardexDetalleMov In Me.LKardexDetalleMov
        If Not mConexion Is Nothing Then
            KardexDetalleMov.Conexion = mConexion
            KardexDetalleMov.Called = True
        End If
        KardexDetalleMov.Delete IdUsuario, MachineName
        Set KardexDetalleMov = Nothing
    Next
    Exit Sub

BLoqueError:
    Set KardexDetalleMov = Nothing
    Err.Raise Err.Number, "[DeleteChild]" & Err.Source, Err.Description
    Exit Sub
End Sub

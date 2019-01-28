VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ELibroCostoTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdLibroCostoTemp As Long
Private mIdAlmacenProceso As Long
Private mIdMovimientoDetalle As Long
Private mTipoMovimiento As String
Private mFechaMovimiento As Date
Private mIdItem As Long
Private mCantidad As Double
Private mCostoUnitario As Double
Private mCostoUnitarioPromedio As Double
Private mCostoPrimo As Double
Private mCostoMOD As Double
Private mCostoCIF As Double
Private F As New SistemaLogica.Funciones
Private mConexion As New ADODB.Connection
Private mCalled As Boolean

Private Sub Class_Initialize()
    mCalled = False
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

Public Property Get IdLibroCostoTemp() As Long
    IdLibroCostoTemp = mIdLibroCostoTemp
End Property
Public Property Let IdLibroCostoTemp(ByVal NewValue As Long)
    mIdLibroCostoTemp = NewValue
End Property

Public Property Get IdAlmacenProceso() As Long
    IdAlmacenProceso = mIdAlmacenProceso
End Property
Public Property Let IdAlmacenProceso(ByVal NewValue As Long)
    mIdAlmacenProceso = NewValue
End Property

Public Property Get IdMovimientoDetalle() As Long
    IdMovimientoDetalle = mIdMovimientoDetalle
End Property
Public Property Let IdMovimientoDetalle(ByVal NewValue As Long)
    mIdMovimientoDetalle = NewValue
End Property

Public Property Get TipoMovimiento() As String
    TipoMovimiento = mTipoMovimiento
End Property
Public Property Let TipoMovimiento(ByVal NewValue As String)
    mTipoMovimiento = NewValue
End Property

Public Property Get FechaMovimiento() As Date
    FechaMovimiento = mFechaMovimiento
End Property
Public Property Let FechaMovimiento(ByVal NewValue As Date)
    mFechaMovimiento = NewValue
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

Public Property Get CostoPrimo() As Double
    CostoPrimo = mCostoPrimo
End Property
Public Property Let CostoPrimo(ByVal NewValue As Double)
    mCostoPrimo = NewValue
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

Public Sub Fetch(IdMovimientoDetalle As Long)
    Dim dataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    dataBase.CommandText = "SELECT con_librocostotemp.idlibrocostotemp, con_librocostotemp.idalmproc, con_librocostotemp.idmovdet, con_librocostotemp.tipmov, con_librocostotemp.fecha, con_librocostotemp.cantidad, con_librocostotemp.costounitario, con_librocostotemp.costounitariopromedio, con_librocostotemp.costoprimo, con_librocostotemp.costomod, con_librocostotemp.costocif " _
            + vbCr + "FROM con_librocostotemp " _
            + vbCr + "WHERE ((con_librocostotemp.idmovdet)=" & IdMovimientoDetalle & ")"
    
    Set mRecord = dataBase.GetRecordset
    If mRecord.RecordCount > 0 Then Me.SetEntity mRecord
    Exit Sub
    
BLoqueError:
    Err.Raise Err.Number, "[Fetch] " & Err.Source, Err.Description
End Sub

Public Sub SetEntity(Record As ADODB.Recordset)
    Me.IdLibroCostoTemp = F.NuloNumeric(Record("idlibrocostotemp"))
    Me.IdAlmacenProceso = F.NuloNumeric(Record("idalmproc"))
    Me.IdMovimientoDetalle = F.NuloNumeric(Record("idmovdet"))
    Me.TipoMovimiento = F.NuloString(Record("tipmov"))
    Me.FechaMovimiento = Record("fecha")
    Me.Cantidad = F.NuloNumeric(Record("cantidad"))
    Me.CostoUnitario = F.NuloNumeric(Record("costounitario"))
    Me.CostoUnitarioPromedio = F.NuloNumeric(Record("costounitariopromedio"))
    Me.CostoPrimo = F.NuloNumeric(Record("costoprimo"))
    Me.CostoMOD = F.NuloNumeric(Record("costomod"))
    Me.CostoCIF = F.NuloNumeric(Record("costocif"))
End Sub

Public Function Save(IdUsuario As Long, MachineName As String) As Boolean
On Error GoTo BLoqueError
    If (F.NuloNumeric(Me.IdLibroCostoTemp) = 0) Then
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
    Dim F As New SistemaLogica.Funciones
    
On Error GoTo BLoqueError
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    ' Se elimina cualquier registro existente del mismo movimiento
    DeleteEquals IdUsuario, MachineName
    
    xId = FData.GetCodeTable("con_librocostotemp", "idlibrocostotemp", mConexion)
    dataBase.ClearParameter

    dataBase.CommandText = "INSERT INTO con_librocostotemp(idlibrocostotemp, idalmproc, idmovdet, tipmov, fecha, cantidad, costounitario, costounitariopromedio, costoprimo, costomod, costocif, fechacreacion, horacreacion, usuariocreacion, placacreacion) " _
            + vbCr + "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

    dataBase.AddParameter "@idlibrocostotemp", adInteger, adParamInput, xId
    dataBase.AddParameter "@idalmproc", adInteger, adParamInput, F.NuloNumeric(Me.IdAlmacenProceso)
    dataBase.AddParameter "@idmovdet", adInteger, adParamInput, F.NuloNumeric(Me.IdMovimientoDetalle)
    dataBase.AddParameter "@tipmov", adVarChar, adParamInput, F.NuloString(Me.TipoMovimiento)
    dataBase.AddParameter "@fecha", adDate, adParamInput, Me.FechaMovimiento
    dataBase.AddParameter "@cantidad", adDouble, adParamInput, F.NuloNumeric(Me.Cantidad)
    dataBase.AddParameter "@costounitario", adDouble, adParamInput, F.NuloNumeric(Me.CostoUnitario)
    dataBase.AddParameter "@costounitariopromedio", adDouble, adParamInput, F.NuloNumeric(Me.CostoUnitarioPromedio)
    dataBase.AddParameter "@costoprimo", adDouble, adParamInput, F.NuloNumeric(Me.CostoPrimo)
    dataBase.AddParameter "@costomod", adDouble, adParamInput, F.NuloNumeric(Me.CostoMOD)
    dataBase.AddParameter "@costocif", adDouble, adParamInput, F.NuloNumeric(Me.CostoCIF)
    ' Parametros de Auditoria
    dataBase.AddParameter "@fechacreacion", adInteger, adParamInput, F.ConvertirFechaANumero(Date)
    dataBase.AddParameter "@horacreacion", adInteger, adParamInput, F.ConvertirHoraANumero(Time)
    dataBase.AddParameter "@usuariocreacion", adVarChar, adParamInput, IdUsuario
    dataBase.AddParameter "@placacreacion", adVarChar, adParamInput, MachineName
    dataBase.Execute
        
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub Update(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    dataBase.ClearParameter
        
    dataBase.CommandText = "UPDATE con_librocostotemp SET idalmproc = ?, idmovdet = ?, tipmov = ?, fecha = ?, cantidad = ?, costounitario = ?, costounitariopromedio = ?, costoprimo = ?, costomod = ?, costocif = ?, fechamodificacion = ?, horamodificacion = ?, usuariomodificacion = ?, placamodificacion = ? " _
            + vbCr + "WHERE(con_librocostotemp.idlibrocostotemp = ?)"

    dataBase.AddParameter "@idalmproc", adInteger, adParamInput, Me.IdAlmacenProceso
    dataBase.AddParameter "@idmovdet", adInteger, adParamInput, Me.IdMovimientoDetalle
    dataBase.AddParameter "@tipmov", adVarChar, adParamInput, Me.TipoMovimiento
    dataBase.AddParameter "@fecha", adDate, adParamInput, Me.FechaMovimiento
    dataBase.AddParameter "@cantidad", adDouble, adParamInput, Me.Cantidad
    dataBase.AddParameter "@costounitario", adDouble, adParamInput, Me.CostoUnitario
    dataBase.AddParameter "@costounitariopromedio", adDouble, adParamInput, Me.CostoUnitarioPromedio
    dataBase.AddParameter "@costoprimo", adDouble, adParamInput, Me.CostoPrimo
    dataBase.AddParameter "@costomod", adDouble, adParamInput, Me.CostoMOD
    dataBase.AddParameter "@costocif", adDouble, adParamInput, Me.CostoCIF
    dataBase.AddParameter "@fechacreacion", adInteger, adParamInput, F.ConvertirFechaANumero(Date)
    dataBase.AddParameter "@horacreacion", adInteger, adParamInput, F.ConvertirHoraANumero(Time)
    dataBase.AddParameter "@usuariocreacion", adVarChar, adParamInput, IdUsuario
    dataBase.AddParameter "@placacreacion", adVarChar, adParamInput, MachineName
    dataBase.AddParameter "@idlibrocostotemp", adInteger, adParamInput, Me.IdLibroCostoTemp
    dataBase.Execute
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Public Sub Delete(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    dataBase.ClearParameter
    
    dataBase.CommandText = "DELETE FROM con_librocostotemp " _
            + vbCr + "WHERE con_librocostotemp.idlibrocostotemp = ?"
    
    dataBase.AddParameter "@idlibrocostotemp", adInteger, adParamInput, Me.IdLibroCostoTemp
    dataBase.Execute
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Private Sub DeleteEquals(IdUsuario As Long, MachineName As String)
    Dim mDataBase As New SistemaData.EDataBase
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set mDataBase.Connection = mConexion
    mDataBase.CommandText = "DELETE FROM con_librocostotemp WHERE idmovdet = " & Me.IdMovimientoDetalle
    mDataBase.Execute
    Exit Sub

BLoqueError:
    Set mDataBase = Nothing
    Err.Raise Err.Number, "[DeleteEquals]" & Err.Source, Err.Description
    Exit Sub
End Sub



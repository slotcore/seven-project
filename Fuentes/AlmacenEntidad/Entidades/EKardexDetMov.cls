VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EKardexDetMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdKardexDetalleMov As Long
Private mIdKardexDetalle As Long
Private mIdMovimientoDetalle As Long
Private mTipoMovimiento As String
Private mFechaMovimiento As Date
Private mIdItem As Long
Private mCantidad As Double
Private mCostoUnitario As Double
Private mCostoUnitarioPromedio As Double
Private F As New SistemaLogica.Funciones
Private mConexion As New ADODB.Connection
Private mCalled As Boolean
Private mUpdateParent As Boolean

Private Sub Class_Initialize()
    mCalled = False
    mUpdateParent = False
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

Public Property Get UpdateParent() As Boolean
    UpdateParent = mUpdateParent
End Property
Public Property Let UpdateParent(ByVal NewValue As Boolean)
    mUpdateParent = NewValue
End Property

Public Property Get IdKardexDetalleMov() As Long
    IdKardexDetalleMov = mIdKardexDetalleMov
End Property
Public Property Let IdKardexDetalleMov(ByVal NewValue As Long)
    mIdKardexDetalleMov = NewValue
End Property

Public Property Get IdKardexDetalle() As Long
    IdKardexDetalle = mIdKardexDetalle
End Property
Public Property Let IdKardexDetalle(ByVal NewValue As Long)
    mIdKardexDetalle = NewValue
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

Public Sub Fetch(IdKardexDetalleMov As Long, _
                    Optional IdKardexDetalle As Long = 0, _
                    Optional IdMovimientoDetalle As Long = 0)
    Dim cWhere As String
    Dim dataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    cWhere = ""
    If IdKardexDetalleMov > 0 Then
        cWhere = "WHERE ((alm_kardexdetmov.idkardexdetmov)=" & IdKardexDetalleMov & ")"
    End If
    If IdKardexDetalle > 0 Then
        If cWhere = "" Then cWhere = "WHERE" Else cWhere = cWhere & " AND"
        cWhere = cWhere & " ((alm_kardexdetmov.idkardexdet)=" & IdKardexDetalle & ")"
    End If
    If IdMovimientoDetalle > 0 Then
        If cWhere = "" Then cWhere = "WHERE" Else cWhere = cWhere & " AND"
        cWhere = cWhere & " ((alm_kardexdetmov.idmovdet)=" & IdMovimientoDetalle & ")"
    End If
    dataBase.CommandText = "SELECT alm_kardexdetmov.idkardexdetmov, alm_kardexdetmov.idkardexdet, alm_kardexdetmov.idmovdet, alm_kardexdetmov.tipmov, alm_kardexdetmov.fecha, alm_kardexdetmov.cantidad, alm_kardexdetmov.costounitario, alm_kardexdetmov.costounitariopromedio " _
            + vbCr + "FROM alm_kardexdetmov " _
            + vbCr + cWhere
    
    Set mRecord = dataBase.GetRecordset
    If mRecord.RecordCount > 0 Then Me.SetEntity mRecord
    Exit Sub
    
BLoqueError:
    Err.Raise Err.Number, "[Fetch] " & Err.Source, Err.Description
End Sub

Public Sub SetEntity(Record As ADODB.Recordset)
    Me.IdKardexDetalleMov = F.NuloNumeric(Record("idkardexdetmov"))
    Me.IdKardexDetalle = F.NuloNumeric(Record("idkardexdet"))
    Me.IdMovimientoDetalle = F.NuloNumeric(Record("idmovdet"))
    Me.TipoMovimiento = F.NuloString(Record("tipmov"))
    Me.FechaMovimiento = Record("fecha")
    Me.Cantidad = F.NuloNumeric(Record("cantidad"))
    Me.CostoUnitario = F.NuloNumeric(Record("costounitario"))
    Me.CostoUnitarioPromedio = F.NuloNumeric(Record("costounitariopromedio"))
End Sub

Public Function Save(IdUsuario As Long, MachineName As String) As Boolean
On Error GoTo BLoqueError
    If (F.NuloNumeric(Me.IdKardexDetalleMov) = 0) Then
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
    
    xId = FData.GetCodeTable("alm_kardexdetmov", "idkardexdetmov", mConexion)
    dataBase.ClearParameter

    dataBase.CommandText = "INSERT INTO alm_kardexdetmov(idkardexdetmov, idkardexdet, idmovdet, tipmov, fecha, cantidad, costounitario, costounitariopromedio, fechacreacion, horacreacion, usuariocreacion, placacreacion) " _
            + vbCr + "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

    dataBase.AddParameter "@idkardexdetmov", adInteger, adParamInput, xId
    dataBase.AddParameter "@idkardexdet", adInteger, adParamInput, F.NuloNumeric(Me.IdKardexDetalle)
    dataBase.AddParameter "@idmovdet", adInteger, adParamInput, F.NuloNumeric(Me.IdMovimientoDetalle)
    dataBase.AddParameter "@tipmov", adVarChar, adParamInput, F.NuloString(Me.TipoMovimiento)
    dataBase.AddParameter "@fecha", adDate, adParamInput, Me.FechaMovimiento
    dataBase.AddParameter "@cantidad", adDouble, adParamInput, F.NuloNumeric(Me.Cantidad)
    dataBase.AddParameter "@costounitario", adDouble, adParamInput, F.NuloNumeric(Me.CostoUnitario)
    dataBase.AddParameter "@costounitariopromedio", adDouble, adParamInput, F.NuloNumeric(Me.CostoUnitarioPromedio)
    '
    dataBase.AddParameter "@fechacreacion", adInteger, adParamInput, F.ConvertirFechaANumero(Date)
    dataBase.AddParameter "@horacreacion", adInteger, adParamInput, F.ConvertirHoraANumero(Time)
    dataBase.AddParameter "@usuariocreacion", adVarChar, adParamInput, IdUsuario
    dataBase.AddParameter "@placacreacion", adVarChar, adParamInput, MachineName
    dataBase.Execute
    
    ' Si se necesita actualizar a los padres
    If mUpdateParent Then ActualizarPadre IdUsuario, MachineName
    
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
        
    dataBase.CommandText = "UPDATE alm_kardexdetmov SET idkardexdet = ?, idmovdet = ?, tipmov = ?, fecha = ?, cantidad = ?, costounitario = ?, costounitariopromedio = ?, fechamodificacion = ?, horamodificacion = ?, usuariomodificacion = ?, placamodificacion = ? " _
            + vbCr + "WHERE(alm_kardexdetmov.idkardexdetmov = ?)"

    dataBase.AddParameter "@idkardexdet", adInteger, adParamInput, F.NuloNumeric(Me.IdKardexDetalle)
    dataBase.AddParameter "@idmovdet", adInteger, adParamInput, F.NuloNumeric(Me.IdMovimientoDetalle)
    dataBase.AddParameter "@tipmov", adVarChar, adParamInput, F.NuloString(Me.TipoMovimiento)
    dataBase.AddParameter "@fecha", adDate, adParamInput, Me.FechaMovimiento
    dataBase.AddParameter "@cantidad", adDouble, adParamInput, F.NuloNumeric(Me.Cantidad)
    dataBase.AddParameter "@costounitario", adDouble, adParamInput, F.NuloNumeric(Me.CostoUnitario)
    dataBase.AddParameter "@costounitariopromedio", adDouble, adParamInput, F.NuloNumeric(Me.CostoUnitarioPromedio)
    dataBase.AddParameter "@fechacreacion", adInteger, adParamInput, F.ConvertirFechaANumero(Date)
    dataBase.AddParameter "@horacreacion", adInteger, adParamInput, F.ConvertirHoraANumero(Time)
    dataBase.AddParameter "@usuariocreacion", adVarChar, adParamInput, IdUsuario
    dataBase.AddParameter "@placacreacion", adVarChar, adParamInput, MachineName
    dataBase.AddParameter "@idkardexdetmov", adInteger, adParamInput, F.NuloNumeric(Me.IdKardexDetalleMov)
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
    
    dataBase.CommandText = "DELETE FROM alm_kardexdetmov " _
            + vbCr + "WHERE alm_kardexdetmov.idkardexdetmov = ?"
    
    dataBase.AddParameter "@idkardexdetmov", adInteger, adParamInput, Me.IdKardexDetalleMov
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

Public Sub ActualizarPadre(IdUsuario As Long, MachineName As String)
    Dim Kardx As New AlmacenEntidad.EKardex
    Dim KardxDet As New AlmacenEntidad.EKardexDet
    Dim ImporteAcumulario As Double
    Dim CantidadAcumulada As Double
    
On Error GoTo BLoqueError
    
    ' Se graba el kardex detalle
    If Not mConexion Is Nothing Then Set KardxDet.Conexion = mConexion
    KardxDet.Called = True
    KardxDet.LoadChild = False
    KardxDet.Fetch Me.IdKardexDetalle
    If KardxDet.IdKardexDetalle = 0 Then Err.Raise &HFFFFFF01, , "No se encontró el elemento padre"
    If Me.TipoMovimiento = "I" Then
        KardxDet.Cantidad = KardxDet.Cantidad + Me.Cantidad
    Else
        KardxDet.Cantidad = KardxDet.Cantidad - Me.Cantidad
    End If
    
    ' Se valida la fecha de ingreso
    If KardxDet.UltimaFecha > Me.FechaMovimiento Then
        Err.Raise &HFFFFFF01, , "La fecha ingresada es menor a la ultima fecha registrada en kardex"
    End If
    
    KardxDet.UltimaFecha = Me.FechaMovimiento
    KardxDet.CostoUnitario = Me.CostoUnitario
    KardxDet.CostoUnitarioPromedio = Me.CostoUnitarioPromedio
    If Not KardxDet.Save(IdUsuario, MachineName) Then Err.Raise &HFFFFFF01, , "Error al intentar grabar el elemento padre"
    
    ' Se graba el kardex
    If Not mConexion Is Nothing Then Set Kardx.Conexion = mConexion
    Kardx.Called = True
    Kardx.LoadChild = False
    Kardx.Fetch KardxDet.IdKardex
    If Kardx.IdKardex = 0 Then Err.Raise &HFFFFFF01, , "No se encontró el elemento padre"
    If Me.TipoMovimiento = "I" Then
        Kardx.Cantidad = Kardx.Cantidad + Me.Cantidad
    Else
        Kardx.Cantidad = Kardx.Cantidad - Me.Cantidad
    End If
    Kardx.CostoUnitario = Me.CostoUnitario
    Kardx.CostoUnitarioPromedio = Me.CostoUnitarioPromedio
    If Not Kardx.Save(IdUsuario, MachineName) Then Err.Raise &HFFFFFF01, , "Error al intentar grabar el elemento padre"
    
    Exit Sub

BLoqueError:
    Err.Raise &HFFFFFF01, Err.Source, Err.Description
    Exit Sub
End Sub

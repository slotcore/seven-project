VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EInventarioIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdInventarioInicial As Long
Private mIdAlmacen As Long
Private mAlmacen As String
Private mIdMoneda As Long
Private mMoneda As String
Private mFechaVigencia As Date
Private mIdTipoDocumentoReferencia As Long
Private mTipoDocumentoReferencia As String
Private mIdDocumentoReferencia As Long
Private mDocumentoReferencia As String
Private mIdEstado As Long
Private mEstado As String
Private mLInventarioIniDet As New LEInventarioIniDet
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

Public Property Get IdInventarioInicial() As Long
    IdInventarioInicial = mIdInventarioInicial
End Property
Public Property Let IdInventarioInicial(ByVal NewValue As Long)
    mIdInventarioInicial = NewValue
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

Public Property Get IdMoneda() As Long
    IdMoneda = mIdMoneda
End Property
Public Property Let IdMoneda(ByVal NewValue As Long)
    mIdMoneda = NewValue
End Property

Public Property Get Moneda() As String
    Moneda = mMoneda
End Property
Public Property Let Moneda(ByVal NewValue As String)
    mMoneda = NewValue
End Property

Public Property Get FechaVigencia() As Date
    FechaVigencia = mFechaVigencia
End Property
Public Property Let FechaVigencia(ByVal NewValue As Date)
    mFechaVigencia = NewValue
End Property

Public Property Get IdTipoDocumentoReferencia() As Long
    IdTipoDocumentoReferencia = mIdTipoDocumentoReferencia
End Property
Public Property Let IdTipoDocumentoReferencia(ByVal NewValue As Long)
    mIdTipoDocumentoReferencia = NewValue
End Property

Public Property Get TipoDocumentoReferencia() As String
    TipoDocumentoReferencia = mTipoDocumentoReferencia
End Property
Public Property Let TipoDocumentoReferencia(ByVal NewValue As String)
    mTipoDocumentoReferencia = NewValue
End Property

Public Property Get IdDocumentoReferencia() As Long
    IdDocumentoReferencia = mIdDocumentoReferencia
End Property
Public Property Let IdDocumentoReferencia(ByVal NewValue As Long)
    mIdDocumentoReferencia = NewValue
End Property

Public Property Get DocumentoReferencia() As String
    DocumentoReferencia = mDocumentoReferencia
End Property
Public Property Let DocumentoReferencia(ByVal NewValue As String)
    mDocumentoReferencia = NewValue
End Property

Public Property Get IdEstado() As Long
    IdEstado = mIdEstado
End Property
Public Property Let IdEstado(ByVal NewValue As Long)
    mIdEstado = NewValue
End Property

Public Property Get Estado() As String
    Estado = mEstado
End Property
Public Property Let Estado(ByVal NewValue As String)
    mEstado = NewValue
End Property

Public Property Get LInventarioIniDet() As LEInventarioIniDet
    Set LInventarioIniDet = mLInventarioIniDet
End Property
Public Property Set LInventarioIniDet(ByVal NewValue As LEInventarioIniDet)
    Set mLInventarioIniDet = NewValue
End Property

Public Sub Fetch(IdInventarioInicial As String)
    Dim dataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
            
    dataBase.CommandText = "SELECT alm_inventarioinicial.idinventarioinicial, alm_inventarioinicial.idalm, alm_almacenes.descripcion AS alm, alm_inventarioinicial.idmon, mae_moneda.simbolo AS mon, alm_inventarioinicial.fechavigencia, alm_inventarioinicial.idtipdocref, mae_documento.descripcion AS tipdocref, alm_inventarioinicial.iddocref, '' AS docref, alm_inventarioinicial.idestado, '' AS estado " _
        + vbCr + "FROM ((alm_inventarioinicial INNER JOIN alm_almacenes ON alm_inventarioinicial.idalm = alm_almacenes.id) INNER JOIN mae_moneda ON alm_inventarioinicial.idmon = mae_moneda.id) INNER JOIN mae_documento ON alm_inventarioinicial.idtipdocref = mae_documento.id " _
        + vbCr + "WHERE (((alm_inventarioinicial.idinventarioinicial)=" & IdInventarioInicial & "))"
        
    Set mRecord = dataBase.GetRecordset
    If mRecord.RecordCount > 0 Then Me.SetEntity mRecord
    Exit Sub
    
BLoqueError:
    Err.Raise &HFFFFFF01, , "[Fetch] " + Trim(Err.Description)
End Sub

Public Sub SetEntity(Record As ADODB.Recordset)
    Me.IdInventarioInicial = F.NuloNumeric(Record("idinventarioinicial"))
    Me.IdAlmacen = F.NuloNumeric(Record("idalm"))
    Me.Almacen = F.NuloString(Record("alm"))
    Me.IdMoneda = F.NuloNumeric(Record("idmon"))
    Me.Moneda = F.NuloString(Record("mon"))
    Me.FechaVigencia = F.ConvertirNumeroAFecha(F.NuloNumeric(Record("fechavigencia")))
    Me.IdTipoDocumentoReferencia = F.NuloNumeric(Record("idtipdocref"))
    Me.TipoDocumentoReferencia = F.NuloString(Record("tipdocref"))
    Me.IdDocumentoReferencia = F.NuloNumeric(Record("iddocref"))
    Me.DocumentoReferencia = F.NuloString(Record("docref"))
    Me.IdEstado = F.NuloNumeric(Record("idestado"))
    Me.Estado = F.NuloString(Record("estado"))
End Sub

Public Function Save(IdUsuario As Long, MachineName As String) As Boolean
On Error GoTo BLoqueError
    If (F.NuloNumeric(Me.IdInventarioInicial) = 0) Then
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
    
    xId = FData.GetCodeTable("alm_inventarioinicial", "idinventarioinicial", mConexion)
    dataBase.ClearParameter
    
    dataBase.CommandText = "INSERT INTO alm_inventarioinicial(idinventarioinicial, idalm, idmon, fechavigencia, idtipdocref, iddocref, idestado, fechacreacion, horacreacion, usuariocreacion, placacreacion) " _
            + vbCr + "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

    dataBase.AddParameter "@idinventarioinicial", adInteger, adParamInput, xId
    dataBase.AddParameter "@idalm", adInteger, adParamInput, F.NuloNumeric(Me.IdAlmacen)
    dataBase.AddParameter "@idmon", adInteger, adParamInput, F.NuloNumeric(Me.IdMoneda)
    dataBase.AddParameter "@fechavigencia", adInteger, adParamInput, F.ConvertirFechaANumero(Me.FechaVigencia)
    dataBase.AddParameter "@idtipdocref", adInteger, adParamInput, F.NuloNumeric(Me.IdTipoDocumentoReferencia)
    dataBase.AddParameter "@iddocref", adInteger, adParamInput, F.NuloNumeric(Me.IdDocumentoReferencia)
    dataBase.AddParameter "@idestado", adInteger, adParamInput, F.NuloNumeric(Me.IdEstado)
    '************************************
    ' Parametros de auditoria
    dataBase.AddParameter "@fechacreacion", adInteger, adParamInput, F.ConvertirFechaANumero(Date)
    dataBase.AddParameter "@horacreacion", adInteger, adParamInput, F.ConvertirHoraANumero(Time)
    dataBase.AddParameter "@usuariocreacion", adVarChar, adParamInput, IdUsuario
    dataBase.AddParameter "@placacreacion", adVarChar, adParamInput, MachineName
    '************************************
    dataBase.Execute
    Me.IdInventarioInicial = xId
    
    Dim InvIniDet As EInventarioIniDet
    For Each InvIniDet In Me.LInventarioIniDet
        If Not mConexion Is Nothing Then
            Set InvIniDet.Conexion = mConexion
            InvIniDet.Called = True
        End If
        InvIniDet.IdInventarioInicial = xId
        InvIniDet.Save IdUsuario, MachineName
    Next
    
    If (Me.IdEstado = F.NuloNumeric(F.KeyValue("EstadoAprobadoInventarioInicial", mConexion))) Then
        ' Agregamos los Movimientos
        InsertMovimientos IdUsuario, MachineName
    End If
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el ParteProduccion por el siguiente motivo :" + Trim(Err.Description)
End Sub

Private Sub InsertMovimientos(IdUsuario As Long, MachineName As String)
    Dim MovIng As New AlmacenEntidad.EMovimiento
        
    ' Ingreso
    MovIng.IdTipoDocumento = 0
    MovIng.FechaMovimiento = Me.FechaVigencia
    MovIng.NumeroSerie = F.NuloString(F.KeyValue("NumeroSerieInventarioInicial", mConexion))
    MovIng.NumeroDocumento = F.HallaNumeroDocumento("alm_ingreso", "'" & MovIng.NumeroSerie & "'", "numser", mConexion)
    MovIng.IdResponsable = 0
    MovIng.IdProveedor = 0
    MovIng.Proveedor = ""
    MovIng.IdEstado = 2
    MovIng.IdTipoMovimiento = -1
    MovIng.IdTipoDocumentoReferencia = F.NuloNumeric(F.KeyValue("IdDocumentoInventarioInicial", mConexion))
    MovIng.IdDocumentoReferencia = Me.IdInventarioInicial
    MovIng.DocumentoReferencia = ""
    MovIng.IdAlmacen = Me.IdAlmacen
    MovIng.Glosa = ""
    MovIng.AnhoTrabajo = Year(Me.FechaVigencia)
    MovIng.MesTrabajo = Month(Me.FechaVigencia)
    ' Detalles
    Dim InvIniDet As EInventarioIniDet
    For Each InvIniDet In Me.LInventarioIniDet
        Dim MovIngDet As New EMovimientoDet
        
        MovIngDet.IdItem = InvIniDet.IdItem
        MovIngDet.Cantidad = InvIniDet.Cantidad
        MovIngDet.CantidadTeorica = InvIniDet.Cantidad
        MovIng.LMovimientoDet.Add MovIngDet
        Set MovIngDet = Nothing
    Next
    
    If Not mConexion Is Nothing Then
        Set MovIng.Conexion = mConexion
        MovIng.Called = True
    End If
    
    ' Grabamos los movimientos
    If (MovIng.LMovimientoDet.Count > 0) Then MovIng.Save IdUsuario, MachineName
End Sub

Private Sub Update(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    dataBase.ClearParameter

    dataBase.CommandText = "UPDATE alm_inventarioinicial SET idalm = ?, idmon = ?, fechavigencia = ?, idtipdocref = ?, iddocref = ? " _
             & ", idestado = ?, fechamodificacion = ?, horamodificacion = ?, usuariomodificacion = ?, placamodificacion = ? " _
            + vbCr + "WHERE (alm_inventarioinicial.idinventarioinicial = ?)"

    dataBase.AddParameter "@idalm", adInteger, adParamInput, F.NuloNumeric(Me.IdAlmacen)
    dataBase.AddParameter "@idmon", adInteger, adParamInput, F.NuloNumeric(Me.IdMoneda)
    dataBase.AddParameter "@fechavigencia", adInteger, adParamInput, F.ConvertirFechaANumero(Me.FechaVigencia)
    dataBase.AddParameter "@idtipdocref", adInteger, adParamInput, F.NuloNumeric(Me.IdTipoDocumentoReferencia)
    dataBase.AddParameter "@iddocref", adInteger, adParamInput, F.NuloNumeric(Me.IdDocumentoReferencia)
    dataBase.AddParameter "@idestado", adInteger, adParamInput, F.NuloNumeric(Me.IdEstado)
    '************************************
    ' Parametros de auditoria
    dataBase.AddParameter "@fechamodificacion", adInteger, adParamInput, F.ConvertirFechaANumero(Date)
    dataBase.AddParameter "@horamodificacion", adInteger, adParamInput, F.ConvertirHoraANumero(Time)
    dataBase.AddParameter "@usuariomodificacion", adVarChar, adParamInput, IdUsuario
    dataBase.AddParameter "@placamodificacion", adVarChar, adParamInput, MachineName
    '************************************
    dataBase.AddParameter "@idinventarioinicial", adInteger, adParamInput, F.NuloNumeric(Me.IdInventarioInicial)
    dataBase.Execute
    
    ' Borramos Hijos
    DeleteChild IdUsuario, MachineName
    
    ' Grabamos Hijos
    Dim InvIniDet As EInventarioIniDet
    For Each InvIniDet In Me.LInventarioIniDet
        If Not mConexion Is Nothing Then
            Set InvIniDet.Conexion = mConexion
            InvIniDet.Called = True
        End If
        InvIniDet.IdInventarioInicial = Me.IdInventarioInicial
        InvIniDet.Save IdUsuario, MachineName
    Next
    
    ' Estado Aprobado
    If (Me.IdEstado = F.NuloNumeric(F.KeyValue("EstadoAprobadoInventarioInicial", mConexion))) Then
        ' Agregamos los Movimientos
        DeleteMovimientos IdUsuario, MachineName
        InsertMovimientos IdUsuario, MachineName
    End If
    ' Estado Anulado
    If (Me.IdEstado = F.NuloNumeric(F.KeyValue("EstadoAnuladoInventarioInicial", mConexion))) Then
        DeleteMovimientos IdUsuario, MachineName
    End If
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el ParteProduccion por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub

Public Sub Delete(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    ' Eliminamos los movimientos
    DeleteMovimientos IdUsuario, MachineName
    ' Eliminamos a los Hijos
    DeleteChild IdUsuario, MachineName
    
    dataBase.ClearParameter
    dataBase.CommandText = "DELETE FROM alm_inventarioinicial " _
            + vbCr + "WHERE alm_inventarioinicial.idinventarioinicial = ?"
    
    dataBase.AddParameter "@idinventarioinicial", adInteger, adParamInput, Me.IdInventarioInicial
    dataBase.Execute
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el ParteProduccion por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub

Public Sub DeleteMovimientos(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    Dim xRs As New ADODB.Recordset
    
On Error GoTo BLoqueError
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    ' Encontramos los movimientos relacionados
    dataBase.ClearParameter
    
    dataBase.CommandText = "SELECT alm_ingreso.id  FROM alm_ingreso " _
            + vbCr + "WHERE alm_ingreso.idtipdocref = ? AND alm_ingreso.iddocref = ?"
            
    dataBase.AddParameter "@idtipdocref", adBigInt, adParamInput, F.NuloNumeric(F.KeyValue("IdDocumentoInventarioInicial", mConexion))
    dataBase.AddParameter "@iddocref", adBigInt, adParamInput, Me.IdInventarioInicial
    Set xRs = dataBase.GetRecordset
    If xRs.State = 0 Then GoTo BLoqueError
    If xRs.RecordCount > 0 Then
        xRs.MoveFirst
        While Not xRs.EOF
            Dim MovObj As New AlmacenEntidad.EMovimiento
            MovObj.IdMovimiento = xRs("id")
            ' Eliminamos los movimientos
            If Not mConexion Is Nothing Then
                Set MovObj.Conexion = mConexion
                MovObj.Called = True
            End If
            MovObj.Delete IdUsuario, MachineName
            
            Set MovObj = Nothing
            xRs.MoveNext
        Wend
    End If
    
    Set dataBase = Nothing
    Set xRs = Nothing
    Exit Sub

BLoqueError:
    Set dataBase = Nothing
    Set xRs = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el Inventario por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub

Public Sub DeleteChild(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    
    dataBase.ClearParameter

    dataBase.CommandText = "DELETE FROM alm_inventarioinicialdet " _
            + vbCr + "WHERE alm_inventarioinicialdet.idinventarioinicial = ?"

    dataBase.AddParameter "@idinventarioinicial", adInteger, adParamInput, Me.IdInventarioInicial
    dataBase.Execute
    
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el Inventario por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub

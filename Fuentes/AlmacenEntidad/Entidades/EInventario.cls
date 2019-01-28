VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EInventario"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdInventario As Long
Private mIdAlmacen As Long
Private mAlmacen As String
Private mFechaInventario As Date
Private mFechaVigencia As Date
Private mNumeroSerie As String
Private mNumeroDocumento As String
Private mIdResponsable As Long
Private mResponsable As String
Private mIdTipoDocumento As Long
Private mIdEstado As Long
Private mEstado As String
Private mIdTipoInventario As Long
Private mTipoInventario As String
Private mDescripcion As String
Private mIdTipoFiltro As Long
Private mTipoFiltro As String
Private mGlosa As String
Private mInventarioDetS As New LEInventarioDet

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

Public Property Get IdInventario() As Long
    IdInventario = mIdInventario
End Property
Public Property Let IdInventario(ByVal NewValue As Long)
    mIdInventario = NewValue
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

Public Property Get FechaInventario() As Date
    FechaInventario = mFechaInventario
End Property
Public Property Let FechaInventario(ByVal NewValue As Date)
    mFechaInventario = NewValue
End Property

Public Property Get FechaVigencia() As Date
    FechaVigencia = mFechaVigencia
End Property
Public Property Let FechaVigencia(ByVal NewValue As Date)
    mFechaVigencia = NewValue
End Property

Public Property Get NumeroSerie() As String
    NumeroSerie = mNumeroSerie
End Property
Public Property Let NumeroSerie(ByVal NewValue As String)
    mNumeroSerie = NewValue
End Property

Public Property Get NumeroDocumento() As String
    NumeroDocumento = mNumeroDocumento
End Property
Public Property Let NumeroDocumento(ByVal NewValue As String)
    mNumeroDocumento = NewValue
End Property

Public Property Get IdResponsable() As Long
    IdResponsable = mIdResponsable
End Property
Public Property Let IdResponsable(ByVal NewValue As Long)
    mIdResponsable = NewValue
End Property

Public Property Get Responsable() As String
    Responsable = mResponsable
End Property
Public Property Let Responsable(ByVal NewValue As String)
    mResponsable = NewValue
End Property

Public Property Get IdTipoDocumento() As Long
    IdTipoDocumento = mIdTipoDocumento
End Property
Public Property Let IdTipoDocumento(ByVal NewValue As Long)
    mIdTipoDocumento = NewValue
End Property

Public Property Get IdEstado() As Long
    IdEstado = mIdEstado
End Property
Public Property Let IdEstado(ByVal NewValue As Long)
    mIdEstado = NewValue
End Property

Public Property Get IdTipoInventario() As Long
    IdTipoInventario = mIdTipoInventario
End Property
Public Property Let IdTipoInventario(ByVal NewValue As Long)
    mIdTipoInventario = NewValue
End Property

Public Property Get TipoInventario() As String
    TipoInventario = mTipoInventario
End Property
Public Property Let TipoInventario(ByVal NewValue As String)
    mTipoInventario = NewValue
End Property

Public Property Get IdTipoFiltro() As Long
    IdTipoFiltro = mIdTipoFiltro
End Property
Public Property Let IdTipoFiltro(ByVal NewValue As Long)
    mIdTipoFiltro = NewValue
End Property

Public Property Get TipoFiltro() As String
    TipoFiltro = mTipoFiltro
End Property
Public Property Let TipoFiltro(ByVal NewValue As String)
    mTipoFiltro = NewValue
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

Public Property Get InventarioDetS() As LEInventarioDet
    Set InventarioDetS = mInventarioDetS
End Property
Public Property Set InventarioDetS(ByVal NewValue As LEInventarioDet)
    Set mInventarioDetS = NewValue
End Property

Public Property Get AnhoTrabajo() As Long
    AnhoTrabajo = mAnhoTrabajo
End Property
Public Property Let AnhoTrabajo(ByVal NewValue As Long)
    mAnhoTrabajo = NewValue
End Property

Public Function Fetch(IdInventario As String) As EInventario

End Function

Public Sub SetEntity(Record As ADODB.Recordset)
    Dim M As New AlmacenEntidad.LEInventarioDet
    
'    Record.MoveFirst
'    Me.IdInventario = F.NuloNumeric(Record("IdInventario"))
'    Me.FechaInventario = F.ConvertirNumeroAFecha(F.NuloNumeric(Record("FechaInventario")))
'    Me.NumeroSerie = F.NuloString(Record("numser"))
'    Me.NumeroDocumento = F.NuloString(Record("numdoc"))
'    Me.IdResponsable = F.NuloNumeric(Record("idresponsable"))
'    Me.Responsable = F.NuloString(Record("responsable"))
'    Me.IdTipoDocumento = F.NuloNumeric(Record("idtipdoc"))
'    Me.IdEstado = F.NuloNumeric(Record("idestado"))
'    Me.IdEstado = F.NuloString(Record("idestado"))
'    Me.IdAlmacenOrigen = F.NuloNumeric(Record("idalmorig"))
'    Me.AlmacenOrigen = F.NuloString(Record("almorig"))
'    Me.IdAlmacenDestino = F.NuloNumeric(Record("idalmdest"))
'    Me.AlmacenDestino = F.NuloString(Record("almdest"))
'    Me.NumeroSerieOrigen = F.NuloString(Record("numserorig"))
'    Me.NumeroDocumentoOrigen = F.NuloString(Record("numdocorig"))
'    Me.NumeroSerieDestino = F.NuloString(Record("numserdest"))
'    Me.NumeroDocumentoDestino = F.NuloString(Record("numdocdest"))
'    Me.MesTrabajo = F.NuloNumeric(Record("idmes"))
'    Me.AnhoTrabajo = F.NuloNumeric(Record("ano"))
    Set Record = Nothing
End Sub

    Public Function Save(IdUsuario As Long, MachineName As String) As Boolean
    On Error GoTo BLoqueError
    If (F.NuloNumeric(Me.IdInventario) = 0) Then
        Insert IdUsuario, MachineName
    Else
        Update IdUsuario, MachineName
    End If
    Save = True
    Exit Function
    
BLoqueError:
    Save = False
    F.MostrarMensajeError "Error al Grabar :" & Err.Description, Err.Source, ""
End Function

Private Sub Insert(IdUsuario As Long, MachineName As String)
    Dim xId As Long
    Dim dataBase As New SistemaData.EDataBase
    Dim FData As New SistemaData.Funciones
    Dim F As New SistemaLogica.Funciones
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    xId = FData.GetCodeTable("alm_tomainventario", "idtomainventario", mConexion)
    Me.IdInventario = xId
                
    dataBase.CommandText = "INSERT INTO alm_tomainventario(idtomainventario, idalm, idresponsable, idtipoinventario, " _
                & "idestadoinventario, numser, numdoc, nombre, descripcion, fchinv, fchvig, tipofiltro, ano) " _
            + vbCr + "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    
    dataBase.AddParameter "@idtomainventario", adBigInt, adParamInput, xId
    dataBase.AddParameter "@idalm", adBigInt, adParamInput, F.NuloNumeric(Me.IdAlmacen)
    dataBase.AddParameter "@idresponsable", adBigInt, adParamInput, F.NuloNumeric(Me.IdResponsable)
    dataBase.AddParameter "@idtipoinventario", adBigInt, adParamInput, F.NuloNumeric(Me.IdTipoInventario)
    dataBase.AddParameter "@idestadoinventario", adBigInt, adParamInput, F.NuloString(Me.IdEstado)
    dataBase.AddParameter "@numser", adVarChar, adParamInput, F.NuloString(Me.NumeroSerie)
    dataBase.AddParameter "@numdoc", adVarChar, adParamInput, F.NuloString(Me.NumeroDocumento)
    dataBase.AddParameter "@nombre", adVarChar, adParamInput, F.NuloString(Me.Descripcion)
    dataBase.AddParameter "@descripcion", adVarChar, adParamInput, F.NuloString(Me.Glosa)
    dataBase.AddParameter "@fchinv", adDate, adParamInput, Me.FechaInventario
    dataBase.AddParameter "@fchvig", adDate, adParamInput, Me.FechaVigencia
    dataBase.AddParameter "@tipofiltro", adBigInt, adParamInput, F.NuloNumeric(Me.IdTipoFiltro)
    dataBase.AddParameter "@ano", adBigInt, adParamInput, F.NuloNumeric(Me.AnhoTrabajo)
    dataBase.Execute
    
    Dim InvDet As EInventarioDet
    For Each InvDet In Me.InventarioDetS
        If Not mConexion Is Nothing Then
            Set InvDet.Conexion = mConexion
            InvDet.Called = True
        End If
        InvDet.IdInventario = xId
        InvDet.Save IdUsuario, MachineName
    Next
    
    If (Me.IdEstado = F.NuloNumeric(F.KeyValue("EstadoAprobadoInventarioInicial", mConexion))) Then
        If (Me.IdTipoInventario = F.NuloNumeric(F.KeyValue("InventarioAjuste", mConexion))) Then
            ' Agregamos los Movimientos
            InsertMovimientos IdUsuario, MachineName
        ElseIf (Me.IdTipoInventario = F.NuloNumeric(F.KeyValue("InventarioInicial", mConexion))) Then
            ' Agregamos inventario inicial
            GrabarInventarioInicial IdUsuario, MachineName
        End If
    End If
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "Insert - Inventario :" + Trim(Err.Description)
End Sub

Private Sub GrabarInventarioInicial(IdUsuario As Long, MachineName As String)
    Dim A As Long
    Dim InvInicial As New EInventarioIni
    Dim xId As Long
    
    xId = F.NuloNumeric(F.BuscaCodigoTabla(Me.IdInventario, "iddocref", "idinventarioinicial", "alm_inventarioinicial", "N", mConexion))
    
    If xId > 0 Then InvInicial.IdInventarioInicial = xId
    
    ' Se carga la cabecera
    InvInicial.IdAlmacen = Me.IdAlmacen
    InvInicial.IdMoneda = F.NuloNumeric(F.KeyValue("IdMonedaSoles", mConexion))
    InvInicial.FechaVigencia = Me.FechaVigencia
    InvInicial.IdTipoDocumentoReferencia = F.NuloNumeric(F.KeyValue("TomaInventario", mConexion))
    InvInicial.IdDocumentoReferencia = Me.IdInventario
    InvInicial.IdEstado = Me.IdEstado
    ' Se llenan los detalles
    Dim InvDet As EInventarioDet
    For Each InvDet In Me.InventarioDetS
        Dim E As New AlmacenEntidad.EInventarioIniDet
        E.IdItem = F.NuloNumeric(InvDet.IdItem)
        E.IdUnidadMedida = F.NuloNumeric(InvDet.IdUnidadMedida)
        E.Cantidad = F.NuloNumeric(InvDet.CantidadCarga)
        E.Costo = F.NuloNumeric(InvDet.CostoCarga)
        InvInicial.LInventarioIniDet.Add E
        Set E = Nothing
    Next
    
    If Not mConexion Is Nothing Then
        Set InvInicial.Conexion = mConexion
        InvInicial.Called = True
    End If
    ' Se graba el inventario inicial
    InvInicial.Save IdUsuario, MachineName
End Sub

Private Sub InsertMovimientos(IdUsuario As Long, MachineName As String)
    Dim MovIng As New AlmacenEntidad.EMovimiento
    Dim MovSal As New AlmacenEntidad.EMovimiento
    Dim MovAjuIng As New AlmacenEntidad.EMovimiento
    Dim MovAjuSal As New AlmacenEntidad.EMovimiento
    Dim cSQL As String
        
    ' Salida
    MovSal.IdTipoDocumento = 0
    MovSal.FechaMovimiento = Me.FechaVigencia
    MovSal.NumeroSerie = Me.NumeroSerie
    MovSal.NumeroDocumento = F.HallaNumeroDocumento("alm_ingreso", "'" & Me.NumeroSerie & "'", "numser", mConexion)
    MovSal.IdResponsable = Me.IdResponsable
    MovSal.IdProveedor = 0
    MovSal.Proveedor = ""
    MovSal.IdEstado = 2
    MovSal.IdTipoMovimiento = 0
    MovSal.IdTipoDocumentoReferencia = F.NuloNumeric(F.KeyValue("TomaInventario", mConexion))
    MovSal.IdDocumentoReferencia = Me.IdInventario
    MovSal.DocumentoReferencia = Me.NumeroSerie & " - " & Me.NumeroDocumento
    MovSal.IdAlmacen = Me.IdAlmacen
    MovSal.Glosa = Me.Glosa
    MovSal.AnhoTrabajo = Me.AnhoTrabajo
    MovSal.MesTrabajo = Month(Me.FechaVigencia)
    ' Ingreso
    MovIng.IdTipoDocumento = 0
    MovIng.FechaMovimiento = Me.FechaVigencia
    MovIng.NumeroSerie = Me.NumeroSerie
    MovIng.NumeroDocumento = Format(F.NuloNumeric(MovSal.NumeroDocumento) + 1, "0000000000")
    MovIng.IdResponsable = Me.IdResponsable
    MovIng.IdProveedor = 0
    MovIng.Proveedor = ""
    MovIng.IdEstado = 2
    MovIng.IdTipoMovimiento = -1
    MovIng.IdTipoDocumentoReferencia = F.NuloNumeric(F.KeyValue("TomaInventario", mConexion))
    MovIng.IdDocumentoReferencia = Me.IdInventario
    MovIng.DocumentoReferencia = Me.NumeroSerie & " - " & Me.NumeroDocumento
    MovIng.IdAlmacen = Me.IdAlmacen
    MovIng.Glosa = Me.Glosa
    MovIng.AnhoTrabajo = Me.AnhoTrabajo
    MovIng.MesTrabajo = Month(Me.FechaVigencia)
    ' Detalles
    Dim InvDet As EInventarioDet
    For Each InvDet In Me.InventarioDetS
        Dim MovSalDet As New EMovimientoDet
        Dim MovIngDet As New EMovimientoDet
        
        If (InvDet.CantidadCarga - InvDet.CantidadInicial) > 0 Then ' Ingreso
            MovIngDet.IdItem = InvDet.IdItem
            MovIngDet.Cantidad = Abs(InvDet.CantidadCarga - InvDet.CantidadInicial)
            MovIngDet.CantidadTeorica = Abs(InvDet.CantidadCarga - InvDet.CantidadInicial)
            MovIng.LMovimientoDet.Add MovIngDet
            Set MovIngDet = Nothing
        ElseIf (InvDet.CantidadCarga - InvDet.CantidadInicial) < 0 Then ' Salida
            MovSalDet.IdItem = InvDet.IdItem
            MovSalDet.Cantidad = Abs(InvDet.CantidadCarga - InvDet.CantidadInicial)
            MovSalDet.CantidadTeorica = Abs(InvDet.CantidadCarga - InvDet.CantidadInicial)
            MovSal.LMovimientoDet.Add MovSalDet
            Set MovSalDet = Nothing
        End If
    Next
        
    If Not mConexion Is Nothing Then
        Set MovSal.Conexion = mConexion
        MovSal.Called = True
        Set MovIng.Conexion = mConexion
        MovIng.Called = True
    End If
    
    ' Grabamos los movimientos
    If (MovSal.LMovimientoDet.Count > 0) Then MovSal.Save IdUsuario, MachineName
    If (MovIng.LMovimientoDet.Count > 0) Then MovIng.Save IdUsuario, MachineName
End Sub

Private Sub Update(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    dataBase.CommandText = "UPDATE alm_tomainventario SET idalm = ?, idresponsable = ?, idtipoinventario = ?, " _
                & "idestadoinventario = ?, numser = ?, numdoc = ?, nombre = ?, descripcion = ?, " _
                & "fchinv = ?, fchvig = ?, tipofiltro = ?, ano = ? " _
            + vbCr + "WHERE (alm_tomainventario.idtomainventario = ?)"
    
    dataBase.AddParameter "@idalm", adBigInt, adParamInput, F.NuloNumeric(Me.IdAlmacen)
    dataBase.AddParameter "@idresponsable", adBigInt, adParamInput, F.NuloNumeric(Me.IdResponsable)
    dataBase.AddParameter "@idtipoinventario", adBigInt, adParamInput, F.NuloNumeric(Me.IdTipoInventario)
    dataBase.AddParameter "@idestadoinventario", adBigInt, adParamInput, F.NuloString(Me.IdEstado)
    dataBase.AddParameter "@numser", adVarChar, adParamInput, F.NuloString(Me.NumeroSerie)
    dataBase.AddParameter "@numdoc", adVarChar, adParamInput, F.NuloString(Me.NumeroDocumento)
    dataBase.AddParameter "@nombre", adVarChar, adParamInput, F.NuloString(Me.Descripcion)
    dataBase.AddParameter "@descripcion", adVarChar, adParamInput, F.NuloString(Me.Glosa)
    dataBase.AddParameter "@fchinv", adDate, adParamInput, Me.FechaInventario
    dataBase.AddParameter "@fchvig", adDate, adParamInput, Me.FechaVigencia
    dataBase.AddParameter "@tipofiltro", adBigInt, adParamInput, F.NuloNumeric(Me.IdTipoFiltro)
    dataBase.AddParameter "@ano", adBigInt, adParamInput, F.NuloNumeric(Me.AnhoTrabajo)
    dataBase.AddParameter "@idtomainventario", adBigInt, adParamInput, Me.IdInventario
    dataBase.Execute
      
    ' Borramos Hijos
    DeleteChild IdUsuario, MachineName
    ' Grabamos Hijos
    Dim InvDet As EInventarioDet
    For Each InvDet In Me.InventarioDetS
        If Not mConexion Is Nothing Then
            Set InvDet.Conexion = mConexion
            InvDet.Called = True
        End If
        InvDet.IdInventarioDet = 0
        InvDet.IdInventario = Me.IdInventario
        InvDet.Save IdUsuario, MachineName
    Next
    
    ' Estado Aprobado
    If (Me.IdEstado = F.NuloNumeric(F.KeyValue("EstadoAprobadoInventario", mConexion))) Then
        If (Me.IdTipoInventario = F.NuloNumeric(F.KeyValue("InventarioAjuste", mConexion))) Then
            ' Agregamos los Movimientos
            DeleteMovimientos IdUsuario, MachineName
            InsertMovimientos IdUsuario, MachineName
        ElseIf (Me.IdTipoInventario = F.NuloNumeric(F.KeyValue("InventarioInicial", mConexion))) Then
            ' Agregamos inventario inicial
            GrabarInventarioInicial IdUsuario, MachineName
        End If
    End If
    ' Estado Anulado
    If (Me.IdEstado = F.NuloNumeric(F.KeyValue("EstadoAnuladoInventario", mConexion))) Then
        If (Me.IdTipoInventario = F.NuloNumeric(F.KeyValue("InventarioAjuste", mConexion))) Then
            DeleteMovimientos IdUsuario, MachineName
        ElseIf (Me.IdTipoInventario = F.NuloNumeric(F.KeyValue("InventarioInicial", mConexion))) Then
            ' Grabamos el inventario inicial
            GrabarInventarioInicial IdUsuario, MachineName
        End If
    End If
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "Update - Inventario :" + Trim(Err.Description)
    Exit Sub
End Sub

Public Function Delete(IdUsuario As Long, MachineName As String) As Boolean
    Dim xId As Double
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    ' Eliminamos Movimientos o inventario inicial
    If (Me.IdEstado = F.NuloNumeric(F.KeyValue("EstadoAprobadoInventario", mConexion))) Then
        If (Me.IdTipoInventario = F.NuloNumeric(F.KeyValue("InventarioAjuste", mConexion))) Then
            ' Eliminamos los movimientos
            DeleteMovimientos IdUsuario, MachineName
        ElseIf (Me.IdTipoInventario = F.NuloNumeric(F.KeyValue("InventarioInicial", mConexion))) Then
            ' Eliminamos inventario inicial
            DeleteInventarioInicial IdUsuario, MachineName
        End If
    End If
        
    ' Eliminamos a los Hijos
    DeleteChild IdUsuario, MachineName
    
    dataBase.ClearParameter
    dataBase.CommandText = "DELETE FROM alm_tomainventario " _
            + vbCr + "WHERE alm_tomainventario.idtomainventario = ?"
    
    dataBase.AddParameter "@idtomainventario", adBigInt, adParamInput, Me.IdInventario
    dataBase.Execute
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Delete = True
    Exit Function

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    F.MostrarMensajeError Err.Description & " - " & Err.Source, "[Delete]"
    Delete = False
End Function

Public Function DeleteMovimientos(IdUsuario As Long, MachineName As String) As Boolean
    Dim dataBase As New SistemaData.EDataBase
    Dim xRs As New ADODB.Recordset
    
On Error GoTo BLoqueError
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    ' Encontramos los movimientos relacionados
    dataBase.ClearParameter
    
    dataBase.CommandText = "SELECT alm_ingreso.id  FROM alm_ingreso " _
            + vbCr + "WHERE alm_ingreso.idtipdocref = ? AND alm_ingreso.iddocref = ?"
            
    dataBase.AddParameter "@idtipdocref", adBigInt, adParamInput, F.NuloNumeric(F.KeyValue("TomaInventario", mConexion))
    dataBase.AddParameter "@iddocref", adBigInt, adParamInput, Me.IdInventario
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
    DeleteMovimientos = True
    Exit Function

BLoqueError:
    Set dataBase = Nothing
    Set xRs = Nothing
    F.MostrarMensajeError Err.Description, Err.Source, "[DeleteMovimientos]"
    DeleteMovimientos = False
End Function


Public Sub DeleteInventarioInicial(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    Dim xRs As New ADODB.Recordset
    
On Error GoTo BLoqueError
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    ' Encontramos los inventarios iniciales relacionados
    dataBase.ClearParameter
    
    dataBase.CommandText = "SELECT alm_inventarioinicial.id  FROM alm_inventarioinicial " _
            + vbCr + "WHERE alm_inventarioinicial.idtipdocref = ? AND alm_inventarioinicial.iddocref = ?"
            
    dataBase.AddParameter "@idtipdocref", adBigInt, adParamInput, F.NuloNumeric(F.KeyValue("IdDocumentoAjusteInventario", mConexion))
    dataBase.AddParameter "@iddocref", adBigInt, adParamInput, Me.IdInventario
    Set xRs = dataBase.GetRecordset
    If xRs.State = 0 Then GoTo BLoqueError
    If xRs.RecordCount > 0 Then
        xRs.MoveFirst
        While Not xRs.EOF
            Dim MovObj As New AlmacenEntidad.EInventarioIni
            MovObj.IdInventarioInicial = xRs("id")
            ' Eliminamos los inventarios iniciales
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

    dataBase.CommandText = "DELETE FROM alm_tomainventariodet " _
            + vbCr + "WHERE alm_tomainventariodet.idtomainventario = ?"

    dataBase.AddParameter "@idtomainventario", adBigInt, adParamInput, Me.IdInventario
    dataBase.Execute
    
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el Inventario por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub





VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ETransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdTransferencia As Long
Private mFechaTransferencia As Date
Private mNumeroSerie As String
Private mNumeroDocumento As String
Private mIdResponsable As Long
Private mResponsable As String
Private mIdTipoDocumento As Long
Private mIdEstado As Long
Private mEstado As String
Private mIdAlmacenOrigen As Long
Private mAlmacenOrigen As String
Private mNumeroSerieOrigen As String
Private mNumeroDocumentoOrigen As String
Private mIdAlmacenDestino As Long
Private mAlmacenDestino As String
Private mNumeroSerieDestino As String
Private mNumeroDocumentoDestino As String
Private mGlosa As String
Private mTransferenciaDetS As New LETransferenciaDet

Private mAnhoTrabajo As Long
Private mMesTrabajo As Long
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

Public Property Get IdTransferencia() As Long
    IdTransferencia = mIdTransferencia
End Property
Public Property Let IdTransferencia(ByVal NewValue As Long)
    mIdTransferencia = NewValue
End Property

Public Property Get FechaTransferencia() As Date
    FechaTransferencia = mFechaTransferencia
End Property
Public Property Let FechaTransferencia(ByVal NewValue As Date)
    mFechaTransferencia = NewValue
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

Public Property Get IdAlmacenOrigen() As Long
    IdAlmacenOrigen = mIdAlmacenOrigen
End Property
Public Property Let IdAlmacenOrigen(ByVal NewValue As Long)
    mIdAlmacenOrigen = NewValue
End Property

Public Property Get AlmacenOrigen() As String
    AlmacenOrigen = mAlmacenOrigen
End Property
Public Property Let AlmacenOrigen(ByVal NewValue As String)
    mAlmacenOrigen = NewValue
End Property

Public Property Get IdAlmacenDestino() As Long
    IdAlmacenDestino = mIdAlmacenDestino
End Property
Public Property Let IdAlmacenDestino(ByVal NewValue As Long)
    mIdAlmacenDestino = NewValue
End Property

Public Property Get AlmacenDestino() As String
    AlmacenDestino = mAlmacenDestino
End Property
Public Property Let AlmacenDestino(ByVal NewValue As String)
    mAlmacenDestino = NewValue
End Property

Public Property Get NumeroSerieOrigen() As String
    NumeroSerieOrigen = mNumeroSerieOrigen
End Property
Public Property Let NumeroSerieOrigen(ByVal NewValue As String)
    mNumeroSerieOrigen = NewValue
End Property

Public Property Get NumeroDocumentoOrigen() As String
    NumeroDocumentoOrigen = mNumeroDocumentoOrigen
End Property
Public Property Let NumeroDocumentoOrigen(ByVal NewValue As String)
    mNumeroDocumentoOrigen = NewValue
End Property

Public Property Get NumeroSerieDestino() As String
    NumeroSerieDestino = mNumeroSerieDestino
End Property
Public Property Let NumeroSerieDestino(ByVal NewValue As String)
    mNumeroSerieDestino = NewValue
End Property

Public Property Get NumeroDocumentoDestino() As String
    NumeroDocumentoDestino = mNumeroDocumentoDestino
End Property
Public Property Let NumeroDocumentoDestino(ByVal NewValue As String)
    mNumeroDocumentoDestino = NewValue
End Property

Public Property Get Glosa() As String
    Glosa = mGlosa
End Property
Public Property Let Glosa(ByVal NewValue As String)
    mGlosa = NewValue
End Property

Public Property Get TransferenciaDetS() As LETransferenciaDet
    Set TransferenciaDetS = mTransferenciaDetS
End Property
Public Property Set TransferenciaDetS(ByVal NewValue As LETransferenciaDet)
    Set mTransferenciaDetS = NewValue
End Property

Public Property Get AnhoTrabajo() As Long
    AnhoTrabajo = mAnhoTrabajo
End Property
Public Property Let AnhoTrabajo(ByVal NewValue As Long)
    mAnhoTrabajo = NewValue
End Property

Public Property Get MesTrabajo() As Long
    MesTrabajo = mMesTrabajo
End Property
Public Property Let MesTrabajo(ByVal NewValue As Long)
    mMesTrabajo = NewValue
End Property

Public Function Fetch(IdTransferencia As String) As EInventario

End Function

Public Sub SetEntity(Record As ADODB.Recordset)
    Dim M As New AlmacenEntidad.LETransferenciaDet
    
    Record.MoveFirst
    Me.IdTransferencia = F.NuloNumeric(Record("IdTransferencia"))
    Me.FechaTransferencia = F.ConvertirNumeroAFecha(F.NuloNumeric(Record("FechaTransferencia")))
    Me.NumeroSerie = F.NuloString(Record("numser"))
    Me.NumeroDocumento = F.NuloString(Record("numdoc"))
    Me.IdResponsable = F.NuloNumeric(Record("idresponsable"))
    Me.Responsable = F.NuloString(Record("responsable"))
    Me.IdTipoDocumento = F.NuloNumeric(Record("idtipdoc"))
    Me.IdEstado = F.NuloNumeric(Record("idestado"))
    Me.IdEstado = F.NuloString(Record("idestado"))
    Me.IdAlmacenOrigen = F.NuloNumeric(Record("idalmorig"))
    Me.AlmacenOrigen = F.NuloString(Record("almorig"))
    Me.IdAlmacenDestino = F.NuloNumeric(Record("idalmdest"))
    Me.AlmacenDestino = F.NuloString(Record("almdest"))
    Me.NumeroSerieOrigen = F.NuloString(Record("numserorig"))
    Me.NumeroDocumentoOrigen = F.NuloString(Record("numdocorig"))
    Me.NumeroSerieDestino = F.NuloString(Record("numserdest"))
    Me.NumeroDocumentoDestino = F.NuloString(Record("numdocdest"))
    Me.MesTrabajo = F.NuloNumeric(Record("idmes"))
    Me.AnhoTrabajo = F.NuloNumeric(Record("ano"))
    Set Record = Nothing
End Sub

Public Sub Save(IdUsuario As Long, MachineName As String)
On Error GoTo BLoqueError
    If (F.NuloNumeric(Me.IdTransferencia) = 0) Then
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
    
    xId = FData.GetCodeTable("alm_transferencia", "idtransferencia", mConexion)
    Me.IdTransferencia = xId
                
    dataBase.CommandText = "INSERT INTO alm_transferencia(idtransferencia, fchdoc, idtipdoc, numser, " _
                & "numdoc, idresponsable, numserorig, numdocorig, idalmorig, numserdest, numdocdest, idalmdest, " _
                & "ano, idmes, idestado, glosa) " _
            + vbCr + "VALUES(?, ?, ?, ?, " _
                & "?, ?, ?, ?, ?, ?, ?, ?, " _
                & "?, ?, ?, ?)"
    
    dataBase.AddParameter "@idtransferencia", adInteger, adParamInput, xId
    dataBase.AddParameter "@fchdoc", adDate, adParamInput, Me.FechaTransferencia
    dataBase.AddParameter "@idtipdoc", adInteger, adParamInput, F.NuloNumeric(Me.IdTipoDocumento)
    dataBase.AddParameter "@numser", adVarChar, adParamInput, F.NuloString(Me.NumeroSerie)
    dataBase.AddParameter "@numdoc", adVarChar, adParamInput, F.NuloString(Me.NumeroDocumento)
    dataBase.AddParameter "@idresponsable", adInteger, adParamInput, F.NuloNumeric(Me.IdResponsable)
    dataBase.AddParameter "@numserorig", adVarChar, adParamInput, F.NuloString(Me.NumeroSerieOrigen)
    dataBase.AddParameter "@numdocorig", adVarChar, adParamInput, F.NuloString(Me.NumeroDocumentoOrigen)
    dataBase.AddParameter "@idalmorig", adInteger, adParamInput, F.NuloNumeric(Me.IdAlmacenOrigen)
    dataBase.AddParameter "@numserdest", adVarChar, adParamInput, F.NuloString(Me.NumeroSerieDestino)
    dataBase.AddParameter "@numdocdest", adVarChar, adParamInput, F.NuloString(Me.NumeroDocumentoDestino)
    dataBase.AddParameter "@idalmdest", adInteger, adParamInput, F.NuloNumeric(Me.IdAlmacenDestino)
    dataBase.AddParameter "@ano", adInteger, adParamInput, F.NuloNumeric(Me.AnhoTrabajo)
    dataBase.AddParameter "@idmes", adInteger, adParamInput, F.NuloNumeric(Me.MesTrabajo)
    dataBase.AddParameter "@idestado", adInteger, adParamInput, F.NuloNumeric(Me.IdEstado)
    dataBase.AddParameter "@glosa", adVarChar, adParamInput, F.NuloString(Me.Glosa)
    dataBase.Execute
    
    Dim TransDet As ETransferenciaDet
    For Each TransDet In Me.TransferenciaDetS
        If Not mConexion Is Nothing Then
            Set TransDet.Conexion = mConexion
            TransDet.Called = True
        End If
        TransDet.IdTransferencia = xId
        TransDet.Save IdUsuario, MachineName
    Next
    
    ' Agregamos los Movimientos
    InsertMovimientos IdUsuario, MachineName
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "Insert - Transferencia :" + Trim(Err.Description)
End Sub

Private Sub InsertMovimientos(IdUsuario As Long, MachineName As String)
    Dim MovIng As New AlmacenEntidad.EMovimiento
    Dim MovSal As New AlmacenEntidad.EMovimiento

    ' Salida
    MovSal.IdTipoDocumento = 0
    MovSal.FechaMovimiento = Me.FechaTransferencia
    MovSal.NumeroSerie = Me.NumeroSerieOrigen
    MovSal.NumeroDocumento = Me.NumeroDocumentoOrigen
    MovSal.IdResponsable = Me.IdResponsable
    MovSal.IdProveedor = 0
    MovSal.Proveedor = ""
    MovSal.IdEstado = 2
    MovSal.IdTipoMovimiento = 0
    MovSal.IdTipoDocumentoReferencia = 119
    MovSal.IdDocumentoReferencia = Me.IdTransferencia
    MovSal.DocumentoReferencia = Me.NumeroSerie & " - " & Me.NumeroDocumento
    MovSal.IdAlmacen = Me.IdAlmacenOrigen
    MovSal.Glosa = Me.Glosa
    MovSal.AnhoTrabajo = Me.AnhoTrabajo
    MovSal.MesTrabajo = Me.MesTrabajo
    
    Dim TransDetMovSal As ETransferenciaDet
    For Each TransDetMovSal In Me.TransferenciaDetS
        Dim MovSalDet As New EMovimientoDet
        
        MovSalDet.IdItem = TransDetMovSal.IdItem
        MovSalDet.Cantidad = TransDetMovSal.Cantidad
        MovSalDet.CantidadTeorica = TransDetMovSal.Cantidad
        MovSal.LMovimientoDet.Add MovSalDet
        Set MovSalDet = Nothing
    Next
    
    ' Ingreso
    MovIng.IdTipoDocumento = 0
    MovIng.FechaMovimiento = Me.FechaTransferencia
    MovIng.NumeroSerie = Me.NumeroSerieDestino
    MovIng.NumeroDocumento = Me.NumeroDocumentoDestino
    MovIng.IdResponsable = Me.IdResponsable
    MovIng.IdProveedor = 0
    MovIng.Proveedor = ""
    MovIng.IdEstado = 2
    MovIng.IdTipoMovimiento = -1
    MovIng.IdTipoDocumentoReferencia = 119
    MovIng.IdDocumentoReferencia = Me.IdTransferencia
    MovIng.DocumentoReferencia = Me.NumeroSerie & " - " & Me.NumeroDocumento
    MovIng.IdAlmacen = Me.IdAlmacenDestino
    MovIng.Glosa = Me.Glosa
    MovIng.AnhoTrabajo = Me.AnhoTrabajo
    MovIng.MesTrabajo = Me.MesTrabajo
    
    Dim TransDetMovIng As ETransferenciaDet
    For Each TransDetMovIng In Me.TransferenciaDetS
        Dim MovIngDet As New EMovimientoDet
        
        MovIngDet.IdItem = TransDetMovIng.IdItem
        MovIngDet.Cantidad = TransDetMovIng.Cantidad
        MovIngDet.CantidadTeorica = TransDetMovIng.Cantidad
        MovIng.LMovimientoDet.Add MovIngDet
        Set MovIngDet = Nothing
    Next
    
    If Not mConexion Is Nothing Then
        Set MovSal.Conexion = mConexion
        MovSal.Called = True
        Set MovIng.Conexion = mConexion
        MovIng.Called = True
    End If
    ' Grabamos los movimientos
    MovSal.Save IdUsuario, MachineName
    MovIng.Save IdUsuario, MachineName
End Sub

Private Sub Update(IdUsuario As Long, MachineName As String)
    Dim xId As Long
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    xId = Me.IdTransferencia '
    
    dataBase.CommandText = "UPDATE alm_transferencia SET fchdoc = ?, idtipdoc = ?, numser = ?, " _
                & "numdoc = ?, idresponsable = ?, numserorig = ?, numdocorig = ?, idalmorig = ?, " _
                & "numserdest = @numserdest, numdocdest = @numdocdest, idalmdest = @idalmdest, " _
                & "ano = ?, idmes = ?, idestado = ?, glosa = ? " _
            + vbCr + "WHERE (alm_transferencia.idtransferencia = ?)"
    
    dataBase.AddParameter "@fchdoc", adDate, adParamInput, Me.FechaTransferencia
    dataBase.AddParameter "@idtipdoc", adInteger, adParamInput, F.NuloNumeric(Me.IdTipoDocumento)
    dataBase.AddParameter "@numser", adVarChar, adParamInput, F.NuloString(Me.NumeroSerie), 4
    dataBase.AddParameter "@numdoc", adVarChar, adParamInput, F.NuloString(Me.NumeroDocumento), 10
    dataBase.AddParameter "@idresponsable", adInteger, adParamInput, F.NuloNumeric(Me.IdResponsable)
    dataBase.AddParameter "@numserorig", adVarChar, adParamInput, F.NuloString(Me.NumeroSerieOrigen), 4
    dataBase.AddParameter "@numdocorig", adVarChar, adParamInput, F.NuloString(Me.NumeroDocumentoOrigen), 10
    dataBase.AddParameter "@idalmorig", adInteger, adParamInput, F.NuloNumeric(Me.IdAlmacenOrigen)
    dataBase.AddParameter "@numserdest", adVarChar, adParamInput, F.NuloString(Me.NumeroSerieDestino), 4
    dataBase.AddParameter "@numdocdest", adVarChar, adParamInput, F.NuloString(Me.NumeroDocumentoDestino), 10
    dataBase.AddParameter "@idalmdest", adInteger, adParamInput, F.NuloNumeric(Me.IdAlmacenDestino)
    dataBase.AddParameter "@ano", adInteger, adParamInput, F.NuloNumeric(Me.AnhoTrabajo)
    dataBase.AddParameter "@idmes", adInteger, adParamInput, F.NuloNumeric(Me.MesTrabajo)
    dataBase.AddParameter "@idestado", adInteger, adParamInput, F.NuloNumeric(Me.IdEstado)
    dataBase.AddParameter "@glosa", adVarChar, adParamInput, F.NuloString(Me.Glosa), 200
    dataBase.AddParameter "@idtransferencia", adInteger, adParamInput, xId
    dataBase.Execute
        
    ' Borramos Hijos
    DeleteChild IdUsuario, MachineName
    ' Grabamos Hijos
    Dim TransDet As ETransferenciaDet
    For Each TransDet In Me.TransferenciaDetS
        If Not mConexion Is Nothing Then
            Set TransDet.Conexion = mConexion
            TransDet.Called = True
        End If
        TransDet.IdTransferenciaDet = 0
        TransDet.IdTransferencia = Me.IdTransferencia
        TransDet.Save IdUsuario, MachineName
    Next
    
    ' Grabamos los movimientos
    DeleteMovimientos IdUsuario, MachineName
    InsertMovimientos IdUsuario, MachineName
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "Update - Transferencia :" + Trim(Err.Description)
    Exit Sub
End Sub

Public Sub Delete(IdUsuario As Long, MachineName As String)
    Dim xId As Double
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    ' Eliminamos los movimientos
    DeleteMovimientos IdUsuario, MachineName
    ' Eliminamos a los Hijos
    DeleteChild IdUsuario, MachineName
    
    dataBase.ClearParameter
    dataBase.CommandText = "DELETE FROM alm_transferencia " _
            + vbCr + "WHERE alm_transferencia.idtransferencia = ?"
    
    dataBase.AddParameter "@idtransferencia", adInteger, adParamInput, Me.IdTransferencia
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

Public Sub DeleteMovimientos(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    Dim xRs As New ADODB.Recordset
    
On Error GoTo BLoqueError
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    ' Encontramos los movimientos relacionados
    dataBase.ClearParameter
    dataBase.CommandText = "SELECT alm_ingreso.id  FROM alm_ingreso " _
            + vbCr + "WHERE alm_ingreso.idtipdocref = ? AND alm_ingreso.iddocref = ?"
            
    dataBase.AddParameter "@idtipdocref", adInteger, adParamInput, F.NuloNumeric(F.KeyValue("Transferencia", mConexion))
    dataBase.AddParameter "@iddocref", adInteger, adParamInput, Me.IdTransferencia
    Set xRs = dataBase.GetRecordset
    If xRs.State = 0 Then GoTo BLoqueError
    If xRs.RecordCount > 0 Then
        xRs.MoveFirst
        While Not xRs.EOF
            Dim MovObj As New AlmacenEntidad.EMovimiento
            MovObj.IdMovimiento = F.NuloNumeric(xRs("id"))
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
    Err.Raise &HFFFFFF01, , "No se pudo registrar el Transferencia por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub

Public Sub DeleteChild(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    
    dataBase.ClearParameter
    dataBase.CommandText = "DELETE FROM alm_transferenciadet " _
            + vbCr + "WHERE alm_transferenciadet.idtransferencia = ?"

    dataBase.AddParameter "@idtransferencia", adInteger, adParamInput, Me.IdTransferencia
    dataBase.Execute
    
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el Transferencia por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub



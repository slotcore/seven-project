VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EMovimientoDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdMovimientoDet As Long
Private mIdMovimiento As Long
Private mIdUnidadMedida As Long
Private mUnidadMedida As String
Private mIdItem As Long
Private mItem As String
Private mCantidad As Double
Private mCantidadTeorica As Double
Private mFechaMovimiento As Date
Private mIdDocumentoReferencia As Long
Private F As New SistemaLogica.Funciones
Private mConexion As New ADODB.Connection
Private mIndex As Long
Private mCalled As Boolean
Private mMarkOld As Boolean
Private mMarkNew As Boolean

Private Sub Class_Initialize()
    mCalled = False
    mMarkOld = False
    mMarkNew = False
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

Public Property Get Index() As Long
    Index = mIndex
End Property
Public Property Let Index(ByVal NewValue As Long)
    mIndex = NewValue
End Property

Public Property Get Called() As Boolean
    Called = mCalled
End Property
Public Property Let Called(ByVal NewValue As Boolean)
    mCalled = NewValue
End Property

Public Property Get MarkOld() As Boolean
    MarkOld = mMarkOld
End Property
Public Property Let MarkOld(ByVal NewValue As Boolean)
    mMarkOld = NewValue
End Property

Public Property Get MarkNew() As Boolean
    MarkNew = mMarkNew
End Property
Public Property Let MarkNew(ByVal NewValue As Boolean)
    mMarkNew = NewValue
End Property

Public Property Get IdMovimientoDet() As Long
    IdMovimientoDet = mIdMovimientoDet
End Property
Public Property Let IdMovimientoDet(ByVal NewValue As Long)
    mIdMovimientoDet = NewValue
End Property

Public Property Get IdMovimiento() As Long
    IdMovimiento = mIdMovimiento
End Property
Public Property Let IdMovimiento(ByVal NewValue As Long)
    mIdMovimiento = NewValue
End Property

Public Property Get IdUnidadMedida() As Long
    IdUnidadMedida = mIdUnidadMedida
End Property
Public Property Let IdUnidadMedida(ByVal NewValue As Long)
    mIdUnidadMedida = NewValue
End Property

Public Property Get UnidadMedida() As String
    UnidadMedida = mUnidadMedida
End Property
Public Property Let UnidadMedida(ByVal NewValue As String)
    mUnidadMedida = NewValue
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

Public Property Get Cantidad() As Double
    Cantidad = mCantidad
End Property
Public Property Let Cantidad(ByVal NewValue As Double)
    mCantidad = NewValue
End Property

Public Property Get CantidadTeorica() As Double
    CantidadTeorica = mCantidadTeorica
End Property
Public Property Let CantidadTeorica(ByVal NewValue As Double)
    mCantidadTeorica = NewValue
End Property

Public Property Get FechaMovimiento() As Date
    FechaMovimiento = mFechaMovimiento
End Property
Public Property Let FechaMovimiento(ByVal NewValue As Date)
    mFechaMovimiento = NewValue
End Property

Public Property Get IdDocumentoReferencia() As Long
    IdDocumentoReferencia = mIdDocumentoReferencia
End Property
Public Property Let IdDocumentoReferencia(ByVal NewValue As Long)
    mIdDocumentoReferencia = NewValue
End Property

Public Sub Fetch(IdMovimientoDetalle As Long)
    Dim dataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
            
    dataBase.CommandText = "SELECT alm_ingresodet.idmovdet, alm_ingresodet.id AS idmov, alm_ingreso.fchdoc AS fchmov, alm_inventario.idunimed, mae_unidades.abrev AS unimed, alm_ingresodet.iditem, alm_inventario.descripcion AS item, alm_ingresodet.cantidad AS canreal, alm_ingresodet.cantteo AS canteo, alm_ingresodet.iddocref " _
        + vbCr + "FROM ((alm_ingresodet INNER JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) INNER JOIN alm_ingreso ON alm_ingresodet.id = alm_ingreso.id " _
        + vbCr + "WHERE (((alm_ingresodet.idmovdet)=" & IdMovimientoDetalle & "))"
    
    Set mRecord = dataBase.GetRecordset
    If mRecord.RecordCount > 0 Then Me.SetEntity mRecord
    Exit Sub
    
BLoqueError:
    Err.Raise Err.Number, "[Fetch] " & Err.Source, Err.Description
End Sub

Public Sub SetEntity(Record As ADODB.Recordset)
    Dim F As New SistemaLogica.Funciones

    Me.IdMovimientoDet = F.NuloNumeric(Record("idmovdet"))
    Me.IdMovimiento = F.NuloNumeric(Record("idmov"))
    Me.IdUnidadMedida = F.NuloNumeric(Record("idunimed"))
    Me.UnidadMedida = F.NuloString(Record("unimed"))
    Me.IdItem = F.NuloNumeric(Record("iditem"))
    Me.Item = F.NuloString(Record("item"))
    Me.Cantidad = F.NuloNumeric(Record("canreal"))
    Me.CantidadTeorica = F.NuloNumeric(Record("canteo"))
    Me.FechaMovimiento = Record("fchmov")
    Me.IdDocumentoReferencia = F.NuloNumeric(Record("iddocref"))
End Sub

Public Function Save(IdUsuario As Long, MachineName As String) As Boolean
On Error GoTo BLoqueError
    If (F.NuloNumeric(Me.IdMovimientoDet) > 0) And (Not Me.MarkNew) Then
        Update IdUsuario, MachineName
    Else
        Insert IdUsuario, MachineName
    End If
    Save = True
    Exit Function
    
BLoqueError:
    Save = False
    F.MostrarMensajeError Err.Description, "Grabar", "[Save] " & Err.Source
End Function

Private Sub Insert(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    Dim FData As New SistemaData.Funciones
    
On Error GoTo BLoqueError
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    If Me.IdMovimientoDet = 0 Then
'        ' Utiliza el gestor de correlativos
'        Dim mCorrelativo As New ECorrelativo
'        Set mCorrelativo.Conexion = mConexion
'        mCorrelativo.GetCorrelativo "ID", IdUsuario, MachineName
'        Me.IdMovimientoDet = mCorrelativo.Correlativo
        
        '************************
        ' Se quita debido a que se encuentran demasiadas concurrencias
        ' De Ids en procesos paralelos
        Me.IdMovimientoDet = FData.GetCodeTable("alm_ingresodet", "idmovdet", mConexion)
        '************************
    End If
    dataBase.ClearParameter

    dataBase.CommandText = "INSERT INTO alm_ingresodet(idmovdet, id, iditem, cantidad, cantteo, iddocref) " _
            + vbCr + "VALUES(?, ?, ?, ?, ?, ?)"

    dataBase.AddParameter "@idmovdet", adInteger, adParamInput, Me.IdMovimientoDet
    dataBase.AddParameter "@id", adInteger, adParamInput, F.NuloNumeric(Me.IdMovimiento)
    dataBase.AddParameter "@iditem", adInteger, adParamInput, F.NuloNumeric(Me.IdItem)
    dataBase.AddParameter "@cantidad", adDouble, adParamInput, F.NuloNumeric(Me.Cantidad)
    dataBase.AddParameter "@cantteo", adDouble, adParamInput, F.NuloNumeric(Me.CantidadTeorica)
    dataBase.AddParameter "@iddocref", adInteger, adParamInput, F.NuloNumeric(Me.IdDocumentoReferencia)
    dataBase.Execute
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el detalle del movimiento por el siguiente motivo :" + Trim(Err.Description)
End Sub

Private Sub Update(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    dataBase.ClearParameter
        
    dataBase.CommandText = "UPDATE alm_ingresodet SET id = ?, iditem = ?, cantidad = ?, cantteo = ?, iddocref = ? " _
            + vbCr + "WHERE(alm_ingresodet.idmovdet = ?)"

    dataBase.AddParameter "@id", adInteger, adParamInput, F.NuloNumeric(Me.IdMovimiento)
    dataBase.AddParameter "@iditem", adInteger, adParamInput, F.NuloNumeric(Me.IdItem)
    dataBase.AddParameter "@cantidad", adDouble, adParamInput, F.NuloNumeric(Me.Cantidad)
    dataBase.AddParameter "@cantteo", adDouble, adParamInput, F.NuloNumeric(Me.CantidadTeorica)
    dataBase.AddParameter "@iddocref", adInteger, adParamInput, F.NuloNumeric(Me.IdDocumentoReferencia)
    dataBase.AddParameter "@idmovdet", adInteger, adParamInput, F.NuloNumeric(Me.IdMovimientoDet)
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

Public Sub Delete(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    dataBase.ClearParameter

    dataBase.CommandText = "DELETE FROM alm_ingresodet " _
            + vbCr + "WHERE alm_ingresodet.idmovdet = ?"
    
    dataBase.AddParameter "@idmovdet", adInteger, adParamInput, Me.IdMovimientoDet
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

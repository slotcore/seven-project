VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EMovimientoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdMovimiento As Long
Private mIdMovimientoDet As Long
Private mIdLibroCostoTemp As Long
Private mIdUnidadMedida As Long
Private mUnidadMedida As String
Private mIdItem As Long
Private mCodigoItem As String
Private mItem As String
Private mIdAlmacen As Long
Private mAlmacen As String
Private mCantidad As Double
Private mIdTipoMovimiento As Long
Private mTipoMovimiento As String
Private mCostoUnitario As Double
Private mCostoUnitarioPromedio As Double
Private mCosto As Double
Private mCostoPrimo As Double
Private mCostoMOD As Double
Private mCostoCIF As Double
Private mNumeroDocumento As String
Private mIdTipoDocumentoReferencia As Long
Private mTipoDocumentoReferencia As String
Private mIdDocumentoReferencia As Long
Private mDocumentoReferencia As String
Private mIdTipoDocumentoAnexado As Long
Private mTipoDocumentoAnexado As String
Private mIdDocumentoAnexado As Long
Private mDocumentoAnexado As String
Private mFechaMovimiento As Date
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

Public Property Get IdMovimiento() As Long
    IdMovimiento = mIdMovimiento
End Property
Public Property Let IdMovimiento(ByVal NewValue As Long)
    mIdMovimiento = NewValue
End Property

Public Property Get IdMovimientoDet() As Long
    IdMovimientoDet = mIdMovimientoDet
End Property
Public Property Let IdMovimientoDet(ByVal NewValue As Long)
    mIdMovimientoDet = NewValue
End Property

Public Property Get IdLibroCostoTemp() As Long
    IdLibroCostoTemp = mIdLibroCostoTemp
End Property
Public Property Let IdLibroCostoTemp(ByVal NewValue As Long)
    mIdLibroCostoTemp = NewValue
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

Public Property Get CodigoItem() As String
    CodigoItem = mCodigoItem
End Property
Public Property Let CodigoItem(ByVal NewValue As String)
    mCodigoItem = NewValue
End Property

Public Property Get Item() As String
    Item = mItem
End Property
Public Property Let Item(ByVal NewValue As String)
    mItem = NewValue
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

Public Property Get Cantidad() As Double
    Cantidad = mCantidad
End Property
Public Property Let Cantidad(ByVal NewValue As Double)
    mCantidad = NewValue
End Property

Public Property Get IdTipoMovimiento() As Long
    IdTipoMovimiento = mIdTipoMovimiento
End Property
Public Property Let IdTipoMovimiento(ByVal NewValue As Long)
    mIdTipoMovimiento = NewValue
End Property

Public Property Get TipoMovimiento() As String
    TipoMovimiento = mTipoMovimiento
End Property
Public Property Let TipoMovimiento(ByVal NewValue As String)
    mTipoMovimiento = NewValue
End Property

Public Property Get NumeroDocumento() As String
    NumeroDocumento = mNumeroDocumento
End Property
Public Property Let NumeroDocumento(ByVal NewValue As String)
    mNumeroDocumento = NewValue
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

Public Property Get IdTipoDocumentoAnexado() As Long
    IdTipoDocumentoAnexado = mIdTipoDocumentoAnexado
End Property
Public Property Let IdTipoDocumentoAnexado(ByVal NewValue As Long)
    mIdTipoDocumentoAnexado = NewValue
End Property

Public Property Get TipoDocumentoAnexado() As String
    TipoDocumentoAnexado = mTipoDocumentoAnexado
End Property
Public Property Let TipoDocumentoAnexado(ByVal NewValue As String)
    mTipoDocumentoAnexado = NewValue
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

Public Property Get IdDocumentoAnexado() As Long
    IdDocumentoAnexado = mIdDocumentoAnexado
End Property
Public Property Let IdDocumentoAnexado(ByVal NewValue As Long)
    mIdDocumentoAnexado = NewValue
End Property

Public Property Get DocumentoAnexado() As String
    DocumentoAnexado = mDocumentoAnexado
End Property
Public Property Let DocumentoAnexado(ByVal NewValue As String)
    mDocumentoAnexado = NewValue
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

Public Property Get Costo() As Double
    Costo = mCosto
End Property
Public Property Let Costo(ByVal NewValue As Double)
    mCosto = NewValue
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

Public Property Get FechaMovimiento() As Date
    FechaMovimiento = mFechaMovimiento
End Property
Public Property Let FechaMovimiento(ByVal NewValue As Date)
    mFechaMovimiento = NewValue
End Property

Public Sub Fetch(IdMovimientoDetalle As Long)
    Dim dataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    Dim cSQL As String
    
On Error GoTo BLoqueError
        
    Set mRecord = F.GeneraRstSQL(F.SQL_MovDetallado(, , , , mConexion, , True, , IdMovimientoDetalle), mConexion)
    If mRecord.RecordCount > 0 Then Me.SetEntity mRecord
    Exit Sub
    
BLoqueError:
    F.MostrarMensajeError "Err.Description", "Fetch", Err.Source
End Sub

Public Sub SetEntity(Record As ADODB.Recordset)
    Me.IdMovimiento = F.NuloNumeric(Record("idmov"))
    Me.IdMovimientoDet = F.NuloNumeric(Record("idmovdet"))
    Me.IdLibroCostoTemp = F.NuloNumeric(Record("idlibrocostotemp"))
    Me.IdUnidadMedida = F.NuloNumeric(Record("idunimed"))
    Me.FechaMovimiento = Record("fchmov")
    Me.UnidadMedida = F.NuloString(Record("unimed"))
    Me.IdItem = F.NuloNumeric(Record("iditem"))
    Me.Item = F.NuloString(Record("item"))
    Me.CodigoItem = F.NuloString(Record("coditem"))
    Me.IdAlmacen = F.NuloNumeric(Record("idalm"))
    Me.Almacen = F.NuloString(Record("alm"))
    Me.Cantidad = F.NuloNumeric(Record("cantidad"))
    Me.IdTipoMovimiento = Record("tipmov")
    Me.TipoMovimiento = F.NuloString(Record("tipmovcad"))
    Me.NumeroDocumento = F.NuloString(Record("numdocconcat"))
    
    Me.CostoUnitario = F.NuloNumeric(Record("costounitario"))
    Me.CostoUnitarioPromedio = F.NuloNumeric(Record("costounitariopromedio"))
    Me.Costo = F.NuloNumeric(Record("costo"))
    Me.CostoPrimo = F.NuloNumeric(Record("costoprimo"))
    Me.CostoMOD = F.NuloNumeric(Record("costomod"))
    Me.CostoCIF = F.NuloNumeric(Record("costocif"))
    
    Me.IdTipoDocumentoReferencia = F.NuloNumeric(Record("idtipdocref"))
    Me.TipoDocumentoReferencia = F.NuloString(Record("tipdocref"))
    Me.IdDocumentoReferencia = F.NuloNumeric(Record("iddocref"))
    Me.DocumentoReferencia = F.NuloString(Record("numdocrefconcat"))
    Me.IdTipoDocumentoAnexado = F.NuloNumeric(Record("idtipdocrefanex"))
    Me.TipoDocumentoAnexado = F.NuloString(Record("tipdocrefanex"))
    Me.IdDocumentoAnexado = F.NuloNumeric(Record("iddocrefanex"))
    Me.DocumentoAnexado = F.NuloString(Record("docrefanex"))
    
End Sub

Public Sub Save(IdUsuario As Long, MachineName As String)
On Error GoTo BLoqueError
    If (F.NuloNumeric(Me.IdMovimientoDet) = 0) Then
        Insert IdUsuario, MachineName
    Else
        Update IdUsuario, MachineName
    End If
    Exit Sub
    
BLoqueError:
    Err.Raise &HFFFFFF01, , "Error al Grabar :" + Trim(Err.Description)
End Sub

Private Sub Insert(IdUsuario As Long, MachineName As String)
End Sub

Private Sub Update(IdUsuario As Long, MachineName As String)
End Sub

Public Sub Delete(IdUsuario As Long, MachineName As String)
End Sub





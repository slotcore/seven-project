VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ECCostoProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mFecha As Date
Private mTipoMovimiento As String
Private mMovimientoAlmacen As String
Private mAlmacen As String
Private mOrdenProduccion As String
Private mParteProduccion As String
Private mTipoItem As String
Private mCodigoItem As String
Private mItem As String
Private mReceta As String
Private mUnidadMedida As String
Private mCantidad As Double
Private mHoraInicio As Date
Private mHoraFin As Date
Private mCostoMP As Double
Private mCostoMOD As Double
Private mCostoPrimo As Double
Private mCostoCIF As Double
Private mCostoTotal As Double
Private mCostoUP As Double
Private mPrecioVenta As Double
Private mImporteVenta As Double
Private mDesviacion As Double
Private mDesviacionPorc As Double
Private mIdParteProduccion As Long
Private mIdParteProduccionDet As Long
Private mIdMovimientoDetalle As Long
Private mIdItem As Long
'*****************************************
' Variables Generales
'*****************************************
Private F As New SistemaLogica.Funciones
Private mConexion As New ADODB.Connection
Private mCalled As Boolean
Private mIsRecursive As Boolean

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

Public Property Get Fecha() As Date
    Fecha = mFecha
End Property
Public Property Let Fecha(ByVal NewValue As Date)
    mFecha = NewValue
End Property

Public Property Get TipoMovimiento() As String
    TipoMovimiento = mTipoMovimiento
End Property
Public Property Let TipoMovimiento(ByVal NewValue As String)
    mTipoMovimiento = NewValue
End Property

Public Property Get MovimientoAlmacen() As String
    MovimientoAlmacen = mMovimientoAlmacen
End Property
Public Property Let MovimientoAlmacen(ByVal NewValue As String)
    mMovimientoAlmacen = NewValue
End Property

Public Property Get Almacen() As String
    Almacen = mAlmacen
End Property
Public Property Let Almacen(ByVal NewValue As String)
    mAlmacen = NewValue
End Property

Public Property Get OrdenProduccion() As String
    OrdenProduccion = mOrdenProduccion
End Property
Public Property Let OrdenProduccion(ByVal NewValue As String)
    mOrdenProduccion = NewValue
End Property

Public Property Get ParteProduccion() As String
    ParteProduccion = mParteProduccion
End Property
Public Property Let ParteProduccion(ByVal NewValue As String)
    mParteProduccion = NewValue
End Property

Public Property Get TipoItem() As String
    TipoItem = mTipoItem
End Property
Public Property Let TipoItem(ByVal NewValue As String)
    mTipoItem = NewValue
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

Public Property Get Receta() As String
    Receta = mReceta
End Property
Public Property Let Receta(ByVal NewValue As String)
    mReceta = NewValue
End Property

Public Property Get UnidadMedida() As String
    UnidadMedida = mUnidadMedida
End Property
Public Property Let UnidadMedida(ByVal NewValue As String)
    mUnidadMedida = NewValue
End Property

Public Property Get Cantidad() As Double
    Cantidad = mCantidad
End Property
Public Property Let Cantidad(ByVal NewValue As Double)
    mCantidad = NewValue
End Property

Public Property Get HoraInicio() As Date
    HoraInicio = mHoraInicio
End Property
Public Property Let HoraInicio(ByVal NewValue As Date)
    mHoraInicio = NewValue
End Property

Public Property Get HoraFin() As Date
    HoraFin = mHoraFin
End Property
Public Property Let HoraFin(ByVal NewValue As Date)
    mHoraFin = NewValue
End Property

Public Property Get CostoMP() As Double
    CostoMP = mCostoMP
End Property
Public Property Let CostoMP(ByVal NewValue As Double)
    mCostoMP = NewValue
End Property

Public Property Get CostoMOD() As Double
    CostoMOD = mCostoMOD
End Property
Public Property Let CostoMOD(ByVal NewValue As Double)
    mCostoMOD = NewValue
End Property

Public Property Get CostoPrimo() As Double
    CostoPrimo = mCostoPrimo
End Property
Public Property Let CostoPrimo(ByVal NewValue As Double)
    mCostoPrimo = NewValue
End Property

Public Property Get CostoCIF() As Double
    CostoCIF = mCostoCIF
End Property
Public Property Let CostoCIF(ByVal NewValue As Double)
    mCostoCIF = NewValue
End Property

Public Property Get CostoTotal() As Double
    CostoTotal = mCostoTotal
End Property
Public Property Let CostoTotal(ByVal NewValue As Double)
    mCostoTotal = NewValue
End Property

Public Property Get CostoUP() As Double
    CostoUP = mCostoUP
End Property
Public Property Let CostoUP(ByVal NewValue As Double)
    mCostoUP = NewValue
End Property

Public Property Get PrecioVenta() As Double
    PrecioVenta = mPrecioVenta
End Property
Public Property Let PrecioVenta(ByVal NewValue As Double)
    mPrecioVenta = NewValue
End Property

Public Property Get ImporteVenta() As Double
    ImporteVenta = mImporteVenta
End Property
Public Property Let ImporteVenta(ByVal NewValue As Double)
    mImporteVenta = NewValue
End Property

Public Property Get Desviacion() As Double
    Desviacion = mDesviacion
End Property
Public Property Let Desviacion(ByVal NewValue As Double)
    mDesviacion = NewValue
End Property

Public Property Get DesviacionPorc() As Double
    DesviacionPorc = mDesviacionPorc
End Property
Public Property Let DesviacionPorc(ByVal NewValue As Double)
    mDesviacionPorc = NewValue
End Property

Public Property Get IdParteProduccion() As Long
    IdParteProduccion = mIdParteProduccion
End Property
Public Property Let IdParteProduccion(ByVal NewValue As Long)
    mIdParteProduccion = NewValue
End Property

Public Property Get IdParteProduccionDet() As Long
    IdParteProduccionDet = mIdParteProduccionDet
End Property
Public Property Let IdParteProduccionDet(ByVal NewValue As Long)
    mIdParteProduccionDet = NewValue
End Property

Public Property Get IdMovimientoDetalle() As Long
    IdMovimientoDetalle = mIdMovimientoDetalle
End Property
Public Property Let IdMovimientoDetalle(ByVal NewValue As Long)
    mIdMovimientoDetalle = NewValue
End Property

Public Property Get IdItem() As Long
    IdItem = mIdItem
End Property
Public Property Let IdItem(ByVal NewValue As Long)
    mIdItem = NewValue
End Property

Private Sub Class_Initialize()
    mCalled = False
    mIsRecursive = False
End Sub

Public Sub Fetch()
End Sub

Public Sub SetEntity(Record As ADODB.Recordset)
On Error GoTo BLoqueError

    Me.Fecha = Record("fchdoc")
    Me.TipoMovimiento = F.NuloString(Record("tipmov"))
    Me.MovimientoAlmacen = F.NuloString(Record("nummov"))
    Me.Almacen = F.NuloString(Record("alm"))
    Me.OrdenProduccion = F.NuloString(Record("numordprod"))
    Me.ParteProduccion = F.NuloString(Record("numparteprod"))
    Me.TipoItem = F.NuloString(Record("tipitem"))
    Me.CodigoItem = F.NuloString(Record("codpro"))
    Me.Item = F.NuloString(Record("item"))
    Me.Receta = F.NuloString(Record("codrec"))
    Me.UnidadMedida = F.NuloString(Record("unimed"))
    Me.Cantidad = F.NuloNumeric(Record("cantidad"))
    Me.HoraInicio = Record("horini")
    Me.HoraFin = Record("horfin")
    Me.CostoMP = F.NuloNumeric(Record("importemp"))
    Me.CostoMOD = F.NuloNumeric(Record("importemod"))
    Me.CostoPrimo = Me.CostoMP + Me.CostoMOD
    Me.CostoCIF = F.NuloNumeric(Record("importecif"))
    Me.CostoTotal = Me.CostoCIF + Me.CostoPrimo
    If Me.Cantidad > 0 Then Me.CostoUP = Me.CostoTotal / Me.Cantidad
    
    Me.ImporteVenta = F.NuloNumeric(Record("impven"))
    If Me.ImporteVenta > 0 Then
        Me.PrecioVenta = Me.ImporteVenta + (Me.ImporteVenta * F.NuloNumeric(Record("igv")) / 100)
        Me.Desviacion = Me.ImporteVenta - Me.CostoUP
        Me.DesviacionPorc = (Me.Desviacion / Me.ImporteVenta) * 100
    Else
        Me.PrecioVenta = 0
        Me.Desviacion = 0
        Me.DesviacionPorc = 0
    End If
    
    Me.IdParteProduccion = F.NuloNumeric(Record("idparteprod"))
    Me.IdParteProduccionDet = F.NuloNumeric(Record("idparteproddet"))
    Me.IdMovimientoDetalle = F.NuloNumeric(Record("idmovdet"))
    Me.IdItem = F.NuloNumeric(Record("iditem"))
    Exit Sub
    
BLoqueError:
    Err.Raise Err.Number, "[SetEntity] " & Err.Source, Err.Description
End Sub

Public Function Save(IdUsuario As Long, MachineName As String) As Boolean
    Save = False
End Function

Private Sub Insert()
End Sub

Private Sub Update()
End Sub

Public Sub Delete()
End Sub

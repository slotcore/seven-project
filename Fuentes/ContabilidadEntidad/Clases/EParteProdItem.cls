VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EParteProdItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdParteProduccionDet As Long
Private mIdParteProduccion As Long
Private mNumeroDocumento As String
Private mFechaParte As Date
Private mIdUnidadMedida As Long
Private mUnidadMedida As String
Private mIdItem As Long
Private mCodigoItem As String
Private mItem As String
Private mIdReceta As Long
Private mReceta As String
Private mCantidadProgramada As Double
Private mCantidadProducida As Double
Private mIdOrdenProduccion As Long
Private mOrdenProduccion As String
Private mLote As String
Private mHoraInicio As Date
Private mHoraFin As Date
Private mGlosa As String
Private mIdMovimientoDetalle As Double
Private mCostoUnitario As Double
Private mCostoUnitarioPromedio As Double
Private mCosto As Double
Private mCostoPrimo As Double
Private mCostoMOD As Double
Private mCostoCIF As Double
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

Public Property Get IdParteProduccionDet() As Long
    IdParteProduccionDet = mIdParteProduccionDet
End Property
Public Property Let IdParteProduccionDet(ByVal NewValue As Long)
    mIdParteProduccionDet = NewValue
End Property

Public Property Get IdParteProduccion() As Long
    IdParteProduccion = mIdParteProduccion
End Property
Public Property Let IdParteProduccion(ByVal NewValue As Long)
    mIdParteProduccion = NewValue
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

Public Property Get IdReceta() As Long
    IdReceta = mIdReceta
End Property
Public Property Let IdReceta(ByVal NewValue As Long)
    mIdReceta = NewValue
End Property

Public Property Get Receta() As String
    Receta = mReceta
End Property
Public Property Let Receta(ByVal NewValue As String)
    mReceta = NewValue
End Property

Public Property Get CantidadProgramada() As Double
    CantidadProgramada = mCantidadProgramada
End Property
Public Property Let CantidadProgramada(ByVal NewValue As Double)
    mCantidadProgramada = NewValue
End Property

Public Property Get CantidadProducida() As Double
    CantidadProducida = mCantidadProducida
End Property
Public Property Let CantidadProducida(ByVal NewValue As Double)
    mCantidadProducida = NewValue
End Property

Public Property Get IdOrdenProduccion() As Long
    IdOrdenProduccion = mIdOrdenProduccion
End Property
Public Property Let IdOrdenProduccion(ByVal NewValue As Long)
    mIdOrdenProduccion = NewValue
End Property

Public Property Get OrdenProduccion() As String
    OrdenProduccion = mOrdenProduccion
End Property
Public Property Let OrdenProduccion(ByVal NewValue As String)
    mOrdenProduccion = NewValue
End Property

Public Property Get Lote() As String
    Lote = mLote
End Property
Public Property Let Lote(ByVal NewValue As String)
    mLote = NewValue
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

Public Property Get Glosa() As String
    Glosa = mGlosa
End Property
Public Property Let Glosa(ByVal NewValue As String)
    mGlosa = NewValue
End Property

Public Property Get NumeroDocumento() As String
    NumeroDocumento = mNumeroDocumento
End Property
Public Property Let NumeroDocumento(ByVal NewValue As String)
    mNumeroDocumento = NewValue
End Property

Public Property Get FechaParte() As Date
    FechaParte = mFechaParte
End Property
Public Property Let FechaParte(ByVal NewValue As Date)
    mFechaParte = NewValue
End Property

Public Property Get IdMovimientoDetalle() As Double
    IdMovimientoDetalle = mIdMovimientoDetalle
End Property
Public Property Let IdMovimientoDetalle(ByVal NewValue As Double)
    mIdMovimientoDetalle = NewValue
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

Public Sub Fetch(IdParteProduccionDet As Long)
End Sub

Public Sub SetEntity(Record As ADODB.Recordset)
On Error GoTo BLoqueError

    Me.IdParteProduccionDet = F.NuloNumeric(Record("idparteproddet"))
    Me.IdParteProduccion = F.NuloNumeric(Record("idparteprod"))
    Me.IdReceta = F.NuloNumeric(Record("idrec"))
    Me.IdItem = F.NuloString(Record("iditem"))
    Me.IdUnidadMedida = F.NuloString(Record("idunimed"))
    Me.CantidadProgramada = F.NuloNumeric(Record("canprog"))
    Me.CantidadProducida = F.NuloNumeric(Record("cantidad"))
    Me.HoraInicio = Record("horini")
    Me.HoraFin = Record("horfin")
    Me.Glosa = F.NuloString(Record("glosa"))
    Me.IdOrdenProduccion = F.NuloNumeric(Record("idord"))
    Me.UnidadMedida = F.NuloString(Record("unimed"))
    Me.Receta = F.NuloString(Record("codrec"))
    Me.OrdenProduccion = F.NuloString(Record("numordprod"))
    Me.Lote = F.NuloString(Record("lote"))
    Me.CodigoItem = F.NuloString(Record("codpro"))
    Me.Item = F.NuloString(Record("item"))
    Me.NumeroDocumento = F.NuloString(Record("numdoc"))
    Me.FechaParte = Record("fchdoc")
    Me.IdMovimientoDetalle = F.NuloNumeric(Record("idmovdet"))
    
    Me.CostoUnitario = F.NuloNumeric(Record("costounitario"))
    Me.CostoUnitarioPromedio = F.NuloNumeric(Record("costounitariopromedio"))
    Me.Costo = F.NuloNumeric(Record("costo"))
    Me.CostoPrimo = F.NuloNumeric(Record("costoprimo"))
    Me.CostoMOD = F.NuloNumeric(Record("costomod"))
    Me.CostoCIF = F.NuloNumeric(Record("costocif"))
    Exit Sub
    
BLoqueError:
    Err.Raise Err.Number, "[SetEntity] " & Err.Source, Err.Description
End Sub

Public Sub Save(IdUsuario As Long, MachineName As String)
On Error GoTo BLoqueError
    If (F.NuloNumeric(Me.IdParteProduccionDet) = 0) Then
        Insert IdUsuario, MachineName
    Else
        Update IdUsuario, MachineName
    End If
    Exit Sub
BLoqueError:
    Err.Raise &HFFFFFF01, , "No se pudo registrar el ParteProduccion por el siguiente motivo :" + Trim(Err.Description)
End Sub

Private Sub Insert(IdUsuario As Long, MachineName As String)
End Sub

Private Sub Update(IdUsuario As Long, MachineName As String)
End Sub

Public Sub Delete(IdUsuario As Long, MachineName As String)
End Sub



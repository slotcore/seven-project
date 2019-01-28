VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EParteProdDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdParteProduccionDet As Long
Private mIdParteProduccion As Long
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
Private mFecha As Date
Private mIdMovimientoDetalle As Double
Private mLParteProduccionDetIns As New LEParteProdDetIns
Private F As New SistemaLogica.Funciones
Private mConexion As New ADODB.Connection
Private mCalled As Boolean
Private mLoadChild As Boolean

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

Public Property Get Fecha() As Date
    Fecha = mFecha
End Property
Public Property Let Fecha(ByVal NewValue As Date)
    mFecha = NewValue
End Property

Public Property Get IdMovimientoDetalle() As Double
    IdMovimientoDetalle = mIdMovimientoDetalle
End Property
Public Property Let IdMovimientoDetalle(ByVal NewValue As Double)
    mIdMovimientoDetalle = NewValue
End Property

Public Property Get LParteProduccionDetIns() As LEParteProdDetIns
    Set LParteProduccionDetIns = mLParteProduccionDetIns
End Property
Public Property Set LParteProduccionDetIns(ByVal NewValue As LEParteProdDetIns)
    Set mLParteProduccionDetIns = NewValue
End Property

Public Sub Fetch(IdParteProduccionDet As Long)
    Dim database As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set database.Connection = mConexion
            
    database.CommandText = "SELECT pro_producciondet.idproddet AS idparteproddet, pro_producciondet.idpro AS idparteprod, pro_producciondet.idrec, pro_producciondet.iditem, pro_producciondet.idunimed, pro_producciondet.canprog, pro_producciondet.cantidad, pro_producciondet.horini, pro_producciondet.horfin, pro_producciondet.obs AS glosa, pro_producciondet.idord, mae_unidades.abrev AS unimed, pro_receta.codrec, [pro_ordenprod].[numser] & '-' & [pro_ordenprod].[numdoc] AS numordprod, pro_ordenprod.lote, alm_inventario.codpro, alm_inventario.descripcion AS item, pro_produccion.fchdoc, alm_ingresodet.idmovdet " _
        + vbCr + "FROM (pro_produccion INNER JOIN ((((pro_producciondet LEFT JOIN mae_unidades ON pro_producciondet.idunimed = mae_unidades.id) LEFT JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN pro_ordenprod ON pro_producciondet.idord = pro_ordenprod.id) LEFT JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id) ON pro_produccion.id = pro_producciondet.idpro) LEFT JOIN alm_ingresodet ON pro_producciondet.idproddet = alm_ingresodet.iddocref " _
        + vbCr + "WHERE (((pro_producciondet.idproddet) = " & IdParteProduccionDet & "));"
        
    Set mRecord = database.GetRecordset
    If mRecord.RecordCount > 0 Then
        Me.SetEntity mRecord
    End If
    Exit Sub
    
BLoqueError:
    Err.Raise &HFFFFFF01, , "[Fetch] " + Trim(Err.Description)
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
    Me.Fecha = Record("fchdoc")
    Me.IdMovimientoDetalle = F.NuloNumeric(Record("idmovdet"))
    If mLoadChild Then
        Set Me.LParteProduccionDetIns.Conexion = mConexion
        Me.LParteProduccionDetIns.Fetch Me.IdParteProduccionDet
    End If
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
    Dim xId As Double
    Dim database As New SistemaData.EDataBase
    Dim FData As New SistemaData.Funciones
    Dim F As New SistemaLogica.Funciones
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set database.Connection = mConexion
    If Not Called Then database.BeginTrans
    
    xId = FData.GetCodeTable("pro_producciondet", "idproddet", mConexion)
    database.ClearParameter
    
    database.CommandText = "INSERT INTO pro_producciondet(idproddet, idpro, iditem, idunimed, cantidad, canprog, horini, horfin, idrec, idord, obs) " _
            + vbCr + "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

    database.AddParameter "@idproddet", adInteger, adParamInput, xId
    database.AddParameter "@idpro", adInteger, adParamInput, F.NuloNumeric(Me.IdParteProduccion)
    database.AddParameter "@iditem", adInteger, adParamInput, F.NuloNumeric(Me.IdItem)
    database.AddParameter "@idunimed", adInteger, adParamInput, F.NuloNumeric(Me.IdUnidadMedida)
    database.AddParameter "@cantidad", adDouble, adParamInput, F.NuloNumeric(Me.CantidadProducida)
    database.AddParameter "@canprog", adDouble, adParamInput, F.NuloNumeric(Me.CantidadProgramada)
    database.AddParameter "@horini", adDate, adParamInput, Me.HoraInicio
    database.AddParameter "@horfin", adDate, adParamInput, Me.HoraFin
    database.AddParameter "@idrec", adInteger, adParamInput, F.NuloNumeric(Me.IdReceta)
    database.AddParameter "@idord", adInteger, adParamInput, F.NuloNumeric(Me.IdOrdenProduccion)
    database.AddParameter "@obs", adVarChar, adParamInput, F.NuloString(Me.Glosa), 200
    database.Execute
    
    If Not Called Then database.CommitTrans
    Set database = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then database.RollbackTrans
    Set database = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el ParteProduccion por el siguiente motivo :" + Trim(Err.Description)
End Sub

Private Sub Update(IdUsuario As Long, MachineName As String)
    Dim database As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set database.Connection = mConexion
    If Not Called Then database.BeginTrans
    
    database.ClearParameter

    database.CommandText = "UPDATE pro_producciondet SET idpro = ?, iditem = ?, idunimed = ?, cantidad = ?, canprog = ? " _
             & ", horini = ?, horfin = ?, idrec = ?, idord = ?, obs = ?" _
            + vbCr + "WHERE(pro_producciondet.idproddet = ?)"

    database.AddParameter "@idpro", adInteger, adParamInput, F.NuloNumeric(Me.IdParteProduccion)
    database.AddParameter "@iditem", adInteger, adParamInput, F.NuloNumeric(Me.IdItem)
    database.AddParameter "@idunimed", adInteger, adParamInput, F.NuloNumeric(Me.IdUnidadMedida)
    database.AddParameter "@cantidad", adDouble, adParamInput, F.NuloNumeric(Me.CantidadProducida)
    database.AddParameter "@canprog", adDouble, adParamInput, F.NuloNumeric(Me.CantidadProgramada)
    database.AddParameter "@horini", adDate, adParamInput, Me.HoraInicio
    database.AddParameter "@horfin", adDate, adParamInput, Me.HoraFin
    database.AddParameter "@idrec", adInteger, adParamInput, F.NuloNumeric(Me.IdReceta)
    database.AddParameter "@idord", adInteger, adParamInput, F.NuloNumeric(Me.IdOrdenProduccion)
    database.AddParameter "@obs", adVarChar, adParamInput, F.NuloString(Me.Glosa), 200
    database.AddParameter "@idproddet", adInteger, adParamInput, F.NuloNumeric(Me.IdParteProduccionDet)
    database.Execute
    
    If Not Called Then database.CommitTrans
    Set database = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then database.RollbackTrans
    Set database = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el ParteProduccion por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub

Public Sub Delete(IdUsuario As Long, MachineName As String)
    Dim database As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then database.Connection = mConexion
    If Not Called Then database.BeginTrans
    
    database.ClearParameter
    database.CommandText = "DELETE FROM pro_producciondet " _
            + vbCr + "WHERE pro_producciondet.idproddet = ?"
    
    database.AddParameter "@idproddet", adInteger, adParamInput, Me.IdParteProduccionDet
    database.Execute
    
    If Not Called Then database.CommitTrans
    Set database = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then database.RollbackTrans
    Set database = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el ParteProduccion por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EParteProdDetIns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdParteProduccionDetIns As Long
Private mIdParteProduccionDet As Long
Private mIdUnidadMedida As Long
Private mUnidadMedida As String
Private mIdItem As Long
Private mCodigoItem As String
Private mItem As String
Private mCantidadTeoricaIngresos As Double
Private mCantidadTeoricaSalidas As Double
Private mCantidadRealIngresos As Double
Private mCantidadRealSalidas As Double
Private mLParteProduccionDetInsMov As New LEParteProdDetInsMov
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

Public Property Get IdParteProduccionDetIns() As Long
    IdParteProduccionDetIns = mIdParteProduccionDetIns
End Property
Public Property Let IdParteProduccionDetIns(ByVal NewValue As Long)
    mIdParteProduccionDetIns = NewValue
End Property

Public Property Get IdParteProduccionDet() As Long
    IdParteProduccionDet = mIdParteProduccionDet
End Property
Public Property Let IdParteProduccionDet(ByVal NewValue As Long)
    mIdParteProduccionDet = NewValue
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

Public Property Get CantidadTeoricaIngresos() As Double
    CantidadTeoricaIngresos = mCantidadTeoricaIngresos
End Property
Public Property Let CantidadTeoricaIngresos(ByVal NewValue As Double)
    mCantidadTeoricaIngresos = NewValue
End Property

Public Property Get CantidadTeoricaSalidas() As Double
    CantidadTeoricaSalidas = mCantidadTeoricaSalidas
End Property
Public Property Let CantidadTeoricaSalidas(ByVal NewValue As Double)
    mCantidadTeoricaSalidas = NewValue
End Property

Public Property Get CantidadRealIngresos() As Double
    CantidadRealIngresos = mCantidadRealIngresos
End Property
Public Property Let CantidadRealIngresos(ByVal NewValue As Double)
    mCantidadRealIngresos = NewValue
End Property

Public Property Get CantidadRealSalidas() As Double
    CantidadRealSalidas = mCantidadRealSalidas
End Property
Public Property Let CantidadRealSalidas(ByVal NewValue As Double)
    mCantidadRealSalidas = NewValue
End Property

Public Property Get LParteProduccionDetInsMov() As LEParteProdDetInsMov
    Set LParteProduccionDetInsMov = mLParteProduccionDetInsMov
End Property
Public Property Set LParteProduccionDetInsMov(ByVal NewValue As LEParteProdDetInsMov)
    Set mLParteProduccionDetInsMov = NewValue
End Property

Public Function Fetch(IdParteProduccionDetIns As String) As EParteProdDetIns
    Dim database As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    Dim EParteProdDetIns As New EParteProdDetIns
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set database.Connection = mConexion
            
    database.CommandText = "SELECT pro_producciondet.idproddet AS idparteproddet, pro_producciondet.idpro AS idparteprod, pro_producciondet.idrec, pro_producciondet.iditem, pro_producciondet.idunimed, pro_producciondet.canprog, pro_producciondet.cantidad, pro_producciondet.horini, pro_producciondet.horfin, pro_producciondet.obs AS glosa, pro_producciondet.idord, mae_unidades.abrev AS unimed, pro_receta.codrec, [pro_ordenprod].[numser] & '-' & [pro_ordenprod].[numdoc] AS numordprod, pro_ordenprod.lote, alm_inventario.codpro, alm_inventario.descripcion AS item " _
        + vbCr + "FROM (((pro_producciondet LEFT JOIN mae_unidades ON pro_producciondet.idunimed = mae_unidades.id) LEFT JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN pro_ordenprod ON pro_producciondet.idord = pro_ordenprod.id) LEFT JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id " _
        + vbCr + "WHERE (((pro_producciondet.idproddet) = " & IdParteProduccionDetIns & "));"
        
    Set mRecord = database.GetRecordset
    If mRecord.RecordCount > 0 Then
        EParteProdDetIns.SetEntity mRecord
    Else
        Set EParteProdDetIns = Nothing
    End If
    
    Set Fetch = EParteProdDetIns
    Exit Function
    
BLoqueError:
    Err.Raise &HFFFFFF01, , "[Fetch] " + Trim(Err.Description)

End Function

Public Sub SetEntity(Record As ADODB.Recordset)
On Error GoTo BLoqueError
    
    Me.IdParteProduccionDetIns = F.NuloNumeric(Record("idprod"))
    Me.IdParteProduccionDet = F.NuloNumeric(Record("idproddet"))
    Me.IdItem = F.NuloString(Record("iditem"))
    Me.IdUnidadMedida = F.NuloString(Record("idunimed"))
    Me.CantidadTeoricaIngresos = F.NuloNumeric(Record("canteoing"))
    Me.CantidadTeoricaSalidas = F.NuloNumeric(Record("canteosal"))
    Me.CantidadRealIngresos = F.NuloNumeric(Record("canrealing"))
    Me.CantidadRealSalidas = F.NuloNumeric(Record("canrealsal"))
    Me.UnidadMedida = F.NuloString(Record("unimed"))
    Me.CodigoItem = F.NuloString(Record("codpro"))
    Me.Item = F.NuloString(Record("item"))
    If mLoadChild Then
        Set Me.LParteProduccionDetInsMov.Conexion = mConexion
        Me.LParteProduccionDetInsMov.Fetch Me.IdParteProduccionDet, Me.IdItem
    End If
    Exit Sub
    
BLoqueError:
    Err.Raise Err.Number, "[SetEntity] " & Err.Source, Err.Description
End Sub

Public Sub Save(IdUsuario As Long, MachineName As String)
On Error GoTo BLoqueError
    If (F.NuloNumeric(Me.IdParteProduccionDetIns) = 0) Then
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

'    database.AddParameter "@idproddet", adInteger, adParamInput, xId
'    database.AddParameter "@idpro", adInteger, adParamInput, F.NuloNumeric(Me.IdParteProduccionDet)
'    database.AddParameter "@iditem", adInteger, adParamInput, F.NuloNumeric(Me.IdItem)
'    database.AddParameter "@idunimed", adInteger, adParamInput, F.NuloNumeric(Me.IdUnidadMedida)
'    database.AddParameter "@cantidad", adDouble, adParamInput, F.NuloNumeric(Me.CantidadTeoricaSalidas)
'    database.AddParameter "@canprog", adDouble, adParamInput, F.NuloNumeric(Me.CantidadTeoricaIngresos)
'    database.AddParameter "@horini", adDate, adParamInput, Me.HoraInicio
'    database.AddParameter "@horfin", adDate, adParamInput, Me.HoraFin
'    database.AddParameter "@idrec", adInteger, adParamInput, F.NuloNumeric(Me.IdReceta)
'    database.AddParameter "@idord", adInteger, adParamInput, F.NuloNumeric(Me.IdOrdenProduccion)
'    database.AddParameter "@obs", adVarChar, adParamInput, F.NuloString(Me.Glosa), 200
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

'    database.AddParameter "@idpro", adInteger, adParamInput, F.NuloNumeric(Me.IdParteProduccionDet)
'    database.AddParameter "@iditem", adInteger, adParamInput, F.NuloNumeric(Me.IdItem)
'    database.AddParameter "@idunimed", adInteger, adParamInput, F.NuloNumeric(Me.IdUnidadMedida)
'    database.AddParameter "@cantidad", adDouble, adParamInput, F.NuloNumeric(Me.CantidadTeoricaSalidas)
'    database.AddParameter "@canprog", adDouble, adParamInput, F.NuloNumeric(Me.CantidadTeoricaIngresos)
'    database.AddParameter "@horini", adDate, adParamInput, Me.HoraInicio
'    database.AddParameter "@horfin", adDate, adParamInput, Me.HoraFin
'    database.AddParameter "@idrec", adInteger, adParamInput, F.NuloNumeric(Me.IdReceta)
'    database.AddParameter "@idord", adInteger, adParamInput, F.NuloNumeric(Me.IdOrdenProduccion)
'    database.AddParameter "@obs", adVarChar, adParamInput, F.NuloString(Me.Glosa), 200
'    database.AddParameter "@idproddet", adInteger, adParamInput, F.NuloNumeric(Me.IdParteProduccionDetIns)
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
    
    database.AddParameter "@idproddet", adInteger, adParamInput, Me.IdParteProduccionDetIns
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



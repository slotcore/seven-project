VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LEParteProdItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mLParteProdDetItem As New Collection
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

Public Sub Add(mParteProdDetItem As EParteProdItem, Optional Key As String)
    If Key <> "" Then
        mLParteProdDetItem.Add mParteProdDetItem, Key
    Else
        mLParteProdDetItem.Add mParteProdDetItem
    End If
End Sub

Public Function Count() As Long
    Count = mLParteProdDetItem.Count
End Function

Public Sub Delete(ByVal Index As Variant)
    mLParteProdDetItem.Remove Index
End Sub

Public Function Item(ByVal Index As Variant) As EParteProdItem
Attribute Item.VB_UserMemId = 0
    Set Item = mLParteProdDetItem.Item(Index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mLParteProdDetItem.[_NewEnum]
End Function

Public Sub Fetch(Optional IdParteProduccion As Long, _
                        Optional IdAlmacen As Long, _
                        Optional FechaInicio As Date, _
                        Optional FechaFin As Date, _
                        Optional FechaInicioMovimientos As Date)
    Dim database As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    Dim mLParteProdDetItem As New LEParteProdItem
    Dim cWhere As String
    
On Error GoTo BLoqueError
    
    If F.NuloNumeric(IdParteProduccion) > 0 Then
        cWhere = "WHERE pro_producciondet.idpro = " & IdParteProduccion
    End If
    If F.NuloNumeric(IdAlmacen) > 0 Then
        If cWhere = "" Then cWhere = "WHERE " Else cWhere = cWhere & " AND "
        cWhere = cWhere & "pro_produccion.idalm = " & IdAlmacen
    End If
    If Not IsNull(FechaInicio) And FechaInicio <> Empty Then
        If cWhere = "" Then cWhere = "WHERE " Else cWhere = cWhere & " AND "
        cWhere = cWhere & "pro_produccion.fchdoc >=CDate('" & FechaInicio & "')"
    End If
    If Not IsNull(FechaFin) And FechaFin <> Empty Then
        If cWhere = "" Then cWhere = "WHERE " Else cWhere = cWhere & " AND "
        cWhere = cWhere & "pro_produccion.fchdoc <=CDate('" & FechaFin & "')"
    End If
    If Not IsNull(FechaInicioMovimientos) And FechaInicioMovimientos <> Empty Then
        If cWhere = "" Then cWhere = "WHERE " Else cWhere = cWhere & " AND "
        cWhere = cWhere & "pro_produccion.fchdoc >=CDate('" & FechaInicioMovimientos & "')"
    End If
    
    If Not mConexion Is Nothing Then Set database.Connection = mConexion
    
    database.CommandText = "SELECT pro_producciondet.idproddet AS idparteproddet, pro_producciondet.idpro AS idparteprod, pro_producciondet.idrec, pro_producciondet.iditem, pro_producciondet.idunimed, pro_producciondet.canprog, pro_producciondet.cantidad, pro_producciondet.horini, pro_producciondet.horfin, pro_producciondet.obs AS glosa, pro_producciondet.idord, mae_unidades.abrev AS unimed, pro_receta.codrec, [pro_ordenprod].[numser] & '-' & [pro_ordenprod].[numdoc] AS numordprod, pro_ordenprod.lote, alm_inventario.codpro, alm_inventario.descripcion AS item, [pro_produccion].[numser] & '-' & [pro_produccion].[numdoc] AS numdoc, pro_produccion.fchdoc, alm_ingresodet.idmovdet, (con_librocostotemp.costoprimo + IIf(con_librocostotemp.costomod Is Null Or con_librocostotemp.costomod = 0, 0, con_librocostotemp.costomod) + IIf(con_librocostotemp.costocif Is Null Or con_librocostotemp.costocif = 0, 0, con_librocostotemp.costocif)) AS costo, con_librocostotemp.costounitariopromedio, " _
                & "con_librocostotemp.costounitario, con_librocostotemp.costoprimo, con_librocostotemp.costomod, con_librocostotemp.costocif, mae_tipoproducto.prefijo AS tippro " _
        + vbCr + "FROM (pro_produccion LEFT JOIN ((((((pro_producciondet LEFT JOIN mae_unidades ON pro_producciondet.idunimed = mae_unidades.id) LEFT JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN pro_ordenprod ON pro_producciondet.idord = pro_ordenprod.id) LEFT JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id) LEFT JOIN alm_ingresodet ON pro_producciondet.idproddet = alm_ingresodet.iddocref) LEFT JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet) ON pro_produccion.id = pro_producciondet.idpro) LEFT JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id " _
        + vbCr + cWhere & " " _
        + vbCr + "ORDER BY mae_tipoproducto.prefijo, pro_producciondet.iditem, pro_produccion.fchdoc"
    
    Set mRecord = database.GetRecordset
    If mRecord.RecordCount > 0 Then
        'mRecord.Sort = "fchdoc"
        mRecord.MoveFirst
        While Not mRecord.EOF
            Dim mParteProdDetItem As New EParteProdItem
            Set mParteProdDetItem.Conexion = mConexion
            mParteProdDetItem.SetEntity mRecord
            Me.Add mParteProdDetItem, F.NuloString(mParteProdDetItem.IdParteProduccionDet)
            Set mParteProdDetItem = Nothing
            mRecord.MoveNext
        Wend
    End If
    Exit Sub
    
BLoqueError:
    Set mParteProdDetItem = Nothing
    Err.Raise Err.Number, "[Fetch] " & Err.Source, Err.Description
End Sub



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LECCostoProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mLCCostoProd As New Collection
Private F As New SistemaLogica.Funciones
Private mConexion As New ADODB.Connection
Private mCalled As Boolean
Private mLoadChild As Boolean

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

Public Property Get LoadChild() As Boolean
    LoadChild = mLoadChild
End Property
Public Property Let LoadChild(ByVal NewValue As Boolean)
    mLoadChild = NewValue
End Property

Public Sub Add(mCCostoProd As ECCostoProd, Optional Key As String)
    If Key <> "" Then
        mLCCostoProd.Add mCCostoProd, Key
    Else
        mLCCostoProd.Add mCCostoProd
    End If
End Sub

Public Function Count() As Long
    Count = mLCCostoProd.Count
End Function

Public Sub Delete(ByVal Index As Variant)
    mLCCostoProd.Remove Index
End Sub

Public Function Item(ByVal Index As Variant) As ECCostoProd
Attribute Item.VB_UserMemId = 0
    Set Item = mLCCostoProd.Item(Index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mLCCostoProd.[_NewEnum]
End Function

Public Sub Fetch(Optional FechaInicio As Date, _
                        Optional FechaFin As Date, _
                        Optional FechaInicioMovimientos As Date, _
                        Optional FiltroWhere As String = "")
    Dim database As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    Dim mLParteProduccionDet As New LECCostoProd
    Dim cWhere As String
    Dim cWhereVentas As String
    
On Error GoTo BLoqueError
    
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
    If FiltroWhere <> "" Then
        If cWhere = "" Then cWhere = "WHERE " Else cWhere = cWhere & " AND "
        cWhere = cWhere & FiltroWhere
    End If
    
    ' Filtro para consulta de Ventas
    If Not IsNull(FechaInicio) And FechaInicio <> Empty Then
        If cWhereVentas = "" Then cWhereVentas = "WHERE " Else cWhereVentas = cWhereVentas & " AND "
        cWhereVentas = cWhereVentas & "vta_ventas.fchdoc >=CDate('" & FechaInicio & "')"
    End If
    If Not IsNull(FechaFin) And FechaFin <> Empty Then
        If cWhereVentas = "" Then cWhereVentas = "WHERE " Else cWhereVentas = cWhereVentas & " AND "
        cWhereVentas = cWhereVentas & "vta_ventas.fchdoc <=CDate('" & FechaFin & "')"
    End If
    If Not IsNull(FechaInicioMovimientos) And FechaInicioMovimientos <> Empty Then
        If cWhereVentas = "" Then cWhereVentas = "WHERE " Else cWhereVentas = cWhereVentas & " AND "
        cWhereVentas = cWhereVentas & "vta_ventas.fchdoc >=CDate('" & FechaInicioMovimientos & "')"
    End If
    
    If Not mConexion Is Nothing Then Set database.Connection = mConexion
    
    database.CommandText = "SELECT pro_producciondet.idproddet AS idparteproddet, pro_producciondet.idpro AS idparteprod, pro_producciondet.idrec, pro_producciondet.iditem, pro_producciondet.idunimed, pro_producciondet.canprog, pro_producciondet.cantidad, pro_producciondet.horini, pro_producciondet.horfin, pro_producciondet.obs AS glosa, pro_producciondet.idord, mae_unidades.abrev AS unimed, pro_receta.codrec, [pro_ordenprod].[numser] & '-' & [pro_ordenprod].[numdoc] AS numordprod, [pro_produccion].[numser] & '-' & [pro_produccion].[numdoc] AS numparteprod, [alm_ingreso].[numser] & '-' & [alm_ingreso].[numdoc] AS nummov, pro_ordenprod.lote, alm_inventario.codpro, " _
                & "alm_inventario.descripcion AS item, pro_produccion.fchdoc, alm_ingresodet.idmovdet, con_librocostotemp.tipmov, con_librocostotemp.costoprimo AS importemp, con_librocostotemp.costomod As importemod, con_librocostotemp.costocif As importecif, mae_tipoproducto.prefijo AS tipitem, CVENTASP.igv, CVENTASP.impven, alm_almacenes.descripcion AS alm " _
        + vbCr + "FROM (alm_ingreso RIGHT JOIN (pro_produccion INNER JOIN ((((((((pro_producciondet LEFT JOIN mae_unidades ON pro_producciondet.idunimed = mae_unidades.id) LEFT JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN pro_ordenprod ON pro_producciondet.idord = pro_ordenprod.id) LEFT JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id) LEFT JOIN alm_ingresodet ON pro_producciondet.idproddet = alm_ingresodet.iddocref) LEFT JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet) LEFT JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id) LEFT JOIN " _
        + vbCr + "( " _
        + vbCr + "SELECT vta_ventasdet.iditem, Last(vta_ventas.tasaigv) AS igv, Last(vta_ventasdet.preuni) AS impven, Last(vta_ventasdet.canpro) AS cantot " _
        + vbCr + "FROM vta_ventas INNER JOIN vta_ventasdet ON vta_ventas.id = vta_ventasdet.idvta " _
        + vbCr + cWhereVentas _
        + vbCr + "GROUP BY vta_ventasdet.iditem" _
        + vbCr + ")  AS CVENTASP ON pro_producciondet.iditem = CVENTASP.iditem) ON pro_produccion.id = pro_producciondet.idpro) ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN alm_almacenes ON alm_ingreso.idalm = alm_almacenes.id " _
        + vbCr + cWhere
    
    Set mRecord = database.GetRecordset
    If mRecord.RecordCount > 0 Then
        mRecord.Sort = "fchdoc"
        mRecord.MoveFirst
        While Not mRecord.EOF
            Dim mCCostoProd As New ECCostoProd
            Set mCCostoProd.Conexion = mConexion
            mCCostoProd.SetEntity mRecord
            Me.Add mCCostoProd, F.NuloString(mCCostoProd.IdParteProduccionDet)
            Set mCCostoProd = Nothing
            mRecord.MoveNext
        Wend
    End If
    Exit Sub
    
BLoqueError:
    Err.Raise Err.Number, "[Fetch] " & Err.Source, Err.Description
End Sub



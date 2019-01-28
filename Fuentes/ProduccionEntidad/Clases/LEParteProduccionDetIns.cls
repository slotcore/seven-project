VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LEParteProdDetIns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mLParteProduccionDetIns As New Collection
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

Public Sub Add(mParteProduccionDetIns As EParteProdDetIns)
    mLParteProduccionDetIns.Add mParteProduccionDetIns
End Sub

Public Function Count() As Long
    Count = mLParteProduccionDetIns.Count
End Function

Public Sub Delete(ByVal Index As Variant)
    mLParteProduccionDetIns.Remove Index
End Sub

Public Function Item(ByVal Index As Variant) As EParteProdDetIns
Attribute Item.VB_UserMemId = 0
    Set Item = mLParteProduccionDetIns.Item(Index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mLParteProduccionDetIns.[_NewEnum]
End Function

Public Sub Fetch(IdParteProduccionDet As String)
    Dim cSQL As String
    Dim cSQLInterna As String
    Dim mIdTipoOrdProduccion As Long
    Dim mIdTipoSolMateriales As Long
    Dim database As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set database.Connection = mConexion
    
    mIdTipoOrdProduccion = F.NuloNumeric(F.KeyValue("OrdenProduccion", mConexion))
    mIdTipoSolMateriales = F.NuloNumeric(F.KeyValue("SolictudMateriales", mConexion))
    
    cSQLInterna = "SELECT pro_ordenprod.id AS idord, pro_solicitudmat.id AS idsol, alm_ingresodet.id AS idmov, alm_ingresodet.iditem, alm_inventario.idunimed, alm_inventario.descripcion AS item, mae_unidades.abrev AS unimed, Sum(alm_ingresodet.cantteo) AS canteoing, 0 AS canteosal, Sum(alm_ingresodet.cantidad) AS canrealing, 0 AS canrealsal " _
        + vbCr + "FROM pro_producciondet INNER JOIN (((((pro_ordenprod INNER JOIN pro_solicitudmat ON pro_ordenprod.id = pro_solicitudmat.iddocref) INNER JOIN alm_ingreso ON pro_solicitudmat.id = alm_ingreso.iddocref) INNER JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id) INNER JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) ON pro_producciondet.idord = pro_ordenprod.id " _
        + vbCr + "WHERE (((pro_solicitudmat.idtipdocref) = " & mIdTipoOrdProduccion & ") And ((alm_ingreso.idtipdocref) = " & mIdTipoSolMateriales & ") And ((alm_ingreso.tipmov) = True) And ((pro_producciondet.idproddet) = " & IdParteProduccionDet & ")) " _
        + vbCr + "GROUP BY pro_ordenprod.id, pro_solicitudmat.id, alm_ingresodet.id, alm_ingresodet.iditem, alm_inventario.idunimed, alm_inventario.descripcion, mae_unidades.abrev, 0 " _
        + vbCr + "UNION " _
        + vbCr + "SELECT pro_ordenprod.id AS idord, pro_solicitudmat_1.id AS idsol, alm_ingresodet.id AS idmov, alm_ingresodet.iditem, alm_inventario.idunimed, alm_inventario.descripcion AS item, mae_unidades.abrev AS unimed, Sum(alm_ingresodet.cantteo) AS canteoing, 0 AS canteosal, Sum(alm_ingresodet.cantidad) AS canrealing, 0 AS canrealsal " _
        + vbCr + "FROM pro_producciondet INNER JOIN (((pro_ordenprod INNER JOIN (pro_solicitudmat INNER JOIN (pro_solicitudmat AS pro_solicitudmat_1 INNER JOIN (alm_ingreso INNER JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id) ON pro_solicitudmat_1.id = alm_ingreso.iddocref) ON pro_solicitudmat.id = pro_solicitudmat_1.iddocref) ON pro_ordenprod.id = pro_solicitudmat.iddocref) INNER JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) ON pro_producciondet.idord = pro_ordenprod.id " _
        + vbCr + "WHERE (((pro_solicitudmat.idtipdocref) = " & mIdTipoOrdProduccion & ") And ((pro_solicitudmat_1.idtipdocref) = " & mIdTipoSolMateriales & ") And ((alm_ingreso.idtipdocref) = " & mIdTipoSolMateriales & ") And ((alm_ingreso.tipmov) = True) And ((pro_producciondet.idproddet) = " & IdParteProduccionDet & ")) " _
        + vbCr + "GROUP BY pro_ordenprod.id, pro_solicitudmat_1.id, alm_ingresodet.id, alm_ingresodet.iditem, alm_inventario.idunimed, alm_inventario.descripcion, mae_unidades.abrev, 0 " _
        + vbCr + "UNION " _
        + vbCr + "SELECT pro_ordenprod.id AS idord, pro_solicitudmat.id AS idsol, alm_ingresodet.id AS idmov, alm_ingresodet.iditem, alm_inventario.idunimed, alm_inventario.descripcion AS item, mae_unidades.abrev AS unimed, 0 AS canteoing, Sum(alm_ingresodet.cantteo) AS canteosal, 0 AS canrealing, Sum(alm_ingresodet.cantidad) AS canrealsal " _
        + vbCr + "FROM pro_producciondet INNER JOIN (((((pro_ordenprod INNER JOIN pro_solicitudmat ON pro_ordenprod.id = pro_solicitudmat.iddocref) INNER JOIN alm_ingreso ON pro_solicitudmat.id = alm_ingreso.iddocref) INNER JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id) INNER JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) ON pro_producciondet.idord = pro_ordenprod.id " _
        + vbCr + "WHERE (((pro_solicitudmat.idtipdocref) = " & mIdTipoOrdProduccion & ") And ((alm_ingreso.tipmov) = False) And ((alm_ingreso.idtipdocref) = " & mIdTipoSolMateriales & ") And ((pro_producciondet.idproddet) = " & IdParteProduccionDet & ")) " _
        + vbCr + "GROUP BY pro_ordenprod.id, pro_solicitudmat.id, alm_ingresodet.id, alm_ingresodet.iditem, alm_inventario.idunimed, alm_inventario.descripcion, mae_unidades.abrev, 0, 0 " _
        + vbCr + "UNION " _
        + vbCr + "SELECT pro_ordenprod.id AS idord, pro_solicitudmat_1.id AS idsol, alm_ingresodet.id AS idmov, alm_ingresodet.iditem, alm_inventario.idunimed, alm_inventario.descripcion AS item, mae_unidades.abrev AS unimed, 0 AS canteoing, Sum(alm_ingresodet.cantteo) AS canteosal, 0 AS canrealing, Sum(alm_ingresodet.cantidad) AS canrealsal " _
        + vbCr + "FROM pro_producciondet INNER JOIN (((pro_ordenprod INNER JOIN (pro_solicitudmat INNER JOIN (pro_solicitudmat AS pro_solicitudmat_1 INNER JOIN (alm_ingreso INNER JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id) ON pro_solicitudmat_1.id = alm_ingreso.iddocref) ON pro_solicitudmat.id = pro_solicitudmat_1.iddocref) ON pro_ordenprod.id = pro_solicitudmat.iddocref) INNER JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) ON pro_producciondet.idord = pro_ordenprod.id " _
        + vbCr + "WHERE (((pro_solicitudmat.idtipdocref) = " & mIdTipoOrdProduccion & ") And ((pro_solicitudmat_1.idtipdocref) = " & mIdTipoSolMateriales & ") And ((alm_ingreso.idtipdocref) = " & mIdTipoSolMateriales & ") And ((alm_ingreso.tipmov) = False) And ((pro_producciondet.idproddet) = " & IdParteProduccionDet & ")) " _
        + vbCr + "GROUP BY pro_ordenprod.id, pro_solicitudmat_1.id, alm_ingresodet.id, alm_ingresodet.iditem, alm_inventario.idunimed, alm_inventario.descripcion, mae_unidades.abrev, 0, 0, 0 "
    
    cSQL = "SELECT pro_produccion.id AS idprod, pro_producciondet.idproddet, pro_produccion.fchdoc AS fecha, pro_producciondet.idrec, pro_producciondet.iditem AS iditemprod, conmov.idunimed, mae_unidades.abrev AS unimed, pro_producciondet.canprog AS canteoprod, pro_producciondet.cantidad AS canrealprod, conmov.iditem, Sum(conmov.canteoing) AS canteoing, Sum(conmov.canteosal) AS canteosal, Sum(conmov.canrealing) AS canrealing, Sum(conmov.canrealsal) AS canrealsal, Sum([conmov].[canteosal]-[conmov].[canteoing]) AS canteomov, Sum([conmov].[canrealsal]-[conmov].[canrealing]) AS canrealmov, alm_inventario.codpro, alm_inventario.descripcion AS item " _
        + vbCr + "FROM (pro_produccion INNER JOIN ((pro_producciondet INNER JOIN " _
        + vbCr + "( " _
        + vbCr + cSQLInterna _
        + vbCr + ") " _
        + vbCr + "AS conmov ON pro_producciondet.idord = conmov.idord) INNER JOIN mae_unidades ON conmov.idunimed = mae_unidades.id) ON pro_produccion.id = pro_producciondet.idpro) INNER JOIN alm_inventario ON conmov.iditem = alm_inventario.id " _
        + vbCr + "GROUP BY pro_produccion.id, pro_producciondet.idproddet, pro_produccion.fchdoc, pro_producciondet.idrec, pro_producciondet.iditem, conmov.idunimed, mae_unidades.abrev, pro_producciondet.canprog, pro_producciondet.cantidad, conmov.iditem, alm_inventario.codpro, alm_inventario.descripcion"
        
    database.CommandText = cSQL
    
    Set mRecord = database.GetRecordset
    If mRecord.RecordCount > 0 Then
        mRecord.MoveFirst
        While Not mRecord.EOF
            Dim mParteProduccionDetIns As New EParteProdDetIns
            Set mParteProduccionDetIns.Conexion = mConexion
            mParteProduccionDetIns.SetEntity mRecord
            Me.Add mParteProduccionDetIns
            Set mParteProduccionDetIns = Nothing
            mRecord.MoveNext
        Wend
    End If
    Exit Sub
    
BLoqueError:
    Err.Raise Err.Number, "[Fetch] " & Err.Source, Err.Description
End Sub



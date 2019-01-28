VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LEParteProdDetInsMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mLEParteProdDetInsMov As New Collection
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

Public Sub Add(mEParteProdDetInsMov As EParteProdDetInsMov)
    mLEParteProdDetInsMov.Add mEParteProdDetInsMov
End Sub

Public Function Count() As Long
    Count = mLEParteProdDetInsMov.Count
End Function

Public Sub Delete(ByVal Index As Variant)
    mLEParteProdDetInsMov.Remove Index
End Sub

Public Function Item(ByVal Index As Variant) As EParteProdDetInsMov
Attribute Item.VB_UserMemId = 0
    Set Item = mLEParteProdDetInsMov.Item(Index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mLEParteProdDetInsMov.[_NewEnum]
End Function

Public Sub Fetch(IdParteProduccionDet As Long, IdItem As Long)
    Dim cSQL As String
    Dim cSQLInterna As String
    Dim mIdTipoOrdProduccion As Long
    Dim database As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set database.Connection = mConexion
    
    mIdTipoOrdProduccion = F.NuloNumeric(F.KeyValue("OrdenProduccion", mConexion))
    
    
    cSQLInterna = "SELECT pro_ordenprod.id AS idord, pro_solicitudmat.id AS idsol, alm_ingresodet.id AS idmov, alm_ingresodet.idmovdet AS idmovdet, alm_ingresodet.iditem, alm_inventario.idunimed, alm_inventario.descripcion AS item, mae_unidades.abrev AS unimed, Sum(alm_ingresodet.cantteo) AS canteoing, 0 AS canteosal, Sum(alm_ingresodet.cantidad) AS canrealing, 0 AS canrealsal " _
        + vbCr + "FROM pro_producciondet INNER JOIN (((((pro_ordenprod INNER JOIN pro_solicitudmat ON pro_ordenprod.id = pro_solicitudmat.iddocref) INNER JOIN alm_ingreso ON pro_solicitudmat.id = alm_ingreso.iddocref) INNER JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id) INNER JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) ON pro_producciondet.idord = pro_ordenprod.id " _
        + vbCr + "WHERE (((alm_ingresodet.cantidad) > 0) AND ((pro_solicitudmat.idtipdocref) = " & mIdTipoOrdProduccion & ") And ((alm_ingreso.tipmov) = True) And ((pro_producciondet.idproddet) = " & IdParteProduccionDet & ")) " _
        + vbCr + "GROUP BY pro_ordenprod.id, pro_solicitudmat.id, alm_ingresodet.id, alm_ingresodet.idmovdet, alm_ingresodet.iditem, alm_inventario.idunimed, alm_inventario.descripcion, mae_unidades.abrev, 0 " _
        + vbCr + "HAVING (((alm_ingresodet.IdItem) = " & IdItem & ")) " _
        + vbCr + "UNION " _
        + vbCr + "SELECT pro_ordenprod.id AS idord, pro_solicitudmat_1.id AS idsol, alm_ingresodet.id AS idmov, alm_ingresodet.idmovdet AS idmovdet, alm_ingresodet.iditem, alm_inventario.idunimed, alm_inventario.descripcion AS item, mae_unidades.abrev AS unimed, Sum(alm_ingresodet.cantteo) AS canteoing, 0 AS canteosal, Sum(alm_ingresodet.cantidad) AS canrealing, 0 AS canrealsal " _
        + vbCr + "FROM pro_producciondet INNER JOIN (((pro_ordenprod INNER JOIN (pro_solicitudmat INNER JOIN (pro_solicitudmat AS pro_solicitudmat_1 INNER JOIN (alm_ingreso INNER JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id) ON pro_solicitudmat_1.id = alm_ingreso.iddocref) ON pro_solicitudmat.id = pro_solicitudmat_1.iddocref) ON pro_ordenprod.id = pro_solicitudmat.iddocref) INNER JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) ON pro_producciondet.idord = pro_ordenprod.id " _
        + vbCr + "WHERE (((alm_ingresodet.cantidad) > 0) AND ((pro_solicitudmat.idtipdocref) = " & mIdTipoOrdProduccion & ") And ((alm_ingreso.tipmov) = True) And ((pro_producciondet.idproddet) = " & IdParteProduccionDet & ")) " _
        + vbCr + "GROUP BY pro_ordenprod.id, pro_solicitudmat_1.id, alm_ingresodet.id, alm_ingresodet.idmovdet, alm_ingresodet.iditem, alm_inventario.idunimed, alm_inventario.descripcion, mae_unidades.abrev, 0 " _
        + vbCr + "HAVING (((alm_ingresodet.IdItem) = " & IdItem & ")) " _
        + vbCr + "UNION " _
        + vbCr + "SELECT pro_ordenprod.id AS idord, pro_solicitudmat.id AS idsol, alm_ingresodet.id AS idmov, alm_ingresodet.idmovdet AS idmovdet, alm_ingresodet.iditem, alm_inventario.idunimed, alm_inventario.descripcion AS item, mae_unidades.abrev AS unimed, 0 AS canteoing, Sum(alm_ingresodet.cantteo) AS canteosal, 0 AS canrealing, Sum(alm_ingresodet.cantidad) AS canrealsal " _
        + vbCr + "FROM pro_producciondet INNER JOIN (((((pro_ordenprod INNER JOIN pro_solicitudmat ON pro_ordenprod.id = pro_solicitudmat.iddocref) INNER JOIN alm_ingreso ON pro_solicitudmat.id = alm_ingreso.iddocref) INNER JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id) INNER JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) ON pro_producciondet.idord = pro_ordenprod.id " _
        + vbCr + "WHERE (((alm_ingresodet.cantidad) > 0) AND ((pro_solicitudmat.idtipdocref) = " & mIdTipoOrdProduccion & ") And ((alm_ingreso.tipmov) = False) And ((pro_producciondet.idproddet) = " & IdParteProduccionDet & ")) " _
        + vbCr + "GROUP BY pro_ordenprod.id, pro_solicitudmat.id, alm_ingresodet.id, alm_ingresodet.idmovdet, alm_ingresodet.iditem, alm_inventario.idunimed, alm_inventario.descripcion, mae_unidades.abrev, 0, 0 " _
        + vbCr + "HAVING (((alm_ingresodet.IdItem) = " & IdItem & ")) " _
        + vbCr + "UNION " _
        + vbCr + "SELECT pro_ordenprod.id AS idord, pro_solicitudmat_1.id AS idsol, alm_ingresodet.id AS idmov, alm_ingresodet.idmovdet AS idmovdet, alm_ingresodet.iditem, alm_inventario.idunimed, alm_inventario.descripcion AS item, mae_unidades.abrev AS unimed, 0 AS canteoing, Sum(alm_ingresodet.cantteo) AS canteosal, 0 AS canrealing, Sum(alm_ingresodet.cantidad) AS canrealsal " _
        + vbCr + "FROM pro_producciondet INNER JOIN (((pro_ordenprod INNER JOIN (pro_solicitudmat INNER JOIN (pro_solicitudmat AS pro_solicitudmat_1 INNER JOIN (alm_ingreso INNER JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id) ON pro_solicitudmat_1.id = alm_ingreso.iddocref) ON pro_solicitudmat.id = pro_solicitudmat_1.iddocref) ON pro_ordenprod.id = pro_solicitudmat.iddocref) INNER JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) ON pro_producciondet.idord = pro_ordenprod.id " _
        + vbCr + "WHERE (((alm_ingresodet.cantidad) > 0) AND ((pro_solicitudmat.idtipdocref) = " & mIdTipoOrdProduccion & ") And ((alm_ingreso.tipmov) = False) And ((pro_producciondet.idproddet) = " & IdParteProduccionDet & ")) " _
        + vbCr + "GROUP BY pro_ordenprod.id, pro_solicitudmat_1.id, alm_ingresodet.id, alm_ingresodet.idmovdet, alm_ingresodet.iditem, alm_inventario.idunimed, alm_inventario.descripcion, mae_unidades.abrev, 0, 0, 0 " _
        + vbCr + "HAVING (((alm_ingresodet.iditem)=" & IdItem & ")) "
        
    cSQL = "SELECT pro_producciondet.idproddet, conmov.idmovdet, conmov.idmov, conmov.idunimed, conmov.unimed, conmov.iditem, alm_inventario.codpro, conmov.item, [canteosal]-[canteoing] AS canteo, [canrealsal]-[canrealing] AS canreal " _
        + vbCr + "FROM pro_produccion INNER JOIN ((pro_producciondet INNER JOIN " _
        + vbCr + "( " _
        + vbCr + cSQLInterna _
        + vbCr + ") " _
        + vbCr + "AS conmov ON pro_producciondet.idord = conmov.idord) INNER JOIN alm_inventario ON conmov.iditem = alm_inventario.id) ON pro_produccion.id = pro_producciondet.idpro"
        
    database.CommandText = cSQL
    
    Set mRecord = database.GetRecordset
    If mRecord.RecordCount > 0 Then
        mRecord.MoveFirst
        While Not mRecord.EOF
            Dim mEParteProdDetInsMov As New EParteProdDetInsMov
            Set mEParteProdDetInsMov.Conexion = mConexion
            mEParteProdDetInsMov.SetEntity mRecord
            Me.Add mEParteProdDetInsMov
            Set mEParteProdDetInsMov = Nothing
            mRecord.MoveNext
        Wend
    End If
    Exit Sub
    
BLoqueError:
    Err.Raise Err.Number, "[Fetch] " & Err.Source, Err.Description
End Sub





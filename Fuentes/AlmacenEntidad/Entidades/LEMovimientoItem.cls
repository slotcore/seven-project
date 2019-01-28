VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LEMovimientoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mLMovimientoItem As New Collection
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

Public Sub Add(mMovimientoItem As EMovimientoItem)
    mLMovimientoItem.Add mMovimientoItem
End Sub

Public Function Count() As Long
    Count = mLMovimientoItem.Count
End Function

Public Sub Delete(ByVal Index As Variant)
    mLMovimientoItem.Remove Index
End Sub

Public Function Item(ByVal Index As Variant) As EMovimientoItem
Attribute Item.VB_UserMemId = 0
    Set Item = mLMovimientoItem.Item(Index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mLMovimientoItem.[_NewEnum]
End Function

Public Sub Fetch(IdItem As Long, IdAlmacen As Long, FechaInicio As Date, Optional FechaFin As Date)
    Dim cSQL As String
    Dim cWhere As String
    Dim cSQLInterna As String
    Dim dataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    cWhere = ""
    If Not IsNull(FechaInicio) And FechaInicio <> Empty Then
        cWhere = cWhere & " AND ((alm_ingreso.fchdoc) >= CDate('" & FechaInicio & "'))"
    End If
    If Not IsNull(FechaFin) And FechaFin <> Empty Then
        cWhere = cWhere & " AND ((alm_ingreso.fchdoc) <= CDate('" & FechaFin & "'))"
    End If

    ' Se valida fecha de inicio de movimientos del item
    cWhere = cWhere & " AND ((alm_ingreso.fchdoc) >= CDate('" & F.FechaInicioMovimientos(IdAlmacen, mConexion) & "'))"

    cSQL = "SELECT alm_ingresodet.idmovdet, con_librocostotemp.idlibrocostotemp, alm_inventario.idunimed, alm_ingreso.fchdoc AS fchmov, mae_unidades.abrev AS unimed, alm_ingresodet.iditem, alm_inventario.descripcion AS item, alm_ingresodet.cantidad, alm_ingreso.tipmov AS idtipmov, IIf([alm_ingreso].[tipmov]=-1,'I','S') AS tipmov, con_librocostotemp.costounitario, con_librocostotemp.costounitariopromedio, con_librocostotemp.costoprimo, con_librocostotemp.costomod, con_librocostotemp.costocif, alm_ingreso.idtipdocref AS idtipdocrefpadre, mae_documento.descripcion AS tipdocrefpadre, alm_ingreso.iddocref AS iddocrefpadre, '' AS docrefpadre, '' AS idtipdocref, '' AS tipdocref, alm_ingresodet.iddocref, '' AS docref, alm_ingreso.idalm, alm_almacenes.descripcion AS alm, [alm_ingreso].[numser] & '-' & [alm_ingreso].[numdoc] AS numdoc " _
        + vbCr + "FROM (((((alm_ingreso INNER JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet) INNER JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) INNER JOIN mae_documento ON alm_ingreso.idtipdocref = mae_documento.id) INNER JOIN alm_almacenes ON alm_ingreso.idalm = alm_almacenes.id " _
        + vbCr + "WHERE ((alm_ingresodet.cantidad) > 0) AND ((alm_ingresodet.iditem)=" & IdItem & ") AND ((alm_ingreso.idalm)=" & IdAlmacen & ")" & cWhere _
        + vbCr + "ORDER BY alm_ingreso.fchdoc"

    dataBase.CommandText = cSQL

    Set mRecord = dataBase.GetRecordset
    ' Se ordena el recordset
    mRecord.Sort = "fchmov, tipmov"

    If mRecord.RecordCount > 0 Then
        mRecord.MoveFirst
        While Not mRecord.EOF
            Dim mParteProduccionDetIns As New EMovimientoItem
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





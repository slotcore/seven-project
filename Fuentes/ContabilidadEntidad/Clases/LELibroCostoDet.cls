VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LELibroCostoDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mLLibroCostoDet As New Collection
Private F As New SistemaLogica.Funciones
Private mConexion As New ADODB.Connection
Private mCalled As Boolean

Private Sub Class_Initialize()
    LastIndex = 1
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

Public Sub Add(mLibroCostoDetalle As ELibroCostoDet, Optional Key As String)
    If Key <> "" Then
        mLLibroCostoDet.Add mLibroCostoDetalle, Key
    Else
        mLLibroCostoDet.Add mLibroCostoDetalle
    End If
End Sub

Public Function Count() As Long
    Count = mLLibroCostoDet.Count
End Function

Public Sub Delete(ByVal Index As Variant)
    mLLibroCostoDet.Remove Index
End Sub

Public Function Item(ByVal Index As Variant) As ELibroCostoDet
Attribute Item.VB_UserMemId = 0
    Set Item = mLLibroCostoDet.Item(Index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mLLibroCostoDet.[_NewEnum]
End Function

Public Sub Fetch(IdLibroCosto As Long)
    Dim dataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
On Error GoTo BLoqueError
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    
    dataBase.CommandText = "SELECT con_librocostodet.id, con_librocostodet.idlibro, pro_produccion.id AS idparteprod, [pro_produccion].[numser] & '-' & [pro_produccion].[numdoc] AS parteprod, pro_producciondet.idproddet AS idpartedet, alm_ingresodet.idmovdet, pro_producciondet.iditem, alm_inventario.descripcion AS item, pro_produccion.fchdoc AS fecha, con_librocostodet.tipo, pro_receta.codrec AS receta, pro_producciondet.idunimed, mae_unidades.abrev AS unimed, con_librocostodet.cantidad, pro_producciondet.horini, pro_producciondet.horfin, con_librocostodet.impmprima, con_librocostodet.impmanobr, con_librocostodet.impgasfab, con_librocostodet.facdist " _
            + vbCr + "FROM (((((pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro) INNER JOIN con_librocostodet ON pro_producciondet.idproddet = con_librocostodet.idproddet) LEFT JOIN alm_ingresodet ON pro_producciondet.idproddet = alm_ingresodet.iddocref) INNER JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id) INNER JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) INNER JOIN mae_unidades ON pro_producciondet.idunimed = mae_unidades.id " _
            + vbCr + "WHERE (((con_librocostodet.idlibro)=" & IdLibroCosto & "))"
    
    Set mRecord = dataBase.GetRecordset
    If mRecord.RecordCount > 0 Then
        mRecord.Sort = "fecha"
        mRecord.MoveFirst
        While Not mRecord.EOF
            Dim mLibroCostoDet As New ELibroCostoDet
            Set mLibroCostoDet.Conexion = mConexion
            mLibroCostoDet.SetEntity mRecord
            Me.Add mLibroCostoDet, F.NuloString(mLibroCostoDet.IdLibroCostoDet)
            Set mLibroCostoDet = Nothing
            mRecord.MoveNext
        Wend
    End If
    Exit Sub
    
BLoqueError:
    Err.Raise Err.Number, "[Fetch] " & Err.Source, Err.Description
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LEInventarioIniDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mLInventarioIniDet As New Collection
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

Public Sub Add(mInventarioIniDet As EInventarioIniDet)
    mLInventarioIniDet.Add mInventarioIniDet
End Sub

Public Function Count() As Long
    Count = mLInventarioIniDet.Count
End Function

Public Sub Delete(ByVal Index As Variant)
    mLInventarioIniDet.Remove Index
End Sub

Public Function Item(ByVal Index As Variant) As EInventarioIniDet
Attribute Item.VB_UserMemId = 0
    Set Item = mLInventarioIniDet.Item(Index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mLInventarioIniDet.[_NewEnum]
End Function

Public Sub Fetch(IdInventarioInicial As Long)
    Dim cSQL As String
    Dim cWhere As String
    Dim cSQLInterna As String
    Dim mIdTipoOrdProduccion As Long
    Dim dataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
        
    cSQL = "SELECT alm_inventarioinicialdet.idinventarioinicialdet, alm_inventarioinicialdet.idinventarioinicial, alm_inventarioinicialdet.iditem, alm_inventario.codpro, alm_inventario.descripcion AS item, alm_inventarioinicialdet.idunimed, mae_unidades.abrev AS unimed, alm_inventarioinicialdet.cantidad, alm_inventarioinicialdet.costo " _
        + vbCr + "FROM (alm_inventarioinicialdet INNER JOIN alm_inventario ON alm_inventarioinicialdet.iditem = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventarioinicialdet.idunimed = mae_unidades.id " _
        + vbCr + "WHERE (((alm_inventarioinicialdet.idinventarioinicial)=" & IdInventarioInicial & "))"
    
    dataBase.CommandText = cSQL
    
    Set mRecord = dataBase.GetRecordset
    If mRecord.RecordCount > 0 Then
        mRecord.MoveFirst
        While Not mRecord.EOF
            Dim mInventarioIniDet As New EInventarioIniDet
            mInventarioIniDet.SetEntity mRecord
            Me.Add mInventarioIniDet
            Set mInventarioIniDet = Nothing
            mRecord.MoveNext
        Wend
    End If
    Exit Sub
    
BLoqueError:
    Set dataBase = Nothing
    MsgBox Err.Description, vbInformation + vbOKOnly + vbDefaultButton1, "[Fetch] " & Err.Source
End Sub





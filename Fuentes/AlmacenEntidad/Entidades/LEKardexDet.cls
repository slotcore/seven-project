VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LEKardexDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mLKardexDetalle As New Collection
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

Public Sub Add(mKardexDetalle As EKardexDet)
    mLKardexDetalle.Add mKardexDetalle
End Sub

Public Function Count() As Long
    Count = mLKardexDetalle.Count
End Function

Public Sub Delete(ByVal Index As Variant)
    mLKardexDetalle.Remove Index
End Sub

Public Function Item(ByVal Index As Variant) As EKardexDet
Attribute Item.VB_UserMemId = 0
    Set Item = mLKardexDetalle.Item(Index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mLKardexDetalle.[_NewEnum]
End Function

Public Sub Fetch(IdKardex As Long)
    Dim cSQL As String
    Dim cWhere As String
    Dim cSQLInterna As String
    Dim mIdTipoOrdProduccion As Long
    Dim dataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
        
    cSQL = "SELECT alm_kardexdet.idkardexdet, alm_kardexdet.idkardex, alm_kardexdet.idalm, alm_kardexdet.fecha, alm_kardexdet.cantidad, alm_kardexdet.costounitario, alm_kardexdet.costounitariopromedio " _
        + vbCr + "FROM alm_kardexdet " _
        + vbCr + "WHERE (((alm_kardexdet.idkardex)=" & IdKardex & "))"
    
    dataBase.CommandText = cSQL
    
    Set mRecord = dataBase.GetRecordset
    If mRecord.RecordCount > 0 Then
        mRecord.MoveFirst
        While Not mRecord.EOF
            Dim mKardexDetalle As New EKardexDet
            mKardexDetalle.LoadChild = mLoadChild
            mKardexDetalle.SetEntity mRecord
            Me.Add mKardexDetalle
            Set mKardexDetalle = Nothing
            mRecord.MoveNext
        Wend
    End If
    Exit Sub
    
BLoqueError:
    Set dataBase = Nothing
    MsgBox Err.Description, vbInformation + vbOKOnly + vbDefaultButton1, "[Fetch] " & Err.Source
End Sub



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LEParteProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mLParteProd As New Collection
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

Public Sub Add(mParteProd As EParteProd)
    mLParteProd.Add mParteProd
End Sub

Public Function Count() As Long
    Count = mLParteProd.Count
End Function

Public Sub Delete(ByVal Index As Variant)
    mLParteProd.Remove Index
End Sub

Public Function Item(ByVal Index As Variant) As EParteProd
Attribute Item.VB_UserMemId = 0
    Set Item = mLParteProd.Item(Index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mLParteProd.[_NewEnum]
End Function

Public Sub Fetch(Optional IdParteProd As String = "", _
                                        Optional FechaInicio As Date, _
                                        Optional FechaFin As Date)
    Dim cSQL As String
    Dim cWhere As String
    Dim cSQLInterna As String
    Dim mIdTipoOrdProduccion As Long
    Dim database As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set database.Connection = mConexion
    
    cWhere = ""
    If IdParteProd <> "" Then
        If cWhere <> "" Then cWhere = " AND "
        cWhere = cWhere & " AND ((pro_produccion.id)=" & IdParteProd & ")"
    Else
        If Not IsNull(FechaInicio) Then
            If cWhere <> "" Then cWhere = cWhere & " AND "
            cWhere = cWhere & "((pro_produccion.fchdoc) >= CDate('" & FechaInicio & "'))"
        End If
        If Not IsNull(FechaFin) Then
            If cWhere <> "" Then cWhere = cWhere & " AND "
            cWhere = cWhere & "((pro_produccion.fchdoc) <= CDate('" & FechaFin & "'))"
        End If
    End If
    cWhere = "WHERE " & cWhere
        
    cSQL = "SELECT pro_produccion.id AS idparteprod, pro_produccion.idresponsable, pro_produccion.idalm, pro_produccion.fchdoc, pro_produccion.numser, pro_produccion.numdoc, pla_empleados.nombre AS responsable, alm_almacenes.descripcion AS almacen, pro_produccion.glosa, pro_produccion.idmes AS mestrabajo, pro_produccion.ano AS anhotrabajo " _
        + vbCr + "FROM (pro_produccion INNER JOIN pla_empleados ON pro_produccion.idresponsable = pla_empleados.id) INNER JOIN alm_almacenes ON pro_produccion.idalm = alm_almacenes.id " _
        + vbCr + cWhere
        
    database.CommandText = cSQL
    
    Set mRecord = database.GetRecordset
    If mRecord.RecordCount > 0 Then
        mRecord.MoveFirst
        While Not mRecord.EOF
            Dim mParteProd As New EParteProd
            mParteProd.LoadChild = mLoadChild
            mParteProd.SetEntity mRecord
            Me.Add mParteProd
            Set mParteProd = Nothing
            mRecord.MoveNext
        Wend
    End If
    Exit Sub
    
BLoqueError:
    Err.Raise Err.Number, "[Fetch] " & Err.Source, Err.Description
End Sub

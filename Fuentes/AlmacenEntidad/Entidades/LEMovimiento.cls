VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LEMovimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mLMovimiento As New Collection
Private F As New SistemaLogica.Funciones
Private mConexion As New ADODB.Connection
Private mCalled As Boolean
Private mLoadChild As Boolean
Private mLastIndex As Long

Private Sub Class_Initialize()
    mCalled = False
    mLoadChild = True
    LastIndex = 1
End Sub

Private Sub Class_Terminate()
    'CloseConnection
End Sub

Public Property Get LastIndex() As Long
    LastIndex = mLastIndex
End Property
Public Property Let LastIndex(ByVal NewValue As Long)
    mLastIndex = NewValue
End Property


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

Public Sub Add(mMovimiento As EMovimiento)
    mMovimiento.Index = LastIndex
    mLMovimiento.Add mMovimiento
    LastIndex = LastIndex + 1
End Sub

Public Function Count() As Long
    Count = mLMovimiento.Count
End Function

Public Sub Delete(ByVal Index As Variant)
    Dim mIndice As Long
    mLMovimiento.Remove Index
    ' Se actualizan los indices de la coleccion
    For mIndice = 1 To mLMovimiento.Count
        mLMovimiento(mIndice).Index = mIndice
    Next
End Sub

Public Function Item(ByVal Index As Variant) As EMovimiento
    Set Item = mLMovimiento.Item(Index)
End Function

Public Function NewEnum() As IUnknown
    Set NewEnum = mLMovimiento.[_NewEnum]
End Function

Public Sub Fetch(IdMes As Long, IdAnho As Long)
    Dim cSQL As String
    Dim cWhere As String
    Dim cSQLInterna As String
    Dim mIdTipoOrdProduccion As Long
    Dim dataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
        
    cSQL = "SELECT alm_ingreso.id, alm_ingreso.fching AS fchmov, IIf(alm_ingreso!tipmov=-1,'ING.','SAL.') AS movi, [alm_ingreso]![numser]+'-'+[alm_ingreso]![numdoc] AS numdoc, alm_ingreso.nombre, mae_documento.abrev AS destipdocref, alm_ingreso.desdocref AS numdocref, IIf([alm_ingreso].[idope]=1,'RECEPCION',IIf([alm_ingreso].[idope]=2,'DESPACHO',IIf([alm_ingreso].[idope]=3,'ENTRADA PRODUCCION',IIf([alm_ingreso].[idope]=4,'SALIDA PRODUCCION','')))) AS desope, UCase([mae_estados].[descripcion]) AS desestado " _
        + vbCr + "FROM (alm_ingreso LEFT JOIN mae_estados ON alm_ingreso.estado = mae_estados.id) LEFT JOIN mae_documento ON alm_ingreso.idtipdocref = mae_documento.id " _
        + vbCr + "WHERE (((alm_ingreso.ano) = " & IdAnho & ") And ((alm_ingreso.idmes) = " & IdMes & ")) " _
        + vbCr + "ORDER BY Format([alm_ingreso].[fching],'Short Date') DESC;"
    
    dataBase.CommandText = cSQL
    
    Set mRecord = dataBase.GetRecordset
    If mRecord.RecordCount > 0 Then
        mRecord.MoveFirst
        While Not mRecord.EOF
            Dim mMovimiento As New EMovimiento
            mMovimiento.LoadChild = mLoadChild
            mMovimiento.SetEntity mRecord
            Me.Add mMovimiento
            Set mMovimiento = Nothing
            mRecord.MoveNext
        Wend
    End If
    Exit Sub
    
BLoqueError:
    Set dataBase = Nothing
    MsgBox Err.Description, vbInformation + vbOKOnly + vbDefaultButton1, "[Fetch] " & Err.Source
End Sub





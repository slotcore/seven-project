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
Private mSaldoInicial As Double
Private mCostoInicial As Double
Private mCostoUnitarioPromedioInicial As Double
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

Public Property Get SaldoInicial() As Double
    SaldoInicial = mSaldoInicial
End Property
Public Property Let SaldoInicial(ByVal NewValue As Double)
    mSaldoInicial = NewValue
End Property

Public Property Get CostoInicial() As Double
    CostoInicial = mCostoInicial
End Property
Public Property Let CostoInicial(ByVal NewValue As Double)
    mCostoInicial = NewValue
End Property

Public Property Get CostoUnitarioPromedioInicial() As Double
    CostoUnitarioPromedioInicial = mCostoUnitarioPromedioInicial
End Property
Public Property Let CostoUnitarioPromedioInicial(ByVal NewValue As Double)
    mCostoUnitarioPromedioInicial = NewValue
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
    
    mSaldoInicial = 0
    mCostoInicial = 0
    mCostoUnitarioPromedioInicial = 0
    
    ' Se verifica si la fecha de inicio es fecha de inicio de movimientos
    If FechaInicio = F.FechaInicioMovimientos(IdAlmacen, mConexion) Then
        ' Saldo Inicial
        mSaldoInicial = 0
        mCostoInicial = 0
        mCostoUnitarioPromedioInicial = 0
        ' Movimientos
        Set mRecord = Nothing
        Set mRecord = F.GeneraRstSQL(F.SQL_MovDetallado(CStr(IdItem), IdAlmacen, FechaInicio, FechaFin, mConexion), mConexion)
    Else
        '********************
        ' Saldo Inicial
        '********************
        Set mRecord = Nothing
        Set mRecord = F.GeneraRstSQL(F.SQL_MovHistoricoTotalizado(IdAlmacen, FechaInicio - 1, CStr(IdItem), mConexion), mConexion)
        If mRecord.RecordCount > 0 Then
            mRecord.MoveFirst
            mSaldoInicial = F.NuloNumeric(mRecord("canini")) + F.NuloNumeric(mRecord("canent")) - F.NuloNumeric(mRecord("cansal"))
            mCostoInicial = F.NuloNumeric(mRecord("costoini")) + F.NuloNumeric(mRecord("costoent")) - F.NuloNumeric(mRecord("costosal"))
            mCostoUnitarioPromedioInicial = F.NuloNumeric(mRecord("costouniprom"))
        End If
        '*************
        ' Movimientos
        '*************
        Set mRecord = Nothing
        Set mRecord = F.GeneraRstSQL(F.SQL_MovDetallado(CStr(IdItem), IdAlmacen, FechaInicio, FechaFin, mConexion, , False), mConexion)
    End If
    
    If mRecord.RecordCount > 0 Then
        mRecord.MoveFirst
        While Not mRecord.EOF
            Dim mMovimientoItem As New EMovimientoItem
            Set mMovimientoItem.Conexion = mConexion
            mMovimientoItem.SetEntity mRecord
            Me.Add mMovimientoItem
            Set mMovimientoItem = Nothing
            mRecord.MoveNext
        Wend
    End If
    Exit Sub
    
BLoqueError:
    Err.Raise Err.Number, "[Fetch] " & Err.Source, Err.Description
End Sub

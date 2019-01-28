VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ETransferenciaDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdTransferenciaDet As Long
Private mIdTransferencia As Long
Private mIdUnidadMedida As Long
Private mUnidadMedida As String
Private mIdItem As Long
Private mItem As String
Private mCantidad As Long
Private mLote As String
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

Public Property Get IdTransferenciaDet() As Long
    IdTransferenciaDet = mIdTransferenciaDet
End Property
Public Property Let IdTransferenciaDet(ByVal NewValue As Long)
    mIdTransferenciaDet = NewValue
End Property

Public Property Get IdTransferencia() As Long
    IdTransferencia = mIdTransferencia
End Property
Public Property Let IdTransferencia(ByVal NewValue As Long)
    mIdTransferencia = NewValue
End Property

Public Property Get IdUnidadMedida() As Long
    IdUnidadMedida = mIdUnidadMedida
End Property
Public Property Let IdUnidadMedida(ByVal NewValue As Long)
    mIdUnidadMedida = NewValue
End Property

Public Property Get UnidadMedida() As String
    UnidadMedida = mUnidadMedida
End Property
Public Property Let UnidadMedida(ByVal NewValue As String)
    mUnidadMedida = NewValue
End Property

Public Property Get IdItem() As Long
    IdItem = mIdItem
End Property
Public Property Let IdItem(ByVal NewValue As Long)
    mIdItem = NewValue
End Property

Public Property Get Item() As String
    Item = mItem
End Property
Public Property Let Item(ByVal NewValue As String)
    mItem = NewValue
End Property

Public Property Get Cantidad() As Long
    Cantidad = mCantidad
End Property
Public Property Let Cantidad(ByVal NewValue As Long)
    mCantidad = NewValue
End Property

Public Property Get Lote() As String
    Lote = mLote
End Property
Public Property Let Lote(ByVal NewValue As String)
    mLote = NewValue
End Property

Public Function Fetch(IdTransferencia As String) As ETransferenciaDet

End Function

Public Sub SetEntity(Record As ADODB.Recordset)
    Dim F As New SistemaLogica.Funciones

    Record.MoveFirst
    Me.IdTransferenciaDet = F.NuloString(Record("idtransferenciadet"))
    Me.IdTransferencia = F.NuloString(Record("idtransferencia"))
    Me.IdUnidadMedida = F.NuloString(Record("idunimed"))
    Me.UnidadMedida = F.NuloString(Record("unimed"))
    Me.IdItem = F.NuloString(Record("iditem"))
    Me.Item = F.NuloString(Record("item"))
    Me.Cantidad = F.NuloString(Record("cantidad"))
    Me.Lote = F.NuloString(Record("lote"))
    Set Record = Nothing
End Sub

Public Sub Save(IdUsuario As Long, MachineName As String)
On Error GoTo BLoqueError
    If (F.NuloNumeric(Me.IdTransferenciaDet) = 0) Then
        Insert IdUsuario, MachineName
    Else
        Update IdUsuario, MachineName
    End If
    Exit Sub
BLoqueError:
    Err.Raise &HFFFFFF01, , "No se pudo registrar el Transferencia por el siguiente motivo :" + Trim(Err.Description)
End Sub

Private Sub Insert(IdUsuario As Long, MachineName As String)
    Dim xId As Double
    Dim dataBase As New SistemaData.EDataBase
    Dim FData As New SistemaData.Funciones
    Dim F As New SistemaLogica.Funciones
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    xId = FData.GetCodeTable("alm_transferenciadet", "idtransferenciadet", mConexion)
    dataBase.ClearParameter

'    dataBase.CommandText = "PARAMETERS @idtransferenciadet Long, @idtransferencia Long, @iditem Long, @idunimed Long, @cantidad Double, @lote Text;" _
'            + vbCr + "INSERT INTO alm_transferenciadet(idtransferencia, iditem, idunimed, cantidad, lote) " _
'            + vbCr + "VALUES(@idtransferencia, @iditem, @idunimed, @cantidad, @lote)"

    dataBase.CommandText = "INSERT INTO alm_transferenciadet(idtransferenciadet, idtransferencia, iditem, idunimed, cantidad, lote) " _
            + vbCr + "VALUES(?, ?, ?, ?, ?, ?)"

    dataBase.AddParameter "@idtransferenciadet", adInteger, adParamInput, xId
    dataBase.AddParameter "@idtransferencia", adInteger, adParamInput, F.NuloNumeric(Me.IdTransferencia)
    dataBase.AddParameter "@iditem", adInteger, adParamInput, F.NuloNumeric(Me.IdItem)
    dataBase.AddParameter "@idunimed", adInteger, adParamInput, F.NuloNumeric(Me.IdUnidadMedida)
    dataBase.AddParameter "@cantidad", adDouble, adParamInput, F.NuloNumeric(Me.Cantidad)
    dataBase.AddParameter "@lote", adVarChar, adParamInput, F.NuloNumeric(Me.Lote)
    dataBase.Execute
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el Transferencia por el siguiente motivo :" + Trim(Err.Description)
End Sub

Private Sub Update(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    dataBase.ClearParameter

'    dataBase.CommandText = "PARAMETERS @idtransferencia Long, @iditem Long, @idunimed Long, @cantidad Double, @lote Text;" _
'            + vbCr + "UPDATE alm_transferenciadet SET iditem = @iditem, idunimed = @idunimed, cantidad = @cantidad, lote = @lote " _
'            + vbCr + "WHERE(alm_transferenciadet.idtransferenciadet = @idtransferenciadet)"
    
    dataBase.CommandText = "UPDATE alm_transferenciadet SET idtransferencia = ?, iditem = ?, idunimed = ?, cantidad = ?, lote = ? " _
            + vbCr + "WHERE(alm_transferenciadet.idtransferenciadet = @idtransferenciadet)"

    dataBase.AddParameter "@idtransferencia", adInteger, adParamInput, F.NuloNumeric(Me.IdTransferencia)
    dataBase.AddParameter "@iditem", adInteger, adParamInput, F.NuloNumeric(Me.IdItem)
    dataBase.AddParameter "@idunimed", adInteger, adParamInput, F.NuloNumeric(Me.IdUnidadMedida)
    dataBase.AddParameter "@cantidad", adDate, adParamInput, F.NuloNumeric(Me.Cantidad)
    dataBase.AddParameter "@lote", adVarChar, adParamInput, F.NuloString(Me.Lote)
    dataBase.AddParameter "@idtransferenciadet", adInteger, adParamInput, F.NuloNumeric(Me.IdTransferenciaDet)
    dataBase.Execute
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el Transferencia por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub

Public Sub Delete(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    dataBase.ClearParameter
    dataBase.CommandText = "PARAMETERS @idtransferenciadet Long;" _
            + vbCr + "DELETE FROM alm_transferenciadet " _
            + vbCr + "WHERE alm_transferenciadet.idtransferenciadet = @idtransferenciadet"
    
    dataBase.AddParameter "@idtransferenciadet", adInteger, adParamInput, Me.IdTransferenciaDet
    dataBase.Execute
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el Transferencia por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub


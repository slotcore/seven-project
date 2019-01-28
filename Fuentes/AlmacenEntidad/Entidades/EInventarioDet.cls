VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EInventarioDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdInventarioDet As Long
Private mIdInventario As Long
Private mIdUnidadMedida As Long
Private mUnidadMedida As String
Private mIdItem As Long
Private mItem As String
Private mCantidadInicial As Double
Private mCantidadCarga As Double
Private mCostoInicial As Double
Private mCostoCarga As Double
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

Public Property Get IdInventarioDet() As Long
    IdInventarioDet = mIdInventarioDet
End Property
Public Property Let IdInventarioDet(ByVal NewValue As Long)
    mIdInventarioDet = NewValue
End Property

Public Property Get IdInventario() As Long
    IdInventario = mIdInventario
End Property
Public Property Let IdInventario(ByVal NewValue As Long)
    mIdInventario = NewValue
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

Public Property Get CantidadInicial() As Double
    CantidadInicial = mCantidadInicial
End Property
Public Property Let CantidadInicial(ByVal NewValue As Double)
    mCantidadInicial = NewValue
End Property

Public Property Get CantidadCarga() As Double
    CantidadCarga = mCantidadCarga
End Property
Public Property Let CantidadCarga(ByVal NewValue As Double)
    mCantidadCarga = NewValue
End Property

Public Property Get CostoInicial() As Double
    CostoInicial = mCostoInicial
End Property
Public Property Let CostoInicial(ByVal NewValue As Double)
    mCostoInicial = NewValue
End Property

Public Property Get CostoCarga() As Double
    CostoCarga = mCostoCarga
End Property
Public Property Let CostoCarga(ByVal NewValue As Double)
    mCostoCarga = NewValue
End Property

Public Property Get Lote() As String
    Lote = mLote
End Property
Public Property Let Lote(ByVal NewValue As String)
    mLote = NewValue
End Property

Public Function Fetch(IdInventario As String) As EInventarioDet

End Function

Public Sub SetEntity(Record As ADODB.Recordset)
    Dim F As New SistemaLogica.Funciones

'    Record.MoveFirst
'    Me.IdInventarioDet = F.NuloString(Record("idInventariodet"))
'    Me.IdInventario = F.NuloString(Record("idInventario"))
'    Me.IdUnidadMedida = F.NuloString(Record("idunimed"))
'    Me.UnidadMedida = F.NuloString(Record("unimed"))
'    Me.IdItem = F.NuloString(Record("iditem"))
'    Me.Item = F.NuloString(Record("item"))
'    Me.Cantidad = F.NuloString(Record("cantidad"))
'    Me.Lote = F.NuloString(Record("lote"))
    Set Record = Nothing
End Sub

Public Sub Save(IdUsuario As Long, MachineName As String)
On Error GoTo BLoqueError
    If (F.NuloNumeric(Me.IdInventarioDet) = 0) Then
        Insert IdUsuario, MachineName
    Else
        Update IdUsuario, MachineName
    End If
    Exit Sub
BLoqueError:
    'Err.Raise &HFFFFFF01, , "No se pudo registrar el Inventario por el siguiente motivo :" + Trim(Err.Description)
    F.MostrarMensajeError Err.Description, Err.Source, ""
End Sub

Private Sub Insert(IdUsuario As Long, MachineName As String)
    Dim xId As Double
    Dim dataBase As New SistemaData.EDataBase
    Dim FData As New SistemaData.Funciones
    Dim F As New SistemaLogica.Funciones
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    xId = FData.GetCodeTable("alm_tomainventariodet", "idtomainventariodet", mConexion)
    dataBase.ClearParameter

    dataBase.CommandText = "INSERT INTO alm_tomainventariodet(idtomainventariodet, idtomainventario, iditem, idunimed, stockactual, preuniactual, cantidad, preuni, lote) " _
            + vbCr + "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?)"

    dataBase.AddParameter "@idtomainventariodet", adInteger, adParamInput, xId
    dataBase.AddParameter "@idtomainventario", adInteger, adParamInput, F.NuloNumeric(Me.IdInventario)
    dataBase.AddParameter "@iditem", adInteger, adParamInput, F.NuloNumeric(Me.IdItem)
    dataBase.AddParameter "@idunimed", adInteger, adParamInput, F.NuloNumeric(Me.IdUnidadMedida)
    dataBase.AddParameter "@stockactual", adDouble, adParamInput, F.NuloNumeric(Me.CantidadInicial)
    dataBase.AddParameter "@preuniactual", adDouble, adParamInput, F.NuloNumeric(Me.CostoInicial)
    dataBase.AddParameter "@cantidad", adDouble, adParamInput, F.NuloNumeric(Me.CantidadCarga)
    dataBase.AddParameter "@preuni", adDouble, adParamInput, F.NuloNumeric(Me.CostoCarga)
    dataBase.AddParameter "@lote", adVarChar, adParamInput, F.NuloString(Me.Lote)
    dataBase.Execute
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el Inventario por el siguiente motivo :" + Trim(Err.Description)
End Sub

Private Sub Update(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    dataBase.ClearParameter
    
    dataBase.CommandText = "UPDATE alm_tomainventariodet SET idtomainventario = ?, iditem = ?, idunimed = ?, stockactual = ?, preuniactual = ?, cantidad = ?, preuni = ?, lote = ? " _
            + vbCr + "WHERE(alm_tomainventariodet.idtomainventariodet = @idtomainventariodet)"

    dataBase.AddParameter "@idtomainventario", adInteger, adParamInput, F.NuloNumeric(Me.IdInventario)
    dataBase.AddParameter "@iditem", adInteger, adParamInput, F.NuloNumeric(Me.IdItem)
    dataBase.AddParameter "@idunimed", adInteger, adParamInput, F.NuloNumeric(Me.IdUnidadMedida)
    dataBase.AddParameter "@stockactual", adDouble, adParamInput, F.NuloNumeric(Me.CantidadInicial)
    dataBase.AddParameter "@preuniactual", adDouble, adParamInput, F.NuloNumeric(Me.CostoInicial)
    dataBase.AddParameter "@cantidad", adDouble, adParamInput, F.NuloNumeric(Me.CantidadCarga)
    dataBase.AddParameter "@preuni", adDouble, adParamInput, F.NuloNumeric(Me.CostoCarga)
    dataBase.AddParameter "@lote", adVarChar, adParamInput, F.NuloString(Me.Lote)
    dataBase.AddParameter "@idtomainventariodet", adInteger, adParamInput, F.NuloNumeric(Me.IdInventarioDet)
    dataBase.Execute
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el Inventario por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub

Public Sub Delete(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    dataBase.ClearParameter
    dataBase.CommandText = "DELETE FROM alm_tomainventariodet " _
            + vbCr + "WHERE alm_tomainventariodet.idtomainventariodet = ?"
    
    dataBase.AddParameter "@idtomainventariodet", adInteger, adParamInput, Me.IdInventarioDet
    dataBase.Execute
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el Inventario por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub




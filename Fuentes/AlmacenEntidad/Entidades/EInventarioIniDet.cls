VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EInventarioIniDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdInventarioInicialDet As Long
Private mIdInventarioInicial As Long
Private mIdUnidadMedida As Long
Private mUnidadMedida As String
Private mIdItem As Long
Private mCodigoItem As String
Private mItem As String
Private mCosto As Double
Private mCantidad As Double
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

Public Property Get IdInventarioInicialDet() As Long
    IdInventarioInicialDet = mIdInventarioInicialDet
End Property
Public Property Let IdInventarioInicialDet(ByVal NewValue As Long)
    mIdInventarioInicialDet = NewValue
End Property

Public Property Get IdInventarioInicial() As Long
    IdInventarioInicial = mIdInventarioInicial
End Property
Public Property Let IdInventarioInicial(ByVal NewValue As Long)
    mIdInventarioInicial = NewValue
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

Public Property Get CodigoItem() As String
    CodigoItem = mCodigoItem
End Property
Public Property Let CodigoItem(ByVal NewValue As String)
    mCodigoItem = NewValue
End Property

Public Property Get Item() As String
    Item = mItem
End Property
Public Property Let Item(ByVal NewValue As String)
    mItem = NewValue
End Property

Public Property Get Costo() As Double
    Costo = mCosto
End Property
Public Property Let Costo(ByVal NewValue As Double)
    mCosto = NewValue
End Property

Public Property Get Cantidad() As Double
    Cantidad = mCantidad
End Property
Public Property Let Cantidad(ByVal NewValue As Double)
    mCantidad = NewValue
End Property

Public Sub Fetch(IdInventarioInicialDet As String)
    Dim dataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
            
    dataBase.CommandText = "SELECT alm_inventarioinicialdet.idinventarioinicialdet, alm_inventarioinicialdet.idinventarioinicial, alm_inventarioinicialdet.iditem, alm_inventario.codpro, alm_inventario.descripcion AS item, alm_inventarioinicialdet.idunimed, mae_unidades.abrev AS unimed, alm_inventarioinicialdet.cantidad, alm_inventarioinicialdet.costo " _
        + vbCr + "FROM (alm_inventarioinicialdet INNER JOIN alm_inventario ON alm_inventarioinicialdet.iditem = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventarioinicialdet.idunimed = mae_unidades.id " _
        + vbCr + "WHERE (((alm_inventarioinicialdet.idinventarioinicialdet)=" & IdInventarioInicialDet & "))"
        
    Set mRecord = dataBase.GetRecordset
    If mRecord.RecordCount > 0 Then Me.SetEntity mRecord
    Exit Sub
    
BLoqueError:
    Err.Raise &HFFFFFF01, , "[Fetch] " + Trim(Err.Description)

End Sub

Public Sub SetEntity(Record As ADODB.Recordset)
    Me.IdInventarioInicialDet = F.NuloNumeric(Record("idinventarioinicialdet"))
    Me.IdInventarioInicial = F.NuloNumeric(Record("idinventarioinicial"))
    Me.IdItem = F.NuloNumeric(Record("iditem"))
    Me.CodigoItem = F.NuloString(Record("codpro"))
    Me.Item = F.NuloString(Record("item"))
    Me.IdUnidadMedida = F.NuloNumeric(Record("idunimed"))
    Me.UnidadMedida = F.NuloString(Record("unimed"))
    Me.Costo = F.NuloNumeric(Record("costo"))
    Me.Cantidad = F.NuloNumeric(Record("cantidad"))
End Sub

Public Function Save(IdUsuario As Long, MachineName As String) As Boolean
On Error GoTo BLoqueError
    If (F.NuloNumeric(Me.IdInventarioInicialDet) = 0) Then
        Insert IdUsuario, MachineName
    Else
        Update IdUsuario, MachineName
    End If
    Save = True
    Exit Function
    
BLoqueError:
    Save = False
    MsgBox Err.Description, vbInformation + vbOKOnly + vbDefaultButton1, "[Save] " & Err.Source
End Function

Private Sub Insert(IdUsuario As Long, MachineName As String)
    Dim xId As Double
    Dim dataBase As New SistemaData.EDataBase
    Dim FData As New SistemaData.Funciones
    Dim F As New SistemaLogica.Funciones
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    xId = FData.GetCodeTable("alm_inventarioinicialdet", "idinventarioinicialdet", mConexion)
    dataBase.ClearParameter
    
    dataBase.CommandText = "INSERT INTO alm_inventarioinicialdet(idinventarioinicialdet, idinventarioinicial, iditem, idunimed, cantidad, costo, fechacreacion, horacreacion, usuariocreacion, placacreacion) " _
            + vbCr + "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

    dataBase.AddParameter "@idinventarioinicialdet", adInteger, adParamInput, xId
    dataBase.AddParameter "@idinventarioinicial", adInteger, adParamInput, F.NuloNumeric(Me.IdInventarioInicial)
    dataBase.AddParameter "@iditem", adInteger, adParamInput, F.NuloNumeric(Me.IdItem)
    dataBase.AddParameter "@idunimed", adInteger, adParamInput, F.NuloNumeric(Me.IdUnidadMedida)
    dataBase.AddParameter "@cantidad", adDouble, adParamInput, F.NuloNumeric(Me.Cantidad)
    dataBase.AddParameter "@costo", adDouble, adParamInput, F.NuloNumeric(Me.Costo)
    '************************************
    ' Parametros de auditoria
    dataBase.AddParameter "@fechacreacion", adInteger, adParamInput, F.ConvertirFechaANumero(Date)
    dataBase.AddParameter "@horacreacion", adInteger, adParamInput, F.ConvertirHoraANumero(Time)
    dataBase.AddParameter "@usuariocreacion", adVarChar, adParamInput, IdUsuario
    dataBase.AddParameter "@placacreacion", adVarChar, adParamInput, MachineName
    '************************************
    dataBase.Execute
    Me.IdInventarioInicialDet = xId
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el ParteProduccion por el siguiente motivo :" + Trim(Err.Description)
End Sub

Private Sub Update(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    dataBase.ClearParameter

    dataBase.CommandText = "UPDATE alm_inventarioinicialdet SET idinventarioinicial = ?, iditem = ?, idunimed = ?, cantidad = ? " _
             & ", costo = ?, fechamodificacion = ?, horamodificacion = ?, usuariomodificacion = ?, placamodificacion = ? " _
            + vbCr + "WHERE (alm_inventarioinicialdet.idinventarioinicialdet = ?)"

    dataBase.AddParameter "@idinventarioinicial", adInteger, adParamInput, F.NuloNumeric(Me.IdInventarioInicial)
    dataBase.AddParameter "@iditem", adInteger, adParamInput, F.NuloNumeric(Me.IdItem)
    dataBase.AddParameter "@idunimed", adInteger, adParamInput, F.NuloNumeric(Me.IdUnidadMedida)
    dataBase.AddParameter "@cantidad", adDouble, adParamInput, F.NuloNumeric(Me.Cantidad)
    dataBase.AddParameter "@costo", adDouble, adParamInput, F.NuloNumeric(Me.Costo)
    '************************************
    ' Parametros de auditoria
    dataBase.AddParameter "@fechamodificacion", adInteger, adParamInput, F.ConvertirFechaANumero(Date)
    dataBase.AddParameter "@horamodificacion", adInteger, adParamInput, F.ConvertirHoraANumero(Time)
    dataBase.AddParameter "@usuariomodificacion", adVarChar, adParamInput, IdUsuario
    dataBase.AddParameter "@placamodificacion", adVarChar, adParamInput, MachineName
    '************************************
    dataBase.AddParameter "@idinventarioinicialdet", adInteger, adParamInput, F.NuloNumeric(Me.IdInventarioInicialDet)
    
    dataBase.Execute
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el ParteProduccion por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub

Public Sub Delete(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    dataBase.ClearParameter
    dataBase.CommandText = "DELETE FROM alm_inventarioinicialdet " _
            + vbCr + "WHERE alm_inventarioinicialdet.idinventarioinicialdet = ?"
    
    dataBase.AddParameter "@inventarioinicialdet", adInteger, adParamInput, Me.IdInventarioInicialDet
    dataBase.Execute
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el ParteProduccion por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub



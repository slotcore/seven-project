VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdKardex As Long
Private mIdItem As Long
Private mCantidad As Double
Private mCostoUnitario As Double
Private mCostoUnitarioPromedio As Double
Private mLKardexDetalle As New LEKardexDet

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

Public Property Get IdKardex() As Long
    IdKardex = mIdKardex
End Property
Public Property Let IdKardex(ByVal NewValue As Long)
    mIdKardex = NewValue
End Property

Public Property Get IdItem() As Long
    IdItem = mIdItem
End Property
Public Property Let IdItem(ByVal NewValue As Long)
    mIdItem = NewValue
End Property

Public Property Get Cantidad() As Double
    Cantidad = mCantidad
End Property
Public Property Let Cantidad(ByVal NewValue As Double)
    mCantidad = NewValue
End Property

Public Property Get CostoUnitario() As Double
    CostoUnitario = mCostoUnitario
End Property
Public Property Let CostoUnitario(ByVal NewValue As Double)
    mCostoUnitario = NewValue
End Property

Public Property Get CostoUnitarioPromedio() As Double
    CostoUnitarioPromedio = mCostoUnitarioPromedio
End Property
Public Property Let CostoUnitarioPromedio(ByVal NewValue As Double)
    mCostoUnitarioPromedio = NewValue
End Property

Public Property Get LKardexDetalle() As LEKardexDet
    Set LKardexDetalle = mLKardexDetalle
End Property
Public Property Set LKardexDetalle(ByVal NewValue As LEKardexDet)
    Set mLKardexDetalle = NewValue
End Property

Public Sub Fetch(IdKardex As Long, Optional IdItem As Long)
    Dim cWhere As String
    Dim dataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    cWhere = ""
    If IdItem > 0 Then
        cWhere = "WHERE ((alm_kardex.iditem)=" & IdItem & ")"
    Else
        cWhere = "WHERE ((alm_kardex.idkardex)=" & IdKardex & ")"
    End If
    dataBase.CommandText = "SELECT alm_kardex.idkardex, alm_kardex.iditem, alm_kardex.cantidad, alm_kardex.costounitario, alm_kardex.costounitariopromedio " _
        + vbCr + "FROM alm_kardex " _
        + vbCr + cWhere
    
    Set mRecord = dataBase.GetRecordset
    If mRecord.RecordCount > 0 Then Me.SetEntity mRecord
    Exit Sub
    
BLoqueError:
    Err.Raise Err.Number, "[Fetch] " & Err.Source, Err.Description
End Sub

Public Sub SetEntity(Record As ADODB.Recordset)
    Me.IdKardex = F.NuloNumeric(Record("IdKardex"))
    Me.IdItem = F.NuloNumeric(Record("IdItem"))
    Me.Cantidad = F.NuloNumeric(Record("Cantidad"))
    Me.CostoUnitario = F.NuloNumeric(Record("CostoUnitario"))
    Me.CostoUnitarioPromedio = F.NuloNumeric(Record("CostoUnitarioPromedio"))
    If mLoadChild Then
        Set Me.LKardexDetalle.Conexion = mConexion
        Me.LKardexDetalle.Fetch Me.IdKardex
    End If
End Sub

Public Function Save(IdUsuario As Long, MachineName As String) As Boolean
On Error GoTo BLoqueError
    If (F.NuloNumeric(Me.IdKardex) = 0) Then
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
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    xId = FData.GetCodeTable("alm_kardex", "idkardex", mConexion)
    
    Me.IdKardex = xId
                
    dataBase.CommandText = "INSERT INTO alm_kardex(idkardex, iditem, cantidad, costounitario, " _
                & "costounitariopromedio, fechacreacion, horacreacion, usuariocreacion, placacreacion) " _
            + vbCr + "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?)"
    
    dataBase.AddParameter "@idkardex", adInteger, adParamInput, xId
    dataBase.AddParameter "@iditem", adInteger, adParamInput, F.NuloNumeric(Me.IdItem)
    dataBase.AddParameter "@cantidad", adDouble, adParamInput, F.NuloNumeric(Me.Cantidad)
    dataBase.AddParameter "@costounitario", adDouble, adParamInput, F.NuloNumeric(Me.CostoUnitario)
    dataBase.AddParameter "@costounitariopromedio", adDouble, adParamInput, F.NuloNumeric(Me.CostoUnitarioPromedio)
    '************************************
    ' Parametros de auditoria
    dataBase.AddParameter "@fechacreacion", adInteger, adParamInput, F.ConvertirFechaANumero(Date)
    dataBase.AddParameter "@horacreacion", adInteger, adParamInput, F.ConvertirHoraANumero(Time)
    dataBase.AddParameter "@usuariocreacion", adVarChar, adParamInput, IdUsuario
    dataBase.AddParameter "@placacreacion", adVarChar, adParamInput, MachineName
    '************************************
    dataBase.Execute
    
    Dim KardexDet As EKardexDet
    For Each KardexDet In Me.LKardexDetalle
        If Not mConexion Is Nothing Then
            Set KardexDet.Conexion = mConexion
            KardexDet.Called = True
        End If
        KardexDet.IdKardex = xId
        If Not KardexDet.Save(IdUsuario, MachineName) Then Err.Raise &HFFFFFF01, "[Kardex - Insert] " & Err.Source, "Error al intentar grabar detalle"
        Set KardexDet = Nothing
    Next
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise Err.Number, "[Insert] " & Err.Source, Err.Description
End Sub

Private Sub Update(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    dataBase.CommandText = "UPDATE alm_kardex SET iditem = ?, cantidad = ?, costounitario = ?, " _
                & "costounitariopromedio = ?, fechamodificacion = ?, horamodificacion = ?, usuariomodificacion = ?, placamodificacion = ? " _
            + vbCr + "WHERE (alm_kardex.idkardex = ?)"
    
    dataBase.AddParameter "@iditem", adInteger, adParamInput, F.NuloNumeric(Me.IdItem)
    dataBase.AddParameter "@cantidad", adDouble, adParamInput, F.NuloNumeric(Me.Cantidad)
    dataBase.AddParameter "@costounitario", adDouble, adParamInput, F.NuloNumeric(Me.CostoUnitario)
    dataBase.AddParameter "@costounitariopromedio", adDouble, adParamInput, F.NuloNumeric(Me.CostoUnitarioPromedio)
    '************************************
    ' Parametros de auditoria
    dataBase.AddParameter "@fechamodificacion", adInteger, adParamInput, F.ConvertirFechaANumero(Date)
    dataBase.AddParameter "@horamodificacion", adInteger, adParamInput, F.ConvertirHoraANumero(Time)
    dataBase.AddParameter "@usuariomodificacion", adVarChar, adParamInput, IdUsuario
    dataBase.AddParameter "@placamodificacion", adVarChar, adParamInput, MachineName
    '************************************
    dataBase.AddParameter "@idkardex", adInteger, adParamInput, F.NuloNumeric(Me.IdKardex)
    dataBase.Execute
    
    Dim KardexDet As EKardexDet
    For Each KardexDet In Me.LKardexDetalle
        If Not mConexion Is Nothing Then
            KardexDet.Conexion = mConexion
            KardexDet.Called = True
        End If
        If Not KardexDet.Save(IdUsuario, MachineName) Then Err.Raise &HFFFFFF01, "[Kardex - Insert] " & Err.Source, "Error al intentar grabar detalle"
        Set KardexDet = Nothing
    Next
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise Err.Number, "[Update] " & Err.Source, Err.Description
    Exit Sub
End Sub

Public Sub Delete(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    ' Eliminamos a los Hijos
    DeleteChild IdUsuario, MachineName
    
    dataBase.ClearParameter
    dataBase.CommandText = "DELETE FROM alm_kardex " _
            + vbCr + "WHERE alm_kardex.idkardex = ?"
    dataBase.AddParameter "@idkardex", adInteger, adParamInput, Me.IdKardex
    dataBase.Execute
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    MsgBox Err.Description, vbInformation + vbOKOnly + vbDefaultButton1, "[Delete] " & Err.Source
    Exit Sub
End Sub

Private Sub DeleteChild(IdUsuario As Long, MachineName As String)
On Error GoTo BLoqueError

    Dim KardexDet As EKardexDet
    For Each KardexDet In Me.LKardexDetalle
        If Not mConexion Is Nothing Then
            KardexDet.Conexion = mConexion
            KardexDet.Called = True
        End If
        KardexDet.Delete IdUsuario, MachineName
        Set KardexDet = Nothing
    Next
    Exit Sub

BLoqueError:
    Set KardexDet = Nothing
    Err.Raise Err.Number, "[DeleteChild]" & Err.Source, Err.Description
    Exit Sub
End Sub




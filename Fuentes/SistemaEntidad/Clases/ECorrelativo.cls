VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ECorrelativo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdCorrelativo As Long
Private mIdTipoCorrelativo As Long
Private mTipoCorrelativo As String
Private mCorrelativo As Long

Private F As New SistemaLogica.Funciones
Private mIndex As Long
Private mConexion As New ADODB.Connection
Private mCalled As Boolean
Private mIsRecursive As Boolean
Private mLoadChild As Boolean

Public Property Get Index() As Long
    Index = mIndex
End Property
Public Property Let Index(ByVal NewValue As Long)
    mIndex = NewValue
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

Public Property Get IsRecursive() As Boolean
    IsRecursive = mIsRecursive
End Property
Public Property Let IsRecursive(ByVal NewValue As Boolean)
    mIsRecursive = NewValue
End Property

Public Property Get LoadChild() As Boolean
    LoadChild = mLoadChild
End Property
Public Property Let LoadChild(ByVal NewValue As Boolean)
    mLoadChild = NewValue
End Property

Public Property Get IdCorrelativo() As Long
    IdCorrelativo = mIdCorrelativo
End Property
Public Property Let IdCorrelativo(ByVal NewValue As Long)
    mIdCorrelativo = NewValue
End Property

Public Property Get IdTipoCorrelativo() As Long
    IdTipoCorrelativo = mIdTipoCorrelativo
End Property
Public Property Let IdTipoCorrelativo(ByVal NewValue As Long)
    mIdTipoCorrelativo = NewValue
End Property

Public Property Get TipoCorrelativo() As String
    TipoCorrelativo = mTipoCorrelativo
End Property
Public Property Let TipoCorrelativo(ByVal NewValue As String)
    mTipoCorrelativo = NewValue
End Property

Public Property Get Correlativo() As Long
    Correlativo = mCorrelativo
End Property
Public Property Let Correlativo(ByVal NewValue As Long)
    mCorrelativo = NewValue
End Property

Private Sub Class_Initialize()
    mCalled = False
    mIsRecursive = False
    mIndex = 0
End Sub

Public Sub Fetch(IdCorrelativo As String)
    Dim dataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BloqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
            
    dataBase.CommandText = "SELECT mae_correlativo.idcorrelativo, mae_correlativo.idtipocorrelativo, mae_correlativo.correlativo, mae_tipocorrelativo.descripcion AS tipocorrelativo " _
        + vbCr + "FROM mae_correlativo INNER JOIN mae_tipocorrelativo ON mae_tipocorrelativo.idtipocorrelativo = mae_correlativo.idtipocorrelativo " _
        + vbCr + "WHERE (((mae_correlativo.idcorrelativo)=" & IdCorrelativo & "))"
    
    Set mRecord = dataBase.GetRecordset
    If mRecord.RecordCount > 0 Then Me.SetEntity mRecord
    Exit Sub
    
BloqueError:
    Err.Raise Err.Number, "[Fetch] " & Err.Source, Err.Description
End Sub

Public Sub GetCorrelativo(CodigoCorrelativo As String, IdUsuario As Long, MachineName As String, Optional Aumento As Long = 1)
    Dim dataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
On Error GoTo BloqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
            
    dataBase.CommandText = "SELECT mae_correlativo.idcorrelativo, mae_correlativo.idtipocorrelativo, mae_correlativo.correlativo, mae_tipocorrelativo.descripcion AS tipocorrelativo " _
        + vbCr + "FROM mae_correlativo INNER JOIN mae_tipocorrelativo ON mae_tipocorrelativo.idtipocorrelativo = mae_correlativo.idtipocorrelativo " _
        + vbCr + "WHERE (((mae_tipocorrelativo.codigo) = '" & CodigoCorrelativo & "'))"
    
    Set mRecord = dataBase.GetRecordset
    If mRecord.RecordCount > 0 Then
        ' Se carga el objeto
        Me.SetEntity mRecord
        ' Se actualiza el correlativo
        Dim mCorrelativo As New ECorrelativo
        Set mCorrelativo.Conexion = mConexion
        mCorrelativo.Fetch Me.IdCorrelativo
        mCorrelativo.Correlativo = mCorrelativo.Correlativo + Aumento
        If Not mCorrelativo.Save(IdUsuario, MachineName) Then
            Err.Raise &HFFFFFF01, , "No se pudo actualizar el correlativo, IdCorrelativo: " & IdCorrelativo
        End If
    End If
    Exit Sub
    
BloqueError:
    Err.Raise Err.Number, "[GetCorrelativo] " & Err.Source, Err.Description
End Sub

Public Sub SetEntity(Record As ADODB.Recordset)
    Record.MoveFirst
    Me.IdCorrelativo = F.NuloNumeric(Record("IdCorrelativo"))
    Me.IdTipoCorrelativo = F.NuloNumeric(Record("IdTipoCorrelativo"))
    Me.TipoCorrelativo = F.NuloString(Record("TipoCorrelativo"))
    Me.Correlativo = F.NuloNumeric(Record("correlativo"))
    Set Record = Nothing
End Sub

Public Function Save(IdUsuario As Long, MachineName As String) As Boolean
On Error GoTo BloqueError
    If (F.NuloNumeric(Me.IdCorrelativo) = 0) Then
        Insert IdUsuario, MachineName
    Else
        Update IdUsuario, MachineName
    End If
    Save = True
    Exit Function
    
BloqueError:
    Save = False
    F.MostrarMensajeError Err.Description, "Grabar", "[Save] " & Err.Source
End Function

Private Sub Insert(IdUsuario As Long, MachineName As String)
    Dim xId As Double
    Dim dataBase As New SistemaData.EDataBase
    Dim FData As New SistemaData.Funciones
    Dim F As New SistemaLogica.Funciones
    
On Error GoTo BloqueError
    
'    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
'    If Not Called Then dataBase.BeginTrans
    
    dataBase.BeginTrans
    
    xId = FData.GetCodeTable("mae_correlativo", "idcorrelativo", mConexion)
    Me.IdCorrelativo = xId
    dataBase.CommandText = "INSERT INTO mae_correlativo " _
                & "(idcorrelativo, idtipocorrelativo, correlativo, " _
                & "fechacreacion, horacreacion, usuariocreacion, placacreacion) " _
            + vbCr + "VALUES(?, ?, ?, " _
                & "?, ?, ?, ?)"
    
    dataBase.AddParameter "@idcorrelativo", adInteger, adParamInput, xId
    dataBase.AddParameter "@idtipocorrelativo", adInteger, adParamInput, Me.IdTipoCorrelativo
    dataBase.AddParameter "@correlativo", adInteger, adParamInput, Me.Correlativo
    '************************************
    ' Parametros de auditoria
    dataBase.AddParameter "@fechacreacion", adInteger, adParamInput, F.ConvertirFechaANumero(Date)
    dataBase.AddParameter "@horacreacion", adInteger, adParamInput, F.ConvertirHoraANumero(Time)
    dataBase.AddParameter "@usuariocreacion", adVarChar, adParamInput, IdUsuario
    dataBase.AddParameter "@placacreacion", adVarChar, adParamInput, F.GetMachineName
    '************************************
    dataBase.Execute
    
    'If Not Called Then dataBase.CommitTrans
    
    dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BloqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , Err.Description
End Sub

Private Sub Update(IdUsuario As Long, MachineName As String)
    Dim xId As Double
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BloqueError

'    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
'    If Not Called Then dataBase.BeginTrans
    
    dataBase.BeginTrans
    xId = Me.IdCorrelativo
    
    dataBase.CommandText = "UPDATE mae_correlativo SET idtipocorrelativo = ?, correlativo = ?, " _
            & "fechamodificacion = ?, horamodificacion = ?, usuariomodificacion = ?, placamodificacion = ? " _
            + vbCr + "WHERE (mae_correlativo.idcorrelativo = ?)"
    
    dataBase.AddParameter "@idtipocorrelativo", adInteger, adParamInput, Me.IdTipoCorrelativo
    dataBase.AddParameter "@correlativo", adInteger, adParamInput, Me.Correlativo
    '************************************
    ' Parametros de auditoria
    dataBase.AddParameter "@fechamodificacion", adInteger, adParamInput, F.ConvertirFechaANumero(Date)
    dataBase.AddParameter "@horamodificacion", adInteger, adParamInput, F.ConvertirHoraANumero(Time)
    dataBase.AddParameter "@usuariomodificacion", adVarChar, adParamInput, IdUsuario
    dataBase.AddParameter "@placamodificacion", adVarChar, adParamInput, F.GetMachineName
    '************************************
    dataBase.AddParameter "@idcorrelativo", adInteger, adParamInput, xId
    dataBase.Execute
        
    'If Not Called Then dataBase.CommitTrans
    dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BloqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el Correlativo por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub

Public Sub Delete(IdUsuario As Long, MachineName As String)
    Dim xId As Double
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BloqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    dataBase.ClearParameter
    dataBase.CommandText = "DELETE FROM mae_correlativo " _
            + vbCr + "WHERE mae_correlativo.idcorrelativo = ?"
    
    dataBase.AddParameter "@idcorrelativo", adInteger, adParamInput, Me.IdCorrelativo
    dataBase.Execute
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BloqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "Delete, IdCorrelativo: " + Me.IdCorrelativo + ", Motivo :" + Trim(Err.Description)
    Exit Sub
End Sub



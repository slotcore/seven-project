VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EMetodoVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdMetodoValorizacion As Long
Private mAbreviacion As String
Private mDescripcion As String
Private mActivo As Integer
Private mIdFactorDistribucion As Long
Private mCodigoFactorDistribucion As String
Private mFactorDistribucion As String
Private mColumnaFactorDistribucion As String
Private mIdTipoDistribucion As Long
Private mCodigoTipoDistribucion As String
Private mTipoDistribucion As String
Private mIdTipoItemDistribucion As Long
Private mCodigoTipoItemDistribucion As String
Private mTipoItemDistribucion As String
Private mProcesaMP As Boolean
Private mProcesaMOD As Boolean
Private mProcesaCIF As Boolean

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

Public Property Get IdMetodoValorizacion() As Long
    IdMetodoValorizacion = mIdMetodoValorizacion
End Property
Public Property Let IdMetodoValorizacion(ByVal NewValue As Long)
    mIdMetodoValorizacion = NewValue
End Property

Public Property Get Abreviacion() As String
    Abreviacion = mAbreviacion
End Property
Public Property Let Abreviacion(ByVal NewValue As String)
    mAbreviacion = NewValue
End Property

Public Property Get Activo() As Integer
    Activo = mActivo
End Property
Public Property Let Activo(ByVal NewValue As Integer)
    mActivo = NewValue
End Property

Public Property Get Descripcion() As String
    Descripcion = mDescripcion
End Property
Public Property Let Descripcion(ByVal NewValue As String)
    mDescripcion = NewValue
End Property

Public Property Get IdFactorDistribucion() As Long
    IdFactorDistribucion = mIdFactorDistribucion
End Property
Public Property Let IdFactorDistribucion(ByVal NewValue As Long)
    mIdFactorDistribucion = NewValue
End Property

Public Property Get CodigoFactorDistribucion() As String
    CodigoFactorDistribucion = mCodigoFactorDistribucion
End Property
Public Property Let CodigoFactorDistribucion(ByVal NewValue As String)
    mCodigoFactorDistribucion = NewValue
End Property

Public Property Get FactorDistribucion() As String
    FactorDistribucion = mFactorDistribucion
End Property
Public Property Let FactorDistribucion(ByVal NewValue As String)
    mFactorDistribucion = NewValue
End Property

Public Property Get ColumnaFactorDistribucion() As String
    ColumnaFactorDistribucion = mColumnaFactorDistribucion
End Property
Public Property Let ColumnaFactorDistribucion(ByVal NewValue As String)
    mColumnaFactorDistribucion = NewValue
End Property

Public Property Get IdTipoDistribucion() As Long
    IdTipoDistribucion = mIdTipoDistribucion
End Property
Public Property Let IdTipoDistribucion(ByVal NewValue As Long)
    mIdTipoDistribucion = NewValue
End Property

Public Property Get CodigoTipoDistribucion() As String
    CodigoTipoDistribucion = mCodigoTipoDistribucion
End Property
Public Property Let CodigoTipoDistribucion(ByVal NewValue As String)
    mCodigoTipoDistribucion = NewValue
End Property

Public Property Get TipoDistribucion() As String
    TipoDistribucion = mTipoDistribucion
End Property
Public Property Let TipoDistribucion(ByVal NewValue As String)
    mTipoDistribucion = NewValue
End Property

Public Property Get IdTipoItemDistribucion() As Long
    IdTipoItemDistribucion = mIdTipoItemDistribucion
End Property
Public Property Let IdTipoItemDistribucion(ByVal NewValue As Long)
    mIdTipoItemDistribucion = NewValue
End Property

Public Property Get CodigoTipoItemDistribucion() As String
    CodigoTipoItemDistribucion = mCodigoTipoItemDistribucion
End Property
Public Property Let CodigoTipoItemDistribucion(ByVal NewValue As String)
    mCodigoTipoItemDistribucion = NewValue
End Property

Public Property Get TipoItemDistribucion() As String
    TipoItemDistribucion = mTipoItemDistribucion
End Property
Public Property Let TipoItemDistribucion(ByVal NewValue As String)
    mTipoItemDistribucion = NewValue
End Property

Public Property Get ProcesaMP() As Boolean
    ProcesaMP = mProcesaMP
End Property
Public Property Let ProcesaMP(ByVal NewValue As Boolean)
    mProcesaMP = NewValue
End Property

Public Property Get ProcesaMOD() As Boolean
    ProcesaMOD = mProcesaMOD
End Property
Public Property Let ProcesaMOD(ByVal NewValue As Boolean)
    mProcesaMOD = NewValue
End Property

Public Property Get ProcesaCIF() As Boolean
    ProcesaCIF = mProcesaCIF
End Property
Public Property Let ProcesaCIF(ByVal NewValue As Boolean)
    mProcesaCIF = NewValue
End Property

Public Sub Fetch(IdMetodoValorizacion As Long, Optional Activo As Long)
    Dim cWhere As String
    Dim dataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    cWhere = ""
    If Activo > 0 Then
        cWhere = "WHERE ((mae_metodoval.activo)=" & Activo & ")"
    Else
        cWhere = "WHERE ((mae_metodoval.id)=" & IdMetodoValorizacion & ")"
    End If
    dataBase.CommandText = "SELECT mae_metodoval.id, mae_metodoval.descripcion, mae_metodoval.abrev, mae_metodoval.activo, mae_metodoval.idfacdist, mae_metodovalfacdist.codigo AS codfacdist, mae_metodovalfacdist.descripcion AS desfacdist, mae_metodovalfacdist.columna AS colfacdist, mae_metodoval.idtipdist, mae_metodovaltipdist.codigo AS codtipdist, mae_metodovaltipdist.descripcion AS destipdist, mae_metodoval.idtipitemdist, mae_metodovaltipitemdist.codigo AS codtipitemdist, mae_metodovaltipitemdist.descripcion AS destipitemdist, mae_metodoval.procesamp, mae_metodoval.procesamod, mae_metodoval.procesacif " _
        + vbCr + "FROM ((mae_metodoval LEFT JOIN mae_metodovalfacdist ON mae_metodoval.idfacdist = mae_metodovalfacdist.id) LEFT JOIN mae_metodovaltipdist ON mae_metodoval.idtipdist = mae_metodovaltipdist.id) LEFT JOIN mae_metodovaltipitemdist ON mae_metodoval.idtipitemdist = mae_metodovaltipitemdist.id " _
        + vbCr + cWhere
    
    Set mRecord = dataBase.GetRecordset
    If mRecord.RecordCount > 0 Then Me.SetEntity mRecord
    Exit Sub
    
BLoqueError:
    Err.Raise Err.Number, "[Fetch] " & Err.Source, Err.Description
End Sub

Public Sub SetEntity(Record As ADODB.Recordset)
    Me.IdMetodoValorizacion = F.NuloNumeric(Record("id"))
    Me.Abreviacion = F.NuloString(Record("abrev"))
    Me.Descripcion = F.NuloString(Record("descripcion"))
    Me.Activo = Record("activo")
    Me.IdFactorDistribucion = F.NuloNumeric(Record("idfacdist"))
    Me.CodigoFactorDistribucion = F.NuloString(Record("codfacdist"))
    Me.FactorDistribucion = F.NuloString(Record("desfacdist"))
    Me.ColumnaFactorDistribucion = F.NuloString(Record("colfacdist"))
    Me.IdTipoDistribucion = F.NuloNumeric(Record("idtipdist"))
    Me.CodigoTipoDistribucion = F.NuloString(Record("codtipdist"))
    Me.TipoDistribucion = F.NuloString(Record("destipdist"))
    Me.IdTipoItemDistribucion = F.NuloNumeric(Record("idtipitemdist"))
    Me.CodigoTipoItemDistribucion = F.NuloString(Record("codtipitemdist"))
    Me.TipoItemDistribucion = F.NuloString(Record("destipitemdist"))
    Me.ProcesaMP = Record("procesamp")
    Me.ProcesaMOD = Record("procesamod")
    Me.ProcesaCIF = Record("procesacif")
End Sub

Public Function Save(IdUsuario As Long, MachineName As String) As Boolean
On Error GoTo BLoqueError
    If (F.NuloNumeric(Me.IdMetodoValorizacion) = 0) Then
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
    xId = FData.GetCodeTable("mae_metodoval", "id", mConexion)
    
    Me.IdMetodoValorizacion = xId
                
    dataBase.CommandText = "INSERT INTO mae_metodoval(id, descripcion, abrev, activo, idfacdist, idtipdist, idtipitemdist, procesamp, procesamod, procesacif, fechacreacion, horacreacion, usuariocreacion, placacreacion) " _
            + vbCr + "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    
    dataBase.AddParameter "@id", adInteger, adParamInput, xId
    dataBase.AddParameter "@descripcion", adVarChar, adParamInput, F.NuloString(Me.Descripcion)
    dataBase.AddParameter "@abrev", adVarChar, adParamInput, F.NuloString(Me.Abreviacion)
    dataBase.AddParameter "@activo", adInteger, adParamInput, F.NuloNumeric(Me.Activo)
    dataBase.AddParameter "@idfacdist", adInteger, adParamInput, F.NuloNumeric(Me.IdFactorDistribucion)
    dataBase.AddParameter "@idtipdist", adInteger, adParamInput, F.NuloNumeric(Me.IdTipoDistribucion)
    dataBase.AddParameter "@idtipitemdist", adInteger, adParamInput, F.NuloNumeric(Me.IdTipoItemDistribucion)
    dataBase.AddParameter "@procesamp", adBoolean, adParamInput, Me.ProcesaMP
    dataBase.AddParameter "@procesamod", adBoolean, adParamInput, Me.ProcesaMOD
    dataBase.AddParameter "@procesacif", adBoolean, adParamInput, Me.ProcesaCIF
    '************************************
    ' Parametros de auditoria
    dataBase.AddParameter "@fechacreacion", adInteger, adParamInput, F.ConvertirFechaANumero(Date)
    dataBase.AddParameter "@horacreacion", adInteger, adParamInput, F.ConvertirHoraANumero(Time)
    dataBase.AddParameter "@usuariocreacion", adVarChar, adParamInput, IdUsuario
    dataBase.AddParameter "@placacreacion", adVarChar, adParamInput, MachineName
    '************************************
    dataBase.Execute
    
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
    
    dataBase.CommandText = "UPDATE mae_metodoval SET descripcion = ?, abrev = ?, activo = ?, idfacdist = ?, idtipdist = ?, idtipitemdist = ?, procesamp = ?, procesamod = ?, procesacif = ?, fechamodificacion = ?, horamodificacion = ?, usuariomodificacion = ?, placamodificacion = ? " _
            + vbCr + "WHERE (mae_metodoval.id = ?)"
    
    dataBase.AddParameter "@descripcion", adVarChar, adParamInput, F.NuloString(Me.Descripcion)
    dataBase.AddParameter "@abrev", adVarChar, adParamInput, F.NuloString(Me.Abreviacion)
    dataBase.AddParameter "@activo", adInteger, adParamInput, F.NuloNumeric(Me.Activo)
    dataBase.AddParameter "@idfacdist", adInteger, adParamInput, F.NuloNumeric(Me.IdFactorDistribucion)
    dataBase.AddParameter "@idtipdist", adInteger, adParamInput, F.NuloNumeric(Me.IdTipoDistribucion)
    dataBase.AddParameter "@idtipitemdist", adInteger, adParamInput, F.NuloNumeric(Me.IdTipoItemDistribucion)
    dataBase.AddParameter "@procesamp", adBoolean, adParamInput, Me.ProcesaMP
    dataBase.AddParameter "@procesamod", adBoolean, adParamInput, Me.ProcesaMOD
    dataBase.AddParameter "@procesacif", adBoolean, adParamInput, Me.ProcesaCIF
    '************************************
    ' Parametros de auditoria
    dataBase.AddParameter "@fechamodificacion", adInteger, adParamInput, F.ConvertirFechaANumero(Date)
    dataBase.AddParameter "@horamodificacion", adInteger, adParamInput, F.ConvertirHoraANumero(Time)
    dataBase.AddParameter "@usuariomodificacion", adVarChar, adParamInput, IdUsuario
    dataBase.AddParameter "@placamodificacion", adVarChar, adParamInput, MachineName
    '************************************
    dataBase.AddParameter "@id", adInteger, adParamInput, F.NuloNumeric(Me.IdMetodoValorizacion)
    dataBase.Execute
    
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
    
    dataBase.ClearParameter
    dataBase.CommandText = "DELETE FROM mae_metodoval " _
            + vbCr + "WHERE mae_metodoval.id = ?"
    dataBase.AddParameter "@id", adInteger, adParamInput, Me.IdMetodoValorizacion
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



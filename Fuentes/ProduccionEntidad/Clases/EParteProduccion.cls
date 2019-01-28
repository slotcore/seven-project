VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EParteProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdParteProduccion As Long
Private mFechaParteProduccion As Date
Private mNumeroSerie As String
Private mNumeroDocumento As String
Private mIdResponsable As Long
Private mResponsable As String
Private mIdTipoDocumento As Long
Private mIdEstado As Long
Private mEstado As String
Private mIdAlmacen As Long
Private mAlmacen As String
Private mGlosa As String
Private mAnhoTrabajo As Long
Private mMesTrabajo As Long
Private mParteProduccionDetS As New LEParteProdDet
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

Public Property Get IdParteProduccion() As Long
    IdParteProduccion = mIdParteProduccion
End Property
Public Property Let IdParteProduccion(ByVal NewValue As Long)
    mIdParteProduccion = NewValue
End Property

Public Property Get FechaParteProduccion() As Date
    FechaParteProduccion = mFechaParteProduccion
End Property
Public Property Let FechaParteProduccion(ByVal NewValue As Date)
    mFechaParteProduccion = NewValue
End Property

Public Property Get NumeroSerie() As String
    NumeroSerie = mNumeroSerie
End Property
Public Property Let NumeroSerie(ByVal NewValue As String)
    mNumeroSerie = NewValue
End Property

Public Property Get NumeroDocumento() As String
    NumeroDocumento = mNumeroDocumento
End Property
Public Property Let NumeroDocumento(ByVal NewValue As String)
    mNumeroDocumento = NewValue
End Property

Public Property Get IdResponsable() As Long
    IdResponsable = mIdResponsable
End Property
Public Property Let IdResponsable(ByVal NewValue As Long)
    mIdResponsable = NewValue
End Property

Public Property Get Responsable() As String
    Responsable = mResponsable
End Property
Public Property Let Responsable(ByVal NewValue As String)
    mResponsable = NewValue
End Property

Public Property Get IdTipoDocumento() As Long
    IdTipoDocumento = mIdTipoDocumento
End Property
Public Property Let IdTipoDocumento(ByVal NewValue As Long)
    mIdTipoDocumento = NewValue
End Property

Public Property Get IdAlmacen() As Long
    IdAlmacen = mIdAlmacen
End Property
Public Property Let IdAlmacen(ByVal NewValue As Long)
    mIdAlmacen = NewValue
End Property

Public Property Get Almacen() As String
    Almacen = mAlmacen
End Property
Public Property Let Almacen(ByVal NewValue As String)
    mAlmacen = NewValue
End Property

Public Property Get IdEstado() As Long
    IdEstado = mIdEstado
End Property
Public Property Let IdEstado(ByVal NewValue As Long)
    mIdEstado = NewValue
End Property

Public Property Get Glosa() As String
    Glosa = mGlosa
End Property
Public Property Let Glosa(ByVal NewValue As String)
    mGlosa = NewValue
End Property

Public Property Get LParteProdDet() As LEParteProdDet
    Set LParteProdDet = mParteProduccionDetS
End Property
Public Property Set LParteProdDet(ByVal NewValue As LEParteProdDet)
    Set mParteProduccionDetS = NewValue
End Property

Public Property Get AnhoTrabajo() As Long
    AnhoTrabajo = mAnhoTrabajo
End Property
Public Property Let AnhoTrabajo(ByVal NewValue As Long)
    mAnhoTrabajo = NewValue
End Property

Public Property Get MesTrabajo() As Long
    MesTrabajo = mMesTrabajo
End Property
Public Property Let MesTrabajo(ByVal NewValue As Long)
    mMesTrabajo = NewValue
End Property

Public Sub Fetch(IdParteProduccion As Long)
    Dim database As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    Dim mParteProduccion As New EParteProd
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set database.Connection = mConexion
            
    database.CommandText = "SELECT pro_produccion.id AS idparteprod, pro_produccion.idresponsable, pro_produccion.idalm, pro_produccion.fchdoc, pro_produccion.numser, pro_produccion.numdoc, pla_empleados.nombre AS responsable, alm_almacenes.descripcion AS almacen, pro_produccion.glosa, pro_produccion.idmes AS mestrabajo, pro_produccion.ano AS anhotrabajo " _
            + vbCr + "FROM (pro_produccion INNER JOIN pla_empleados ON pro_produccion.idresponsable = pla_empleados.id) INNER JOIN alm_almacenes ON pro_produccion.idalm = alm_almacenes.id " _
            + vbCr + "WHERE (((pro_produccion.id)=" & IdParteProduccion & "))"
    
    Set mRecord = database.GetRecordset
    If mRecord.RecordCount > 0 Then Me.SetEntity mRecord
    Exit Sub
    
BLoqueError:
    Err.Raise Err.Number, "[Fetch] " & Err.Source, Err.Description
End Sub

Public Sub SetEntity(Record As ADODB.Recordset)
On Error GoTo BLoqueError
    
    Me.IdParteProduccion = F.NuloNumeric(Record("idparteprod"))
    Me.IdResponsable = F.NuloNumeric(Record("idresponsable"))
    Me.IdAlmacen = F.NuloNumeric(Record("idalm"))
    Me.FechaParteProduccion = Record("fchdoc")
    Me.NumeroSerie = F.NuloString(Record("numser"))
    Me.NumeroDocumento = F.NuloString(Record("numdoc"))
    Me.Responsable = F.NuloString(Record("responsable"))
    Me.Almacen = F.NuloString(Record("almacen"))
    Me.Glosa = F.NuloString(Record("glosa"))
    Me.MesTrabajo = F.NuloNumeric(Record("mestrabajo"))
    Me.AnhoTrabajo = F.NuloNumeric(Record("anhotrabajo"))
    If mLoadChild Then
        Set Me.LParteProdDet.Conexion = mConexion
        Me.LParteProdDet.Fetch Me.IdParteProduccion
    End If
    Exit Sub
    
BLoqueError:
    Err.Raise Err.Number, "[SetEntity] " & Err.Source, Err.Description
End Sub

Public Function Save(IdUsuario As Long, MachineName As String) As Boolean
On Error GoTo BLoqueError
    If (F.NuloNumeric(Me.IdParteProduccion) = 0) Then
        Insert IdUsuario, MachineName
    Else
        Update IdUsuario, MachineName
    End If
    Save = True
    Exit Function
    
BLoqueError:
    Save = False
    F.MostrarMensajeError "Error al Grabar :" + Err.Description, Err.Source
End Function

Private Sub Insert(IdUsuario As Long, MachineName As String)
    Dim xId As Long
    Dim database As New SistemaData.EDataBase
    Dim FData As New SistemaData.Funciones
    Dim F As New SistemaLogica.Funciones
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set database.Connection = mConexion
    If Not Called Then database.BeginTrans
    
    xId = FData.GetCodeTable("pro_produccion", "id", mConexion)
    Me.IdParteProduccion = xId
              
    database.CommandText = "INSERT INTO pro_produccion(id, fchdoc, idtipdoc, numser, " _
                & "numdoc, idresponsable, ano, idmes, idestado, glosa, idalm) " _
            + vbCr + "VALUES(?, ?, ?, ?, " _
                & "?, ?, ?, ?, ?, ?, ?)"
    
    database.AddParameter "@id", adInteger, adParamInput, xId
    database.AddParameter "@fchdoc", adDate, adParamInput, Me.FechaParteProduccion
    database.AddParameter "@idtipdoc", adInteger, adParamInput, F.NuloNumeric(Me.IdTipoDocumento)
    database.AddParameter "@numser", adVarChar, adParamInput, F.NuloString(Me.NumeroSerie)
    database.AddParameter "@numdoc", adVarChar, adParamInput, F.NuloString(Me.NumeroDocumento)
    database.AddParameter "@idresponsable", adInteger, adParamInput, F.NuloNumeric(Me.IdResponsable)
    database.AddParameter "@ano", adInteger, adParamInput, F.NuloNumeric(Me.AnhoTrabajo)
    database.AddParameter "@idmes", adInteger, adParamInput, F.NuloNumeric(Me.MesTrabajo)
    database.AddParameter "@idestado", adInteger, adParamInput, F.NuloNumeric(Me.IdEstado)
    database.AddParameter "@glosa", adVarChar, adParamInput, F.NuloString(Me.Glosa)
    database.AddParameter "@idalm", adInteger, adParamInput, F.NuloNumeric(Me.IdAlmacen)
    database.Execute
    
    Dim ParteProdDet As EParteProdDet
    For Each ParteProdDet In Me.LParteProdDet
        If Not mConexion Is Nothing Then
            Set ParteProdDet.Conexion = mConexion
            ParteProdDet.Called = True
        End If
        ParteProdDet.IdParteProduccion = xId
        ParteProdDet.Save IdUsuario, MachineName
    Next
    
    If Not Called Then database.CommitTrans
    Set database = Nothing
    Exit Sub
    
BLoqueError:
    If Not Called Then database.RollbackTrans
    Set database = Nothing
    Err.Raise &HFFFFFF01, , "Insert - ParteProduccion :" + Trim(Err.Description)
End Sub

Private Sub Update(IdUsuario As Long, MachineName As String)
    Dim xId As Long
    Dim database As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set database.Connection = mConexion
    If Not Called Then database.BeginTrans
    xId = Me.IdParteProduccion
    
    database.CommandText = "UPDATE pro_produccion SET fchdoc = ?, idtipdoc = ?, numser = ?, " _
                & "numdoc = ?, idresponsable = ?, ano = ?, idmes = ?, idestado = ?, glosa = ?, idalm = ? " _
            + vbCr + "WHERE (pro_produccion.id = ?)"
    
    database.AddParameter "@fchdoc", adDate, adParamInput, Me.FechaParteProduccion
    database.AddParameter "@idtipdoc", adInteger, adParamInput, F.NuloNumeric(Me.IdTipoDocumento)
    database.AddParameter "@numser", adVarChar, adParamInput, F.NuloString(Me.NumeroSerie), 4
    database.AddParameter "@numdoc", adVarChar, adParamInput, F.NuloString(Me.NumeroDocumento), 10
    database.AddParameter "@idresponsable", adInteger, adParamInput, F.NuloNumeric(Me.IdResponsable)
    database.AddParameter "@ano", adInteger, adParamInput, F.NuloNumeric(Me.AnhoTrabajo)
    database.AddParameter "@idmes", adInteger, adParamInput, F.NuloNumeric(Me.MesTrabajo)
    database.AddParameter "@idestado", adInteger, adParamInput, F.NuloNumeric(Me.IdEstado)
    database.AddParameter "@glosa", adVarChar, adParamInput, F.NuloString(Me.Glosa), 200
    database.AddParameter "@idalm", adInteger, adParamInput, F.NuloNumeric(Me.IdAlmacen)
    database.AddParameter "@id", adInteger, adParamInput, xId
    database.Execute
    
    ' Borramos a los hijos
    DeleteChild IdUsuario, MachineName
    ' Grabamos a los Hijos
    Dim ParteProdDet As EParteProdDet
    For Each ParteProdDet In Me.LParteProdDet
        If Not mConexion Is Nothing Then
            Set ParteProdDet.Conexion = mConexion
            ParteProdDet.Called = True
        End If
        ParteProdDet.IdParteProduccion = Me.IdParteProduccion
        ParteProdDet.Save IdUsuario, MachineName
    Next
    
    If Not Called Then database.CommitTrans
    Set database = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then database.RollbackTrans
    Set database = Nothing
    Err.Raise &HFFFFFF01, , "Update - ParteProduccion :" + Trim(Err.Description)
    Exit Sub
End Sub

Public Function Delete(IdUsuario As Long, MachineName As String) As Boolean
    Dim xId As Double
    Dim database As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set database.Connection = mConexion
    If Not Called Then database.BeginTrans
    
    ' Eliminamos a los Hijos
    DeleteChild IdUsuario, MachineName
    
    database.ClearParameter
    database.CommandText = "DELETE FROM pro_produccion " _
            + vbCr + "WHERE pro_produccion.id = ?"
    
    database.AddParameter "@id", adInteger, adParamInput, Me.IdParteProduccion
    database.Execute
    
    If Not Called Then database.CommitTrans
    Set database = Nothing
    Delete = True
    Exit Function

BLoqueError:
    If Not Called Then database.RollbackTrans
    Set database = Nothing
    Delete = False
    F.MostrarMensajeError "No se pudo registrar el ParteProduccion por el siguiente motivo :" + Err.Description, Err.Source, Err.Source
    Exit Function
End Function

Public Sub DeleteChild(IdUsuario As Long, MachineName As String)
    Dim database As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set database.Connection = mConexion
    
    database.ClearParameter

    database.CommandText = "DELETE FROM pro_producciondet " _
            + vbCr + "WHERE pro_producciondet.idpro = ?"

    database.AddParameter "@idpro", adInteger, adParamInput, Me.IdParteProduccion
    database.Execute
    
    Set database = Nothing
    Exit Sub

BLoqueError:
    Set database = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el ParteProduccion por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub





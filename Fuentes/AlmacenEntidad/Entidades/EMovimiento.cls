VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EMovimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdMovimiento As Long
Private mIdTipoMovimiento As Long
Private mTipoMovimiento As String
Private mFechaMovimiento As Date
Private mNumeroSerie As String
Private mNumeroDocumento As String
Private mIdResponsable As Long
Private mResponsable As String
Private mIdProveedor As Long
Private mProveedor As String
Private mIdTipoDocumento As Long
Private mIdTipoDocumentoReferencia As Long
Private mIdDocumentoReferencia As Long
Private mDocumentoReferencia As String
Private mIdTipoDocumentoReferencia2 As Long
Private mIdDocumentoReferencia2 As Long
Private mDocumentoReferencia2 As String
Private mIdEstado As Long
Private mEstado As String
Private mIdAlmacen As Long
Private mAlmacen As String
Private mGlosa As String
Private mLMovimientoDet As New LEMovimientoDet
Private mAnhoTrabajo As Long
Private mMesTrabajo As Long

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

Public Property Get IdMovimiento() As Long
    IdMovimiento = mIdMovimiento
End Property
Public Property Let IdMovimiento(ByVal NewValue As Long)
    mIdMovimiento = NewValue
End Property

Public Property Get IdTipoMovimiento() As Long
    IdTipoMovimiento = mIdTipoMovimiento
End Property
Public Property Let IdTipoMovimiento(ByVal NewValue As Long)
    mIdTipoMovimiento = NewValue
End Property

Public Property Get TipoMovimiento() As String
    TipoMovimiento = mTipoMovimiento
End Property
Public Property Let TipoMovimiento(ByVal NewValue As String)
    mTipoMovimiento = NewValue
End Property

Public Property Get FechaMovimiento() As Date
    FechaMovimiento = mFechaMovimiento
End Property
Public Property Let FechaMovimiento(ByVal NewValue As Date)
    mFechaMovimiento = NewValue
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

Public Property Get IdProveedor() As Long
    IdProveedor = mIdProveedor
End Property
Public Property Let IdProveedor(ByVal NewValue As Long)
    mIdProveedor = NewValue
End Property

Public Property Get Proveedor() As String
    Proveedor = mProveedor
End Property
Public Property Let Proveedor(ByVal NewValue As String)
    mProveedor = NewValue
End Property

Public Property Get IdTipoDocumento() As Long
    IdTipoDocumento = mIdTipoDocumento
End Property
Public Property Let IdTipoDocumento(ByVal NewValue As Long)
    mIdTipoDocumento = NewValue
End Property

Public Property Get IdTipoDocumentoReferencia() As Long
    IdTipoDocumentoReferencia = mIdTipoDocumentoReferencia
End Property
Public Property Let IdTipoDocumentoReferencia(ByVal NewValue As Long)
    mIdTipoDocumentoReferencia = NewValue
End Property

Public Property Get IdDocumentoReferencia() As Long
    IdDocumentoReferencia = mIdDocumentoReferencia
End Property
Public Property Let IdDocumentoReferencia(ByVal NewValue As Long)
    mIdDocumentoReferencia = NewValue
End Property

Public Property Get DocumentoReferencia() As String
    DocumentoReferencia = mDocumentoReferencia
End Property
Public Property Let DocumentoReferencia(ByVal NewValue As String)
    mDocumentoReferencia = NewValue
End Property

Public Property Get IdTipoDocumentoReferencia2() As Long
    IdTipoDocumentoReferencia2 = mIdTipoDocumentoReferencia2
End Property
Public Property Let IdTipoDocumentoReferencia2(ByVal NewValue As Long)
    mIdTipoDocumentoReferencia2 = NewValue
End Property

Public Property Get IdDocumentoReferencia2() As Long
    IdDocumentoReferencia2 = mIdDocumentoReferencia2
End Property
Public Property Let IdDocumentoReferencia2(ByVal NewValue As Long)
    mIdDocumentoReferencia2 = NewValue
End Property

Public Property Get DocumentoReferencia2() As String
    DocumentoReferencia2 = mDocumentoReferencia2
End Property
Public Property Let DocumentoReferencia2(ByVal NewValue As String)
    mDocumentoReferencia2 = NewValue
End Property

Public Property Get IdEstado() As Long
    IdEstado = mIdEstado
End Property
Public Property Let IdEstado(ByVal NewValue As Long)
    mIdEstado = NewValue
End Property

Public Property Get Estado() As String
    Estado = mEstado
End Property
Public Property Let Estado(ByVal NewValue As String)
    mEstado = NewValue
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

Public Property Get Glosa() As String
    Glosa = mGlosa
End Property
Public Property Let Glosa(ByVal NewValue As String)
    mGlosa = NewValue
End Property

Public Property Get LMovimientoDet() As LEMovimientoDet
    Set LMovimientoDet = mLMovimientoDet
End Property
Public Property Set LMovimientoDet(ByVal NewValue As LEMovimientoDet)
    Set mLMovimientoDet = NewValue
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

Private Sub Class_Initialize()
    mCalled = False
    mIsRecursive = False
    mIndex = 0
End Sub

Public Function Fetch(IdMovimiento As String) As EInventario

End Function

Public Sub SetEntity(Record As ADODB.Recordset)
    Dim M As New AlmacenEntidad.LEMovimientoDet
    
    Record.MoveFirst
    Me.IdMovimiento = F.NuloNumeric(Record("IdMovimiento"))
    Me.IdTipoMovimiento = F.NuloNumeric(Record("IdTipoMovimiento"))
    Me.FechaMovimiento = F.ConvertirNumeroAFecha(F.NuloNumeric(Record("FechaMovimiento")))
    Me.NumeroSerie = F.NuloString(Record("NumeroSerie"))
    Me.NumeroDocumento = F.NuloString(Record("NumeroDocumento"))
    Me.IdResponsable = F.NuloNumeric(Record("IdResponsable"))
    Me.Responsable = F.NuloString(Record("Responsable"))
    Me.IdProveedor = F.NuloNumeric(Record("Responsable"))
    Me.Proveedor = F.NuloString(Record("Responsable"))
    Me.IdTipoDocumento = F.NuloNumeric(Record("Responsable"))
    Me.IdTipoDocumentoReferencia = F.NuloNumeric(Record("Responsable"))
    Me.DocumentoReferencia = F.NuloString(Record("Responsable"))
    Me.IdEstado = F.NuloNumeric(Record("Responsable"))
    Me.Estado = F.NuloString(Record("Responsable"))
    Me.IdAlmacen = F.NuloNumeric(Record("Responsable"))
    Me.Almacen = F.NuloString(Record("Responsable"))
    Me.MesTrabajo = F.NuloNumeric(Record("mes"))
    Me.AnhoTrabajo = F.NuloNumeric(Record("anho"))
    Set Me.LMovimientoDet = M.Fetch(TransferenciaE, Me.IdMovimiento)
    Set Record = Nothing
End Sub

Public Function Save(IdUsuario As Long, MachineName As String) As Boolean
On Error GoTo BLoqueError
    If (F.NuloNumeric(Me.IdMovimiento) = 0) Then
        Insert IdUsuario, MachineName
    Else
        Update IdUsuario, MachineName
    End If
    Save = True
    Exit Function
    
BLoqueError:
    Save = False
    F.MostrarMensajeError Err.Description, "Grabar", "[Save] " & Err.Source
End Function

Private Sub Insert(IdUsuario As Long, MachineName As String)
    Dim xId As Double
    Dim IdMovDet As Long
    Dim dataBase As New SistemaData.EDataBase
    Dim FData As New SistemaData.Funciones
    Dim F As New SistemaLogica.Funciones
    
On Error GoTo BLoqueError
    
    '**************************
'    ' Utiliza el gestor de correlativos
'    ' Se coloca este fragmento de codigo antes de que se inicie BeginTrans
'    Dim mCorrelativo As New ECorrelativo
'    Set mCorrelativo.Conexion = mConexion
'    mCorrelativo.GetCorrelativo "IN", IdUsuario, MachineName
'    xId = mCorrelativo.Correlativo
'
'    Dim mCorrelativoDet As New ECorrelativo
'    Set mCorrelativoDet.Conexion = mConexion
'    mCorrelativoDet.GetCorrelativo "ID", IdUsuario, MachineName, Me.LMovimientoDet.Count
'    IdMovDet = mCorrelativoDet.Correlativo
    '**************************
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
        
    '************************
    ' Se quita debido a que se encuentran demasiadas concurrencias
    ' De Ids en procesos paralelos
    xId = FData.GetCodeTable("alm_ingreso", "id", mConexion)
    '************************
    
    Me.IdMovimiento = xId
    dataBase.CommandText = "INSERT INTO alm_ingreso(id, tipdoc, fching, fchdoc, " _
                & "numser, numdoc, idres, idpro, nombre, estado, tipmov, idtipdocref, " _
                & "iddocref, idtipdocref2, iddocref2, idalm, glosa, desdocref, ano, idmes) " _
            + vbCr + "VALUES(?, ?, ?, ?, " _
                & "?, ?, ?, ?, ?, ?, ?, ?, " _
                & "?, ?, ?, ?, ?, ?, ?, ?)"
    
    dataBase.AddParameter "@id", adInteger, adParamInput, xId
    dataBase.AddParameter "@tipdoc", adInteger, adParamInput, F.NuloNumeric(Me.IdTipoDocumento)
    dataBase.AddParameter "@fching", adDate, adParamInput, Me.FechaMovimiento
    dataBase.AddParameter "@fchdoc", adDate, adParamInput, Me.FechaMovimiento
    dataBase.AddParameter "@numser", adVarChar, adParamInput, F.NuloString(Me.NumeroSerie)
    dataBase.AddParameter "@numdoc", adVarChar, adParamInput, F.NuloString(Me.NumeroDocumento)
    dataBase.AddParameter "@idres", adInteger, adParamInput, F.NuloNumeric(Me.IdResponsable)
    dataBase.AddParameter "@idpro", adInteger, adParamInput, F.NuloNumeric(Me.IdProveedor)
    dataBase.AddParameter "@nombre", adVarChar, adParamInput, F.NuloString(Me.Proveedor)
    dataBase.AddParameter "@estado", adInteger, adParamInput, F.NuloNumeric(Me.IdEstado)
    dataBase.AddParameter "@tipmov", adInteger, adParamInput, F.NuloNumeric(Me.IdTipoMovimiento)
    dataBase.AddParameter "@idtipdocref", adInteger, adParamInput, F.NuloNumeric(Me.IdTipoDocumentoReferencia)
    dataBase.AddParameter "@iddocref", adInteger, adParamInput, F.NuloNumeric(Me.IdDocumentoReferencia)
    dataBase.AddParameter "@idtipdocref2", adInteger, adParamInput, F.NuloNumeric(Me.IdTipoDocumentoReferencia2)
    dataBase.AddParameter "@iddocref2", adInteger, adParamInput, F.NuloNumeric(Me.IdDocumentoReferencia2)
    dataBase.AddParameter "@idalm", adInteger, adParamInput, F.NuloNumeric(Me.IdAlmacen)
    dataBase.AddParameter "@glosa", adVarChar, adParamInput, F.NuloString(Me.Glosa)
    dataBase.AddParameter "@desdocref", adVarChar, adParamInput, F.NuloString(Me.DocumentoReferencia)
    dataBase.AddParameter "@ano", adInteger, adParamInput, F.NuloNumeric(Me.AnhoTrabajo)
    dataBase.AddParameter "@idmes", adInteger, adParamInput, F.NuloNumeric(Me.MesTrabajo)
    dataBase.Execute
    
    ' Genera Almacenaje Automatico
    If F.NuloNumeric(F.KeyValue("GeneraAlmacenajeAutomatico", mConexion)) = -1 Then
        If Not IsRecursive Then GeneraAlmacenajeAutomatico IdUsuario, MachineName
    End If
    
    Dim MovDet As EMovimientoDet
    '************************
    ' Se quita debido a que se encuentran demasiadas concurrencias
    ' De Ids en procesos paralelos
    IdMovDet = F.NuloNumeric(FData.GetCodeTable("alm_ingresodet", "idmovdet", mConexion))
    '************************
    For Each MovDet In Me.LMovimientoDet
        If Not mConexion Is Nothing Then
            Set MovDet.Conexion = mConexion
            MovDet.Called = True
        End If
        MovDet.MarkNew = True
        MovDet.IdMovimientoDet = IdMovDet
        MovDet.IdMovimiento = xId
        If Not MovDet.Save(IdUsuario, MachineName) Then Err.Raise &HFFFFFF01, , "Error al intentar grabar el movimiento"
        IdMovDet = IdMovDet + 1
    Next
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , Err.Description
End Sub

Private Sub Update(IdUsuario As Long, MachineName As String)
    Dim xId As Double
    Dim IdMovDet As Long
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim dataBase As New SistemaData.EDataBase
    Dim FData As New SistemaData.Funciones
    
On Error GoTo BLoqueError
    
'    Dim mCorrelativoDet As New ECorrelativo
'    Set mCorrelativoDet.Conexion = mConexion
'    mCorrelativoDet.GetCorrelativo "ID", IdUsuario, MachineName, Me.LMovimientoDet.Count
'    IdMovDet = mCorrelativoDet.Correlativo
'
    
    '************************
    ' Se quita debido a que se encuentran demasiadas concurrencias
    ' De Ids en procesos paralelos
    IdMovDet = F.NuloNumeric(FData.GetCodeTable("alm_ingresodet", "idmovdet", mConexion))
    '************************

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    xId = Me.IdMovimiento
    
    dataBase.CommandText = "UPDATE alm_ingreso SET tipdoc = ?, fching = ?, fchdoc = ?, " _
                & "numser = ?, numdoc = ?, idres = ?, idpro = ?, nombre = ?, " _
                & "estado = ?, tipmov = ?, idtipdocref = ?, " _
                & "iddocref = ?, idtipdocref2 = ?, iddocref2 = ?, " _
                & "idalm = ?, glosa = ?, desdocref = ?, ano = ?, idmes = ? " _
            + vbCr + "WHERE (alm_ingreso.id = ?)"
    
    dataBase.AddParameter "@tipdoc", adInteger, adParamInput, F.NuloNumeric(Me.IdTipoDocumento)
    dataBase.AddParameter "@fching", adDate, adParamInput, Me.FechaMovimiento
    dataBase.AddParameter "@fchdoc", adDate, adParamInput, Me.FechaMovimiento
    dataBase.AddParameter "@numser", adVarChar, adParamInput, F.NuloString(Me.NumeroSerie)
    dataBase.AddParameter "@numdoc", adVarChar, adParamInput, F.NuloString(Me.NumeroDocumento)
    dataBase.AddParameter "@idres", adInteger, adParamInput, F.NuloNumeric(Me.IdResponsable)
    dataBase.AddParameter "@idpro", adInteger, adParamInput, F.NuloNumeric(Me.IdProveedor)
    dataBase.AddParameter "@nombre", adVarChar, adParamInput, F.NuloString(Me.Proveedor)
    dataBase.AddParameter "@estado", adInteger, adParamInput, F.NuloNumeric(Me.IdEstado)
    dataBase.AddParameter "@tipmov", adInteger, adParamInput, F.NuloNumeric(Me.IdTipoMovimiento)
    dataBase.AddParameter "@idtipdocref", adInteger, adParamInput, F.NuloNumeric(Me.IdTipoDocumentoReferencia)
    dataBase.AddParameter "@iddocref", adInteger, adParamInput, F.NuloNumeric(Me.IdDocumentoReferencia)
    dataBase.AddParameter "@idtipdocref2", adInteger, adParamInput, F.NuloNumeric(Me.IdTipoDocumentoReferencia2)
    dataBase.AddParameter "@iddocref2", adInteger, adParamInput, F.NuloNumeric(Me.IdDocumentoReferencia2)
    dataBase.AddParameter "@idalm", adInteger, adParamInput, F.NuloNumeric(Me.IdAlmacen)
    dataBase.AddParameter "@glosa", adVarChar, adParamInput, F.NuloString(Me.Glosa)
    dataBase.AddParameter "@desdocref", adVarChar, adParamInput, F.NuloNumeric(Me.DocumentoReferencia)
    dataBase.AddParameter "@ano", adInteger, adParamInput, F.NuloNumeric(Me.AnhoTrabajo)
    dataBase.AddParameter "@idmes", adInteger, adParamInput, F.NuloNumeric(Me.MesTrabajo)
    dataBase.AddParameter "@id", adInteger, adParamInput, xId
    dataBase.Execute
    
    ' Borramos Hijos
    DeleteChild IdUsuario, MachineName
    
    ' Genera Almacenaje Automatico
    If F.NuloNumeric(F.KeyValue("GeneraAlmacenajeAutomatico", mConexion)) = -1 Then
        If Not IsRecursive Then GeneraAlmacenajeAutomatico IdUsuario, MachineName
    End If
    
    ' Grabamos Hijos
    Dim MovDet As EMovimientoDet
    For Each MovDet In Me.LMovimientoDet
        If Not mConexion Is Nothing Then
            Set MovDet.Conexion = mConexion
        End If
        MovDet.Called = True
        MovDet.MarkNew = True
        MovDet.IdMovimientoDet = IdMovDet
        MovDet.IdMovimiento = Me.IdMovimiento
        If Not MovDet.Save(IdUsuario, MachineName) Then Err.Raise &HFFFFFF01, , "Error al intentar grabar el movimiento"
        IdMovDet = IdMovDet + 1
    Next
    
'    ' Grabamos Hijos
'    Dim MovDet As EMovimientoDet
'    For Each MovDet In Me.LMovimientoDet
'        If Not mConexion Is Nothing Then
'            Set MovDet.Conexion = mConexion
'        End If
'        MovDet.Called = True
'        MovDet.IdMovimiento = Me.IdMovimiento
'        MovDet.Save IdUsuario, MachineName
'    Next
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el movimiento por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub

Private Sub GeneraAlmacenajeAutomatico(IdUsuario As Long, MachineName As String)
    Dim IdAlmacenItem As Long
    Dim MovDet As EMovimientoDet
    Dim ListaMov As New Collection
    Dim mIndice As Long
            
    For Each MovDet In Me.LMovimientoDet
        mIndice = MovDet.Index
        ' Se verifica si el item tiene almacenaje automatico
        IdAlmacenItem = F.AlmacenaEn(MovDet.IdItem, mConexion)
        If IdAlmacenItem <> 0 And IdAlmacenItem <> Me.IdAlmacen Then ' Si no coincide con el almacen actual y esta direccionado
            
            ' Se verifica si en la coleccion ya existe un movimiento para ese almacen
            If F.GetItemCollection(ListaMov, CStr(IdAlmacenItem)) Is Nothing Then
                ' Se crea un movimiento adicional para el item
                Dim MovAdicional As New AlmacenEntidad.EMovimiento
                MovAdicional.IdMovimiento = 0
                MovAdicional.FechaMovimiento = Me.FechaMovimiento
                MovAdicional.NumeroSerie = Me.NumeroSerie
                MovAdicional.NumeroDocumento = F.HallaNumeroDocumento("alm_ingreso", "'" & Me.NumeroSerie & "'", "numser", mConexion)
                MovAdicional.IdEstado = Me.IdEstado
                MovAdicional.IdAlmacen = IdAlmacenItem
                MovAdicional.IdProveedor = Me.IdProveedor
                MovAdicional.Proveedor = Me.Proveedor
                MovAdicional.Glosa = Me.Glosa
                MovAdicional.IdTipoMovimiento = Me.IdTipoMovimiento
                MovAdicional.IdTipoDocumentoReferencia = Me.IdTipoDocumentoReferencia
                MovAdicional.IdDocumentoReferencia = Me.IdDocumentoReferencia
                MovAdicional.DocumentoReferencia = Me.DocumentoReferencia
                ' Se agrega como segundo doc ref a la SM que lo crea
                MovAdicional.IdTipoDocumentoReferencia2 = F.NuloNumeric(F.KeyValue("IdDocumentoSolictudMateriales", mConexion))
                MovAdicional.IdDocumentoReferencia2 = Me.IdMovimiento
                '
                MovAdicional.MesTrabajo = Me.MesTrabajo
                MovAdicional.AnhoTrabajo = Me.AnhoTrabajo
                MovAdicional.IsRecursive = True
                ' Se agrega el movimiento detalle
                MovAdicional.LMovimientoDet.Add MovDet
                ' Se agrega a la coleccion global
                ListaMov.Add MovAdicional, CStr(IdAlmacenItem)
                Set MovAdicional = Nothing
            Else
                ' Se agrega el detalle al movimiento existente
                ListaMov(CStr(IdAlmacenItem)).LMovimientoDet.Add MovDet
            End If
            ' Se quita del objeto el detalle
            Me.LMovimientoDet.Delete mIndice
        End If
    Next
    ' Se graban los movimientos generados
    Dim Mov As EMovimiento
    For Each Mov In ListaMov
        If Not mConexion Is Nothing Then
            Set Mov.Conexion = mConexion
            Mov.Called = True
        End If
        Mov.Save IdUsuario, MachineName
    Next
    Set Mov = Nothing
End Sub

Public Sub Delete(IdUsuario As Long, MachineName As String)
    Dim xId As Double
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    ' Eliminamos a los Hijos
    DeleteChild IdUsuario, MachineName
    
    dataBase.ClearParameter
    dataBase.CommandText = "DELETE FROM alm_ingreso " _
            + vbCr + "WHERE alm_ingreso.id = ?"
    
    dataBase.AddParameter "@id", adInteger, adParamInput, Me.IdMovimiento
    dataBase.Execute
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "Delete, IdMovimiento: " + Me.IdMovimiento + ", Motivo :" + Trim(Err.Description)
    Exit Sub
End Sub

Private Sub DeleteChild(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    
    dataBase.ClearParameter
    dataBase.CommandText = "DELETE FROM alm_ingresodet " _
            + vbCr + "WHERE alm_ingresodet.id = ?"

    dataBase.AddParameter "@id", adInteger, adParamInput, Me.IdMovimiento
    dataBase.Execute
    
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "DeleteChild, IdMovimiento: " + Me.IdMovimiento + ", Motivo :" + Trim(Err.Description)
    Exit Sub
End Sub


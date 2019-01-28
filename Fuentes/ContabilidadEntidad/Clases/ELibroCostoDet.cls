VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ELibroCostoDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdLibroCostoDet As Long
Private mIdLibroCosto As Long
Private mIdParteDetalle As Long
Private mIdMovimientoDetalle As Long
Private mIdItem As Long
Private mItem As String
Private mFecha As Date
Private mIdParteProd As Long
Private mParteProd As String
Private mIdReceta As Long
Private mReceta As String
Private mIdUniMed As Long
Private mUniMed As String
Private mCantidad As Double
Private mHoraInicio As Date
Private mHoraFin As Date
Private mTotalHoras As Date
Private mFactorDistribucion As Double
Private mImporteMP As Double
Private mImporteMOD As Double
Private mImporteCIF As Double
Private mTipo As String
Private F As New SistemaLogica.Funciones
Private mConexion As New ADODB.Connection
Private mIndex As Long
Private mCalled As Boolean
Private mMarkOld As Boolean
Private mMarkNew As Boolean

Private Sub Class_Initialize()
    mCalled = False
    mMarkOld = False
    mMarkNew = False
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

Public Property Get Index() As Long
    Index = mIndex
End Property
Public Property Let Index(ByVal NewValue As Long)
    mIndex = NewValue
End Property

Public Property Get Called() As Boolean
    Called = mCalled
End Property
Public Property Let Called(ByVal NewValue As Boolean)
    mCalled = NewValue
End Property

Public Property Get MarkOld() As Boolean
    MarkOld = mMarkOld
End Property
Public Property Let MarkOld(ByVal NewValue As Boolean)
    mMarkOld = NewValue
End Property

Public Property Get MarkNew() As Boolean
    MarkNew = mMarkNew
End Property
Public Property Let MarkNew(ByVal NewValue As Boolean)
    mMarkNew = NewValue
End Property

Public Property Get IdLibroCostoDet() As Long
    IdLibroCostoDet = mIdLibroCostoDet
End Property
Public Property Let IdLibroCostoDet(ByVal NewValue As Long)
    mIdLibroCostoDet = NewValue
End Property

Public Property Get IdLibroCosto() As Long
    IdLibroCosto = mIdLibroCosto
End Property
Public Property Let IdLibroCosto(ByVal NewValue As Long)
    mIdLibroCosto = NewValue
End Property

Public Property Get IdParteDetalle() As Long
    IdParteDetalle = mIdParteDetalle
End Property
Public Property Let IdParteDetalle(ByVal NewValue As Long)
    mIdParteDetalle = NewValue
End Property

Public Property Get IdMovimientoDetalle() As Long
    IdMovimientoDetalle = mIdMovimientoDetalle
End Property
Public Property Let IdMovimientoDetalle(ByVal NewValue As Long)
    mIdMovimientoDetalle = NewValue
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

Public Property Get Fecha() As Date
    Fecha = mFecha
End Property
Public Property Let Fecha(ByVal NewValue As Date)
    mFecha = NewValue
End Property

Public Property Get IdParteProd() As Long
    IdParteProd = mIdParteProd
End Property
Public Property Let IdParteProd(ByVal NewValue As Long)
    mIdParteProd = NewValue
End Property

Public Property Get ParteProd() As String
    ParteProd = mParteProd
End Property
Public Property Let ParteProd(ByVal NewValue As String)
    mParteProd = NewValue
End Property

Public Property Get IdReceta() As Long
    IdReceta = mIdReceta
End Property
Public Property Let IdReceta(ByVal NewValue As Long)
    mIdReceta = NewValue
End Property

Public Property Get Receta() As String
    Receta = mReceta
End Property
Public Property Let Receta(ByVal NewValue As String)
    mReceta = NewValue
End Property

Public Property Get IdUniMed() As Long
    IdUniMed = mIdUniMed
End Property
Public Property Let IdUniMed(ByVal NewValue As Long)
    mIdUniMed = NewValue
End Property

Public Property Get UniMed() As String
    UniMed = mUniMed
End Property
Public Property Let UniMed(ByVal NewValue As String)
    mUniMed = NewValue
End Property

Public Property Get Cantidad() As Double
    Cantidad = mCantidad
End Property
Public Property Let Cantidad(ByVal NewValue As Double)
    mCantidad = NewValue
End Property

Public Property Get HoraInicio() As Date
    HoraInicio = mHoraInicio
End Property
Public Property Let HoraInicio(ByVal NewValue As Date)
    mHoraInicio = NewValue
End Property

Public Property Get HoraFin() As Date
    HoraFin = mHoraFin
End Property
Public Property Let HoraFin(ByVal NewValue As Date)
    mHoraFin = NewValue
End Property

Public Property Get TotalHoras() As Date
    TotalHoras = mTotalHoras
End Property
Public Property Let TotalHoras(ByVal NewValue As Date)
    mTotalHoras = NewValue
End Property

Public Property Get FactorDistribucion() As Double
    FactorDistribucion = mFactorDistribucion
End Property
Public Property Let FactorDistribucion(ByVal NewValue As Double)
    mFactorDistribucion = NewValue
End Property

Public Property Get ImporteMP() As Double
    ImporteMP = mImporteMP
End Property
Public Property Let ImporteMP(ByVal NewValue As Double)
    mImporteMP = NewValue
End Property

Public Property Get ImporteMOD() As Double
    ImporteMOD = mImporteMOD
End Property
Public Property Let ImporteMOD(ByVal NewValue As Double)
    mImporteMOD = NewValue
End Property

Public Property Get ImporteCIF() As Double
    ImporteCIF = mImporteCIF
End Property
Public Property Let ImporteCIF(ByVal NewValue As Double)
    mImporteCIF = NewValue
End Property

Public Property Get Tipo() As String
    Tipo = mTipo
End Property
Public Property Let Tipo(ByVal NewValue As String)
    mTipo = NewValue
End Property

Public Sub Fetch(IdLibroCostoDetalle As Long)
    Dim dataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BLoqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
            
    dataBase.CommandText = "SELECT con_librocostodet.idmovdet, con_librocostodet.id AS idmov, con_librocosto.fchdoc AS fchmov, alm_inventario.idunimed, mae_unidades.abrev AS unimed, con_librocostodet.iditem, alm_inventario.descripcion AS item, con_librocostodet.cantidad AS canreal, con_librocostodet.cantteo AS canteo, con_librocostodet.iddocref, con_librocostodet.facdist " _
        + vbCr + "FROM ((con_librocostodet INNER JOIN alm_inventario ON con_librocostodet.iditem = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) INNER JOIN con_librocosto ON con_librocostodet.id = con_librocosto.id " _
        + vbCr + "WHERE (((con_librocostodet.idmovdet)=" & IdLibroCostoDetalle & "))"
    
    Set mRecord = dataBase.GetRecordset
    If mRecord.RecordCount > 0 Then Me.SetEntity mRecord
    Exit Sub
    
BLoqueError:
    Err.Raise Err.Number, "[Fetch] " & Err.Source, Err.Description
End Sub

Public Sub SetEntity(Record As ADODB.Recordset)
    Me.IdLibroCostoDet = F.NuloNumeric(Record("id"))
    Me.IdLibroCosto = F.NuloNumeric(Record("idlibro"))
    Me.IdParteProd = F.NuloNumeric(Record("idpartedet"))
    Me.ParteProd = F.NuloString(Record("parteprod"))
    Me.IdParteDetalle = F.NuloNumeric(Record("idpartedet"))
    Me.IdMovimientoDetalle = F.NuloNumeric(Record("idmovdet"))
    Me.IdItem = F.NuloNumeric(Record("iditem"))
    Me.Item = F.NuloString(Record("item"))
    Me.Fecha = Record("fecha")
    Me.Tipo = F.NuloString(Record("tipo"))
    Me.Receta = F.NuloString(Record("receta"))
    Me.IdUniMed = F.NuloNumeric(Record("idunimed"))
    Me.UniMed = F.NuloString(Record("unimed"))
    Me.Cantidad = F.NuloNumeric(Record("cantidad"))
    Me.HoraInicio = Record("horini")
    Me.HoraFin = Record("horfin")
    Me.TotalHoras = Me.HoraFin - Me.HoraInicio
    Me.ImporteMP = F.NuloNumeric(Record("impmprima"))
    Me.ImporteMOD = F.NuloNumeric(Record("impmanobr"))
    Me.ImporteCIF = F.NuloNumeric(Record("impgasfab"))
    Me.FactorDistribucion = F.NuloNumeric(Record("facdist"))
End Sub

Public Function Save(IdUsuario As Long, MachineName As String) As Boolean
On Error GoTo BLoqueError
    If (F.NuloNumeric(Me.IdLibroCostoDet) > 0) And (Not Me.MarkNew) Then
        Update IdUsuario, MachineName
    Else
        Insert IdUsuario, MachineName
    End If
    Save = True
    Exit Function
    
BLoqueError:
    Save = False
    F.MostrarMensajeError Err.Description, "Grabar", "[Save] " & Err.Source
End Function

Private Sub Insert(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    Dim FData As New SistemaData.Funciones
    
On Error GoTo BLoqueError
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    If Me.IdLibroCostoDet = 0 Then
        Me.IdLibroCostoDet = FData.GetCodeTable("con_librocostodet", "id", mConexion)
    End If
    dataBase.ClearParameter

    dataBase.CommandText = "INSERT INTO con_librocostodet(id, idlibro, iditem, idproddet, idmovdet, cantidad, impmprima, impmanobr, impgasfab, tipo, facdist) " _
            + vbCr + "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

    dataBase.AddParameter "@id", adInteger, adParamInput, Me.IdLibroCostoDet
    dataBase.AddParameter "@idlibro", adInteger, adParamInput, F.NuloNumeric(Me.IdLibroCosto)
    dataBase.AddParameter "@iditem", adInteger, adParamInput, F.NuloNumeric(Me.IdItem)
    dataBase.AddParameter "@idproddet", adInteger, adParamInput, F.NuloNumeric(Me.IdParteDetalle)
    dataBase.AddParameter "@idmovdet", adInteger, adParamInput, F.NuloNumeric(Me.IdMovimientoDetalle)
    dataBase.AddParameter "@cantidad", adDouble, adParamInput, F.NuloNumeric(Me.Cantidad)
    dataBase.AddParameter "@impmprima", adDouble, adParamInput, F.NuloNumeric(Me.ImporteMP)
    dataBase.AddParameter "@impmanobr", adDouble, adParamInput, F.NuloNumeric(Me.ImporteMOD)
    dataBase.AddParameter "@impgasfab", adDouble, adParamInput, F.NuloNumeric(Me.ImporteCIF)
    dataBase.AddParameter "@tipo", adVarChar, adParamInput, F.NuloString(Me.Tipo)
    dataBase.AddParameter "@facdist", adDouble, adParamInput, F.NuloNumeric(Me.FactorDistribucion)
    dataBase.Execute
    
    ' Se graba el gasto
    GrabaGasto IdUsuario, MachineName
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el detalle del LibroCosto por el siguiente motivo :" + Trim(Err.Description)
End Sub

Private Sub Update(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
    
    dataBase.ClearParameter
        
    dataBase.CommandText = "UPDATE con_librocostodet SET idlibro = ?, iditem = ?, idproddet = ?, idmovdet = ?, cantidad = ?, impmprima = ?, impmanobr = ?, impgasfab = ?, tipo = ?, facdist = ? " _
            + vbCr + "WHERE(con_librocostodet.id = ?)"

    dataBase.AddParameter "@idlibro", adInteger, adParamInput, F.NuloNumeric(Me.IdLibroCosto)
    dataBase.AddParameter "@iditem", adInteger, adParamInput, F.NuloNumeric(Me.IdItem)
    dataBase.AddParameter "@idproddet", adInteger, adParamInput, F.NuloNumeric(Me.IdParteDetalle)
    dataBase.AddParameter "@idmovdet", adInteger, adParamInput, F.NuloNumeric(Me.IdMovimientoDetalle)
    dataBase.AddParameter "@cantidad", adDouble, adParamInput, F.NuloNumeric(Me.Cantidad)
    dataBase.AddParameter "@impmprima", adDouble, adParamInput, F.NuloNumeric(Me.ImporteMP)
    dataBase.AddParameter "@impmanobr", adDouble, adParamInput, F.NuloNumeric(Me.ImporteMOD)
    dataBase.AddParameter "@impgasfab", adDouble, adParamInput, F.NuloNumeric(Me.ImporteCIF)
    dataBase.AddParameter "@tipo", adVarChar, adParamInput, F.NuloString(Me.Tipo)
    dataBase.AddParameter "@facdist", adDouble, adParamInput, F.NuloNumeric(Me.FactorDistribucion)
    dataBase.AddParameter "@id", adInteger, adParamInput, Me.IdLibroCostoDet
    dataBase.Execute
    
    ' Se graba el gasto
    GrabaGasto IdUsuario, MachineName
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "No se pudo registrar el Transferencia por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub

Private Sub GrabaGasto(IdUsuario As Long, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
    
On Error GoTo BLoqueError

    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    dataBase.ClearParameter

    dataBase.CommandText = "UPDATE alm_kardexdetmov SET importemod = ?, importecif = ? " _
            + vbCr + "WHERE alm_kardexdetmov.idmovdet = ?"
    
    dataBase.AddParameter "@importemod", adDouble, adParamInput, Me.ImporteMOD
    dataBase.AddParameter "@importecif", adDouble, adParamInput, Me.ImporteCIF
    dataBase.AddParameter "@idmovdet", adInteger, adParamInput, Me.IdMovimientoDetalle
    dataBase.Execute
    
    Set dataBase = Nothing
    Exit Sub

BLoqueError:
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

    dataBase.CommandText = "DELETE FROM con_librocostodet " _
            + vbCr + "WHERE con_librocostodet.id = ?"
    
    dataBase.AddParameter "@id", adInteger, adParamInput, Me.IdLibroCostoDet
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



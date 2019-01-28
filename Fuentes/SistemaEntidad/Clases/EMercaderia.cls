VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EMercaderia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mIdMercaderia As Integer
Private mCodigo As String
Private mDescripcion As String
Private mDescripcionTecnica As String
Private mStockInicial As Double
Private mPrecioUnitarioInicial As Double
Private mStockActual As Double
Private mPrecioUnitarioActual As Double
Private mIdMoneda As Integer
Private mMoneda As String
Private mIdUnidadMedida As Integer
Private mUnidadMedida As String
Private mIdTipoProducto As Integer
Private mTipoProducto As String
Private mIdFamilia As Integer
Private mFamilia As String
Private mIdClase As Integer
Private mClase As String
Private mIdSubClase As Integer
Private mSubClase As String
Private F As New SistemaLogica.Funciones

Public Property Get IdMercaderia() As Integer
    IdMercaderia = mIdMercaderia
End Property
Public Property Let IdMercaderia(ByVal NewValue As Integer)
    mIdMercaderia = NewValue
End Property

Public Property Get Codigo() As String
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal NewValue As String)
    mCodigo = NewValue
End Property

Public Property Get Descripcion() As String
    Descripcion = mDescripcion
End Property
Public Property Let Descripcion(ByVal NewValue As String)
    mDescripcion = NewValue
End Property

Public Property Get DescripcionTecnica() As String
    DescripcionTecnica = mDescripcionTecnica
End Property
Public Property Let DescripcionTecnica(ByVal NewValue As String)
    mDescripcionTecnica = NewValue
End Property

Public Property Get StockInicial() As Double
    StockInicial = mStockInicial
End Property
Public Property Let StockInicial(ByVal NewValue As Double)
    mStockInicial = NewValue
End Property

Public Property Get PrecioUnitarioInicial() As Double
    PrecioUnitarioInicial = mPrecioUnitarioInicial
End Property
Public Property Let PrecioUnitarioInicial(ByVal NewValue As Double)
    mPrecioUnitarioInicial = NewValue
End Property

Public Property Get StockActual() As Double
    StockActual = mStockActual
End Property
Public Property Let StockActual(ByVal NewValue As Double)
    mStockActual = NewValue
End Property

Public Property Get PrecioUnitarioActual() As Double
    PrecioUnitarioActual = mPrecioUnitarioActual
End Property
Public Property Let PrecioUnitarioActual(ByVal NewValue As Double)
    mPrecioUnitarioActual = NewValue
End Property

Public Property Get IdMoneda() As Integer
    IdMoneda = mIdMoneda
End Property
Public Property Let IdMoneda(ByVal NewValue As Integer)
    mIdMoneda = NewValue
End Property

Public Property Get Moneda() As String
    Moneda = mMoneda
End Property
Public Property Let Moneda(ByVal NewValue As String)
    mMoneda = NewValue
End Property

Public Property Get IdUnidadMedida() As Integer
    IdUnidadMedida = mIdUnidadMedida
End Property
Public Property Let IdUnidadMedida(ByVal NewValue As Integer)
    mIdUnidadMedida = NewValue
End Property

Public Property Get UnidadMedida() As String
    UnidadMedida = mUnidadMedida
End Property
Public Property Let UnidadMedida(ByVal NewValue As String)
    mUnidadMedida = NewValue
End Property

Public Property Get IdTipoProducto() As Integer
    IdTipoProducto = mIdTipoProducto
End Property
Public Property Let IdTipoProducto(ByVal NewValue As Integer)
    mIdTipoProducto = NewValue
End Property

Public Property Get TipoProducto() As String
    TipoProducto = mTipoProducto
End Property
Public Property Let TipoProducto(ByVal NewValue As String)
    mTipoProducto = NewValue
End Property

Public Property Get IdFamilia() As Integer
    IdFamilia = mIdFamilia
End Property
Public Property Let IdFamilia(ByVal NewValue As Integer)
    mIdFamilia = NewValue
End Property

Public Property Get Familia() As String
    Familia = mFamilia
End Property
Public Property Let Familia(ByVal NewValue As String)
    mFamilia = NewValue
End Property

Public Property Get IdClase() As Integer
    IdClase = mIdClase
End Property
Public Property Let IdClase(ByVal NewValue As Integer)
    mIdClase = NewValue
End Property

Public Property Get Clase() As String
    Clase = mClase
End Property
Public Property Let Clase(ByVal NewValue As String)
    mClase = NewValue
End Property

Public Property Get IdSubClase() As Integer
    IdSubClase = mIdSubClase
End Property
Public Property Let IdSubClase(ByVal NewValue As Integer)
    mIdSubClase = NewValue
End Property

Public Property Get SubClase() As String
    SubClase = mSubClase
End Property
Public Property Let SubClase(ByVal NewValue As String)
    mSubClase = NewValue
End Property



Public Function Fetch(IdMovimiento As String) As EMercaderia

End Function

Public Sub SetEntity(Record As ADODB.Recordset)
    Record.MoveFirst
    Me.IdMercaderia = F.NuloNumeric(Record("Id"))
    Set Record = Nothing
End Sub

Public Sub Save(IdUsuario As Integer, MachineName As String)
    If (F.NuloNumeric(Me.IdMercaderia) = 0) Then
        Insert IdUsuario, MachineName
    Else
        Update IdUsuario, MachineName
    End If
End Sub

Private Sub Insert(IdUsuario As Integer, MachineName As String)
    
End Sub

Private Sub Update(IdUsuario As Integer, MachineName As String)

End Sub

Public Sub Delete(IdUsuario As Integer, MachineName As String)

End Sub

Public Sub SaveInventarioInicial(IdUsuario As Integer, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
        
On Error GoTo BloqueError

    dataBase.BeginTrans
                
    dataBase.CommandText = "PARAMETERS pid Integer, ppreini Double, ppreuni Double, pstckini Double, pstckact Double;" _
            + vbCr + "UPDATE alm_inventario " _
            + vbCr + "SET preini= ppreini, preuni=ppreuni, stckini=pstckini, stckact = pstckact " _
            + vbCr + "WHERE id = pid"
    
    dataBase.AddParameter "pid", adInteger, adParamInput, F.NuloNumeric(Me.IdMercaderia)
    dataBase.AddParameter "ppreini", adDouble, adParamInput, F.NuloNumeric(Me.PrecioUnitarioInicial)
    dataBase.AddParameter "ppreuni", adDouble, adParamInput, F.NuloNumeric(Me.PrecioUnitarioActual)
    dataBase.AddParameter "pstckini", adDouble, adParamInput, F.NuloNumeric(Me.StockInicial)
    dataBase.AddParameter "pstckact", adDouble, adParamInput, F.NuloNumeric(Me.StockActual)
    dataBase.Execute
    dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub
    
BloqueError:
    dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "Ocurrio un error al intentar generar el inventario inicial, " + Trim(Err.Description)
End Sub


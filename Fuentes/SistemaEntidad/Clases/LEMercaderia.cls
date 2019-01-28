VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LEMercaderia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mMercaderias As New Collection
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

Public Sub Add(mMercaderia As EMercaderia)
    mMercaderias.Add mMercaderia, CStr(mMercaderia.IdMercaderia)
End Sub

Public Function Count() As Long
    Count = mMercaderias.Count
End Function

Public Sub Delete(ByVal Index As Variant)
    mMercaderias.Remove Index
    
End Sub

Public Function Item(ByVal Index As Variant) As EMercaderia
Attribute Item.VB_UserMemId = 0
    Set Item = mMercaderias.Item(Index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mMercaderias.[_NewEnum]
End Function

Public Function Fetch(Opcion As MercaderiaCollectionEnum, _
                        Optional IdInventario As String, _
                        Optional Codigo As String, _
                        Optional Descripcion As String, _
                        Optional FechaCreacionRegistroInicio As Date, _
                        Optional FechaCreacionRegistroFin As Date, _
                        Optional FechaUltimaActualizacionInicio As Date, _
                        Optional FechaUltimaActualizacionFin As Date) As LEMercaderia
    
End Function

Public Sub SaveInventarioInicial(IdUsuario As Integer, MachineName As String)
    Dim dataBase As New SistemaData.EDataBase
        
On Error GoTo BloqueError
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    If Not Called Then dataBase.BeginTrans
                
    Dim merc As EMercaderia
    For Each merc In mMercaderias
        dataBase.ClearParameter
        dataBase.CommandText = "UPDATE alm_inventario " _
                + vbCr + "SET preini= ?, stckini=? " _
                + vbCr + "WHERE id = ?"
        
        dataBase.AddParameter "@preini", adDouble, adParamInput, F.NuloNumeric(merc.PrecioUnitarioActual)
        dataBase.AddParameter "@stckini", adDouble, adParamInput, F.NuloNumeric(merc.StockActual)
        dataBase.AddParameter "@id", adInteger, adParamInput, F.NuloNumeric(merc.IdMercaderia)
        dataBase.Execute
    Next
    
    If Not Called Then dataBase.CommitTrans
    Set dataBase = Nothing
    Exit Sub
    
BloqueError:
    If Not Called Then dataBase.RollbackTrans
    Set dataBase = Nothing
    Err.Raise &HFFFFFF01, , "Ocurrio un error al intentar generar el inventario inicial, " + Trim(Err.Description)
End Sub



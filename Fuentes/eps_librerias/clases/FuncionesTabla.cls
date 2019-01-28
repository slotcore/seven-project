VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FuncionesTabla"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Explicit
'
'Enum TipoDocumento
'    Factura = 1
'    Boleta = 2
'    Recibo = 3
'    Otros = 4
'End Enum
'
'Function HallaNumeroRecibo(Tabla As String, Campo As String, NumSerie As String, Tipo As TipoDocumento, Coneccion As ADODB.Connection) As String
'    Dim xRst As New ADODB.Recordset
'
'    RST_Busq xRst, "SELECT * FROM " & Tabla & " WHERE numserie = '" & NumSerie & "' AND tipodoc = " & Tipo & " ORDER BY " & Campo & " ", Coneccion
'
'    If xRst.RecordCount = 0 Then
'        HallaNumeroRecibo = "0000001"
'    Else
'        xRst.MoveLast
'        HallaNumeroRecibo = Format(Val(xRst(Campo)) + 1, "0000000")
'    End If
'    Set xRst = Nothing
'End Function

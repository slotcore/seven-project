VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Imprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Cabecera1 As String
Public Cabecera2 As String
Public Fecha As String
Public Titulo1 As String
Public Titulo2 As String
Public TamañoFuente As Integer
Public TamañoCabecera As Integer
Public FuenteCabecera As String
Public Tamaño_Hoja As TamañoHoja
Public Posicion_Hoja As Posicion
Public TextoConsiderar As String
Public TextoConsiderarAncho As Integer

Public Enum TamañoHoja
    A_4 = 1
    Carta = 2
    Oficio = 3
End Enum

Public Enum Posicion
    Vertical = 1
    Horizontal = 2
End Enum

Sub ImprimirArray(xArray() As String, RstCabecera As ADODB.Recordset)
    ArrayPrin = xArray
    
    Prin_Cabecera1 = Cabecera1
    Prin_Cabecera2 = Cabecera2
    Prin_Fecha = Fecha
    Prin_Titulo1 = Titulo1
    Prin_Titulo2 = Titulo2
    xTamañoFuente = TamañoFuente
    
    Prin_TamañoHoja = Tamaño_Hoja
    Prin_OrientacionHoja = Posicion_Hoja
    Prin_TamañoCabecera = TamañoCabecera
    Prin_FuenteCabecera = FuenteCabecera
    
    Prin_TextoConsiderar = TextoConsiderar
    Prin_TextoConsiderarAncho = TextoConsiderarAncho
    
    Set RstPrin = RstCabecera
    FrmPrinter.Show vbModal
End Sub

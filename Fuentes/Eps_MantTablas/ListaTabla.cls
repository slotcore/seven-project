VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListaTabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Titulo As String
Public SQLCAD As String
Public RstCampos As ADODB.Recordset

Sub CargaLista(Con As ADODB.Connection)
    Set xCon = Con
    FrmListaTabla.LblTitulo.Caption = Titulo
    FrmListaTabla.Caption = Titulo
    FrmListaTabla.xSqlCad = SQLCAD
    Set FrmListaTabla.xRstCampos = RstCampos
    FrmListaTabla.Show
End Sub


Public Property Get xhWnd() As Long
   xhWnd = FrmListaTabla.hWnd
End Property

'Public Property Let Ancho(ByVal xAncho As Integer)
'   'FrmListaTabla.Left = 0
'   'FrmListaTabla.Top = 0
'   FrmListaTabla.Width = xAncho
'   FrmListaTabla.Refresh
'End Property
''
'Public Property Let Alto(ByVal xAlto As Integer)
'   'FrmListaTabla.Left = 0
'   'FrmListaTabla.Top = 0
'   FrmListaTabla.Height = xAlto
'   FrmListaTabla.Refresh
'End Property


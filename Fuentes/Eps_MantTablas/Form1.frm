VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#12.0#0"; "Codejock.CommandBars.v12.0.0.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Begin SizerOneLibCtl.ElasticOne EO 
      Height          =   8025
      Left            =   15
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   645
      Width           =   10320
      _cx             =   18203
      _cy             =   14155
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      AutoSizeChildren=   8
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   3
      GridCols        =   2
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"Form1.frx":0000
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   375
         Left            =   2565
         TabIndex        =   2
         Top             =   90
         Width           =   7665
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nombre del Formulario"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   90
            TabIndex        =   3
            Top             =   75
            Width           =   2835
         End
      End
      Begin VB.PictureBox Pic 
         BackColor       =   &H00C0C0FF&
         Height          =   7005
         Left            =   2565
         ScaleHeight     =   347.25
         ScaleMode       =   2  'Point
         ScaleWidth      =   380.25
         TabIndex        =   1
         Top             =   525
         Width           =   7665
      End
   End
   Begin XtremeCommandBars.CommandBars CommandBars1 
      Left            =   0
      Top             =   0
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Left            =   1275
      Top             =   120
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "Form1.frx":005B
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TopEO As Integer
Dim TAMAÑO_TOOL As TOOL_TAMAÑO_ICO
Dim xF As New Eps_MantTablas.ListaTabla
Dim res As Long

Const BTN_NEW = 1  ' NUEVO
Const BTN_MOD = 2  ' MOFICAR
Const BTN_BUS = 3  ' BUSCAR
Const BTN_EXC = 4  ' EXPORTAR EXCEL
Const BTN_IMP = 5  ' IMPRIMIR
Const BTN_CAL = 6  ' CALENDARIO
Const BTN_CON = 7  ' CONFIGURAR
Const BTN_SAL = 8  ' SALIR

Function CargarTabla(IdMantenimiento As Integer) As ADODB.Recordset
    Dim RstCab As New ADODB.Recordset
    
    RST_Busq RstCab, "SELECT * FROM mae_manformularios WHERE id = " & IdMantenimiento & "", xCon
    If RstCab.RecordCount = 0 Then
        Set CargarTabla = Nothing
    Else
        Set CargarTabla = RstCab
    End If
End Function

Function CargarTablaCampos(IdMantenimiento As Integer) As ADODB.Recordset
    Dim RstCab As New ADODB.Recordset
    
    RST_Busq RstCab, "SELECT * FROM mae_manformulariosdet WHERE id = " & IdMantenimiento & " ORDER BY corr", xCon
    If RstCab.RecordCount = 0 Then
        Set CargarTablaCampos = Nothing
    Else
        Set CargarTablaCampos = RstCab
    End If
End Function

Private Sub Form_Activate()
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    
    Set RstCab = CargarTabla(8)
    Set RstDet = CargarTablaCampos(8)
    
    xF.Titulo = RstCab("titulo")
    xF.SQLCad = RstCab("sqlcad")
    Set xF.RSTCAMPOS = RstDet
    'xF.Alto = Pic.Height
    'xF.Ancho = Pic.Width
    
    xF.CargaLista xCon
    res = SetParent(xF.xhWnd, Pic.hwnd)
    res = ShowWindow(xF.xhWnd, SW_HIDE)
    res = ShowWindow(xF.xhWnd, SW_RESTORE)
    
    'res = ShowWindow(xF.xhWnd, SHOWMAXIMIZED_eSW)
End Sub

Private Sub Form_Load()
    CrearTool
    
    If TAMAÑO_TOOL = I16x16 Then EO.Top = 400: TopEO = 400
    If TAMAÑO_TOOL = I24x24 Then EO.Top = 520: TopEO = 520
    If TAMAÑO_TOOL = I32x32 Then EO.Top = 640: TopEO = 640
    If TAMAÑO_TOOL = I48x48 Then EO.Top = 890: TopEO = 890
End Sub

Sub CrearTool()
    'CREAMOS EL TOOLBAR
    Dim Opciones(7, 3) As String
    
    Opciones(0, 0) = Str(BTN_NEW):    Opciones(0, 1) = "Nuevo Registro":              Opciones(0, 2) = "0":      Opciones(0, 3) = "Nuevo Registro"
    Opciones(1, 0) = Str(BTN_MOD):    Opciones(1, 1) = "Modificar Registro":          Opciones(1, 2) = "0":      Opciones(1, 3) = "Modificar Registro"
    Opciones(2, 0) = Str(BTN_BUS):    Opciones(2, 1) = "Buscar Registro":             Opciones(2, 2) = "1":      Opciones(2, 3) = "Buscar Registro"
    Opciones(3, 0) = Str(BTN_EXC):    Opciones(3, 1) = "Exportar Excel":              Opciones(3, 2) = "0":      Opciones(3, 3) = "Exportar Excel"
    Opciones(4, 0) = Str(BTN_IMP):    Opciones(4, 1) = "Imprimir":                    Opciones(4, 2) = "0":      Opciones(4, 3) = "Imprimir"
    Opciones(5, 0) = Str(BTN_CAL):    Opciones(5, 1) = "Calendario":                  Opciones(5, 2) = "0":      Opciones(5, 3) = "Calendario"
    Opciones(6, 0) = Str(BTN_CON):    Opciones(6, 1) = "Configurar Formulario":       Opciones(6, 2) = "1":      Opciones(6, 3) = "Configurar Formulario"
    Opciones(7, 0) = Str(BTN_SAL):    Opciones(7, 1) = "Salir":                       Opciones(7, 2) = "1":      Opciones(7, 3) = "Salir"
        
    Dim xFun As New eps_librerias.Codejock
    'PocisionarContenedor
    xFun.BORRARMENU = True
    TAMAÑO_TOOL = I24x24
    xFun.CrearToolBar Opciones, CommandBars1, ImageManager1, TAMAÑO_TOOL
    Set xFun = Nothing
End Sub

Private Sub Form_Resize()
    ' RECONFIGURAMOS EL TAMAÑOS DE LOS CONTROLES CUANDO SE MODIFIQUE EL TAMAÑO DEL FORMULARIO
    If Me.WindowState = 1 Then Exit Sub
    EO.Width = Me.Width - 130
    If Me.Height <= (TopEO + 2375) Then
        Me.Height = (TopEO + 2375)
    Else
        EO.Height = (Me.Height - (TopEO + 400))
    End If
    Me.Refresh
End Sub

Private Sub Pic_Resize()
    'SW_HIDE
    res = ShowWindow(xF.xhWnd, SW_HIDE)
    res = ShowWindow(xF.xhWnd, SW_RESTORE)
    'xF.Ancho = Pic.Width
    'xF.Alto = Pic.Height
    res = ShowWindow(xF.xhWnd, SHOWMAXIMIZED_eSW)
'    'xF.CargaLista xCon
    'res = SetParent(xF.xhWnd, Pic.hwnd)
End Sub

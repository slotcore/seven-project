VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#12.0#0"; "CODEJO~2.OCX"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#12.0#0"; "CODEJO~1.OCX"
Begin VB.Form FrmCalendario 
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin SizerOneLibCtl.ElasticOne EO1 
      Height          =   5265
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   645
      Width           =   9600
      _cx             =   16933
      _cy             =   9287
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
      BorderWidth     =   2
      ChildSpacing    =   2
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
      GridRows        =   2
      GridCols        =   2
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmCalendario.frx":0000
      Begin XtremeCalendarControl.CalendarControl CalendarControl 
         Height          =   4890
         Left            =   30
         TabIndex        =   1
         Top             =   345
         Width           =   6915
         _Version        =   786432
         _ExtentX        =   12197
         _ExtentY        =   8625
         _StockProps     =   64
      End
      Begin XtremeCalendarControl.DatePicker wndDatePicker 
         Height          =   4890
         Left            =   6975
         TabIndex        =   2
         Top             =   345
         Width           =   2595
         _Version        =   786432
         _ExtentX        =   4577
         _ExtentY        =   8625
         _StockProps     =   64
         Show3DBorder    =   0
         RowCount        =   2
      End
      Begin VB.Label lblCaption 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         Caption         =   " Calendario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   6915
      End
   End
   Begin XtremeCommandBars.ImageManager ImageManager 
      Left            =   2220
      Top             =   0
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "FrmCalendario.frx":004F
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   0
      Top             =   0
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Menu menu_01 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu_01_01 
         Caption         =   "Ver Entrega"
      End
   End
   Begin VB.Menu menu_02 
      Caption         =   "menu2"
      Visible         =   0   'False
      Begin VB.Menu menu_02_01 
         Caption         =   "Ver Entregas"
      End
   End
End
Attribute VB_Name = "FrmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const CALENDARIO_1 = 1
Const CALENDARIO_2 = 2
Const CALENDARIO_3 = 3
Const CALENDARIO_4 = 4
Const SALIR = 5
Const IMPRESORA = 6
Public TAMAÑO_TOOL As TOOL_TAMAÑO_ICO
Dim ContextEvent As CalendarEvent

Private Sub CalendarControl_DblClick()
    menu_02_01_Click
End Sub

Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error Resume Next
    
    Select Case Control.Id
        Case CALENDARIO_1:
            CalendarControl.ViewType = xtpCalendarDayView
        Case CALENDARIO_2:
            CalendarControl.ViewType = xtpCalendarWorkWeekView
        Case CALENDARIO_3:
            CalendarControl.ViewType = xtpCalendarWeekView
        Case CALENDARIO_4:
            CalendarControl.ViewType = xtpCalendarMonthView
        Case SALIR:
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    
    Me.Height = 8500
    Me.Width = 12000
    CentrarFormulario Me
    
    'CREAMOS EL TOOLBAR
    Dim Opciones(5, 3) As String
    
    Opciones(0, 0) = Str(CALENDARIO_1):    Opciones(0, 1) = "Dia":               Opciones(0, 2) = "0":      Opciones(0, 3) = "Ver Un Dia"
    Opciones(1, 0) = Str(CALENDARIO_2):    Opciones(1, 1) = "Semanana Laboral":  Opciones(1, 2) = "0":      Opciones(1, 3) = "Ver Semana de Trabajo"
    Opciones(2, 0) = Str(CALENDARIO_3):    Opciones(2, 1) = "Semana":            Opciones(2, 2) = "0":      Opciones(2, 3) = "Ver Semana Completa"
    Opciones(3, 0) = Str(CALENDARIO_4):    Opciones(3, 1) = "Mes":               Opciones(3, 2) = "0":      Opciones(3, 3) = "Ver Mes Completo"
    Opciones(4, 0) = Str(IMPRESORA):       Opciones(4, 1) = "Imprimir":          Opciones(4, 2) = "1":      Opciones(4, 3) = "Imprimir"
    Opciones(5, 0) = Str(SALIR):           Opciones(5, 1) = "Salir":             Opciones(5, 2) = "1":      Opciones(5, 3) = "Salir"
    
    Dim xFun As New eps_librerias.Codejock
    PocisionarContenedor
    xFun.BORRARMENU = True
    xFun.CrearToolBar Opciones, CommandBars, ImageManager, TAMAÑO_TOOL
    Set xFun = Nothing
    '
    'LlenarDatos
    AbriConeccion
    CargarData
    Set xConTMP = Nothing
End Sub

Sub PocisionarContenedor()
    EO1.Left = 0
    
    If TAMAÑO_TOOL = TOOL_TAMAÑO_ICO.I48x48 Then EO1.Top = 870
    If TAMAÑO_TOOL = TOOL_TAMAÑO_ICO.I32x32 Then EO1.Top = 630
    If TAMAÑO_TOOL = TOOL_TAMAÑO_ICO.I24x24 Then EO1.Top = 500
    If TAMAÑO_TOOL = TOOL_TAMAÑO_ICO.I16x16 Then EO1.Top = 380
End Sub

Sub AbriConeccion()
    Dim xFun As New eps_librerias.FuncionesData
    
    xFun.F_BASEDATOS = xRutaData
    xFun.F_GRUPOTRABAJO = xRutaFileTrabajo
    xFun.F_PASSWORD = Eps_Pass
    xFun.F_USUARIO = Eps_User
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"
    Set xConTMP = xFun.AbrirConeccion
End Sub

Sub CargarData()
    CalendarControl.SetDataProvider xConTMP.ConnectionString

    CalendarControl.DataProvider.Create

    If Not CalendarControl.DataProvider.Open Then
        CalendarControl.DataProvider.Create
    End If
    CalendarControl.DataProvider.Open

    CalendarControl.ActiveView.ShowDay Date
    CalendarControl.ViewType = xtpCalendarWorkWeekView
    CalendarControl.DayView.ScrollToWorkDayBegin

    wndDatePicker.AttachToCalendar CalendarControl
    CalendarControl.Populate
    CalendarControl.RedrawControl
End Sub

Private Sub Form_Resize()
    EO1.Width = Me.Width - 125
    

    If TAMAÑO_TOOL = TOOL_TAMAÑO_ICO.I48x48 Then
        If Me.Height <= (400 + 870) Then
            Me.Height = (400 + 870)
            Exit Sub
        End If
        EO1.Height = (Me.Height - (400 + 870)) 'EO1.Top = 870
    End If
    
    If TAMAÑO_TOOL = TOOL_TAMAÑO_ICO.I32x32 Then
        If Me.Height <= (400 + 630) Then
            Me.Height = (400 + 630)
            Exit Sub
        End If
        EO1.Height = (Me.Height - (400 + 630))
    End If
    
    If TAMAÑO_TOOL = TOOL_TAMAÑO_ICO.I24x24 Then
        If Me.Height <= (400 + 500) Then
            Me.Height = (400 + 500)
            Exit Sub
        End If
        EO1.Height = (Me.Height - (400 + 500)) 'EO1.Top = '500
    End If
    
    If TAMAÑO_TOOL = TOOL_TAMAÑO_ICO.I16x16 Then
        If Me.Height <= (400 + 380) Then
            Me.Height = (400 + 380)
            Exit Sub
        End If
        EO1.Height = (Me.Height - (400 + 380)) 'EO1.Top = 380
    End If
End Sub

Private Sub menu_02_01_Click()
    FrmAsunto.Show vbModal
End Sub



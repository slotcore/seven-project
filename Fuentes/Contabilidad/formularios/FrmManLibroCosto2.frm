VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmManLibroCosto2 
   Caption         =   "Contabilidad - Libro de Costos de Producción"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frm4 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   3200
      Left            =   60
      TabIndex        =   41
      Top             =   8010
      Visible         =   0   'False
      Width           =   7620
      Begin VB.CommandButton cmd 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   330
         Index           =   14
         Left            =   1320
         TabIndex        =   55
         ToolTipText     =   "Agregar Personal"
         Top             =   2800
         Width           =   1200
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Aceptar"
         Enabled         =   0   'False
         Height          =   330
         Index           =   13
         Left            =   90
         TabIndex        =   54
         ToolTipText     =   "Agregar Personal"
         Top             =   2800
         Width           =   1200
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Elimi&nar"
         Enabled         =   0   'False
         Height          =   330
         Index           =   7
         Left            =   6330
         TabIndex        =   53
         TabStop         =   0   'False
         ToolTipText     =   "Eliminar Personal"
         Top             =   1920
         Width           =   1200
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Eliminar Todo"
         Enabled         =   0   'False
         Height          =   330
         Index           =   8
         Left            =   6330
         TabIndex        =   52
         ToolTipText     =   "Agregar Personal"
         Top             =   2280
         Width           =   1200
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Sel. Anterior"
         Enabled         =   0   'False
         Height          =   330
         Index           =   6
         Left            =   6330
         TabIndex        =   51
         ToolTipText     =   "Agregar Personal"
         Top             =   1320
         Width           =   1200
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Agregar"
         Enabled         =   0   'False
         Height          =   330
         Index           =   4
         Left            =   6330
         TabIndex        =   50
         ToolTipText     =   "Agregar Personal"
         Top             =   375
         Width           =   1200
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Seleccionar"
         Enabled         =   0   'False
         Height          =   330
         Index           =   5
         Left            =   6330
         TabIndex        =   49
         TabStop         =   0   'False
         ToolTipText     =   "Eliminar Personal"
         Top             =   750
         Width           =   1200
      End
      Begin VB.PictureBox PbCerrar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   7350
         Picture         =   "FrmManLibroCosto2.frx":0000
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   42
         ToolTipText     =   "Cerrar"
         Top             =   50
         Width           =   195
      End
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   2370
         Index           =   1
         Left            =   90
         TabIndex        =   43
         Top             =   360
         Width           =   6195
         _cx             =   10927
         _cy             =   4180
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   128
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmManLibroCosto2.frx":02EC
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Shape Shape2 
         Height          =   345
         Left            =   4665
         Top             =   2760
         Width           =   1635
      End
      Begin VB.Label lblTotalGr 
         Alignment       =   2  'Center
         Caption         =   "lblTotalGr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   195
         Left            =   4695
         TabIndex        =   45
         Top             =   2850
         Width           =   1590
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   4
         X1              =   7590
         X2              =   7590
         Y1              =   0
         Y2              =   3170
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   6
         X1              =   0
         X2              =   8295
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   7
         X1              =   0
         X2              =   7590
         Y1              =   3170
         Y2              =   3170
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opciones de Distribución de Gastos de Fábrica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   105
         TabIndex        =   44
         Top             =   45
         Width           =   4020
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   45
         Top             =   30
         Width           =   7515
      End
   End
   Begin VB.Frame FraProgreso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   7710
      TabIndex        =   36
      Top             =   8040
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   225
         TabIndex        =   37
         Top             =   420
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cancelar = ESC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   4470
         TabIndex        =   40
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Procesando:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   39
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label LblProg 
         AutoSize        =   -1  'True
         Caption         =   "LblProg"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1350
         TabIndex        =   38
         Top             =   180
         Width           =   525
      End
      Begin VB.Shape Shape1 
         Height          =   885
         Index           =   0
         Left            =   60
         Top             =   90
         Width           =   5805
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6570
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto2.frx":0385
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto2.frx":08C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto2.frx":0C5B
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto2.frx":0DDF
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto2.frx":1233
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto2.frx":134B
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto2.frx":188F
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto2.frx":1DD3
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto2.frx":1EE7
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto2.frx":1FFB
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto2.frx":244F
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto2.frx":25BB
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto2.frx":2B03
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Solicitud de Materiales"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Solicitud de Linea"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7590
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   11925
      _cx             =   21034
      _cy             =   13388
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   2
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   8388608
      Caption         =   "  &Consulta  |   &Detalle  "
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   0
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7170
         Left            =   45
         TabIndex        =   7
         Top             =   375
         Width           =   11835
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6555
            Left            =   30
            TabIndex        =   9
            Top             =   480
            Width           =   11760
            _ExtentX        =   20743
            _ExtentY        =   11562
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Id"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Mes"
            Columns(1).DataField=   "desmes"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Descripcion"
            Columns(2).DataField=   "descripcion"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Método Valorización"
            Columns(3).DataField=   "desmetval"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Aplica Gas.Fab."
            Columns(4).DataField=   "desaplgasfab"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Tip. Dist. Gas. Fab."
            Columns(5).DataField=   "destipdisgasfab"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1455"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1376"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2064"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1984"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=5345"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=5265"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=4974"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=4895"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=3493"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=3413"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=3731"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=3651"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   0
            DefColWidth     =   0
            HeadLines       =   1.5
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12632256
            RowDividerColor =   12632256
            RowSubDividerColor=   12632256
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HE0FEFE&,.fgcolor=&H0&,.bold=0"
            _StyleDefs(7)   =   ":id=1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H800000&"
            _StyleDefs(11)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80&"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13,.alignment=3"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14,.alignment=2"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15,.alignment=3"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=70,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=67,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=68,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=69,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14,.alignment=2"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Named:id=33:Normal"
            _StyleDefs(61)  =   ":id=33,.parent=0"
            _StyleDefs(62)  =   "Named:id=34:Heading"
            _StyleDefs(63)  =   ":id=34,.parent=33,.alignment=3,.valignment=2,.bgcolor=&H8000000F&"
            _StyleDefs(64)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1"
            _StyleDefs(65)  =   "Named:id=35:Footing"
            _StyleDefs(66)  =   ":id=35,.parent=33,.alignment=3,.valignment=2,.bgcolor=&H8000000F&"
            _StyleDefs(67)  =   ":id=35,.fgcolor=&H80000012&"
            _StyleDefs(68)  =   "Named:id=36:Selected"
            _StyleDefs(69)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(70)  =   "Named:id=37:Caption"
            _StyleDefs(71)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(72)  =   "Named:id=38:HighlightRow"
            _StyleDefs(73)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(74)  =   "Named:id=39:EvenRow"
            _StyleDefs(75)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(76)  =   "Named:id=40:OddRow"
            _StyleDefs(77)  =   ":id=40,.parent=33"
            _StyleDefs(78)  =   "Named:id=41:RecordSelector"
            _StyleDefs(79)  =   ":id=41,.parent=34"
            _StyleDefs(80)  =   "Named:id=42:FilterBar"
            _StyleDefs(81)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Libro de Costo de Producción"
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
            Index           =   0
            Left            =   45
            TabIndex        =   8
            Top             =   45
            Width           =   11685
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7170
         Left            =   12570
         TabIndex        =   5
         Top             =   375
         Width           =   11835
         Begin VB.CommandButton cmd 
            Caption         =   "&Procesar"
            Enabled         =   0   'False
            Height          =   350
            Index           =   3
            Left            =   10230
            TabIndex        =   35
            ToolTipText     =   "Procesar las Tareas del Producto Seleccionado"
            Top             =   1110
            Width           =   1400
         End
         Begin VB.CommandButton cmd 
            Caption         =   "C&onsultar"
            Enabled         =   0   'False
            Height          =   350
            Index           =   2
            Left            =   10230
            TabIndex        =   34
            ToolTipText     =   "Procesar las Tareas del Producto Seleccionado"
            Top             =   750
            Width           =   1400
         End
         Begin VB.CommandButton cmd 
            Caption         =   "&Config. Distrib."
            Height          =   350
            Index           =   1
            Left            =   10230
            TabIndex        =   33
            ToolTipText     =   "Procesar las Tareas del Producto Seleccionado"
            Top             =   390
            Width           =   1400
         End
         Begin VB.Frame Frame9 
            Caption         =   "[ Gastos de Fábrica ]"
            ForeColor       =   &H00800000&
            Height          =   1035
            Left            =   5910
            TabIndex        =   25
            Top             =   330
            Width           =   4185
            Begin VB.Frame Frame3 
               BorderStyle     =   0  'None
               Caption         =   "Frame3"
               Height          =   465
               Left            =   90
               TabIndex        =   46
               Top             =   510
               Width           =   1785
               Begin VB.OptionButton optdisgasfab 
                  Caption         =   "Todos"
                  Height          =   225
                  Index           =   0
                  Left            =   90
                  TabIndex        =   48
                  Top             =   120
                  Width           =   795
               End
               Begin VB.OptionButton optdisgasfab 
                  Caption         =   "Ventas"
                  Height          =   225
                  Index           =   1
                  Left            =   975
                  TabIndex        =   47
                  Top             =   120
                  Width           =   795
               End
            End
            Begin VB.OptionButton opttipdiscta 
               Caption         =   "Distribuida"
               Height          =   225
               Index           =   1
               Left            =   3020
               TabIndex        =   29
               Top             =   650
               Width           =   1065
            End
            Begin VB.OptionButton opttipdiscta 
               Caption         =   "Global"
               Height          =   225
               Index           =   0
               Left            =   2090
               TabIndex        =   28
               Top             =   650
               Width           =   885
            End
            Begin VB.Line Line4 
               X1              =   1980
               X2              =   1980
               Y1              =   210
               Y2              =   950
            End
            Begin VB.Line Line1 
               BorderColor     =   &H80000005&
               X1              =   2000
               X2              =   2000
               Y1              =   210
               Y2              =   950
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Distribubión de Cta.:"
               Height          =   195
               Left            =   2150
               TabIndex        =   27
               Top             =   330
               Width           =   2010
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Aplicar Distribucion a:"
               Height          =   195
               Left            =   120
               TabIndex        =   26
               Top             =   330
               Width           =   1530
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "[ Datos de Producción ]"
            Height          =   5685
            Left            =   0
            TabIndex        =   15
            Top             =   1380
            Width           =   11775
            Begin SizerOneLibCtl.TabOne TabOne2 
               Height          =   3735
               Left            =   60
               TabIndex        =   16
               Top             =   1830
               Width           =   11655
               _cx             =   20558
               _cy             =   6588
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
               Appearance      =   2
               MousePointer    =   0
               _ConvInfo       =   1
               Version         =   700
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               FrontTabColor   =   -2147483633
               BackTabColor    =   12632256
               TabOutlineColor =   -2147483632
               FrontTabForeColor=   -2147483630
               Caption         =   "   &Mat. Pri.  |    &Man. Obr.   |  &Gas. fab.  "
               Align           =   0
               CurrTab         =   2
               FirstTab        =   0
               Style           =   0
               Position        =   1
               AutoSwitch      =   -1  'True
               AutoScroll      =   -1  'True
               TabPreview      =   -1  'True
               ShowFocusRect   =   -1  'True
               TabsPerPage     =   0
               BorderWidth     =   0
               BoldCurrent     =   -1  'True
               DogEars         =   -1  'True
               MultiRow        =   0   'False
               MultiRowOffset  =   200
               CaptionStyle    =   0
               TabHeight       =   0
               TabCaptionPos   =   4
               TabPicturePos   =   0
               CaptionEmpty    =   ""
               Separators      =   0   'False
               Begin VB.Frame Frame6 
                  Caption         =   "[ Importe de Gastos de Fábrica ]"
                  Height          =   3360
                  Left            =   45
                  TabIndex        =   19
                  Top             =   45
                  Width           =   11565
                  Begin VB.CommandButton cmd 
                     Caption         =   "Eliminar Todos"
                     Enabled         =   0   'False
                     Height          =   330
                     Index           =   12
                     Left            =   10110
                     TabIndex        =   23
                     ToolTipText     =   "Eliminar Todos"
                     Top             =   2220
                     Width           =   1400
                  End
                  Begin VB.CommandButton cmd 
                     Caption         =   "&Seleccionar"
                     Enabled         =   0   'False
                     Height          =   330
                     Index           =   10
                     Left            =   10110
                     TabIndex        =   22
                     TabStop         =   0   'False
                     ToolTipText     =   "Agregar Personal de una Lista"
                     Top             =   600
                     Width           =   1400
                  End
                  Begin VB.CommandButton cmd 
                     Caption         =   "&Eliminar"
                     Enabled         =   0   'False
                     Height          =   330
                     Index           =   11
                     Left            =   10110
                     TabIndex        =   21
                     TabStop         =   0   'False
                     ToolTipText     =   "Eliminar Personal"
                     Top             =   1860
                     Width           =   1400
                  End
                  Begin VB.CommandButton cmd 
                     Caption         =   "Agregar"
                     Enabled         =   0   'False
                     Height          =   330
                     Index           =   9
                     Left            =   10110
                     TabIndex        =   20
                     ToolTipText     =   "Agregar Personal"
                     Top             =   240
                     Width           =   1400
                  End
                  Begin VSFlex7Ctl.VSFlexGrid fg 
                     Height          =   2970
                     Index           =   2
                     Left            =   60
                     TabIndex        =   24
                     Top             =   270
                     Width           =   9945
                     _cx             =   17542
                     _cy             =   5239
                     _ConvInfo       =   1
                     Appearance      =   0
                     BorderStyle     =   1
                     Enabled         =   -1  'True
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MousePointer    =   0
                     BackColor       =   -2147483643
                     ForeColor       =   -2147483640
                     BackColorFixed  =   -2147483633
                     ForeColorFixed  =   -2147483630
                     BackColorSel    =   128
                     ForeColorSel    =   -2147483634
                     BackColorBkg    =   -2147483636
                     BackColorAlternate=   -2147483643
                     GridColor       =   -2147483633
                     GridColorFixed  =   -2147483632
                     TreeColor       =   -2147483632
                     FloodColor      =   192
                     SheetBorder     =   -2147483642
                     FocusRect       =   1
                     HighLight       =   1
                     AllowSelection  =   -1  'True
                     AllowBigSelection=   -1  'True
                     AllowUserResizing=   1
                     SelectionMode   =   0
                     GridLines       =   1
                     GridLinesFixed  =   2
                     GridLineWidth   =   1
                     Rows            =   2
                     Cols            =   5
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   0
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   0   'False
                     FormatString    =   $"FrmManLibroCosto2.frx":2E95
                     ScrollTrack     =   0   'False
                     ScrollBars      =   3
                     ScrollTips      =   0   'False
                     MergeCells      =   0
                     MergeCompare    =   0
                     AutoResize      =   -1  'True
                     AutoSizeMode    =   0
                     AutoSearch      =   0
                     AutoSearchDelay =   2
                     MultiTotals     =   -1  'True
                     SubtotalPosition=   1
                     OutlineBar      =   0
                     OutlineCol      =   0
                     Ellipsis        =   0
                     ExplorerBar     =   0
                     PicturesOver    =   0   'False
                     FillStyle       =   0
                     RightToLeft     =   0   'False
                     PictureType     =   0
                     TabBehavior     =   0
                     OwnerDraw       =   0
                     Editable        =   0
                     ShowComboButton =   1
                     WordWrap        =   0   'False
                     TextStyle       =   0
                     TextStyleFixed  =   0
                     OleDragMode     =   0
                     OleDropMode     =   0
                     DataMode        =   0
                     VirtualData     =   -1  'True
                     DataMember      =   ""
                     ComboSearch     =   3
                     AutoSizeMouse   =   -1  'True
                     FrozenRows      =   0
                     FrozenCols      =   0
                     AllowUserFreezing=   0
                     BackColorFrozen =   0
                     ForeColorFrozen =   0
                     WallPaperAlignment=   9
                  End
               End
               Begin VB.Frame Frame8 
                  Caption         =   "[ Importe de Mano de Obra ]"
                  ForeColor       =   &H00800000&
                  Height          =   3360
                  Left            =   -12210
                  TabIndex        =   18
                  Top             =   45
                  Width           =   11565
                  Begin VSFlex7Ctl.VSFlexGrid fg 
                     Height          =   3015
                     Index           =   4
                     Left            =   60
                     TabIndex        =   31
                     Top             =   240
                     Width           =   11415
                     _cx             =   20135
                     _cy             =   5318
                     _ConvInfo       =   1
                     Appearance      =   0
                     BorderStyle     =   1
                     Enabled         =   -1  'True
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MousePointer    =   0
                     BackColor       =   -2147483643
                     ForeColor       =   -2147483640
                     BackColorFixed  =   -2147483633
                     ForeColorFixed  =   -2147483630
                     BackColorSel    =   128
                     ForeColorSel    =   -2147483634
                     BackColorBkg    =   -2147483636
                     BackColorAlternate=   -2147483643
                     GridColor       =   -2147483633
                     GridColorFixed  =   -2147483632
                     TreeColor       =   -2147483632
                     FloodColor      =   192
                     SheetBorder     =   -2147483642
                     FocusRect       =   1
                     HighLight       =   1
                     AllowSelection  =   -1  'True
                     AllowBigSelection=   -1  'True
                     AllowUserResizing=   1
                     SelectionMode   =   0
                     GridLines       =   1
                     GridLinesFixed  =   2
                     GridLineWidth   =   1
                     Rows            =   2
                     Cols            =   7
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   0
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   0   'False
                     FormatString    =   $"FrmManLibroCosto2.frx":2F33
                     ScrollTrack     =   0   'False
                     ScrollBars      =   3
                     ScrollTips      =   0   'False
                     MergeCells      =   0
                     MergeCompare    =   0
                     AutoResize      =   -1  'True
                     AutoSizeMode    =   0
                     AutoSearch      =   0
                     AutoSearchDelay =   2
                     MultiTotals     =   -1  'True
                     SubtotalPosition=   1
                     OutlineBar      =   0
                     OutlineCol      =   0
                     Ellipsis        =   0
                     ExplorerBar     =   0
                     PicturesOver    =   0   'False
                     FillStyle       =   0
                     RightToLeft     =   0   'False
                     PictureType     =   0
                     TabBehavior     =   0
                     OwnerDraw       =   0
                     Editable        =   0
                     ShowComboButton =   1
                     WordWrap        =   0   'False
                     TextStyle       =   0
                     TextStyleFixed  =   0
                     OleDragMode     =   0
                     OleDropMode     =   0
                     DataMode        =   0
                     VirtualData     =   -1  'True
                     DataMember      =   ""
                     ComboSearch     =   3
                     AutoSizeMouse   =   -1  'True
                     FrozenRows      =   0
                     FrozenCols      =   0
                     AllowUserFreezing=   0
                     BackColorFrozen =   0
                     ForeColorFrozen =   0
                     WallPaperAlignment=   9
                  End
               End
               Begin VB.Frame Frame7 
                  Caption         =   "[ Importe de Materia Prima ]"
                  ForeColor       =   &H00800000&
                  Height          =   3360
                  Left            =   -12510
                  TabIndex        =   17
                  Top             =   45
                  Width           =   11565
                  Begin VSFlex7Ctl.VSFlexGrid fg 
                     Height          =   3015
                     Index           =   3
                     Left            =   90
                     TabIndex        =   30
                     Top             =   240
                     Width           =   11370
                     _cx             =   20055
                     _cy             =   5318
                     _ConvInfo       =   1
                     Appearance      =   0
                     BorderStyle     =   1
                     Enabled         =   -1  'True
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MousePointer    =   0
                     BackColor       =   -2147483643
                     ForeColor       =   -2147483640
                     BackColorFixed  =   -2147483633
                     ForeColorFixed  =   -2147483630
                     BackColorSel    =   -2147483635
                     ForeColorSel    =   -2147483634
                     BackColorBkg    =   -2147483636
                     BackColorAlternate=   -2147483643
                     GridColor       =   -2147483633
                     GridColorFixed  =   -2147483632
                     TreeColor       =   -2147483632
                     FloodColor      =   192
                     SheetBorder     =   -2147483642
                     FocusRect       =   1
                     HighLight       =   1
                     AllowSelection  =   -1  'True
                     AllowBigSelection=   -1  'True
                     AllowUserResizing=   0
                     SelectionMode   =   0
                     GridLines       =   1
                     GridLinesFixed  =   2
                     GridLineWidth   =   1
                     Rows            =   2
                     Cols            =   8
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   0
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   0   'False
                     FormatString    =   $"FrmManLibroCosto2.frx":3006
                     ScrollTrack     =   0   'False
                     ScrollBars      =   3
                     ScrollTips      =   0   'False
                     MergeCells      =   0
                     MergeCompare    =   0
                     AutoResize      =   -1  'True
                     AutoSizeMode    =   0
                     AutoSearch      =   0
                     AutoSearchDelay =   2
                     MultiTotals     =   -1  'True
                     SubtotalPosition=   1
                     OutlineBar      =   0
                     OutlineCol      =   0
                     Ellipsis        =   0
                     ExplorerBar     =   0
                     PicturesOver    =   0   'False
                     FillStyle       =   0
                     RightToLeft     =   0   'False
                     PictureType     =   0
                     TabBehavior     =   0
                     OwnerDraw       =   0
                     Editable        =   0
                     ShowComboButton =   1
                     WordWrap        =   0   'False
                     TextStyle       =   0
                     TextStyleFixed  =   0
                     OleDragMode     =   0
                     OleDropMode     =   0
                     DataMode        =   0
                     VirtualData     =   -1  'True
                     DataMember      =   ""
                     ComboSearch     =   3
                     AutoSizeMouse   =   -1  'True
                     FrozenRows      =   0
                     FrozenCols      =   0
                     AllowUserFreezing=   0
                     BackColorFrozen =   0
                     ForeColorFrozen =   0
                     WallPaperAlignment=   9
                  End
               End
            End
            Begin VSFlex7Ctl.VSFlexGrid fg 
               Height          =   1455
               Index           =   0
               Left            =   60
               TabIndex        =   32
               Top             =   300
               Width           =   11625
               _cx             =   20505
               _cy             =   2566
               _ConvInfo       =   1
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483642
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   25
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManLibroCosto2.frx":30FC
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   0
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   0   'False
               PictureType     =   0
               TabBehavior     =   0
               OwnerDraw       =   0
               Editable        =   0
               ShowComboButton =   1
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               DataMode        =   0
               VirtualData     =   -1  'True
               DataMember      =   ""
               ComboSearch     =   3
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
            End
         End
         Begin VB.TextBox txtdescripcion 
            Height          =   300
            Left            =   1170
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   1
            Text            =   "txtdescripcion"
            Top             =   690
            Width           =   4635
         End
         Begin VB.CommandButton cmd 
            Height          =   240
            Index           =   0
            Left            =   1830
            Picture         =   "FrmManLibroCosto2.frx":33C0
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   1050
            Width           =   240
         End
         Begin VB.ComboBox cbMes 
            Height          =   315
            Left            =   1170
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   390
            Width           =   4665
         End
         Begin VB.TextBox txtidmetval 
            Height          =   300
            Left            =   1170
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   2
            Text            =   "txtidmetval"
            Top             =   1020
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Met. Val."
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   1065
            Width           =   630
         End
         Begin VB.Label lblmetval 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblmetval"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   2115
            TabIndex        =   13
            Top             =   1020
            Width           =   3690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Mes"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   11
            Top             =   420
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   765
            Width           =   840
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Libro de Costo de Producción"
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
            Left            =   60
            TabIndex        =   6
            Top             =   75
            Width           =   11685
         End
      End
   End
   Begin VB.Menu menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Insertar"
      End
      Begin VB.Menu menu1_2 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Ver Receta"
      End
   End
End
Attribute VB_Name = "FrmManLibroCosto2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------VARIABLES DE ESTADO DE FORMULARIO
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim Agregando As Boolean               ' para saber cuando se este agregando FILAS AL CONTROL grid de productos
Dim IdMenuActivo As Integer            ' INDICA EL CODIGO DEL MENU ACTIVO
Dim xHorIni As Date                    ' ESPECIFICA LA HORA DE INICIO
Dim mIdRegistro&                       ' identificador del registro
Dim mMesActivo As Integer              ' indica el mes activo
Dim OrigFX As Long
Dim OrigFY As Long
Dim fOrdenLista As Boolean              ' especfica el orden de la lista de la consulta
'***********************************************
'-----------------------VARIABLES DE FORMULARIO
'***********************************************
Dim Rst As New ADODB.Recordset
Dim RstLibro As New ADODB.Recordset
Dim cSQL As String
Dim ESTADOANTERIOR_ As Double

Dim BANDERA_ As Boolean
Dim RECORDSETPREUNI_ As New ADODB.Recordset
Dim RECORDSETMOBRA_ As New ADODB.Recordset
Dim RECORDSETGFABRICA_ As New ADODB.Recordset
Dim RECORDSETERRORES_ As New ADODB.Recordset

Dim RSTCABECERA As New ADODB.Recordset
Dim RSTCTAGASFAB As New ADODB.Recordset
Dim RSTDETALLEMATPRI As New ADODB.Recordset
Dim RSTDETALLEMANOBR As New ADODB.Recordset
Dim RSTDETALLEGASFAB As New ADODB.Recordset
Dim CORRELATIVO_ As Double
Dim FILAINICIAL_ As Integer

'-----------------------PROPIEDADES DE PROCESADO
' -----ESTRUCTURA
Private Type PROPIEDADESPROCESADO_
    MODOTAREA_  As Integer
    PORCENTAJE_  As Double
    MINUTOS_ As Date
    INCLUIRREFRIGERIO_ As Boolean
    HORINIREFRIGERIO_ As Date
    HORFINREFRIGERIO_  As Date
    LIMITARNUMEROTAREAS_ As Boolean
    LIMITARNUMEROPERSONAL_ As Boolean
    LIMITARSELPERSONAL_ As Boolean
End Type
' -----TIPO
Dim PROPIEDADES_ As PROPIEDADESPROCESADO_
' ----------------------DEFINICION DE COLUMNAS
Private Enum COLUMNACABECERA_
    COLUMNAFECHA_ = 1
    COLUMNAREGPROD_
    COLUMNATIPO_
    COLUMNAPROCESO_
    COLUMNAITEM_
    COLUMNARECETA_
    COLUMNARESPONSABLE_
    COLUMNAUNIMED_
    COLUMNACANTIDAD_
    COLUMNAHORINI_
    COLUMNAHORFIN_
    COLUMNACOSTOMP_
    COLUMNACOSTOMOBRA_
    COLUMNACOSTOPRIMO_
    COLUMNACOSTOFABRICA_
    COLUMNACOSTOTOTAL_
    COLUMNACOSTOUNIPRODUCCION_
    COLUMNAPRECIOVENTA_
    COLUMNAIMPORTEVENTA_
    COLUMNADESVIACION_
    COLUMNADESVIACIONPORC_
    COLUMNAIDPROD_
    COLUMNAIDITEM_
    COLUMNACORRELATIVO_
End Enum

Private Enum COLUMNADETALLETAREA_
    SEL_ = 1
    TAREA_
    DURACION_
    HORINI_
    HORFIN_
    NUMOP_
    CANTIDADSUM_
    CANTIDADPROC_
    FCHINI_
    FCHFIN_
    AREA_
    RESPONSABLE_
    IDTAR_
    IDAREA_
    IDRESP_
End Enum

Private Enum COLUMNADETALLEPERS_
    DNI_ = 1
    NOMBRE_
    IDPER_
End Enum

Private Enum COLUMNADETALLEREPROC_
    LOTE_ = 1
    ALMACEN_
    CANTIDAD_
    IDLOTE_
    IDLOTEDET_
    IDALM_
End Enum

Private Enum COLUMNADETALLEINSUMOS_
    INSUMO_ = 1
    UNIMED_
    CANTIDAD_
    IDINSUMO_
    IDUNIMED_
End Enum
' ----------------------DEFINICION DE ESTADOS
Const ESTADOPENDIENTE_ = 1
Const ESTADOPROCESADO_ = 2
Const ESTADOAPROBADO_ = 3
Const ESTADOANULADO_ = 4

Private Function aplicarCambios(PROCESO_ As Integer, MESATRABAJAR_ As Integer, _
                                Optional ESGASFAB_ As Boolean = False, _
                                Optional IDPROD_ As Integer, _
                                Optional PREUNIGASFAB_ As Double) As Boolean
    Dim xId As Double
    Dim xIdDet As Double
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    
    Dim ULTIMODIAMES_ As Date
    Dim PRIMERDIAMES_ As Date
    Dim ANIOACTUAL_ As Double
    Dim MESACTUAL_ As Double
    Dim nSQLId As String
    
    ANIOACTUAL_ = AnoTra
    MESACTUAL_ = MESATRABAJAR_
    ' Se encuentra el primer dia del mes actual
    PRIMERDIAMES_ = CDate("01/" & MESATRABAJAR_ & "/" & ANIOACTUAL_ & "")
    ' Se encuentra el ultimo dia del mes actual
    If MESACTUAL_ = 12 Then MESACTUAL_ = 0: ANIOACTUAL_ = ANIOACTUAL_ + 1
    ULTIMODIAMES_ = CDate("01/" & MESACTUAL_ + 1 & "/" & ANIOACTUAL_ & "") - 1
    ' Si es que haya habido algun cambio se regresan a su estado inicial
    ANIOACTUAL_ = AnoTra
    MESACTUAL_ = MESATRABAJAR_
    
On Error GoTo LaCague
    ' GRABA EL REGISTRO
    xCon.BeginTrans
    
    If ESGASFAB_ Then
        RST_Busq RstCab, "SELECT * FROM con_centrocostopreuni WHERE idprod=" & IDPROD_, xCon
        If RstCab.RecordCount = 0 Then
            aplicarCambios = False
            xCon.RollbackTrans
            Set RstCab = Nothing
            Set RstDet = Nothing
            Exit Function
        End If
        
        RstCab.MoveFirst
        While Not RstCab.EOF
            RstCab("pregfabrica") = PREUNIGASFAB_
            RstCab.Update
            RstCab.MoveNext
        Wend
    Else
        ' SE ELIMINA LOS COSTOS REGISTRADOS
        xCon.Execute "DELETE * FROM con_centrocostopreuni WHERE ((proceso=" & PROCESO_ & ") AND ((fecha>=CDate('" & PRIMERDIAMES_ & "')) AND (fecha<=CDate('" & ULTIMODIAMES_ & "'))))"
        RST_Busq RstCab, "SELECT TOP 1 * FROM con_centrocostopreuni", xCon
        
        RECORDSETPREUNI_.Filter = adFilterNone
        RECORDSETMOBRA_.Filter = adFilterNone
        If RECORDSETPREUNI_.RecordCount = 0 Then
            aplicarCambios = False
            xCon.RollbackTrans
            Set RstCab = Nothing
            Set RstDet = Nothing
            Exit Function
        End If
        
        RECORDSETPREUNI_.MoveFirst
        While Not RECORDSETPREUNI_.EOF
            RstCab.AddNew
            RstCab("idprod") = RECORDSETPREUNI_("idprod")
            RstCab("proceso") = PROCESO_
            RstCab("iditem") = RECORDSETPREUNI_("iditem")
            RstCab("fecha") = RECORDSETPREUNI_("fecha")
            RstCab("premprima") = NulosN(RECORDSETPREUNI_("preuni"))
            ' -----------------MANO DE OBRA
            RECORDSETMOBRA_.Filter = "idprod=" & NulosN(RECORDSETPREUNI_("idprod"))
            If RECORDSETMOBRA_.RecordCount = 0 Then
                RstCab("premobra") = 0
            Else
                RstCab("premobra") = NulosN(RECORDSETMOBRA_("preuni"))
            End If
            ' -----------------GASTOS DE FABRICA
            If RECORDSETGFABRICA_.State = 0 Then GoTo SIGUIENTE_
            RECORDSETGFABRICA_.Filter = "idprod=" & NulosN(RECORDSETPREUNI_("idprod"))
            If RECORDSETGFABRICA_.RecordCount = 0 Then
                RstCab("pregfabrica") = 0
            Else
                RstCab("pregfabrica") = NulosN(RECORDSETGFABRICA_("preuni"))
            End If
            RstCab("preuni") = NulosN(RstCab("premprima")) + NulosN(RstCab("premobra")) + NulosN(RstCab("pregfabrica"))
            
            RstCab("horini") = RECORDSETPREUNI_("horini")
            RstCab("horfin") = RECORDSETPREUNI_("horfin")
                    
SIGUIENTE_:
            RstCab.Update
            RECORDSETPREUNI_.MoveNext
        Wend
    End If
    xCon.CommitTrans
    Set RstCab = Nothing
    aplicarCambios = True
    Exit Function
    
LaCague:
    'Resume
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    aplicarCambios = False
End Function

Private Sub pExportar()
    Dim nSQL As String
    Dim oExport As New SGI2_funciones.formularios
    Dim xRs  As New ADODB.Recordset
    Dim xCampos() As String
    Dim TITULO_ As String
    
    ReDim xCampos(5, 3) As String
    xCampos(0, 0) = "Documento":                    xCampos(0, 1) = "numdoc":       xCampos(0, 2) = 0:      xCampos(0, 3) = "1500"
    xCampos(1, 0) = "Ítem":                         xCampos(1, 1) = "item":         xCampos(1, 2) = 0:      xCampos(1, 3) = "3500"
    xCampos(2, 0) = "Precio/Importe/Cantidad":      xCampos(2, 1) = "preuni":       xCampos(2, 2) = 0:      xCampos(2, 3) = "1200"
    xCampos(3, 0) = "Detalle Error":                xCampos(3, 1) = "detalle":      xCampos(3, 2) = 0:      xCampos(3, 3) = "3500"
    xCampos(4, 0) = "Fecha":                        xCampos(4, 1) = "fecha":        xCampos(4, 2) = 0:      xCampos(4, 3) = "1200"
    xCampos(5, 0) = "Insumo":                       xCampos(5, 1) = "insumo":       xCampos(5, 2) = 0:      xCampos(5, 3) = "3500"
    
    TITULO_ = "ERRORES DE PROCESAMIENTO DE COSTO"
    RECORDSETERRORES_.Filter = adFilterNone
    Set xRs = RECORDSETERRORES_
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , TITULO_, "", "", TITULO_, xRs, xCampos
    Set oExport = Nothing
    Set xRs = Nothing
End Sub

Private Sub pProcesarDatos(MESATRABAJAR_ As Integer)
    Dim xRs As New ADODB.Recordset
    Dim IDITEM_ As Integer
    Dim IDPROD_ As Integer
    Dim FECHA_ As String
    Dim VALOR_ As Double ' unid/hora de cada producto
    Dim TOTALHORAS_ As Double ' Tiempo en horas de cada producto
    Dim ULTIMODIAMES_ As Date
    Dim PRIMERDIAMES_ As Date
    Dim ANIOACTUAL_ As Double
    Dim MESACTUAL_ As Integer
    Dim nSQLId As String
    Dim nSQLIdNot As String
    Dim CONSULTA_ As String
    Dim NUMEROREGISTROS_ As Integer
    Dim PROCESO_ As Integer
    Dim IMPORTEPARCIAL_ As Double
    Dim INDICEFAB_ As Double
    Dim A As Integer
    Dim FILATOPE_ As Integer
    
    Dim IMPORTEMANOBR_ As Double
    Dim IMPORTEMATPRI_ As Double
    
    ' INICIALIZAMOS PROCESO Y NUMERO DE REGISTROS
    PROCESO_ = 0
    NUMEROREGISTROS_ = 1
    
    If RECORDSETERRORES_.RecordCount = 0 Then
        ' SE INICIALIZA LA FILA DE INICIO
        FILAINICIAL_ = fg(0).FixedRows
        ' SE LIMPIA EL GRID
        For A = FILAINICIAL_ To fg(0).Rows - 1
            fg(0).TextMatrix(A, COLUMNACOSTOMP_) = ""
            fg(0).TextMatrix(A, COLUMNACOSTOMOBRA_) = ""
            fg(0).TextMatrix(A, COLUMNACOSTOPRIMO_) = ""
            fg(0).TextMatrix(A, COLUMNACOSTOFABRICA_) = ""
            fg(0).TextMatrix(A, COLUMNACOSTOTOTAL_) = ""
            fg(0).TextMatrix(A, COLUMNACOSTOUNIPRODUCCION_) = ""
        Next A
    Else
        limpiarRST RECORDSETERRORES_
    End If
            
    fg(2).Rows = fg(2).FixedRows
    PgBar.Min = 0
    PgBar.Max = fg(0).Rows - 1
    PgBar.Value = 0
    
    For A = FILAINICIAL_ To fg(0).Rows - 1
        
        VALOR_ = 0
        TOTALHORAS_ = 0
        With fg(0)
            CentrarFrm FraProgreso
            FraProgreso.Visible = True
            lbl(0).Caption = "PROCESO: " & .TextMatrix(A, COLUMNAPROCESO_)
            FILAINICIAL_ = A
            Agregando = True
            DoEvents
            If BANDERA_ Then GoTo SALIR_
            IDITEM_ = NulosN(.TextMatrix(A, COLUMNAIDITEM_))
            If A <= 10 Then
                FILATOPE_ = A
            Else
                FILATOPE_ = A - 10
            End If
            .TopRow = FILATOPE_
            FraProgreso.Refresh
            LblProg.Caption = NulosC(.TextMatrix(A, COLUMNAITEM_))
            PgBar.Value = PgBar.Value + 1
            
            IDPROD_ = NulosN(.TextMatrix(A, COLUMNAIDPROD_))
            FECHA_ = NulosC(.TextMatrix(A, COLUMNAFECHA_))
            
            IMPORTEMATPRI_ = 0
            IMPORTEMANOBR_ = 0
            IMPORTEMATPRI_ = pImporteMateriaPrima(IDITEM_, NulosN(.TextMatrix(A, COLUMNACANTIDAD_)), FECHA_, NulosC(.TextMatrix(A, COLUMNAHORINI_)), NulosC(.TextMatrix(A, COLUMNAHORFIN_)), xCon, 0, "P", IDPROD_, NulosN(.TextMatrix(A, COLUMNACORRELATIVO_)))
            IMPORTEMANOBR_ = pImporteManoObra(IDITEM_, FECHA_, xCon, IDPROD_, NulosN(.TextMatrix(A, COLUMNACORRELATIVO_)))
            
            ' SE ADICIONA LA MANO DE OBRA AL PRECIO PRIMO
            RECORDSETPREUNI_.Filter = "idprod=" & IDPROD_
            If RECORDSETPREUNI_.RecordCount > 0 Then
                RECORDSETPREUNI_("impprimo") = RECORDSETPREUNI_("impprimo") + IMPORTEMANOBR_
            End If
            
            If IMPORTEMANOBR_ < 0 Then
                MsgBox "Ocurrio un error al procesar el costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                RECORDSETERRORES_.AddNew
                RECORDSETERRORES_("numdoc") = NulosC(.TextMatrix(A, COLUMNAREGPROD_))
                RECORDSETERRORES_("item") = IDITEM_
                RECORDSETERRORES_("preuni") = NulosN(.TextMatrix(A, COLUMNACOSTOMOBRA_))
                RECORDSETERRORES_("detalle") = "Mano de Obra - Precio unitario negativo"
                RECORDSETERRORES_.Update
                GoTo SALIR_
            ElseIf IMPORTEMANOBR_ = 0 Then
                MsgBox "Ocurrio un error al procesar el costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                RECORDSETERRORES_.AddNew
                RECORDSETERRORES_("numdoc") = NulosC(.TextMatrix(A, COLUMNAREGPROD_))
                RECORDSETERRORES_("item") = IDITEM_
                RECORDSETERRORES_("preuni") = NulosN(.TextMatrix(A, COLUMNACOSTOMOBRA_))
                RECORDSETERRORES_("detalle") = "Mano de Obra - Precio unitario cero"
                RECORDSETERRORES_.Update
                GoTo SALIR_
            End If
            
            ' SE AGREGAN LOS DATOS AUXILIARES AL LIBRO DE COSTO
            RSTCABECERA.Filter = "id=" & NulosN(.TextMatrix(A, COLUMNACORRELATIVO_))
            RSTCABECERA("impmprima") = IMPORTEMATPRI_
            RSTCABECERA("impmanobr") = IMPORTEMANOBR_
            RSTCABECERA.Update
            
            .TextMatrix(A, COLUMNACOSTOMP_) = Format(IMPORTEMATPRI_, FORMAT_IMPORTEKARDEX)
            .TextMatrix(A, COLUMNACOSTOMOBRA_) = Format(IMPORTEMANOBR_, FORMAT_IMPORTEKARDEX)
            .TextMatrix(A, COLUMNACOSTOPRIMO_) = Format(NulosN(.TextMatrix(A, COLUMNACOSTOMP_)) + NulosN(.TextMatrix(A, COLUMNACOSTOMOBRA_)), "0.0000")
        End With
    Next A
    
GASTOSDEFABRICA_:
    IMPORTEPARCIAL_ = GRID_SUMAR_COL(fg(0), COLUMNACOSTOPRIMO_)
    RSTCTAGASFAB.Filter = adFilterNone
    INDICEFAB_ = NulosN(RST_SUMAR(RSTCTAGASFAB, "importe")) / IMPORTEPARCIAL_
    lbl(2).Caption = "APLICANDO GASTOS DE FABRICA"
    For A = fg(0).FixedRows To fg(0).Rows - 1
        DoEvents
        lbl(0).Caption = "PROCESO: " & fg(0).TextMatrix(A, COLUMNAPROCESO_)
        LblProg.Caption = NulosC(fg(0).TextMatrix(A, COLUMNAITEM_))
        
        fg(0).TopRow = A
        If optdisgasfab(1).Value = True Then '----------- SOLO VENTAS
            If NulosC(fg(0).TextMatrix(A, COLUMNATIPO_)) = "V" Then
                fg(0).TextMatrix(A, COLUMNACOSTOFABRICA_) = Format(fg(0).TextMatrix(A, COLUMNACOSTOPRIMO_) * INDICEFAB_, FORMAT_IMPORTEKARDEX)
            Else
                INDICEFAB_ = 0
                fg(0).TextMatrix(fg(0).Rows - 1, COLUMNACOSTOFABRICA_) = Format(0, FORMAT_IMPORTEKARDEX)
            End If
        Else '------------------------------------------- TODOS
            fg(0).TextMatrix(A, COLUMNACOSTOFABRICA_) = Format(INDICEFAB_ * NulosN(fg(0).TextMatrix(A, COLUMNACOSTOPRIMO_)), FORMAT_IMPORTEKARDEX)
        End If
        
        ' SE AGREGAN LOS DATOS AUXILIARES AL LIBRO DE COSTO
        RSTCABECERA.Filter = "id=" & fg(0).TextMatrix(A, COLUMNACORRELATIVO_)
        RSTCABECERA("impgasfab") = NulosN(fg(0).TextMatrix(A, COLUMNACOSTOPRIMO_)) * INDICEFAB_
        RSTCABECERA.Update
        
        fg(0).TextMatrix(A, COLUMNACOSTOTOTAL_) = Format(NulosN(fg(0).TextMatrix(A, COLUMNACOSTOMP_)) + NulosN(fg(0).TextMatrix(A, COLUMNACOSTOMOBRA_)) + NulosN(fg(0).TextMatrix(A, COLUMNACOSTOFABRICA_)), FORMAT_IMPORTEKARDEX)
        fg(0).TextMatrix(A, COLUMNACOSTOUNIPRODUCCION_) = Format(fg(0).TextMatrix(A, COLUMNACOSTOTOTAL_) / NulosN(fg(0).TextMatrix(A, COLUMNACANTIDAD_)), FORMAT_IMPORTEKARDEX)
    Next A

SALIR_:
    pExportar
    FraProgreso.Visible = False
    Agregando = False
    BANDERA_ = False
End Sub

Private Sub pLlenarDatos(MESATRABAJAR_ As Integer)
    Dim xRs As New ADODB.Recordset
    Dim xRsAux As New ADODB.Recordset
    Dim IDITEM_ As Integer
    Dim IDPROD_ As Integer
    Dim FECHA_ As String
    Dim VALOR_ As Double ' unid/hora de cada producto
    Dim TOTALHORAS_ As Double ' Tiempo en horas de cada producto
    Dim ULTIMODIAMES_ As Date
    Dim PRIMERDIAMES_ As Date
    Dim ANIOACTUAL_ As Double
    Dim MESACTUAL_ As Double
    Dim nSQLId As String
    Dim nSQLIdNot As String
    Dim CONSULTA_ As String
    Dim NUMEROREGISTROS_ As Integer
    Dim PROCESO_ As Integer
    Dim XRSPATRON_ As New ADODB.Recordset
    
    ANIOACTUAL_ = AnoTra
    MESACTUAL_ = MESATRABAJAR_
    ' Se encuentra el primer dia del mes actual
    PRIMERDIAMES_ = CDate("01/" & MESATRABAJAR_ & "/" & ANIOACTUAL_ & "")
    ' Se encuentra el ultimo dia del mes actual
    If MESACTUAL_ = 12 Then MESACTUAL_ = 0: ANIOACTUAL_ = ANIOACTUAL_ + 1
    ULTIMODIAMES_ = CDate("01/" & MESACTUAL_ + 1 & "/" & ANIOACTUAL_ & "") - 1
    ' Si es que haya habido algun cambio se regresan a su estado inicial
    ANIOACTUAL_ = AnoTra
    MESACTUAL_ = MESATRABAJAR_
        
    cSQL = "SELECT pro_recetains.iditem " _
        + vbCr + "FROM pro_recetains INNER JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id " _
        + vbCr + "WHERE (((pro_recetains.canpro) <> 0) And ((alm_inventario.tippro) = 1)) " _
        + vbCr + "GROUP BY pro_recetains.iditem;"
    
    Set XRSPATRON_ = Nothing
    RST_Busq XRSPATRON_, cSQL, xCon
      
    If XRSPATRON_.State = 0 Then GoTo SALIR_
    If XRSPATRON_.RecordCount = 0 Then GoTo SALIR_
    
    ' INICIALIZAMOS PROCESO Y NUMERO DE REGISTROS
    PROCESO_ = 0
    NUMEROREGISTROS_ = 1
    llenarDefinirRST 0, False
        
    fg(3).Rows = fg(3).FixedRows
    fg(4).Rows = fg(4).FixedRows
    fg(2).Rows = fg(2).FixedRows
    fg(0).Rows = fg(0).FixedRows
        
    While NUMEROREGISTROS_ > 0
        PROCESO_ = PROCESO_ + 1
        
        nSQLId = GENERAR_SQL_ID_RST(XRSPATRON_, "iditem", " AND pro_recetains.iditem")
        nSQLIdNot = GENERAR_SQL_ID_RST(XRSPATRON_, "iditem", " AND pro_producciondet.iditem", "NOT IN")
        
        ' HALLAMOS PRODUCTOS DEL PROCESO
        cSQL = "SELECT pro_receta.iditem " _
            + vbCr + "FROM pro_receta INNER JOIN pro_recetains ON pro_receta.id = pro_recetains.idrec " _
            + vbCr + "WHERE ((pro_recetains.canpro)<>0) " & nSQLId _
            + vbCr + "GROUP BY pro_receta.iditem;"
        
        Set XRSPATRON_ = Nothing
        RST_Busq XRSPATRON_, cSQL, xCon
        
        If XRSPATRON_.State = 0 Then GoTo SALIR_
        If XRSPATRON_.RecordCount = 0 Then GoTo SALIR_
        nSQLId = GENERAR_SQL_ID_RST(XRSPATRON_, "iditem", " AND pro_producciondet.iditem")
        
        ' BUSCAMOS PRODUCCION DEL PROCESO
        cSQL = "SELECT pro_produccion.id, pro_produccion.dia AS fchdoc, pro_producciondet.numparte, pro_producciondet.iditem, alm_inventario.descripcion AS item, pro_receta.codrec, pro_producciondet.idres AS idresp, pla_empleados.nombre AS desresp, pro_producciondet.cantidad, mae_unidades.abrev, pro_producciondet.horini, pro_producciondet.horfin, IIf([cPREVEN].[preven]<>0,'V','P') AS tipo, cPREVEN.preven " _
            + vbCr + "FROM (pro_produccion INNER JOIN (((((pro_producciondet INNER JOIN pro_emp ON pro_producciondet.idres = pro_emp.id) INNER JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id) INNER JOIN mae_unidades ON pro_producciondet.idunimed = mae_unidades.id) INNER JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id) ON pro_produccion.id = pro_producciondet.idpro) LEFT JOIN " _
            + vbCr + "( " _
            + vbCr + "SELECT vta_ventasdet.iditem, Avg(vta_ventasdet.preuni) AS preven " _
            + vbCr + "FROM vta_ventas INNER JOIN vta_ventasdet ON vta_ventas.id = vta_ventasdet.idvta " _
            + vbCr + "WHERE (((vta_ventas.fchdoc)>=CDate('" & PRIMERDIAMES_ & "') And (vta_ventas.fchdoc)<=CDate('" & ULTIMODIAMES_ & "'))) " _
            + vbCr + "GROUP BY vta_ventasdet.iditem " _
            + vbCr + ") AS cPREVEN ON pro_producciondet.iditem = cPREVEN.iditem " _
            + vbCr + "WHERE (((pro_producciondet.cantidad)>0) AND ((Month([pro_produccion].[dia]))=" & MESATRABAJAR_ & ") AND ((pro_producciondet.estado) In (2,3))) " & nSQLId & nSQLIdNot _
            + vbCr + "ORDER BY pro_produccion.dia, pro_producciondet.iditem;"
        
        Set xRs = Nothing
        RST_Busq xRs, cSQL, xCon
        
        If xRs.State = 0 Then GoTo SALIR_
        If xRs.RecordCount = 0 Then GoTo SALIR_
        
        ' HALLAMOS NUMERO DE REGISTROS
        NUMEROREGISTROS_ = xRs.RecordCount
        
        If xRs.State = 0 Then Exit Sub
        VALOR_ = 0
        TOTALHORAS_ = 0
        
        With fg(0)
            If NUMEROREGISTROS_ = 0 Then Exit Sub
            
            CentrarFrm FraProgreso
            FraProgreso.Visible = True
            lbl(0).Caption = "PROCESO: " & PROCESO_
            PgBar.Min = 0
            PgBar.Max = xRs.RecordCount
            PgBar.Value = 0
            
            Agregando = True
            xRs.MoveFirst
            While Not xRs.EOF
                DoEvents
                If BANDERA_ Then GoTo SALIR_
                If NUMEROREGISTROS_ = 0 Then GoTo SALIR_
                IDITEM_ = NulosN(xRs("iditem"))
                
                .Rows = .Rows + 1
                .TopRow = .Rows - 1
                FraProgreso.Refresh
                LblProg.Caption = NulosC(xRs("item"))
                PgBar.Value = PgBar.Value + 1
                
                IDPROD_ = NulosN(xRs("id"))
                FECHA_ = NulosC(xRs("fchdoc"))
                
                .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNAFECHA_) = Format(NulosC(xRs("fchdoc")), FORMAT_DATE)
                .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNAREGPROD_) = NulosC(xRs("numparte"))
                .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNATIPO_) = NulosC(xRs("tipo"))
                .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNAPROCESO_) = PROCESO_
                .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNAITEM_) = NulosC(xRs("item"))
                .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNARECETA_) = NulosC(xRs("codrec"))
                .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNARESPONSABLE_) = NulosC(xRs("desresp"))
                .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNAUNIMED_) = NulosC(xRs("abrev"))
                .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNACANTIDAD_) = Format(NulosN(xRs("cantidad")), "0.0000")
                .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNAHORINI_) = Format(NulosC(xRs("horini")), FORMAT_HORA_SIN_SEGUNDO)
                .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNAHORFIN_) = Format(NulosC(xRs("horfin")), FORMAT_HORA_SIN_SEGUNDO)
                .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNAIDPROD_) = IDPROD_
                .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNAIDITEM_) = IDITEM_
                .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNACORRELATIVO_) = CORRELATIVO_
                
                ' SE AGREGA CABECERA
                RSTCABECERA.AddNew
                RSTCABECERA("id") = CORRELATIVO_
                RSTCABECERA("iditem") = IDITEM_
                RSTCABECERA("idprod") = IDPROD_
                RSTCABECERA("proceso") = NulosN(.TextMatrix(.Rows - 1, COLUMNAPROCESO_))
                RSTCABECERA("cantidad") = NulosN(.TextMatrix(.Rows - 1, COLUMNACANTIDAD_))
                RSTCABECERA.Update
                
'                ' SE AGREGA DETALLE DE INSUMOS
'                cSQL = "SELECT pro_producciondetins.iditem AS idins, pro_producciondetins.canutil AS cantidad " _
'                    + vbCr + "FROM (pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro) INNER JOIN pro_producciondetins ON (pro_producciondet.idrec = pro_producciondetins.idrec) AND (pro_producciondet.numparte = pro_producciondetins.numparte) AND (pro_producciondet.idpro = pro_producciondetins.idpro) " _
'                    + vbCr + "WHERE (((pro_producciondetins.canutil)>0) AND ((pro_produccion.id)=" & IDPROD_ & ") AND ((pro_producciondet.iditem)=" & IDITEM_ & "));"
'
'                Set xRsAux = Nothing
'                RST_Busq xRsAux, cSQL, xCon
'                If xRsAux.State = 0 Then GoTo SALIR_
'                If xRsAux.RecordCount = 0 Then GoTo SALIR_
'
'                xRsAux.MoveFirst
'                While Not xRsAux.EOF
'                    ' ---SE AGREGA AL LIBRO DE COSTO
'                    RSTDETALLEMATPRI.AddNew
'                    RSTDETALLEMATPRI("idlibrodet") = CORRELATIVO_
'                    RSTDETALLEMATPRI("iditem") = NulosN(xRsAux("idins"))
'                    RSTDETALLEMATPRI("cantidad") = NulosN(xRsAux("cantidad"))
'                    RSTDETALLEMATPRI.Update
'
'                    xRsAux.MoveNext
'                Wend
                
                CORRELATIVO_ = CORRELATIVO_ + 1
                xRs.MoveNext
            Wend
            xRs.Filter = adFilterNone
        End With
    Wend
    
SALIR_:
    FraProgreso.Visible = False
    Agregando = False
    BANDERA_ = False
End Sub

Sub llenarDefinirRST(IDLIBRO_ As Integer, Optional CARGAR_ As Boolean = True)
    ' TIPO_:0=COSTO MP, TIPO_:1=COSTO MANO OBRA, TIPO_:2=REPORTE DE ERRORES
    Dim xFun As New eps_librerias.FuncionesData
    Dim xRs As New ADODB.Recordset
    Dim nSQLId As String
    Dim xCampos() As String
    ' N: Numerico
    ' D: Double
    ' F: Fecha
    ' C: Caracter
    ' L: Logico
    
    ' SE DEFINE EL RECORDSET AUXILIAR DE INSUMOS
    ReDim xCampos(5, 3) As String
    
    xCampos(0, 0) = "iditem":           xCampos(0, 1) = "N":      xCampos(0, 2) = ""
    xCampos(1, 0) = "fecha":            xCampos(1, 1) = "F":      xCampos(1, 2) = ""
    xCampos(2, 0) = "impprimo":         xCampos(2, 1) = "D":      xCampos(2, 2) = ""
    xCampos(3, 0) = "cantidad":         xCampos(3, 1) = "D":      xCampos(3, 2) = ""
    xCampos(4, 0) = "idprod":           xCampos(4, 1) = "N":      xCampos(4, 2) = ""
    
    Set RECORDSETPREUNI_ = Nothing
    Set RECORDSETPREUNI_ = xFun.CrearRstTMP(xCampos)
    RECORDSETPREUNI_.Open
    
     ' SE DEFINE EL RECORDSET DE ERRORES
    ReDim xCampos(6, 3) As String
    
    xCampos(0, 0) = "numdoc":           xCampos(0, 1) = "C":      xCampos(0, 2) = "20"
    xCampos(1, 0) = "item":             xCampos(1, 1) = "C":      xCampos(1, 2) = "60"
    xCampos(2, 0) = "preuni":           xCampos(2, 1) = "D":      xCampos(2, 2) = ""
    xCampos(3, 0) = "detalle":          xCampos(3, 1) = "C":      xCampos(3, 2) = "40"
    xCampos(4, 0) = "fecha":            xCampos(4, 1) = "F":      xCampos(4, 2) = ""
    xCampos(5, 0) = "insumo":           xCampos(5, 1) = "C":      xCampos(5, 2) = "60"
    
    Set RECORDSETERRORES_ = Nothing
    Set RECORDSETERRORES_ = xFun.CrearRstTMP(xCampos)
    RECORDSETERRORES_.Open
                           
    ' SE DEFINE EL RECORDSET CTAS GASTOS DE FABRICA
    cSQL = "SELECT * FROM con_librocostocta WHERE ((con_librocostocta.idlibro)=" & IDLIBRO_ & ")"
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    If xRs.State = 0 Then Exit Sub
    
    Set RSTCTAGASFAB = Nothing
    DEFINIR_RST_TMP RSTCTAGASFAB, xRs
    If CARGAR_ Then CARGAR_RST_TMP RSTCTAGASFAB, xRs
                     
    ' SE DEFINE EL RECORDSET CABECERA
    cSQL = "SELECT * FROM con_librocostodet WHERE ((con_librocostodet.idlibro)=" & IDLIBRO_ & ")"
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    If xRs.State = 0 Then Exit Sub
    
    Set RSTCABECERA = Nothing
    DEFINIR_RST_TMP RSTCABECERA, xRs
    If CARGAR_ Then CARGAR_RST_TMP RSTCABECERA, xRs
    
    nSQLId = GENERAR_SQL_ID_RST(xRs, "id", "idlibrodet")
    If nSQLId = "" Then nSQLId = "idlibrodet=0"
    ' ---------------------DETALLE MATERIA PRIMA
    cSQL = "SELECT * FROM con_librocostomatpri WHERE " & nSQLId
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    Set RSTDETALLEMATPRI = Nothing
    DEFINIR_RST_TMP RSTDETALLEMATPRI, xRs
    If CARGAR_ Then CARGAR_RST_TMP RSTDETALLEMATPRI, xRs
    
    ' --------------------DETALLE MANO DE OBRA
    cSQL = "SELECT * FROM con_librocostomanobr WHERE " & nSQLId
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    Set RSTDETALLEMANOBR = Nothing
    DEFINIR_RST_TMP RSTDETALLEMANOBR, xRs
    If CARGAR_ Then CARGAR_RST_TMP RSTDETALLEMANOBR, xRs
    
    ' --------------------DETALLE GASTOS DE FABRICA
    cSQL = "SELECT * FROM con_librocostogasfab WHERE " & nSQLId
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    Set RSTDETALLEGASFAB = Nothing
    DEFINIR_RST_TMP RSTDETALLEGASFAB, xRs
    If CARGAR_ Then CARGAR_RST_TMP RSTDETALLEGASFAB, xRs
End Sub

Function hallarConsulta(IDITEM_ As Integer, FCHINI_ As Date, FCHFIN_ As Date) As String
    Dim xCadSQL As String
    Dim xSQLFiltroPS As String '--Util para aplicar un filtro adicional que mostrará solo materia prima en sentencia de "produccion insumos salida"

    If NulosN(AnoTra) >= 2012 Then
        '--Aplicar filtro en produccion de salida para mostrar solo materia prima del 2012 en adelante
        xSQLFiltroPS = " AND alm_inventario.tippro=3  "
    End If

    ' PREPARAMOS LA SELECT PARA ARMAR EL KARDEX
    xCadSQL = "SELECT alm_ingreso.id, alm_ingresodet.iditem, alm_inventario.descripcion, alm_ingreso.fching AS fchdoc, [alm_ingreso]![numser]+'-'+[alm_ingreso]![numdoc] AS numdoc, alm_ingresodet.cantidad AS canpro, alm_ingresodet.preuni, mae_documento.abrev AS descdoc, 'AI' AS tipo, alm_ingreso.nombre AS entidad, 0 AS aa, (SELECT Count(1) AS numdocumentos FROM alm_ingresodoc WHERE (((alm_ingresodoc.id)=alm_ingreso.id))) AS numdocumentos, 'Almacén' & IIf(CStr(numdocumentos)<>'0',' - Compras','') AS modulo, '' AS registro, '' AS ctanum, '' AS ctanom, alm_ingresodet.hora AS horini, alm_ingresodet.hora AS horfin " _
        + vbCr + " FROM (alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id) LEFT JOIN (alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) ON alm_ingreso.id = alm_ingresodet.id " _
        + vbCr + " WHERE (((alm_ingresodet.iditem)=" & IDITEM_ & ") AND ((alm_ingreso.fching)>=CDate('" & FCHINI_ & "') And (alm_ingreso.fching)<=CDate('" & FCHFIN_ & "')) AND ((alm_ingreso.tipmov)=-1)) AND alm_ingreso.estado Not In (1,4,5) AND alm_ingresodet.cantidad<>0 " _
        + vbCr + " UNION ALL " _
        + vbCr + " SELECT alm_ingreso.id, alm_ingresodet.iditem, alm_inventario.descripcion, alm_ingreso.fching AS fchdoc, [alm_ingreso]![numser]+'-'+[alm_ingreso]![numdoc] AS numdoc, alm_ingresodet.cantidad AS canpro, alm_ingresodet.preuni, mae_documento.abrev AS descdoc, 'AS' AS tipo, alm_ingreso.nombre AS entidad, 0 AS aa, (SELECT Count(1) AS numdocumentos FROM alm_ingresodoc WHERE (((alm_ingresodoc.id)=alm_ingreso.id))) AS numdocumentos, 'Almacén' & IIf(CStr(numdocumentos)<>'0',' - Compras','') AS modulo, '' AS registro, '' AS ctanum, '' AS ctanom, alm_ingresodet.hora AS horini, alm_ingresodet.hora AS horfin  " _
        + vbCr + " FROM (alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id) LEFT JOIN (alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) ON alm_ingreso.id = alm_ingresodet.id  " _
        + vbCr + " WHERE (((alm_ingresodet.iditem)=" & IDITEM_ & ") AND ((alm_ingreso.fching)>=CDate('" & FCHINI_ & "') And (alm_ingreso.fching)<=CDate('" & FCHFIN_ & "')) AND ((alm_ingreso.tipmov)=0)) AND alm_ingreso.estado Not In (1,4,5) AND alm_ingresodet.cantidad<>0 " _
        + vbCr + " UNION ALL " _
        + vbCr + " SELECT com_compras.id, com_comprasdet.iditem, alm_inventario.descripcion, com_compras.fchdoc, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, com_comprasdet.canpro, IIf([com_compras]![idmon]=2,[com_comprasdet]![preuni]*[con_tc]![impcom],[com_comprasdet]![preuni]) AS preuni, mae_documento.abrev AS descdoc, 'C' AS Tipo, mae_prov.nombre AS entidad, 0 AS aa, 0 AS numdocumentos, 'Compras' AS modulo, com_compras.numreg AS registro, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctanom, '' AS horini, '' AS horfin " _
        + vbCr + " FROM (alm_inventario RIGHT JOIN (mae_prov LEFT JOIN ((mae_documento RIGHT JOIN (com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) LEFT JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) ON mae_prov.id = com_compras.idpro) ON alm_inventario.id = com_comprasdet.iditem) LEFT JOIN con_planctas ON alm_inventario.idcuenta = con_planctas.id " _
        + vbCr + " WHERE (((com_comprasdet.iditem)=" & IDITEM_ & ") AND ((com_compras.fchdoc)>=CDate('" & FCHINI_ & "') And (com_compras.fchdoc)<=CDate('" & FCHFIN_ & "')) AND ((com_compras.tipcom)=1))"

    xCadSQL = xCadSQL _
        + vbCr + "  UNION ALL" _
        + vbCr + " SELECT vta_guia.id, vta_guiadet.iditem, alm_inventario.descripcion, vta_guia.fecgiro, [vta_guia]![numser]+'-'+[vta_guia]![numdoc] AS numdoc, vta_guiadet.canpro, 0 AS preuni, mae_documento.abrev AS desdoc, 'GR' AS tipo, mae_cliente.nombre AS entidad, 0 AS aa, IIf([vta_guia]![iddocven]<>0,1,0) AS numdocumentos, 'Guia de Remisión' AS modulo, '' AS registro, '' AS ctanum, '' AS ctanom, '' AS horini, '' AS horfin " _
        + vbCr + " FROM ((mae_cliente RIGHT JOIN vta_guia ON mae_cliente.id = vta_guia.idcli) LEFT JOIN mae_documento ON vta_guia.tipdoc = mae_documento.id) LEFT JOIN (vta_guiadet LEFT JOIN alm_inventario ON vta_guiadet.iditem = alm_inventario.id) ON vta_guia.id = vta_guiadet.idgui " _
        + vbCr + " WHERE (((vta_guiadet.iditem)=" & IDITEM_ & ") AND ((vta_guia.fecgiro)>=CDate('" & FCHINI_ & "') And (vta_guia.fecgiro)<=CDate('" & FCHFIN_ & "'))) " _
        + vbCr + " UNION ALL " _
        + vbCr + " SELECT pro_produccion.id, pro_producciondetins.iditem, alm_inventario.descripcion, pro_produccion.dia, pro_producciondetins.numparte, pro_producciondetins.canutil, 0 AS preuni, 'SM' AS desdoc, 'PS' AS tipo, alm_inventario_1.descripcion AS entidad, pro_producciondet.iditem AS aa, 0 AS numdocumentos, 'Producción' AS modulo, '' AS registro, '' AS ctanum, '' AS ctanom, pro_producciondet.horini, pro_producciondet.horfin  " _
        + vbCr + " FROM (((pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro) INNER JOIN (pro_producciondetins LEFT JOIN alm_inventario ON pro_producciondetins.iditem = alm_inventario.id) ON (pro_producciondet.idrec = pro_producciondetins.idrec) AND (pro_producciondet.numparte = pro_producciondetins.numparte) AND (pro_producciondet.idpro = pro_producciondetins.idpro)) LEFT JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN alm_inventario AS alm_inventario_1 ON pro_receta.iditem = alm_inventario_1.id " _
        + vbCr + " WHERE (((pro_producciondetins.iditem)=" & IDITEM_ & ") AND ((pro_produccion.dia)>=CDate('" & FCHINI_ & "') And (pro_produccion.dia)<=CDate('" & FCHFIN_ & "'))) AND pro_producciondet.estado Not In (1,4,5) AND pro_producciondetins.canutil<>0 " & xSQLFiltroPS _
        + vbCr + " UNION ALL " _
        + vbCr & " SELECT pro_produccion.id, pro_producciondet.iditem, alm_inventario.descripcion, pro_produccion.dia, pro_producciondet.numparte, pro_producciondet.cantidad, 0 AS preuni, 'PP' AS desdoc, 'P' AS tipo, 'Producción' AS entidad, pro_producciondet.iditem AS aa, 0 AS numdocumentos, 'Producción' AS modulo, '' AS registro, '' AS ctanum, '' AS ctanom, pro_producciondet.horini, pro_producciondet.horfin  " _
        + vbCr & " FROM pro_produccion INNER JOIN (pro_producciondet LEFT JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id) ON pro_produccion.id = pro_producciondet.idpro " _
        + vbCr & " WHERE (((pro_producciondet.iditem)=" & IDITEM_ & ") AND ((pro_produccion.dia)>=CDate('" & FCHINI_ & "') And (pro_produccion.dia)<=CDate('" & FCHFIN_ & "'))) AND pro_producciondet.estado Not In (1,4,5) AND pro_producciondet.cantidad<>0 "

    xCadSQL = xCadSQL + " UNION All " _
        + vbCr + " SELECT vta_ventas.id, vta_ventasdet.iditem, alm_inventario.descripcion, vta_ventas.fchdoc, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, vta_ventasdet.canpro, vta_ventasdet.preuni, mae_documento.abrev AS descdoc, 'V' AS tipo, mae_cliente.nombre AS entidad, 0 AS aa, 0 AS numdocumentos, 'Ventas' AS modulo, vta_ventas.numreg AS registro, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctanom, '' AS horini, '' AS horfin " _
        + vbCr + " FROM ((mae_cliente RIGHT JOIN (vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) ON mae_cliente.id = vta_ventas.idcli) RIGHT JOIN (vta_ventasdet LEFT JOIN alm_inventario ON vta_ventasdet.iditem = alm_inventario.id) ON vta_ventas.id = vta_ventasdet.idvta) LEFT JOIN con_planctas ON alm_inventario.idcuenta = con_planctas.id " _
        + vbCr + " WHERE (((vta_ventasdet.iditem)=" & IDITEM_ & ") AND ((vta_ventas.fchdoc)>=CDate('" & FCHINI_ & "') And (vta_ventas.fchdoc)<=CDate('" & FCHFIN_ & "')) AND ((vta_ventas.oriitem)=1 Or (vta_ventas.oriitem)=3) AND ((vta_ventas.iddocref) Is Null Or (vta_ventas.iddocref)=0) )" _
        + vbCr + " UNION All " _
        + vbCr + " SELECT vta_ventas.id, vta_ventasdet.iditem, alm_inventario.descripcion, vta_ventas.fchdoc, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, vta_ventasdet.canpro, vta_ventasdet.preuni, mae_documento.abrev AS descdoc, 'V' AS tipo, mae_cliente.nombre AS entidad, 0 AS aa, 0 AS numdocumentos, 'Ventas NC' AS modulo, vta_ventas.numreg AS registro, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctanom, '' AS horini, '' AS horfin " _
        + vbCr + " FROM ((mae_cliente RIGHT JOIN (vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) ON mae_cliente.id = vta_ventas.idcli) RIGHT JOIN (vta_ventasdet LEFT JOIN alm_inventario ON vta_ventasdet.iditem = alm_inventario.id) ON vta_ventas.id = vta_ventasdet.idvta) LEFT JOIN con_planctas ON alm_inventario.idcuenta = con_planctas.id " _
        + vbCr + " WHERE (((vta_ventasdet.iditem)=" & IDITEM_ & ") AND ((vta_ventas.fchdoc)>=CDate('" & FCHINI_ & "') And (vta_ventas.fchdoc)<=CDate('" & FCHFIN_ & "')) AND ((vta_ventas.oriitem)=1 Or (vta_ventas.oriitem)=3) AND ((vta_ventas.iddocref)<>0) AND ((vta_ventas.idmotnotcre)=4))"

    hallarConsulta = xCadSQL
End Function

Private Function GENERAR_SQL_ID_RST(Rst As ADODB.Recordset, nDesc As String, _
                            nCampo As String, Optional nTipoIn As String = "IN", _
                            Optional fEsNumero As Boolean = True) As String
    Dim nSQL As String
    Dim k&
    nSQL = ""
    
    If Rst.State = 0 Then Exit Function
    If Rst.RecordCount = 0 Then Exit Function Else Rst.MoveFirst
    While Not Rst.EOF
        If Trim(CStr(Rst("" & nDesc & ""))) <> "" Then
            If fEsNumero = True Then
                nSQL = nSQL & NulosN(Rst("" & nDesc & "")) & ","
            Else
                nSQL = nSQL & "'" & NulosC(Rst("" & nDesc & "")) & "',"
            End If
        End If
        Rst.MoveNext
    Wend
    
    If nSQL <> "" Then nSQL = " " & nCampo & " " & nTipoIn & " (" + Left(nSQL, Len(nSQL) - 1) & ") "
        
    GENERAR_SQL_ID_RST = nSQL
End Function

Private Function PrecioUni(IdDocumento, IdItem As Double, DondeBuscar As String) As Double
    Dim xRst As New ADODB.Recordset
    Dim nSQL As String
    
    If DondeBuscar = "AI" Then
        nSQL = "SELECT Avg(com_comprasdet.preuni) AS preuniprom " _
            + vbCr + " FROM com_comprasdet INNER JOIN alm_ingresodoc ON com_comprasdet.idcom = alm_ingresodoc.iddoc " _
            + vbCr + " GROUP BY alm_ingresodoc.id, com_comprasdet.iditem " _
            + vbCr + " HAVING (((alm_ingresodoc.id)=" & IdDocumento & ") AND ((com_comprasdet.iditem)=" & IdItem & "))"

    ElseIf DondeBuscar = "GR" Then
        nSQL = "SELECT vta_guia.id, vta_ventasdet.iditem, Avg(vta_ventasdet.preuni) AS preuniprom " _
            + vbCr + " FROM vta_guia INNER JOIN vta_ventasdet ON vta_guia.iddocven = vta_ventasdet.idvta " _
            + vbCr + " GROUP BY vta_guia.id, vta_ventasdet.iditem " _
            + vbCr + " HAVING (((vta_guia.id)=" & IdDocumento & ") AND ((vta_ventasdet.iditem)=" & IdItem & ")); "
       
    Else
        PrecioUni = 0
        Exit Function
    End If
    
    RST_Busq xRst, nSQL, xCon
    
    If xRst.RecordCount <> 0 Then
        PrecioUni = NulosN(xRst("preuniprom"))
    Else
        PrecioUni = 0
    End If
    
    Set xRst = Nothing
    
End Function

Private Sub pMostrarGasFab(IDLIBRO_ As Integer)
    Dim ULTIMODIAMES_ As Date
    Dim PRIMERDIAMES_ As Date
    Dim ANIOACTUAL_ As Double
    Dim MESACTUAL_ As Double
    Dim IMPGASFAB_ As Double
    Dim xRs As New ADODB.Recordset
    Dim nSQLId As String
            
    fg(1).Rows = fg(1).FixedRows
    CentrarFrm Frm4
    Frm4.Visible = True
    
    If RSTCTAGASFAB.State = 0 Then Exit Sub
    RSTCTAGASFAB.Filter = "idlibro=" & IDLIBRO_
    If RSTCTAGASFAB.RecordCount = 0 Then Exit Sub
    
    With fg(1)
        RSTCTAGASFAB.MoveFirst
        While Not RSTCTAGASFAB.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = Busca_Codigo(NulosN(RSTCTAGASFAB("idcuenta")), "id", "cuenta", "con_planctas", "N", xCon)
            .TextMatrix(.Rows - 1, 2) = Busca_Codigo(NulosN(RSTCTAGASFAB("idcuenta")), "id", "descripcion", "con_planctas", "N", xCon)
            .TextMatrix(.Rows - 1, 3) = Format(NulosN(RSTCTAGASFAB("importe")), FORMAT_IMPORTEKARDEX)
            .TextMatrix(.Rows - 1, 4) = NulosN(RSTCTAGASFAB("idcuenta"))
            RSTCTAGASFAB.MoveNext
        Wend
    End With
    
    lblTotalGr.Caption = Format(GRID_SUMAR_COL(fg(1), 3), FORMAT_MONTO)
End Sub

Private Sub llenarDetallePersonal(CORRELATIVO_ As Integer)
    Dim RECORDSET_ As New ADODB.Recordset
    Dim TOTALPRODUCCION_ As Double
    Dim TOTALPLANILLA_ As Double
    Dim FECHA_ As String
    
    fg(4).Rows = fg(4).FixedRows
    RSTDETALLEMANOBR.Filter = adFilterNone
    If RSTDETALLEMANOBR.State = 0 Then Exit Sub
    RSTDETALLEMANOBR.Filter = "idlibrodet=" & CORRELATIVO_
    If RSTDETALLEMANOBR.RecordCount = 0 Then Exit Sub
    With fg(4)
        RSTDETALLEMANOBR.MoveFirst
        Agregando = True
        While Not RSTDETALLEMANOBR.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = Busca_Codigo(NulosN(RSTDETALLEMANOBR("idemp")), "id", "numdoc", "pla_empleados", "N", xCon)
            .TextMatrix(.Rows - 1, 2) = Busca_Codigo(NulosN(RSTDETALLEMANOBR("idemp")), "id", "nombre", "pla_empleados", "N", xCon)
            .TextMatrix(.Rows - 1, 4) = Format(NulosN(RSTDETALLEMANOBR("impmanobr")), FORMAT_IMPORTEKARDEX)
            .TextMatrix(.Rows - 1, 5) = NulosN(RSTDETALLEMANOBR("idemp"))
            .TextMatrix(.Rows - 1, 6) = Busca_Codigo(NulosN(RSTDETALLEMANOBR("idemp")), "id", "idarea", "pla_empleados", "N", xCon)
            .TextMatrix(.Rows - 1, 3) = Busca_Codigo(NulosN(.TextMatrix(.Rows - 1, 6)), "id", "descripcion", "mae_area", "N", xCon)
            
            RSTDETALLEMANOBR.MoveNext
        Wend
        
        .Rows = .Rows + 1
        .Select .Rows - 1, 3
        .CellFontBold = True
        .TextMatrix(.Rows - 1, 3) = "TOTAL"
        .TextMatrix(.Rows - 1, 4) = Format(GRID_SUMAR_COL(fg(4), 4), "0.0000")
        .TopRow = .Rows - 1
        Agregando = False
    End With
End Sub

Private Sub llenarDetalleGasFab(PARCIALGASFAB_ As Double)
    Dim RECORDSET_ As New ADODB.Recordset
    Dim TOTALGASFAB_ As Double
    Dim INDICE_ As Double
    Dim FECHA_ As String
            
    fg(2).Rows = fg(2).FixedRows
    If opttipdiscta(0).Value = True Then
        RSTCTAGASFAB.Filter = adFilterNone
        TOTALGASFAB_ = RST_SUMAR(RSTCTAGASFAB, "importe")
        If TOTALGASFAB_ = 0 Then
            INDICE_ = 0
        Else
            INDICE_ = PARCIALGASFAB_ / TOTALGASFAB_
        End If
        
        RSTCTAGASFAB.Filter = adFilterNone
        If RSTCTAGASFAB.State = 0 Then Exit Sub
        If RSTCTAGASFAB.RecordCount = 0 Then Exit Sub
        With fg(2)
            RSTCTAGASFAB.MoveFirst
            Agregando = True
            While Not RSTCTAGASFAB.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = Busca_Codigo(NulosN(RSTCTAGASFAB("idcuenta")), "id", "cuenta", "con_planctas", "N", xCon)
                .TextMatrix(.Rows - 1, 2) = Busca_Codigo(NulosN(RSTCTAGASFAB("idcuenta")), "id", "descripcion", "con_planctas", "N", xCon)
                .TextMatrix(.Rows - 1, 3) = Format(INDICE_ * NulosN(RSTCTAGASFAB("importe")), FORMAT_IMPORTEKARDEX)
                .TextMatrix(.Rows - 1, 4) = NulosN(RSTCTAGASFAB("idcuenta"))
                RSTCTAGASFAB.MoveNext
            Wend
            
            .Rows = .Rows + 1
            .Select .Rows - 1, 2
            .CellFontBold = True
            .TextMatrix(.Rows - 1, 2) = "TOTAL"
            .TextMatrix(.Rows - 1, 3) = Format(GRID_SUMAR_COL(fg(2), 3), "0.0000")
            .TopRow = .Rows - 1
            Agregando = False
        End With
    End If
End Sub

Private Sub llenarDetalleInsumos(IDLIBRODET_ As Integer)
    Dim IDDOCUMENTO_ As Integer
    Dim IDITEM_ As Integer
    Dim FECHA_ As String
    
    If Agregando Then Exit Sub
    RSTDETALLEMATPRI.Filter = adFilterNone
    With fg(3)
        .Rows = .FixedRows
        If RSTDETALLEMATPRI.State = 0 Then Me.MousePointer = vbDefault: Exit Sub
        RSTDETALLEMATPRI.Filter = "idlibrodet=" & IDLIBRODET_
        If RSTDETALLEMATPRI.RecordCount = 0 Then Me.MousePointer = vbDefault: Exit Sub
        RSTDETALLEMATPRI.MoveFirst
        While Not RSTDETALLEMATPRI.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 6) = NulosN(RSTDETALLEMATPRI("iditem"))
            .TextMatrix(.Rows - 1, 2) = UCase(Busca_Codigo(NulosN(RSTDETALLEMATPRI("iditem")), "id", "descripcion", "alm_inventario", "N", xCon))
            .TextMatrix(.Rows - 1, 7) = Busca_Codigo(NulosN(RSTDETALLEMATPRI("iditem")), "id", "tippro", "alm_inventario", "N", xCon)
            .TextMatrix(.Rows - 1, 1) = UCase(Busca_Codigo(NulosN(.TextMatrix(.Rows - 1, 7)), "id", "descripcion", "mae_tipoproducto", "N", xCon))
            .TextMatrix(.Rows - 1, 3) = Format(NulosN(RSTDETALLEMATPRI("cantidad")), FORMAT_CANTIDAD)
            .TextMatrix(.Rows - 1, 5) = Format(NulosN(RSTDETALLEMATPRI("impmatpri")), FORMAT_IMPORTEKARDEX)
            .TextMatrix(.Rows - 1, 4) = Format(NulosN(RSTDETALLEMATPRI("impmatpri")) / NulosN(RSTDETALLEMATPRI("cantidad")), FORMAT_IMPORTEKARDEX)
SIGUIENTE_:
            RSTDETALLEMATPRI.MoveNext
        Wend
        .Rows = .Rows + 1
        .Select .Rows - 1, 2
        .CellFontBold = True
        .TextMatrix(.Rows - 1, 2) = "TOTAL"
        .TextMatrix(.Rows - 1, 5) = GRID_SUMAR_COL(fg(3), 5)
        
    End With
End Sub

Private Function pImporteManoObra(IDITEM_ As Integer, FECHA_ As String, XCON_ As ADODB.Connection, _
                                        IDDOCUMENTO_ As Integer, Optional CORRELATIVO_ As Integer = 0)
    
    Dim RECORDSET_ As New ADODB.Recordset
    Dim IMPORTEMANOOBRA_ As Double
    Dim DURACPRODUCCION_ As Double
    Dim DURHORASARREGLO() As String
    Dim TOTALPLANILLA_ As Double
    Dim TOTALHORASPRODUCCION_ As Double
    Dim DURHORASNUMERICO_ As Double
    Dim COSTOPROMHORA_ As Double
    '-----------------------------------------
    ' -----------------------COSTO DE PLANILLA
    '-----------------------------------------
    ' ---------------DURACION DE LA PRODUCCION
    cSQL = "SELECT CDate([pro_producciondet].[horfin]-[pro_producciondet].[horini]) AS dur " _
        + vbCr + "FROM pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro " _
        + vbCr + "WHERE (((pro_producciondet.iditem)=" & IDITEM_ & ") AND ((pro_produccion.id)=" & IDDOCUMENTO_ & "));"
        
    Set RECORDSET_ = Nothing
    RST_Busq RECORDSET_, cSQL, XCON_
    
    If RECORDSET_.State = 0 Then pImporteManoObra = 0: Exit Function
    If RECORDSET_.RecordCount = 0 Then pImporteManoObra = 0: Exit Function
    DURHORASARREGLO = Split(Format(RECORDSET_("dur"), "HH:mm"), ":")
    DURACPRODUCCION_ = NulosN(DURHORASARREGLO(0)) + (NulosN(DURHORASARREGLO(1)) / 60)
    
    ' ---------------TOTAL PLANILLA DEL DIA
    cSQL = "SELECT Sum(pro_pagos.imptot) AS montotot " _
        + vbCr + "FROM pro_pagos " _
        + vbCr + "WHERE (((pro_pagos.fchtra)=CDate('" & FECHA_ & "')) AND ((pro_pagos.idarea) In (3,4,8,23)));"
    
    Set RECORDSET_ = Nothing
    RST_Busq RECORDSET_, cSQL, XCON_
    
    If RECORDSET_.State = 0 Then pImporteManoObra = 0: Exit Function
    If RECORDSET_.RecordCount = 0 Then pImporteManoObra = 0: Exit Function
    TOTALPLANILLA_ = NulosN(RECORDSET_("montotot"))
    
    ' ---------------TOTAL HORAS DE PRODUCCION DEL DIA
    cSQL = "SELECT pro_producciondet.iditem, CDate([pro_producciondet].[horfin]-[pro_producciondet].[horini]) AS dur " _
        + vbCr + "FROM pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro " _
        + vbCr + "WHERE (((pro_produccion.dia)=CDate('" & FECHA_ & "')));"
    
    Set RECORDSET_ = Nothing
    RST_Busq RECORDSET_, cSQL, XCON_
    
    If RECORDSET_.State = 0 Then pImporteManoObra = 0: Exit Function
    If RECORDSET_.RecordCount = 0 Then pImporteManoObra = 0: Exit Function
    RECORDSET_.MoveFirst
    While Not RECORDSET_.EOF
        DURHORASARREGLO = Split(Format(RECORDSET_("dur"), "HH:mm"), ":")
        DURHORASNUMERICO_ = NulosN(DURHORASARREGLO(0)) + (NulosN(DURHORASARREGLO(1)) / 60)
        TOTALHORASPRODUCCION_ = TOTALHORASPRODUCCION_ + DURHORASNUMERICO_
        RECORDSET_.MoveNext
    Wend
      
    If CORRELATIVO_ <> 0 Then
        ' -------------PERSONAS INVOLUCRDAS EN LA RODUCCION
        cSQL = "SELECT Sum(pro_pagos.imptot) AS montotot, pro_pagos.idemp, pla_empleados.nombre, pro_pagos.idarea " _
            + vbCr + "FROM pro_pagos INNER JOIN pla_empleados ON pro_pagos.idemp = pla_empleados.id " _
            + vbCr + "WHERE (((pro_pagos.fchtra)=CDate('" & FECHA_ & "')) AND ((pro_pagos.idarea) In (3,4,8,23))) " _
            + vbCr + "GROUP BY pro_pagos.idemp, pla_empleados.nombre, pro_pagos.idarea;"
        
        Set RECORDSET_ = Nothing
        RST_Busq RECORDSET_, cSQL, xCon
        
        If RECORDSET_.State = 0 Then Exit Function
        If RECORDSET_.RecordCount = 0 Then Exit Function
        RECORDSET_.MoveFirst
        While Not RECORDSET_.EOF
            '******************************************************
            '******************************************************
            RSTDETALLEMANOBR.AddNew
            RSTDETALLEMANOBR("idlibrodet") = CORRELATIVO_
            RSTDETALLEMANOBR("idemp") = NulosN(RECORDSET_("idemp"))
            RSTDETALLEMANOBR("impmanobr") = (NulosN(RECORDSET_("montotot")) / TOTALHORASPRODUCCION_) * DURACPRODUCCION_
            RSTDETALLEMANOBR.Update
            '******************************************************
            '******************************************************
            
            RECORDSET_.MoveNext
        Wend
    End If
       
    ' ---------------COSTO PROMEDIO POR HORA
    IMPORTEMANOOBRA_ = (TOTALPLANILLA_ / TOTALHORASPRODUCCION_) * DURACPRODUCCION_
    
    pImporteManoObra = IMPORTEMANOOBRA_
End Function

Private Function pImporteMateriaPrima(IDITEM_ As Integer, CANTIDAD_ As Double, FECHA_ As String, HORINI_ As Date, HORFIN_ As Date, XCON_ As ADODB.Connection, _
                                Optional TIPO_ As Integer = 1, Optional TIPODOCUMENTO_ As String, _
                                Optional IDDOCUMENTO_ As Integer, Optional CORRELATIVO_ As Integer = 0) As Double
    Dim cSQL As String
    Dim PRECIOPROMEDIO_ As Double
    Dim PRECIOUNITARIO_ As Double
    Dim PRECIOMANOOBRA_ As Double
    Dim A As Integer
    Dim STOCKINICIAL_ As Double
    Dim PRECIOINICIAL_ As Double
    Dim TOTALSALIDAS_ As Double
    Dim TOTALENTRADAS_ As Double
    Dim CANTIDADACUMULADA_ As Double
    Dim IMPORTEACUMULADO_ As Double
    Dim TIPOPRODUCTO_ As Integer
    Dim FECHAINICIO_ As String
    Dim RECORDSET_ As New ADODB.Recordset
    
    Dim IMPORTEINSUMO_ As Double
    
    PRECIOPROMEDIO_ = 0
    CANTIDADACUMULADA_ = 0
    IMPORTEACUMULADO_ = 0
    TOTALENTRADAS_ = 0
    TOTALSALIDAS_ = 0
        
    '---------------DETALLE DE MOVIMIENTOS
    RECORDSETPREUNI_.Filter = "iditem=" & IDITEM_ & " And fecha<=" & FECHA_
            
    If RECORDSETPREUNI_.RecordCount = 0 Then
        FECHAINICIO_ = "01/01/" & Year(CDate(FECHA_))
        PRECIOINICIAL_ = NulosN(Busca_Codigo("id", NulosC(IDITEM_), "preini", "alm_inventario", "N", xCon))
        STOCKINICIAL_ = NulosN(Busca_Codigo("id", NulosC(IDITEM_), "stckini", "alm_inventario", "N", xCon))
        
        If STOCKINICIAL_ > 0 And PRECIOINICIAL_ = 0 Then
            MsgBox "Ocurrio un error al procesar el costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            RECORDSETERRORES_.AddNew
            RECORDSETERRORES_("numdoc") = Busca_Codigo(IDDOCUMENTO_, "idpro", "numparte", "pro_producciondet", "N", XCON_)
            RECORDSETERRORES_("item") = Busca_Codigo(IDITEM_, "id", "descripcion", "Alm_inventario", "N", XCON_)
            RECORDSETERRORES_("preuni") = 0
            RECORDSETERRORES_("detalle") = "Costo MP - Precio inicial cero"
            RECORDSETERRORES_("fecha") = FECHAINICIO_
            RECORDSETERRORES_.Update
            BANDERA_ = True
        End If
    Else
        RECORDSETPREUNI_.Sort = "fecha DESC"
        RECORDSETPREUNI_.MoveFirst
        FECHAINICIO_ = RECORDSETPREUNI_("fecha")
        PRECIOINICIAL_ = RECORDSETPREUNI_("impprimo") / RECORDSETPREUNI_("cantidad")
        STOCKINICIAL_ = SaldoActual(CDbl(IDITEM_), "01/01/" & Year(CDate(FECHAINICIO_)), FECHAINICIO_, xCon)
        FECHAINICIO_ = CDate(FECHAINICIO_) + 1
    End If
                              
    cSQL = hallarConsulta(CDbl(IDITEM_), CDate(FECHAINICIO_), CDate(FECHA_))
    
    RST_Busq RECORDSET_, cSQL, xCon
    RECORDSET_.Sort = "fchdoc, Tipo, numdoc"
    
    ' --------------STOCK Y PRECIO INICIAL
    PRECIOPROMEDIO_ = PRECIOINICIAL_
    CANTIDADACUMULADA_ = STOCKINICIAL_
    IMPORTEACUMULADO_ = CANTIDADACUMULADA_ * PRECIOINICIAL_
    TOTALENTRADAS_ = TOTALENTRADAS_ + STOCKINICIAL_
        
    Select Case TIPO_
        Case 0
            ' ----------------------------------------------------------INGRESOS
            If TIPODOCUMENTO_ = "C" Or TIPODOCUMENTO_ = "AI" Or TIPODOCUMENTO_ = "P" Then
                ' -------------------------------------
                ' ----------------------COSTO DE TAREAS
                ' -------------------------------------
                ' ----------------------INSUMOS DE LA PRODUCCION
                cSQL = "SELECT pro_producciondetins.iditem AS idins, pro_producciondetins.canutil AS cantidad " _
                    + vbCr + "FROM (pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro) INNER JOIN pro_producciondetins ON (pro_producciondet.idrec = pro_producciondetins.idrec) AND (pro_producciondet.numparte = pro_producciondetins.numparte) AND (pro_producciondet.idpro = pro_producciondetins.idpro) " _
                    + vbCr + "WHERE (((pro_producciondetins.canutil)>0) AND ((pro_produccion.id)=" & IDDOCUMENTO_ & ") AND ((pro_producciondet.iditem)=" & IDITEM_ & "));"
                
                Set RECORDSET_ = Nothing
                RST_Busq RECORDSET_, cSQL, XCON_
                If RECORDSET_.State = 0 Then pImporteMateriaPrima = 0: Exit Function
                If RECORDSET_.RecordCount = 0 Then pImporteMateriaPrima = 0: Exit Function
                
                RECORDSET_.MoveFirst
                IMPORTEACUMULADO_ = 0
                While Not RECORDSET_.EOF
'                    If NulosN(RECORDSET_("idins")) = 1695 Then
'                        MsgBox "ENTRO"
'                    End If
                    RECORDSETPREUNI_.Filter = "iditem=" & NulosN(RECORDSET_("idins")) & " AND fecha=" & FECHA_
                    If RECORDSETPREUNI_.RecordCount = 0 Then
                        IMPORTEINSUMO_ = pImporteMateriaPrima(RECORDSET_("idins"), NulosN(RECORDSET_("cantidad")), CDate(FECHA_), HORINI_, HORFIN_, XCON_)
                                            
                        If IMPORTEINSUMO_ < 0 Then
                            MsgBox "Ocurrio un error al procesar el costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                            RECORDSETERRORES_.AddNew
                            RECORDSETERRORES_("numdoc") = Busca_Codigo(IDDOCUMENTO_, "idpro", "numparte", "pro_producciondet", "N", XCON_)
                            RECORDSETERRORES_("item") = Busca_Codigo(IDITEM_, "id", "descripcion", "Alm_inventario", "N", XCON_)
                            RECORDSETERRORES_("insumo") = Busca_Codigo(NulosN(RECORDSET_("idins")), "id", "descripcion", "Alm_inventario", "N", XCON_)
                            RECORDSETERRORES_("preuni") = IMPORTEINSUMO_
                            RECORDSETERRORES_("detalle") = "Costo MP - Precio unitario negativo"
                            RECORDSETERRORES_("fecha") = FECHA_
                            RECORDSETERRORES_.Update
                            BANDERA_ = True
                        ElseIf IMPORTEINSUMO_ = 0 Then
                            MsgBox "Ocurrio un error al procesar el costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                            RECORDSETERRORES_.AddNew
                            RECORDSETERRORES_("numdoc") = Busca_Codigo(IDDOCUMENTO_, "idpro", "numparte", "pro_producciondet", "N", XCON_)
                            RECORDSETERRORES_("item") = Busca_Codigo(IDITEM_, "id", "descripcion", "Alm_inventario", "N", XCON_)
                            RECORDSETERRORES_("insumo") = Busca_Codigo(NulosN(RECORDSET_("idins")), "id", "descripcion", "Alm_inventario", "N", XCON_)
                            RECORDSETERRORES_("preuni") = 0
                            RECORDSETERRORES_("detalle") = "Costo MP - Precio unitario cero"
                            RECORDSETERRORES_("fecha") = FECHA_
                            RECORDSETERRORES_.Update
                            BANDERA_ = True
                        End If
                        
                        ' ---SE AGREGA AL RECORDSET AUXILIAR
                        RECORDSETPREUNI_.AddNew
                        RECORDSETPREUNI_("iditem") = RECORDSET_("idins")
                        RECORDSETPREUNI_("fecha") = FECHA_
                        RECORDSETPREUNI_("impprimo") = IMPORTEINSUMO_
                        RECORDSETPREUNI_("cantidad") = NulosN(RECORDSET_("cantidad"))
                        RECORDSETPREUNI_.Update
                        
                        If CORRELATIVO_ <> 0 Then
                            ' ---SE AGREGA AL LIBRO DE COSTO
                            RSTDETALLEMATPRI.AddNew
                            RSTDETALLEMATPRI("idlibrodet") = CORRELATIVO_
                            RSTDETALLEMATPRI("iditem") = NulosN(RECORDSET_("idins"))
                            RSTDETALLEMATPRI("cantidad") = NulosN(RECORDSET_("cantidad"))
                            RSTDETALLEMATPRI("impmatpri") = IMPORTEINSUMO_
                            RSTDETALLEMATPRI.Update
                        End If
                    Else
                        IMPORTEINSUMO_ = (NulosN(RECORDSETPREUNI_("impprimo")) / NulosN(RECORDSETPREUNI_("cantidad"))) * NulosN(RECORDSET_("cantidad"))
                        
                        If CORRELATIVO_ <> 0 Then
                            ' ---SE AGREGA AL LIBRO DE COSTO
                            RSTDETALLEMATPRI.AddNew
                            RSTDETALLEMATPRI("idlibrodet") = CORRELATIVO_
                            RSTDETALLEMATPRI("iditem") = NulosN(RECORDSET_("idins"))
                            RSTDETALLEMATPRI("cantidad") = NulosN(RECORDSET_("cantidad"))
                            RSTDETALLEMATPRI("impmatpri") = IMPORTEINSUMO_
                            RSTDETALLEMATPRI.Update
                        End If
                    End If
                    
                    ' ---SE ACUMULA EL IMPORTE
                    IMPORTEACUMULADO_ = IMPORTEACUMULADO_ + IMPORTEINSUMO_
                    RECORDSET_.MoveNext
                Wend
                
                If IMPORTEACUMULADO_ < 0 Then
                    MsgBox "Ocurrio un error al procesar el costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    RECORDSETERRORES_.AddNew
                    RECORDSETERRORES_("numdoc") = Busca_Codigo(IDDOCUMENTO_, "idpro", "numparte", "pro_producciondet", "N", XCON_)
                    RECORDSETERRORES_("item") = Busca_Codigo(IDITEM_, "id", "descripcion", "Alm_inventario", "N", XCON_)
                    RECORDSETERRORES_("preuni") = IMPORTEACUMULADO_
                    RECORDSETERRORES_("detalle") = "Costo MP - Importe negativas"
                    RECORDSETERRORES_("fecha") = FECHA_
                    RECORDSETERRORES_.Update
                    BANDERA_ = True
                ElseIf IMPORTEACUMULADO_ = 0 Then
                    MsgBox "Ocurrio un error al procesar el costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    RECORDSETERRORES_.AddNew
                    RECORDSETERRORES_("numdoc") = Busca_Codigo(IDDOCUMENTO_, "id", "numparte", "pro_produccion", "N", XCON_)
                    RECORDSETERRORES_("item") = Busca_Codigo(IDITEM_, "id", "dscripcion", "Alm_inventario", "N", XCON_)
                    RECORDSETERRORES_("preuni") = 0
                    RECORDSETERRORES_("detalle") = "Costo MP - Importe acumulado cero"
                    RECORDSETERRORES_("fecha") = FECHA_
                    RECORDSETERRORES_.Update
                    BANDERA_ = True
                End If
                
                ' SE AGREGA AL RECORDSET DE PRECIOS UNITARIOS
'                RECORDSETPREUNI_.AddNew
'                RECORDSETPREUNI_("iditem") = IDITEM_
'                RECORDSETPREUNI_("fecha") = FECHA_
'                RECORDSETPREUNI_("impprimo") = IMPORTEACUMULADO_
'                RECORDSETPREUNI_("cantidad") = CANTIDAD_
'                RECORDSETPREUNI_("idprod") = IDDOCUMENTO_
'                RECORDSETPREUNI_.Update
                
                pImporteMateriaPrima = IMPORTEACUMULADO_
                Exit Function
            ' ----------------------------------------------------------SALIDAS
            Else
            End If
            
        Case 1
            If RECORDSET_.RecordCount = 0 Then pImporteMateriaPrima = PRECIOINICIAL_ * CANTIDAD_: Exit Function
            RECORDSET_.MoveFirst
            While Not RECORDSET_.EOF
                ' HALLAMOS TIPO DE PRODUCTO
                TIPOPRODUCTO_ = Busca_Codigo(NulosN(IDITEM_), "id", "tippro", "alm_inventario", "N", XCON_)
                If TIPOPRODUCTO_ = 3 Then
                    If RECORDSET_("fchdoc") = FECHA_ Then
                        If NulosC(RECORDSET_("horini")) = "" Then
                            GoTo SIGUIENTE_
                        Else
                            If RECORDSET_("horini") >= HORINI_ Then GoTo SIGUIENTE_
                        End If
                    End If
                End If
                
                ' ----------------------------------------------------------INGRESOS
                If RECORDSET_("tipo") = "C" Or RECORDSET_("tipo") = "AI" Or RECORDSET_("tipo") = "P" Then
                    ' --------------------------------SALDO Y TOTALES
                    If RECORDSET_("descdoc") = "NC" Then
                        CANTIDADACUMULADA_ = CANTIDADACUMULADA_ - NulosN(RECORDSET_("canpro"))
                        TOTALSALIDAS_ = TOTALSALIDAS_ + NulosN(RECORDSET_("canpro"))
                    Else
                        CANTIDADACUMULADA_ = CANTIDADACUMULADA_ + NulosN(RECORDSET_("canpro"))
                        TOTALENTRADAS_ = TOTALENTRADAS_ + NulosN(RECORDSET_("canpro"))
                    End If
                    '---------------------------------PRECIO UNITARIO
                    If RECORDSET_("tipo") = "AI" And RECORDSET_("numdocumentos") <> 0 Then
                        IMPORTEINSUMO_ = PrecioUni(RECORDSET_("id"), CDbl(IDITEM_), NulosC(RECORDSET_("tipo"))) * NulosN(RECORDSET_("canpro"))
                        
                        If IMPORTEINSUMO_ < 0 Then
                            MsgBox "Ocurrio un error al procesar el costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                            RECORDSETERRORES_.AddNew
                            RECORDSETERRORES_("numdoc") = NulosC(RECORDSET_("numdoc"))
                            RECORDSETERRORES_("item") = Busca_Codigo(IDITEM_, "id", "descripcion", "Alm_inventario", "N", XCON_)
                            RECORDSETERRORES_("preuni") = PRECIOUNITARIO_
                            RECORDSETERRORES_("detalle") = "Costo MP - Precio unitario negativo"
                            RECORDSETERRORES_("fecha") = RECORDSET_("fchdoc")
                            RECORDSETERRORES_.Update
                            BANDERA_ = True
                        ElseIf IMPORTEINSUMO_ = 0 Then
                            MsgBox "Ocurrio un error al procesar el costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                            RECORDSETERRORES_.AddNew
                            RECORDSETERRORES_("numdoc") = NulosC(RECORDSET_("numdoc"))
                            RECORDSETERRORES_("item") = Busca_Codigo(IDITEM_, "id", "descripcion", "Alm_inventario", "N", XCON_)
                            RECORDSETERRORES_("preuni") = PRECIOUNITARIO_
                            RECORDSETERRORES_("detalle") = "Costo MP - Precio unitario cero"
                            RECORDSETERRORES_("fecha") = RECORDSET_("fchdoc")
                            RECORDSETERRORES_.Update
                            BANDERA_ = True
                        End If
                    Else
                        ' --------------TIPO DE ITEM
                        TIPOPRODUCTO_ = Busca_Codigo(IDITEM_, "id", "tippro", "alm_inventario", "N", XCON_)
                        Select Case TIPOPRODUCTO_
                            Case 3
                                RECORDSETPREUNI_.Filter = "iditem=" & IDITEM_ & " AND idprod=" & NulosN(RECORDSET_("id"))
                                If RECORDSETPREUNI_.RecordCount = 0 Then
                                    ' COSTO DE LA MANO DE LA MATERIA PRIMA
                                    IMPORTEINSUMO_ = pImporteMateriaPrima(IDITEM_, RECORDSET_("canpro"), RECORDSET_("fchdoc"), CDate(RECORDSET_("horini")), CDate(RECORDSET_("horfin")), XCON_, 0, RECORDSET_("tipo"), RECORDSET_("id"))
                                    ' COSTO DE LA MANO DE OBRA
                                    IMPORTEINSUMO_ = IMPORTEINSUMO_ + pImporteManoObra(IDITEM_, RECORDSET_("fchdoc"), xCon, RECORDSET_("id"))
                                    
                                    ' SE AGREGA AL RECORDSET DE PRECIOS UNITARIOS
                                    RECORDSETPREUNI_.AddNew
                                    RECORDSETPREUNI_("iditem") = IDITEM_
                                    RECORDSETPREUNI_("fecha") = RECORDSET_("fchdoc")
                                    RECORDSETPREUNI_("impprimo") = IMPORTEINSUMO_
                                    RECORDSETPREUNI_("cantidad") = RECORDSET_("canpro")
                                    RECORDSETPREUNI_("idprod") = NulosN(RECORDSET_("id"))
                                    RECORDSETPREUNI_.Update
                                Else
                                    IMPORTEINSUMO_ = (NulosN(RECORDSETPREUNI_("impprimo")) / NulosN(RECORDSETPREUNI_("cantidad"))) * RECORDSET_("canpro")
                                End If
                            Case Else
                                RECORDSETPREUNI_.Filter = "iditem=" & IDITEM_ & " AND fecha=" & RECORDSET_("fchdoc")
                                If RECORDSETPREUNI_.RecordCount = 0 Then
                                    IMPORTEINSUMO_ = NulosN(RECORDSET_("preuni")) * NulosN(RECORDSET_("canpro"))
                                    ' SE AGREGA AL RECORDSET DE PRECIOS UNITARIOS
                                    RECORDSETPREUNI_.AddNew
                                    RECORDSETPREUNI_("iditem") = IDITEM_
                                    RECORDSETPREUNI_("fecha") = RECORDSET_("fchdoc")
                                    RECORDSETPREUNI_("impprimo") = IMPORTEINSUMO_
                                    RECORDSETPREUNI_("cantidad") = RECORDSET_("canpro")
                                    RECORDSETPREUNI_.Update
                                Else
                                    IMPORTEINSUMO_ = (NulosN(RECORDSETPREUNI_("impprimo")) / NulosN(RECORDSETPREUNI_("cantidad"))) * RECORDSET_("canpro")
                                End If
                        End Select
                    End If
                    ' --------------------------------IMPORTE ACUMULADO
                    If RECORDSET_("descdoc") = "NC" Then
                        IMPORTEACUMULADO_ = IMPORTEACUMULADO_ - IMPORTEINSUMO_
                    Else
                        IMPORTEACUMULADO_ = IMPORTEACUMULADO_ + IMPORTEINSUMO_
                    End If
                    ' --------------------------------PRECIO PROMEDIO
                    If CANTIDADACUMULADA_ > 0 Then
                        PRECIOPROMEDIO_ = IMPORTEACUMULADO_ / CANTIDADACUMULADA_
                    ElseIf CANTIDADACUMULADA_ < 0 Then
                        MsgBox "Ocurrio un error al procesar el costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                        RECORDSETERRORES_.AddNew
                        RECORDSETERRORES_("numdoc") = NulosC(RECORDSET_("numdoc"))
                        RECORDSETERRORES_("item") = Busca_Codigo(IDITEM_, "id", "descripcion", "Alm_inventario", "N", XCON_)
                        RECORDSETERRORES_("preuni") = CANTIDADACUMULADA_
                        RECORDSETERRORES_("detalle") = "Costo MP - Unidades negativas"
                        RECORDSETERRORES_("fecha") = RECORDSET_("fchdoc")
                        RECORDSETERRORES_.Update
                        BANDERA_ = True
                    ElseIf CANTIDADACUMULADA_ = 0 Then
                        PRECIOPROMEDIO_ = 0
                    End If
                ' ----------------------------------------------------------SALIDAS
                Else
                    ' --------------------------------SALDO Y TOTALES
                    'PRECIOUNITARIO_ = IMPORTEACUMULADO_ / CANTIDADACUMULADA_
                    IMPORTEINSUMO_ = PRECIOPROMEDIO_ * NulosN(RECORDSET_("canpro"))
                    
                    If RECORDSET_("descdoc") = "NC" Then
                        CANTIDADACUMULADA_ = CANTIDADACUMULADA_ + NulosN(RECORDSET_("canpro"))
                        TOTALENTRADAS_ = TOTALENTRADAS_ + NulosN(RECORDSET_("canpro"))
                    Else
                        CANTIDADACUMULADA_ = CANTIDADACUMULADA_ - NulosN(RECORDSET_("canpro"))
                        TOTALSALIDAS_ = TOTALSALIDAS_ + NulosN(RECORDSET_("canpro"))
                    End If
                    ' REDONDEAMOS A 4 DECIMALES
                    CANTIDADACUMULADA_ = Format(CANTIDADACUMULADA_, "0.0000")
                                        
                    If CANTIDADACUMULADA_ < 0 Then
                        MsgBox "Ocurrio un error al procesar el costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                        RECORDSETERRORES_.AddNew
                        RECORDSETERRORES_("numdoc") = NulosC(RECORDSET_("numdoc"))
                        RECORDSETERRORES_("item") = Busca_Codigo(IDITEM_, "id", "descripcion", "Alm_inventario", "N", XCON_)
                        RECORDSETERRORES_("preuni") = CANTIDADACUMULADA_
                        RECORDSETERRORES_("detalle") = "Costo MP - Unidades negativas"
                        RECORDSETERRORES_("fecha") = RECORDSET_("fchdoc")
                        RECORDSETERRORES_.Update
                        BANDERA_ = True
                    End If
                    
                    ' --------------------------------IMPORTE ACUMULADO
                    If RECORDSET_("descdoc") = "NC" Then
                        IMPORTEACUMULADO_ = IMPORTEACUMULADO_ + IMPORTEINSUMO_
                    Else
                        IMPORTEACUMULADO_ = IMPORTEACUMULADO_ - IMPORTEINSUMO_
                    End If
                End If
SIGUIENTE_:
                RECORDSET_.MoveNext
            Wend
    End Select
    
    pImporteMateriaPrima = PRECIOPROMEDIO_ * CANTIDAD_
End Function

Private Sub Cmd_Click(Index As Integer)
    Dim A As Integer
    Dim num As Integer
    Dim Rpta As Integer
    Dim nTitulo As String
    Dim xRs As New ADODB.Recordset
    Dim xCampos() As String
    Dim xform As New eps_librerias.FormSeleccion
    Dim MENSAJE_ As String
    Dim nSQLId As String
    Dim nSQLId2 As String
    Dim NUMEROMAXTRAB_ As Integer
    Dim NUMREGAAGREGAR_ As Integer
    
    Dim ULTIMODIAMES_ As Date
    Dim PRIMERDIAMES_ As Date
    Dim ANIOACTUAL_ As Integer
    Dim MESACTUAL_ As Integer
    
    If Index = 1 Then ' CONFIGURAR DISTRIBUCION GAS. FAB.
        pMostrarGasFab NulosN(RstLibro("id"))
    End If
    
    If QueHace = 3 Then Exit Sub

    ANIOACTUAL_ = AnoTra
    MESACTUAL_ = cbMes.ListIndex + 1
    ' Se encuentra el primer dia del mes actual
    PRIMERDIAMES_ = CDate("01/" & MESACTUAL_ & "/" & ANIOACTUAL_ & "")
    ' Se encuentra el ultimo dia del mes actual
    If MESACTUAL_ = 12 Then MESACTUAL_ = 0: ANIOACTUAL_ = ANIOACTUAL_ + 1
    ULTIMODIAMES_ = CDate("01/" & MESACTUAL_ + 1 & "/" & ANIOACTUAL_ & "") - 1
    ' Si es que haya habido algun cambio se regresan a su estado inicial
    ANIOACTUAL_ = AnoTra
    MESACTUAL_ = cbMes.ListIndex + 1
            
    Select Case Index
        Case 0 ' METODO DE VALORIZACION
            ReDim xCampos(2, 4) As String

            xCampos(0, 0) = "Id":               xCampos(0, 1) = "id":               xCampos(0, 2) = "1000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            xCampos(1, 0) = "Descripción":      xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "4500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
                        
            cSQL = "SELECT * FROM mae_metodoval;"
                
            nTitulo = "Buscando Metodos de valorizacion"
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            txtidmetval.Text = NulosN(xRs("id"))
            lblmetval.Caption = NulosC(xRs("descripcion"))
            cmd(2).SetFocus
            
        Case 2 ' CONSULTAR
            pLlenarDatos MESACTUAL_
            
        Case 3 ' PROCESAR
            pProcesarDatos MESACTUAL_
                    
        Case 4 ' AGREGAR CUENTA
            ReDim xCampos(3, 4) As String

            xCampos(0, 0) = "Cuenta":           xCampos(0, 1) = "cuenta":           xCampos(0, 2) = "1000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            xCampos(1, 0) = "Descripción":      xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "4500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
            xCampos(2, 0) = "Importe":          xCampos(2, 1) = "importe":          xCampos(2, 2) = "1200":     xCampos(2, 3) = "N":    xCampos(2, 4) = "N"
            
            nSQLId = " AND (Left(con_planctas.cuenta, 1) In ('9'))"
            nSQLId = nSQLId & GENERAR_SQL_ID(fg(1), 4, " AND con_planctas.id", "NOT IN", True)
            
            cSQL = "SELECT con_planctas.cuenta, con_planctas.descripcion, IIf(MovPeriodo.DebSol Is Null,0,MovPeriodo.DebSol) AS importe, con_planctas.id AS idcuenta " _
                + vbCr + "FROM con_planctas LEFT JOIN " _
                + vbCr + "( " _
                + vbCr + "SELECT con_planctas.id AS IdCta, con_planctas.cuenta, con_planctas.descripcion, " _
                    + vbCr + "Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebdol=0,0,con_diario.impdebdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS DebSol, " _
                    + vbCr + "Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabdol=0,0,con_diario.imphabdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS HabSol, " _
                    + vbCr + "Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebsol=0,0,con_diario.impdebsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebdol)) AS DebDol, " _
                    + vbCr + "Sum(IIf(con_diario.idmon = 1, IIf(IIf(con_diario.aplicatc = -1, con_diario.tc, IIf(con_tc.impven Is Null, 0, con_tc.impven)) = 0 Or con_diario.imphabsol = 0, 0, con_diario.imphabsol / (IIf(con_diario.aplicatc = -1, con_diario.tc, con_tc.impven))), con_diario.imphabdol)) As HabDol " _
                + vbCr + "FROM (con_planctas RIGHT JOIN con_diario ON con_planctas.id=con_diario.idcue) LEFT JOIN con_tc ON con_diario.fchdoc=con_tc.fecha " _
                + vbCr + "WHERE (((con_diario.fchasi) Between CDate('" & PRIMERDIAMES_ & "') And CDate('" & ULTIMODIAMES_ & "')))  AND (con_diario.ajuste in (0, 1) ) " _
                + vbCr + "GROUP BY con_planctas.id, con_planctas.cuenta, con_planctas.descripcion " _
                + vbCr + "ORDER BY con_planctas.cuenta " _
                + vbCr + ")  AS MovPeriodo ON con_planctas.id = MovPeriodo.IdCta " _
                + vbCr + "WHERE (((con_planctas.id) In (SELECT con_diario.idcue FROM con_diario WHERE  (con_diario.ajuste in (0, 1) )  AND (  (((con_diario.fchasi) Between CDate('" & PRIMERDIAMES_ & "') And CDate('" & ULTIMODIAMES_ & "')))  OR  (con_diario.fchasi)<CDate('" & PRIMERDIAMES_ & "')  OR  (con_diario.fchasi) is null  )   ))) " & nSQLId _
                + vbCr + "ORDER BY con_planctas.cuenta;"
                
            nTitulo = "Buscando Cuentas"
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "cuenta", "cuenta", Principio
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub

            Agregando = True
            With fg(1)
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = NulosC(xRs("cuenta"))
                .TextMatrix(.Rows - 1, 2) = NulosC(xRs("descripcion"))
                .TextMatrix(.Rows - 1, 3) = Format(NulosN(xRs("importe")), FORMAT_IMPORTEKARDEX)
                .TextMatrix(.Rows - 1, 4) = NulosC(xRs("idcuenta"))
            End With
            lblTotalGr.Caption = Format(GRID_SUMAR_COL(fg(1), 3), FORMAT_MONTO)
            Agregando = False
        
        Case 5 ' SELECCIONAR CUENTA
            ReDim xCampos(3, 4) As String

            xCampos(0, 0) = "Cuenta":           xCampos(0, 1) = "cuenta":           xCampos(0, 2) = "1000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            xCampos(1, 0) = "Descripción":      xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "4500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
            xCampos(2, 0) = "Importe":          xCampos(2, 1) = "importe":          xCampos(2, 2) = "1200":     xCampos(2, 3) = "N":    xCampos(2, 4) = "N"
            
            nSQLId = " AND (Left(con_planctas.cuenta, 1) In ('9'))"
            nSQLId = nSQLId & GENERAR_SQL_ID(fg(1), 4, " AND con_planctas.id", "NOT IN", True)
            
            cSQL = "SELECT 0 AS xsel, con_planctas.cuenta, con_planctas.descripcion, IIf(MovPeriodo.DebSol Is Null,0,MovPeriodo.DebSol) AS importe, con_planctas.id AS idcuenta " _
                + vbCr + "FROM con_planctas LEFT JOIN " _
                + vbCr + "( " _
                + vbCr + "SELECT con_planctas.id AS IdCta, con_planctas.cuenta, con_planctas.descripcion, " _
                    + vbCr + "Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebdol=0,0,con_diario.impdebdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS DebSol, " _
                    + vbCr + "Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabdol=0,0,con_diario.imphabdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS HabSol, " _
                    + vbCr + "Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebsol=0,0,con_diario.impdebsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebdol)) AS DebDol, " _
                    + vbCr + "Sum(IIf(con_diario.idmon = 1, IIf(IIf(con_diario.aplicatc = -1, con_diario.tc, IIf(con_tc.impven Is Null, 0, con_tc.impven)) = 0 Or con_diario.imphabsol = 0, 0, con_diario.imphabsol / (IIf(con_diario.aplicatc = -1, con_diario.tc, con_tc.impven))), con_diario.imphabdol)) As HabDol " _
                + vbCr + "FROM (con_planctas RIGHT JOIN con_diario ON con_planctas.id=con_diario.idcue) LEFT JOIN con_tc ON con_diario.fchdoc=con_tc.fecha " _
                + vbCr + "WHERE (((con_diario.fchasi) Between CDate('" & PRIMERDIAMES_ & "') And CDate('" & ULTIMODIAMES_ & "')))  AND (con_diario.ajuste in (0, 1) ) " _
                + vbCr + "GROUP BY con_planctas.id, con_planctas.cuenta, con_planctas.descripcion " _
                + vbCr + "ORDER BY con_planctas.cuenta " _
                + vbCr + ")  AS MovPeriodo ON con_planctas.id = MovPeriodo.IdCta " _
                + vbCr + "WHERE (((con_planctas.id) In (SELECT con_diario.idcue FROM con_diario WHERE  (con_diario.ajuste in (0, 1) )  AND (  (((con_diario.fchasi) Between CDate('" & PRIMERDIAMES_ & "') And CDate('" & ULTIMODIAMES_ & "')))  OR  (con_diario.fchasi)<CDate('" & PRIMERDIAMES_ & "')  OR  (con_diario.fchasi) is null  )   ))) " & nSQLId _
                + vbCr + "ORDER BY con_planctas.cuenta;"
                
            xform.SqlCad = cSQL
            xform.Titulo = "Seleccionando Cuentas"
            Set xform.Coneccion = xCon
            Set xRs = Nothing
            Set xRs = xform.Seleccionar(xCampos)
            Set xform = Nothing
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub

            Agregando = True
            With fg(1)
                xRs.MoveFirst
                While Not xRs.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = NulosC(xRs("cuenta"))
                    .TextMatrix(.Rows - 1, 2) = NulosC(xRs("descripcion"))
                    .TextMatrix(.Rows - 1, 3) = Format(NulosN(xRs("importe")), FORMAT_IMPORTEKARDEX)
                    .TextMatrix(.Rows - 1, 4) = NulosC(xRs("idcuenta"))
                    xRs.MoveNext
                Wend
            End With
            lblTotalGr.Caption = Format(GRID_SUMAR_COL(fg(1), 3), FORMAT_MONTO)
            Agregando = False
            
        Case 6 ' SELECCIONAR CUENTAS ANTERIORES
            ReDim xCampos(3, 4) As String

            xCampos(0, 0) = "Cuenta":           xCampos(0, 1) = "cuenta":           xCampos(0, 2) = "1000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            xCampos(1, 0) = "Descripción":      xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "4500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
            xCampos(2, 0) = "Mes":              xCampos(2, 1) = "desmes":           xCampos(2, 2) = "2000":     xCampos(2, 3) = "C":    xCampos(2, 4) = "C"
                        
            cSQL = "SELECT 0 AS xsel, con_librocostocta.idlibro, con_librocostocta.idcuenta, con_planctas.cuenta, con_planctas.descripcion, con_librocostocta.importe, con_meses.descripcion AS desmes, con_librocosto.idmes " _
                + vbCr + "FROM ((con_librocosto INNER JOIN con_librocostocta ON con_librocosto.id = con_librocostocta.idlibro) INNER JOIN con_planctas ON con_librocostocta.idcuenta = con_planctas.id) INNER JOIN con_meses ON con_librocosto.idmes = con_meses.id " _
                + vbCr + "WHERE ((con_librocostocta.importe)>0) " _
                + vbCr + "ORDER BY con_librocosto.idmes;"
            
            xform.SqlCad = cSQL
            xform.Titulo = "Seleccionando Cuentas"
            Set xform.Coneccion = xCon
            Set xRs = Nothing
            Set xRs = xform.Seleccionar(xCampos)
            Set xform = Nothing
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            nSQLId = nSQLId & GENERAR_SQL_ID_RST(xRs, "idcuenta", " AND con_planctas.id", "IN", True)
                        
            cSQL = "SELECT 0 AS xsel, con_planctas.cuenta, con_planctas.descripcion, IIf(MovPeriodo.DebSol Is Null,0,MovPeriodo.DebSol) AS importe, con_planctas.id AS idcuenta " _
                + vbCr + "FROM con_planctas LEFT JOIN " _
                + vbCr + "( " _
                + vbCr + "SELECT con_planctas.id AS IdCta, con_planctas.cuenta, con_planctas.descripcion, " _
                    + vbCr + "Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebdol=0,0,con_diario.impdebdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS DebSol, " _
                    + vbCr + "Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabdol=0,0,con_diario.imphabdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS HabSol, " _
                    + vbCr + "Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebsol=0,0,con_diario.impdebsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebdol)) AS DebDol, " _
                    + vbCr + "Sum(IIf(con_diario.idmon = 1, IIf(IIf(con_diario.aplicatc = -1, con_diario.tc, IIf(con_tc.impven Is Null, 0, con_tc.impven)) = 0 Or con_diario.imphabsol = 0, 0, con_diario.imphabsol / (IIf(con_diario.aplicatc = -1, con_diario.tc, con_tc.impven))), con_diario.imphabdol)) As HabDol " _
                + vbCr + "FROM (con_planctas RIGHT JOIN con_diario ON con_planctas.id=con_diario.idcue) LEFT JOIN con_tc ON con_diario.fchdoc=con_tc.fecha " _
                + vbCr + "WHERE (((con_diario.fchasi) Between CDate('" & PRIMERDIAMES_ & "') And CDate('" & ULTIMODIAMES_ & "')))  AND (con_diario.ajuste in (0, 1) ) " _
                + vbCr + "GROUP BY con_planctas.id, con_planctas.cuenta, con_planctas.descripcion " _
                + vbCr + "ORDER BY con_planctas.cuenta " _
                + vbCr + ")  AS MovPeriodo ON con_planctas.id = MovPeriodo.IdCta " _
                + vbCr + "WHERE (((con_planctas.id) In (SELECT con_diario.idcue FROM con_diario WHERE  (con_diario.ajuste in (0, 1) )  AND (  (((con_diario.fchasi) Between CDate('" & PRIMERDIAMES_ & "') And CDate('" & ULTIMODIAMES_ & "')))  OR  (con_diario.fchasi)<CDate('" & PRIMERDIAMES_ & "')  OR  (con_diario.fchasi) is null  )   ))) " & nSQLId _
                + vbCr + "ORDER BY con_planctas.cuenta;"
            
            Set xRs = Nothing
            RST_Busq xRs, cSQL, xCon
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            fg(1).Rows = fg(1).FixedRows
                        
                        
                        
            Agregando = True
            With fg(1)
                xRs.MoveFirst
                While Not xRs.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = NulosC(xRs("cuenta"))
                    .TextMatrix(.Rows - 1, 2) = NulosC(xRs("descripcion"))
                    .TextMatrix(.Rows - 1, 3) = Format(NulosN(xRs("importe")), FORMAT_IMPORTEKARDEX)
                    .TextMatrix(.Rows - 1, 4) = NulosC(xRs("idcuenta"))
                    xRs.MoveNext
                Wend
            End With
            lblTotalGr.Caption = Format(GRID_SUMAR_COL(fg(1), 3), FORMAT_MONTO)
            Agregando = False
        
        Case 7 ' ELIMINAR CUENTA
            If fg(1).Rows <= fg(1).FixedRows Then Exit Sub
            Rpta = MsgBox("¿Está seguro de eliminar el registro actual?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
            If Rpta = vbYes Then fg(1).RemoveItem fg(1).Row: lblTotalGr.Caption = Format(GRID_SUMAR_COL(fg(1), 3), FORMAT_MONTO)
            
        Case 8 ' ELIMINAR TODAS CUENTAS
            If fg(1).Rows <= fg(1).FixedRows Then Exit Sub
            Rpta = MsgBox("¿Está seguro de eliminar todos los registros?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
            If Rpta = vbYes Then fg(1).Rows = fg(1).FixedRows: lblTotalGr.Caption = Format(GRID_SUMAR_COL(fg(1), 3), FORMAT_MONTO)
            
        Case 9 ' AGREGAR CUENTA GAS FAB
            
        Case 10 ' SELECCIONAR CUENTA GAS FAB
'
        Case 11 ' ELIMINAR CUENTA GAS FAB
'
        Case 12 ' ELIMINAR TODOS CUENTA GAS FAB

        Case 13 ' ACEPTAR DISTRIBUCION DE CTA GAS FAB
            limpiarRST RSTCTAGASFAB
            For A = 1 To fg(1).Rows - 1
                RSTCTAGASFAB.AddNew
                RSTCTAGASFAB("idcuenta") = NulosN(fg(1).TextMatrix(A, 4))
                RSTCTAGASFAB("importe") = NulosN(fg(1).TextMatrix(A, 3))
                RSTCTAGASFAB.Update
            Next A
            Frm4.Visible = False
            
        Case 14 ' CANCELAR DISTRIBUCION GAS FAB
            Frm4.Visible = False
        
    End Select
End Sub

Private Function RST_SUMAR(RECORDSET_ As ADODB.Recordset, CAMPORSUMA_ As String) As Double
    Dim SUMA_ As Double
    
    If RECORDSET_.State = 0 Then RST_SUMAR = 0: Exit Function
    If RECORDSET_.RecordCount = 0 Then RST_SUMAR = 0: Exit Function
    
    RECORDSET_.MoveFirst
    SUMA_ = 0
    While Not RECORDSET_.EOF
        SUMA_ = SUMA_ + RECORDSET_(CAMPORSUMA_)
        RECORDSET_.MoveNext
    Wend
    
    RST_SUMAR = SUMA_
End Function

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstLibro("id")), xCon
    End If
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstLibro
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA EN FORMA ASCENDENTE O DECENDETE LAS COLUMNAS DEL CONTROL Dg3
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstLibro.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub fg_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Function GRID_SUMARSEL(ByRef GRID_ As VSFlexGrid, COLUMNASEL_ As Integer, COLUMNASUMA_ As Integer, _
                        Optional FILAINICIO_ As Integer = 0, Optional FILAFIN_ As Integer = 0) As Double
    Dim ACUMULADO_ As Double
    Dim A As Integer
    
    ACUMULADO_ = 0
    If FILAINICIO_ = 0 Then FILAINICIO_ = GRID_.FixedRows
    If FILAFIN_ = 0 Then FILAFIN_ = GRID_.Rows - 1
    
    With GRID_
        For A = FILAINICIO_ To FILAFIN_
            If NulosN(.TextMatrix(A, COLUMNASEL_)) = -1 Then
                ACUMULADO_ = ACUMULADO_ + NulosN(.TextMatrix(A, COLUMNASUMA_))
            End If
        Next A
    End With
    
    GRID_SUMARSEL = ACUMULADO_
End Function

Private Sub fg_DblClick(Index As Integer)
    If Index <> 0 Then Exit Sub
    If Agregando Then Exit Sub
    If fg(0).Row = fg(0).Rows - 1 Then Exit Sub
    
    Me.MousePointer = vbHourglass
    llenarDetalleInsumos NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.COLUMNACORRELATIVO_))
    llenarDetallePersonal NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.COLUMNACORRELATIVO_))
    llenarDetalleGasFab NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACABECERA_.COLUMNACOSTOFABRICA_))
    Me.MousePointer = vbDefault
End Sub

Private Sub cbMes_DropDown()
    If Agregando Then Exit Sub
    ESTADOANTERIOR_ = cbMes.ItemData(cbMes.ListIndex)
End Sub

Private Sub Anular()
    Dim MENSAJE_ As String
    Dim xRs As New ADODB.Recordset
    
    If verificarCambioEstado(NulosN(RstLibro("id")), MENSAJE_) Then
        ' ----------------------------------------SE CAMBIA DE ESTADO A LA SOLICITUD DE MATERIALES
        cSQL = "UPDATE pro_solicitudmat SET pro_solicitudmat.estado = " & ESTADOANULADO_ & " " _
            + vbCr + "WHERE (((pro_solicitudmat.idtipdocref)=115) AND ((pro_solicitudmat.iddocref)=" & NulosN(RstLibro("id")) & "));"
        ' --------------EJECUTA COMANDO
        xCon.Execute cSQL
        ' --------------ACTUALIZA VAR_EDICION
        cSQL = "SELECT pro_solicitudmat.id " _
            + vbCr + "FROM pro_solicitudmat " _
            + vbCr + "WHERE (((pro_solicitudmat.idtipdocref)=115) And ((pro_solicitudmat.iddocref)=" & NulosN(RstLibro("ID")) & "))"
        
        Set xRs = Nothing
        RST_Busq xRs, cSQL, xCon
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        xRs.MoveFirst
        While Not xRs.EOF
            GrabarOperacion xIdUsuario, 54, 7, xHorIni, Time, Date, xCon, NulosN(xRs("id"))
            xRs.MoveNext
        Wend
        ' ----------------------------------------SE CAMBIA DE ESTADO AL REGISTRO
        xCon.Execute "UPDATE pro_ordenprod SET pro_ordenprod.estado = " & ESTADOANULADO_ & " WHERE (((pro_ordenprod.id) = " & NulosN(RstLibro("id")) & "))"
        MsgBox "El registro se anuló con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstLibro.Requery
        Dg1.Refresh
    Else
        MsgBox MENSAJE_, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

Private Function verificarCambioEstado(IDORD_ As Integer, ByRef MENSAJE_ As String) As Boolean
    Dim xRs As New ADODB.Recordset
    
    ' -------------------------------------SOLICITUD DE MATERIALES
    cSQL = "SELECT * " _
        + vbCr + "FROM pro_solicitudmat " _
        + vbCr + "WHERE (((pro_solicitudmat.idtipdocref)=115) AND ((pro_solicitudmat.iddocref)=" & IDORD_ & ") AND ((pro_solicitudmat.estado)=" & ESTADOPROCESADO_ & "));"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    MENSAJE_ = "Solicitud de Materiales"
    
    If xRs.State = 0 Then verificarCambioEstado = False: GoTo SALIR_
    If xRs.RecordCount > 0 Then verificarCambioEstado = False: GoTo SALIR_
    
    verificarCambioEstado = True
    Exit Function
    
SALIR_:
    MENSAJE_ = "Se han encontrado " & MENSAJE_ & " que se encuentran en un estado no modificable; " _
    & vbCr & "verifique la condición de dichos Registros para completar esta acción."
End Function

Private Function cambiarEstadoRelacionados(IDORDDET_ As Double, ESTADO_ As Double) As Boolean
    Dim ID_ As Double
    
    On Error GoTo ERROR_
    ' Salidas de Almacen
    cSQL = "UPDATE alm_ingreso SET alm_ingreso.estado = " & ESTADO_ & " " _
        + vbCr + "WHERE (((alm_ingreso.idorddet)=" & IDORDDET_ & "));"

    xCon.Execute cSQL
    
    ' GRABAMOS LOS MOVIMIENTOS
    ' INGRESOS Y SALIDAS DE ALMACEN
    ID_ = Busca_Codigo(IDORDDET_, "idorddet", "id", "alm_ingreso", "N", xCon)
    GrabarOperacion xIdUsuario, 8, 7, xHorIni, Time, Date, xCon, ID_
        
    cambiarEstadoRelacionados = True
    Exit Function
    
ERROR_:
    MsgBox "Ha ocurrido un error al tratar de cambiar de estado", vbInformation, xTitulo
    cambiarEstadoRelacionados = False
End Function

Private Sub fg_EnterCell(Index As Integer)
    If QueHace = 3 Or Index <> 1 Then
        fg(Index).Editable = flexEDNone
        Exit Sub
    End If
    fg(Index).Editable = flexEDKbdMouse
End Sub

Private Sub fg_RowColChange(Index As Integer)
    If Agregando Then Exit Sub
    If Index <> 0 Then Exit Sub
    
    fg(3).Rows = fg(3).FixedRows
    fg(4).Rows = fg(4).FixedRows
    fg(2).Rows = fg(2).FixedRows
End Sub

Private Sub Form_Load()
    QueHace = 3
    TabOne1.CurrTab = 0
    SeEjecuto = False
    Agregando = False
    iniciarCampos
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDOS E CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        mMesActivo = xMes
            
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
        pCargarGrid
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        '--interrumpir
        BANDERA_ = True
    End If
End Sub

Private Sub Form_Resize()
    ' Si esta minimizado no se hace nada
    If Me.WindowState = 1 Then Exit Sub

    If Me.Width <= 12000 Then Me.Width = 12000
    If Me.Height <= 4000 Then Me.Height = 4000

    ' Se dimensiona el Contenido
    TabOne1.Width = Me.Width - 150
    TabOne1.Height = Me.Height - 750
    
    Label4(0).Width = Me.Width - 100
    Dg1.Width = TabOne1.Width - 100
    Dg1.Height = TabOne1.Height - 1000
    
    ' Se dimensiona el Detalle
    Label5.Width = Me.Width - 100
    
    Frame4.Width = TabOne1.Width - 150
    Frame4.Height = TabOne1.Height - 1905
    
    fg(0).Width = Frame4.Width - 150
    fg(0).Height = Frame4.Height - 4230
    
    TabOne2.Top = Frame4.Height - 3855
    TabOne2.Width = Frame4.Width - 120
    fg(3).Width = TabOne2.Width - 285
    fg(4).Width = TabOne2.Width - 285
    fg(2).Width = TabOne2.Width - 1710
    cmd(9).Left = TabOne2.Width - 1575
    cmd(10).Left = TabOne2.Width - 1575
    cmd(11).Left = TabOne2.Width - 1575
    cmd(12).Left = TabOne2.Width - 1575
End Sub

Private Sub iniciarCampos()
    TabOne1.CurrTab = 0
    TabOne2.CurrTab = 0
    
    ' -------------------------PROPIEDADES DE PROCESADO
    PROPIEDADES_.MODOTAREA_ = 3
    PROPIEDADES_.PORCENTAJE_ = 10
    PROPIEDADES_.MINUTOS_ = "00:10"
    PROPIEDADES_.INCLUIRREFRIGERIO_ = True
    PROPIEDADES_.HORINIREFRIGERIO_ = "13:00"
    PROPIEDADES_.HORFINREFRIGERIO_ = "14:00"
    PROPIEDADES_.LIMITARNUMEROPERSONAL_ = True
    PROPIEDADES_.LIMITARNUMEROTAREAS_ = True
    PROPIEDADES_.LIMITARSELPERSONAL_ = True
    
    '**********************
    ' CONFIGURACIONES GRID
    '**********************
    fg(0).AllowUserResizing = flexResizeColumns
    fg(0).AutoSearch = flexSearchFromTop
    fg(0).ExplorerBar = flexExSortShow
    fg(0).SelectionMode = flexSelectionByRow
    fg(0).ForeColorSel = &H80000005
    fg(0).BackColorSel = &H80&
    fg(0).ColWidth(COLUMNAIDITEM_) = 0
    fg(0).ColWidth(COLUMNAIDPROD_) = 0
    fg(0).ColWidth(COLUMNACORRELATIVO_) = 0
    fg(0).Rows = fg(0).FixedRows
    fg(0).FrozenCols = COLUMNAITEM_
    
    
    fg(2).AllowUserResizing = flexResizeColumns
    fg(2).AutoSearch = flexSearchFromTop
    fg(2).ExplorerBar = flexExSortShow
    fg(2).ForeColorSel = &H80000005
    fg(2).BackColorSel = &H80&
    fg(2).Editable = flexEDKbdMouse
    fg(2).Rows = fg(2).FixedRows
    fg(2).ColWidth(4) = 0
    
    fg(3).AllowUserResizing = flexResizeColumns
    fg(3).AutoSearch = flexSearchFromTop
    fg(3).ExplorerBar = flexExSortShow
    fg(3).ForeColorSel = &H80000005
    fg(3).BackColorSel = &H80&
    fg(3).Editable = flexEDKbdMouse
    fg(3).Rows = fg(3).FixedRows
    fg(3).ColWidth(6) = 0
    fg(3).ColWidth(7) = 0
    
    fg(4).AllowUserResizing = flexResizeColumns
    fg(4).AutoSearch = flexSearchFromTop
    fg(4).ExplorerBar = flexExSortShow
    fg(4).ForeColorSel = &H80000005
    fg(4).BackColorSel = &H80&
    fg(4).Editable = flexEDKbdMouse
    fg(4).ColWidth(5) = 0
    fg(4).ColWidth(6) = 0
    fg(4).Rows = fg(4).FixedRows
    
    CORRELATIVO_ = -9999
    Llenar_Mes cbMes
    FILAINICIAL_ = fg(0).FixedRows
End Sub

Sub ActivaTool()
    Dim A As Integer
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Private Sub pCargarGrid()
    Dim cSQL  As String
    Dim Rpta As Integer
    
    TDB_FiltroLimpiar Dg1
    
    cSQL = "SELECT con_librocosto.*, con_meses.descripcion AS desmes, mae_metodoval.descripcion AS desmetval, IIf([con_librocosto].[aplvtas]=0,'TODOS','VENTAS') AS desaplgasfab, IIf([con_librocosto].[tipo]=0,'GLOBAL','DISTRIBUIDO') AS destipdisgasfab " _
        + vbCr + "FROM (con_librocosto LEFT JOIN mae_metodoval ON con_librocosto.idmetodoval = mae_metodoval.id) LEFT JOIN con_meses ON con_librocosto.idmes = con_meses.id " _
        + vbCr + "ORDER BY con_librocosto.idmes;"
        
    Me.MousePointer = vbHourglass
    
    RST_Busq RstLibro, cSQL, xCon
    Set Dg1.DataSource = RstLibro
    
    Me.MousePointer = vbDefault
    If RstLibro.State = 0 Then Exit Sub
End Sub

Private Sub MuestraSegundoTab()
    Dim Rst As New ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    Dim Rpta As Integer
    Dim ULTIMODIAMES_ As Date
    Dim PRIMERDIAMES_ As Date
    Dim ANIOACTUAL_ As Double
    Dim MESACTUAL_ As Double
    
    On Error Resume Next
    
    ANIOACTUAL_ = AnoTra
    MESACTUAL_ = NulosN(RstLibro("idmes"))
    ' Se encuentra el primer dia del mes actual
    PRIMERDIAMES_ = CDate("01/" & MESACTUAL_ & "/" & ANIOACTUAL_ & "")
    ' Se encuentra el ultimo dia del mes actual
    If MESACTUAL_ = 12 Then MESACTUAL_ = 0: ANIOACTUAL_ = ANIOACTUAL_ + 1
    ULTIMODIAMES_ = CDate("01/" & MESACTUAL_ + 1 & "/" & ANIOACTUAL_ & "") - 1
    ' Si es que haya habido algun cambio se regresan a su estado inicial
    ANIOACTUAL_ = AnoTra
    MESACTUAL_ = NulosN(RstLibro("idmes")) - 1
    
    Agregando = True
    Blanquea
    
    If RstLibro.RecordCount = 0 Then Exit Sub
    If RstLibro.EOF = True Then Exit Sub
     
    Agregando = True
    cbMes.ListIndex = NulosN(RstLibro("idmes")) - 1
    txtdescripcion.Text = NulosC(RstLibro("descripcion"))
    txtidmetval.Text = NulosN(RstLibro("idmetodoval"))
    lblmetval.Caption = NulosC(RstLibro("desmetval"))
    If NulosN(RstLibro("aplvtas")) = 1 Then
        optdisgasfab(1).Value = True
    Else
        optdisgasfab(0).Value = True
    End If
    If NulosN(RstLibro("tipo")) = 0 Then
        opttipdiscta(0).Value = True
    Else
        opttipdiscta(1).Value = True
    End If
    
    ' SE LLENAN LOS RECORDSET ASOCIADOS
    llenarDefinirRST NulosN(RstLibro("id"))
    
    ' BUSCAMOS PRODUCCION DEL PROCESO
    cSQL = "SELECT pro_produccion.dia AS fecha, pro_producciondet.numparte AS numprod, con_librocostodet.proceso, alm_inventario.descripcion AS desitem, pro_receta.codrec, pla_empleados.nombre, mae_unidades.abrev, con_librocostodet.cantidad, pro_producciondet.horini, pro_producciondet.horfin, con_librocostodet.impmprima, con_librocostodet.impmanobr, con_librocostodet.impgasfab, IIf([cPREVEN].[preven]<>0,'V','P') AS tipo, cPREVEN.preven, con_librocostodet.iditem, con_librocostodet.idprod, con_librocostodet.id " _
        + vbCr + "FROM ((((((con_librocostodet LEFT JOIN alm_inventario ON con_librocostodet.iditem = alm_inventario.id) LEFT JOIN (pro_produccion RIGHT JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro) ON (con_librocostodet.iditem = pro_producciondet.iditem) AND (con_librocostodet.idprod = pro_producciondet.idpro)) LEFT JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_producciondet.idunimed = mae_unidades.id) LEFT JOIN pro_emp ON pro_producciondet.idres = pro_emp.id) LEFT JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id) LEFT JOIN " _
        + vbCr + "( " _
        + vbCr + "SELECT vta_ventasdet.iditem, Avg(vta_ventasdet.preuni) AS preven " _
        + vbCr + "FROM vta_ventas INNER JOIN vta_ventasdet ON vta_ventas.id = vta_ventasdet.idvta " _
        + vbCr + "WHERE (((vta_ventas.fchdoc)>=CDate('" & PRIMERDIAMES_ & "') And (vta_ventas.fchdoc)<=CDate('" & ULTIMODIAMES_ & "'))) " _
        + vbCr + "GROUP BY vta_ventasdet.iditem " _
        + vbCr + ") AS cPREVEN ON pro_producciondet.iditem = cPREVEN.iditem " _
        + vbCr + "WHERE (((con_librocostodet.idlibro)=" & NulosN(RstLibro("id")) & ")) " _
        + vbCr + "ORDER BY pro_produccion.dia;"
        
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    fg(0).Rows = fg(0).FixedRows
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    
    xRs.MoveFirst
    Dim IMPORTEPRODUCCION_ As Double
    Dim IMPORTEVENTA_ As Double
    With fg(0)
        While Not xRs.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNAFECHA_) = Format(xRs("fecha"), FORMAT_DATE)
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNAREGPROD_) = NulosC(xRs("numprod"))
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNATIPO_) = NulosC(xRs("tipo"))
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNAPROCESO_) = NulosN(xRs("proceso"))
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNAITEM_) = NulosC(xRs("desitem"))
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNARECETA_) = NulosC(xRs("codrec"))
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNARESPONSABLE_) = NulosC(xRs("nombre"))
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNAUNIMED_) = NulosC(xRs("abrev"))
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNACANTIDAD_) = Format(NulosN(xRs("cantidad")), FORMAT_CANTIDAD)
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNAHORINI_) = Format(xRs("horini"), FORMAT_HORA_SIN_SEGUNDO)
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNAHORFIN_) = Format(xRs("horfin"), FORMAT_HORA_SIN_SEGUNDO)
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNACOSTOMP_) = Format(NulosN(xRs("impmprima")), FORMAT_MONTO)
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNACOSTOMOBRA_) = Format(NulosN(xRs("impmanobr")), FORMAT_MONTO)
            IMPORTEPRODUCCION_ = NulosN(xRs("impmprima")) + NulosN(xRs("impmanobr"))
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNACOSTOPRIMO_) = Format(IMPORTEPRODUCCION_, FORMAT_MONTO)
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNACOSTOFABRICA_) = Format(NulosN(xRs("impgasfab")), FORMAT_IMPORTEKARDEX)
            IMPORTEPRODUCCION_ = IMPORTEPRODUCCION_ + NulosN(xRs("impgasfab"))
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNACOSTOTOTAL_) = Format(IMPORTEPRODUCCION_, FORMAT_MONTO)
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNACOSTOUNIPRODUCCION_) = Format(IMPORTEPRODUCCION_ / NulosN(xRs("cantidad")), FORMAT_MONTO)
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNAPRECIOVENTA_) = Format(NulosN(xRs("preven")), FORMAT_MONTO)
            IMPORTEVENTA_ = NulosN(xRs("cantidad")) * NulosN(xRs("preven"))
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNAIMPORTEVENTA_) = Format(IMPORTEVENTA_, FORMAT_MONTO)
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNADESVIACION_) = Format(IMPORTEPRODUCCION_ - IMPORTEVENTA_, FORMAT_MONTO)
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNADESVIACIONPORC_) = Format((IMPORTEPRODUCCION_ - IMPORTEVENTA_) / IMPORTEPRODUCCION_ * 100, FORMAT_MONTO)
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNAIDPROD_) = NulosN(xRs("idprod"))
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNAIDITEM_) = NulosN(xRs("iditem"))
            .TextMatrix(.Rows - 1, COLUMNACABECERA_.COLUMNACORRELATIVO_) = NulosN(xRs("id"))
            xRs.MoveNext
        Wend
        
        .Rows = .Rows + 1
        FORMATO_CELDA fg(0), .Rows - 1, COLUMNAHORFIN_, , True, , "TOTAL"
        .TextMatrix(.Rows - 1, COLUMNACOSTOMP_) = Format(GRID_SUMAR_COL(fg(0), COLUMNACOSTOMP_), FORMAT_MONTO)
        .TextMatrix(.Rows - 1, COLUMNACOSTOMOBRA_) = Format(GRID_SUMAR_COL(fg(0), COLUMNACOSTOMOBRA_), FORMAT_MONTO)
        .TextMatrix(.Rows - 1, COLUMNACOSTOPRIMO_) = Format(GRID_SUMAR_COL(fg(0), COLUMNACOSTOPRIMO_), FORMAT_MONTO)
        .TextMatrix(.Rows - 1, COLUMNACOSTOFABRICA_) = Format(GRID_SUMAR_COL(fg(0), COLUMNACOSTOFABRICA_), FORMAT_MONTO)
        .TextMatrix(.Rows - 1, COLUMNACOSTOTOTAL_) = Format(GRID_SUMAR_COL(fg(0), COLUMNACOSTOTOTAL_), FORMAT_MONTO)
        .TextMatrix(.Rows - 1, COLUMNAIMPORTEVENTA_) = Format(GRID_SUMAR_COL(fg(0), COLUMNAIMPORTEVENTA_), FORMAT_MONTO)
        .TopRow = .Rows - 1
        
    End With
    Agregando = False
End Sub

Sub Cancelar()
    Bloquea
    Label5.Caption = "Detalle de Orden de Producción"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
     
    ActivaTool
    QueHace = 3
    Dg1.SetFocus
End Sub

Sub Nuevo()
    QueHace = 1
    xHorIni = Time
    Bloquea
    Blanquea
    fg(0).Rows = fg(0).FixedRows
    fg(1).Rows = fg(1).FixedRows
    fg(2).Rows = fg(2).FixedRows
    fg(3).Rows = fg(3).FixedRows
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    
    Label5.Caption = "Agregando Libro de Costo de Producción"
    fg(0).Editable = flexEDKbdMouse
    fg(0).SelectionMode = flexSelectionFree
End Sub

Sub Bloquea()
    cbMes.Locked = Not cbMes.Locked
    txtdescripcion.Locked = Not txtdescripcion.Locked
    txtidmetval.Locked = Not txtidmetval.Locked
    habilitar cmd, Not txtdescripcion.Locked
    habilitar optdisgasfab, Not txtdescripcion.Locked
    habilitar opttipdiscta, Not txtdescripcion.Locked
    cmd(1).Enabled = True
End Sub

Sub Blanquea()
    txtdescripcion.Text = ""
    txtidmetval.Text = ""
    lblmetval.Caption = ""
End Sub

Function Grabar() As Boolean
    Dim IDLIBRO_ As Integer
    Dim DESCRIPCION_ As String
    Dim IDMES_ As Integer
    Dim IDMETODOVAL_ As Integer
    Dim APLVTAS_ As Integer
    Dim TIPO_ As Integer
    Dim A As Integer
    Dim NUMSEL_ As Integer
    
    Dim IDREC_ As Integer
    Dim IDUNIMED_ As Integer
    Dim CANTIDAD_ As Double
    Dim IDLINEA_ As Integer
    Dim EFIC_ As Integer
    Dim HORFIN_ As String
    Dim FCHFIN_ As String
    Dim NUMOP_ As Integer
    Dim REPROC_ As Boolean
    Dim IDESTADO_ As Integer
    Dim xRs As New ADODB.Recordset
    
    Dim xRsTar As New ADODB.Recordset
    Dim xRsPer As New ADODB.Recordset
    Dim xRsRep As New ADODB.Recordset
    
    Dim xRsAux As New ADODB.Recordset
    
    ' VERIFICAMOS QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
    If txtdescripcion.Text = "" Then
        MsgBox "No ha especificado una descripcion para el libro actual", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        txtdescripcion.SetFocus
        Exit Function
    End If
    
    If txtidmetval.Text = "" Then
        MsgBox "No ha especificado el metodo de valorizacion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        txtidmetval.SetFocus
        Exit Function
    End If
    
    If opttipdiscta(0).Value = True Then
        If fg(1).Rows = fg(1).FixedRows Then
            MsgBox "No ha especificado cuentas para la distribución de gastos de fábrica", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Cmd_Click 1
            Exit Function
        End If
    End If
    
    If fg(0).Rows = fg(0).FixedRows Then
        MsgBox "No se han procesado datos de producción para el libro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Function
    End If
        
    ' Se llenan los detalles
    If QueHace = 1 Then IDLIBRO_ = 0 Else IDLIBRO_ = NulosN(RstLibro("id"))
    If optdisgasfab(0).Value = True Then
        APLVTAS_ = 0
    Else
        APLVTAS_ = 1
    End If
    If opttipdiscta(0).Value = True Then
        TIPO_ = 0
    Else
        TIPO_ = 1
    End If
    IDMES_ = cbMes.ListIndex + 1
    IDMETODOVAL_ = NulosN(txtidmetval.Text)
    DESCRIPCION_ = NulosC(txtdescripcion.Text)
    
    If RSTCTAGASFAB.State = 0 Then Grabar = False: Exit Function
    If RSTCABECERA.State = 0 Then Grabar = False: Exit Function
    If RSTDETALLEMATPRI.State = 0 Then Grabar = False: Exit Function
    If RSTDETALLEMANOBR.State = 0 Then Grabar = False: Exit Function
    If RSTDETALLEGASFAB.State = 0 Then Grabar = False: Exit Function
        
    RSTCTAGASFAB.Filter = adFilterNone
    RSTCABECERA.Filter = adFilterNone
    RSTDETALLEMATPRI.Filter = adFilterNone
    RSTDETALLEMANOBR.Filter = adFilterNone
    RSTDETALLEGASFAB.Filter = adFilterNone
    
    ' Se graba el movimiento
    Grabar = grabarLibCosPro(IDMES_, DESCRIPCION_, IDMETODOVAL_, APLVTAS_, TIPO_, RSTCABECERA, RSTCTAGASFAB, RSTDETALLEMATPRI, _
                                    RSTDETALLEMANOBR, RSTDETALLEGASFAB, IDLIBRO_, CInt(AnoTra), 50, QueHace, xHorIni)
    
    mIdRegistro = IDLIBRO_
End Function

Function grabarLibCosPro(MES_ As Integer, DESCRIPCION_ As String, _
                                    IDMETODOVAL_ As Integer, APLVTAS_ As Integer, _
                                    TIPO_ As Integer, RSTDET_ As ADODB.Recordset, _
                                    RSTCTAGASFAB_ As ADODB.Recordset, RSTMATPRI_ As ADODB.Recordset, _
                                    RSTMANOBR_ As ADODB.Recordset, RSTGASFAB_ As ADODB.Recordset, _
                                    Optional ByRef IDLIBRO_ As Integer, Optional ANIO_ As Integer, _
                                    Optional IDFORM_ As Integer, Optional QUEHACE_ As Integer, _
                                    Optional HORINIOPE_ As Date) As Boolean
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstManObr As New ADODB.Recordset
    Dim RstMatPri As New ADODB.Recordset
    Dim RstGasFab As New ADODB.Recordset
    Dim RSTCTAGASFAB As New ADODB.Recordset
    Dim xId As Integer
    Dim xIdDet As Integer
    
On Error GoTo LaCague
    xCon.BeginTrans

    If IDLIBRO_ = 0 Then
        ' Obtenemos el Id del registro
        xId = HallaCodigoTabla("con_librocosto", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM con_librocosto", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = IDLIBRO_
        RST_Busq RstCab, "SELECT * FROM con_librocosto WHERE id=" & xId, xCon
        ' ELIMINAMOS DETALLES
        xCon.Execute "DELETE * FROM con_librocostogasfab WHERE idlibro=" & xId
        xCon.Execute "DELETE * FROM con_librocostomanobr WHERE idlibro=" & xId
        xCon.Execute "DELETE * FROM con_librocostomatpri WHERE idlibro=" & xId
        xCon.Execute "DELETE * FROM con_librocostocta WHERE idlibro=" & xId
        xCon.Execute "DELETE * FROM con_librocostodet WHERE idlibro=" & xId
    End If
    
    RST_Busq RstGasFab, "SELECT TOP 1 * FROM con_librocostogasfab", xCon
    RST_Busq RstManObr, "SELECT TOP 1 * FROM con_librocostomanobr", xCon
    RST_Busq RstMatPri, "SELECT TOP 1 * FROM con_librocostomatpri", xCon
    RST_Busq RSTCTAGASFAB, "SELECT TOP 1 * FROM con_librocostocta", xCon
    RST_Busq RstDet, "SELECT TOP 1 * FROM con_librocostodet", xCon
        
    ' ---------------------------------------CABECERA
    RstCab("descripcion") = DESCRIPCION_
    RstCab("idmes") = MES_
    RstCab("idmetodoval") = IDMETODOVAL_
    RstCab("aplvtas") = APLVTAS_
    RstCab("tipo") = TIPO_
    RstCab.Update
        
        
    PgBar.Min = 0
    PgBar.Max = RSTDET_.RecordCount
    PgBar.Value = 0
    lbl(0).Caption = "GRABANDO ESPERE POR FAVOR"
    LblProg.Caption = ""
    lbl(2).Caption = ""
    CentrarFrm FraProgreso
    FraProgreso.Visible = True
        
    ' --------------------------------------CUENTAS GASTOS DE FABRICA
    If RSTCTAGASFAB_.State = 0 Then grabarLibCosPro = False: Exit Function
    If RSTCTAGASFAB_.RecordCount > 0 Then
        RSTCTAGASFAB_.MoveFirst
        While Not RSTCTAGASFAB_.EOF
            RSTCTAGASFAB.AddNew
            RSTCTAGASFAB("idlibro") = xId
            RSTCTAGASFAB("idcuenta") = NulosN(RSTCTAGASFAB_("idcuenta"))
            RSTCTAGASFAB("importe") = NulosN(RSTCTAGASFAB_("importe"))
            RSTCTAGASFAB.Update
            RSTCTAGASFAB_.MoveNext
        Wend
    End If
    ' ---------------------------------------DETALLE
    If RSTDET_.State = 0 Then grabarLibCosPro = False: Exit Function
    If RSTDET_.RecordCount = 0 Then grabarLibCosPro = False: Exit Function
    xIdDet = HallaCodigoTabla("con_librocostodet", xCon, "id")
    RSTDET_.MoveFirst
    While Not RSTDET_.EOF
        DoEvents
        PgBar.Value = PgBar.Value + 1
            
        RstDet.AddNew
        RstDet("id") = xIdDet
        RstDet("idlibro") = xId
        RstDet("iditem") = NulosN(RSTDET_("iditem"))
        RstDet("idprod") = NulosN(RSTDET_("idprod"))
        RstDet("proceso") = NulosN(RSTDET_("proceso"))
        RstDet("impmprima") = NulosN(RSTDET_("impmprima"))
        RstDet("impmanobr") = NulosN(RSTDET_("impmanobr"))
        RstDet("impgasfab") = RSTDET_("impgasfab")
        RstDet("cantidad") = RSTDET_("cantidad")
        RstDet.Update
        
AGREGARMATERIAPRIMA_:
        ' --------------------------------------MATERIA PRIMA
        If RSTMATPRI_.State = 0 Then grabarLibCosPro = False: Exit Function
        RSTMATPRI_.Filter = "idlibrodet=" & RSTDET_("id")
        If RSTMATPRI_.RecordCount > 0 Then
            RSTMATPRI_.MoveFirst
            While Not RSTMATPRI_.EOF
                RstMatPri.AddNew
                RstMatPri("idlibrodet") = xIdDet
                RstMatPri("idlibro") = xId
                RstMatPri("iditem") = NulosN(RSTMATPRI_("iditem"))
                RstMatPri("cantidad") = NulosN(RSTMATPRI_("cantidad"))
                RstMatPri("impmatpri") = NulosN(RSTMATPRI_("impmatpri"))
                RstMatPri.Update
                RSTMATPRI_.MoveNext
            Wend
        End If
AGREGARMANODEOBRA_:
        ' -------------------------------------MANO DE OBRA
        If RSTMANOBR_.State = 0 Then grabarLibCosPro = False: Exit Function
        RSTMANOBR_.Filter = "idlibrodet=" & RSTDET_("id")
        If RSTMANOBR_.RecordCount > 0 Then
            RSTMANOBR_.MoveFirst
            While Not RSTMANOBR_.EOF
                RstManObr.AddNew
                RstManObr("idlibrodet") = xIdDet
                RstManObr("idlibro") = xId
                RstManObr("idemp") = NulosN(RSTMANOBR_("idemp"))
                RstManObr("impmanobr") = NulosN(RSTMANOBR_("impmanobr"))
                RstManObr.Update
                RSTMANOBR_.MoveNext
            Wend
        End If
AGREGARGASTOSDEFABRICA_:
        ' -------------------------------------GASTOS DE FABRICA
        If RSTGASFAB_.State = 0 Then grabarLibCosPro = False: Exit Function
        RSTGASFAB_.Filter = "idlibrodet=" & RSTDET_("id")
        If RSTGASFAB_.RecordCount > 0 Then
            RSTGASFAB_.MoveFirst
            While Not RSTGASFAB_.EOF
                RstGasFab.AddNew
                RstGasFab("idlibrodet") = xIdDet
                RstGasFab("idlibro") = xId
                RstGasFab("idper") = NulosN(RSTGASFAB_("idper"))
                RstGasFab.Update
                RSTGASFAB_.MoveNext
            Wend
        End If
        
        RSTDET_.MoveNext
        xIdDet = xIdDet + 1
    Wend
TERMINAR_:
    IDLIBRO_ = xId
    ' Grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IDFORM_, QUEHACE_, HORINIOPE_, Time, Date, xCon, CDbl(xId)
    FraProgreso.Visible = False
   
    xCon.CommitTrans
    MsgBox "La operación se registró con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RSTCTAGASFAB = Nothing
    Set RstGasFab = Nothing
    Set RstManObr = Nothing
    Set RstMatPri = Nothing
    grabarLibCosPro = True
    Exit Function
    
LaCague:
    'Resume
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RSTCTAGASFAB = Nothing
    Set RstMatPri = Nothing
    Set RstManObr = Nothing
    Set RstGasFab = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    grabarLibCosPro = False
End Function

Sub Modificar()
    If RstLibro.RecordCount = 0 Then
        MsgBox "No hay Registros para Modificar", vbInformation, xTitulo
        Exit Sub
    End If
   
    QueHace = 2
    xHorIni = Time
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    
    Label5.Caption = "Modificando Libro de Costo de Producción"
    fg(0).Editable = flexEDKbdMouse
    fg(0).SelectionMode = flexSelectionFree
    
    xHorIni = Time
    cbMes.SetFocus
End Sub

Sub Eliminar()
'    Dim Rpta As Integer
'    Dim xRs As New ADODB.Recordset
'
'    If RstLibro.RecordCount = 0 Then
'        MsgBox "No hay documentos para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Exit Sub
'    End If
'
'    TabOne1.CurrTab = 0
'    Rpta = MsgBox("¿ Esta seguro de eliminar el Registro seleccionado ?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
'
'    If Rpta = vbYes Then
'        xCon.Execute "DELETE * FROM pro_ordenprodreproc WHERE idord = " & NulosN(RstLibro("id"))
'        xCon.Execute "DELETE * FROM pro_ordenprodpers WHERE idord = " & NulosN(RstLibro("id"))
'        xCon.Execute "DELETE * FROM pro_ordenprodtar WHERE idord = " & NulosN(RstLibro("id"))
'        xCon.Execute "DELETE * FROM pro_ordenprod WHERE id = " & NulosN(RstLibro("id"))
'
'        'Eliminar historial del registro
'        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & NulosN(RstLibro("id")) & " AND idform = " & IdMenuActivo
'
'        MsgBox "El registro se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        RstLibro.Requery
'        Dg1.Refresh
'    End If
End Sub

Private Sub limpiarRST(Rst As ADODB.Recordset, Optional TODO As Boolean = True)
    With Rst
        If .State <> 0 Then
            If TODO Then .Filter = adFilterNone
            If .RecordCount <> 0 Then
                .MoveFirst
                While Not .EOF
                    .Delete
                    .MoveNext
                Wend
            End If
        End If
    End With
End Sub

Private Sub PbCerrar_Click(Index As Integer)
    Frm4.Visible = False
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 1 Then Exit Sub
        MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then
        If RstLibro.RecordCount = 0 Then
            MsgBox "No se han registardos ventas para realizar esta opcion", vbInformation, Me.Caption
            Exit Sub
        End If
        Modificar
    End If
    
    If Button.Index = 3 Then
        If RstLibro.RecordCount = 0 Then
            MsgBox "No se han registrados Pedidos para realizar esta opción", vbInformation, Me.Caption
            Exit Sub
        End If
        Eliminar
    End If
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstLibro.Requery
            Dg1.Refresh
            If RstLibro.RecordCount <> 0 Then
                RstLibro.MoveFirst
                RstLibro.Find "id=" & mIdRegistro
                If RstLibro.EOF = True Then RstLibro.MoveFirst
            End If
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then
        If TabOne1.CurrTab = 0 Then RstLibro.Filter = "": TDB_FiltroLimpiar Dg1
    End If
        
    If Button.Index = 14 Then ExportarExcel fg(0)
    
    If Button.Index = 17 Then Unload Me
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.Index = 2 Then
        If ButtonMenu.Index = 1 Then ' ANULAR REGISTRO
            If TabOne1.CurrTab = 1 Then TabOne1.CurrTab = 0
            Anular
        End If
    End If
End Sub

Sub ExportarExcel(ByRef GRID_ As VSFlexGrid)
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios
    Dim TITULO_ As String
    
    TITULO_ = "REPORTE DE PRODUCCIÓN"
    
    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, GRID_, TITULO_, "", "", TITULO_
    Set X_EXPORT = Nothing
    MousePointer = vbDefault
    Exit Sub
error:
    MousePointer = vbDefault
    SHOW_ERROR Name, "Exportar"
End Sub

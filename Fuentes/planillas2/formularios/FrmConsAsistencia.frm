VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsAsistencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planilas - Consulta de Asistencia"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   30
      TabIndex        =   9
      Top             =   345
      Width           =   11790
      Begin VB.CommandButton cmd_periodo 
         Height          =   255
         Index           =   0
         Left            =   2340
         Picture         =   "FrmConsAsistencia.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   210
         Width           =   285
      End
      Begin VB.CommandButton cb 
         Height          =   225
         Index           =   0
         Left            =   4440
         Picture         =   "FrmConsAsistencia.frx":0382
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Seleccione el Personal"
         Top             =   210
         Width           =   210
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   0
         Left            =   3900
         MaxLength       =   20
         TabIndex        =   12
         Text            =   "txt_cb(0)"
         Top             =   180
         Width           =   780
      End
      Begin VB.Label lbl_periodo 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_periodo(0)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   0
         Left            =   960
         TabIndex        =   16
         Top             =   180
         Width           =   1710
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   15
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lbl_capt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Personal"
         Height          =   195
         Index           =   0
         Left            =   3015
         TabIndex        =   14
         Top             =   270
         Width           =   615
      End
      Begin VB.Label lbl_cod 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cod(0)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   0
         Left            =   8100
         TabIndex        =   13
         Top             =   195
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   2895
         X2              =   2895
         Y1              =   135
         Y2              =   585
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         X1              =   2895
         X2              =   2895
         Y1              =   135
         Y2              =   600
      End
      Begin VB.Label lbl_cb 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cb(0)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   4680
         TabIndex        =   17
         Top             =   180
         Width           =   4485
      End
   End
   Begin VB.Frame fra_barra 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   795
      Left            =   2775
      TabIndex        =   1
      Top             =   3195
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar barra 
         Height          =   285
         Left            =   105
         TabIndex        =   2
         Top             =   330
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblbarra 
         Caption         =   "Interrumpir = ESC"
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
         Height          =   255
         Index           =   1
         Left            =   4365
         TabIndex        =   4
         Top             =   90
         Width           =   1530
      End
      Begin VB.Label lblbarra 
         AutoSize        =   -1  'True
         Caption         =   "Procesando Asistencia"
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
         Left            =   165
         TabIndex        =   3
         Top             =   90
         Width           =   1905
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   960
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   5925
         X2              =   5925
         Y1              =   -15
         Y2              =   915
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   -30
         X2              =   5940
         Y1              =   15
         Y2              =   30
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   5910
         Y1              =   780
         Y2              =   765
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   0
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
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4830
         Top             =   135
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483637
         ImageWidth      =   16
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsAsistencia.frx":04B4
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsAsistencia.frx":09F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsAsistencia.frx":0D8A
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsAsistencia.frx":0F0E
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsAsistencia.frx":1362
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsAsistencia.frx":147A
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsAsistencia.frx":19BE
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsAsistencia.frx":1F02
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsAsistencia.frx":2016
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsAsistencia.frx":212A
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsAsistencia.frx":257E
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsAsistencia.frx":26EA
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsAsistencia.frx":2C32
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsAsistencia.frx":2F4C
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6525
      Left            =   -15
      TabIndex        =   5
      Top             =   1005
      Width           =   11910
      _cx             =   21008
      _cy             =   11509
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
      FrontTabForeColor=   -2147483630
      Caption         =   "      Detalle    |     Resumen     "
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   0
      Position        =   1
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
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6105
         Left            =   45
         TabIndex        =   7
         Top             =   45
         Width           =   11820
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   6045
            Left            =   45
            TabIndex        =   22
            Top             =   15
            Width           =   11715
            _cx             =   20664
            _cy             =   10663
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   0
            MousePointer    =   0
            _ConvInfo       =   1
            Version         =   700
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            FrontTabColor   =   -2147483633
            BackTabColor    =   -2147483633
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483630
            Caption         =   "  Por Registro  |Por Tipo de Horas|  Por Personal  "
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   0
            Position        =   2
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
            Begin VB.Frame Frame9 
               BorderStyle     =   0  'None
               Height          =   6015
               Left            =   12945
               TabIndex        =   27
               Top             =   15
               Width           =   11370
               Begin VSFlex7Ctl.VSFlexGrid Fg 
                  Height          =   5745
                  Index           =   2
                  Left            =   105
                  TabIndex        =   28
                  Top             =   135
                  Width           =   11190
                  _cx             =   19738
                  _cy             =   10134
                  _ConvInfo       =   1
                  Appearance      =   2
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
                  ForeColorSel    =   16777215
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
                  Rows            =   1
                  Cols            =   1
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsAsistencia.frx":32DE
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
               BorderStyle     =   0  'None
               Height          =   6015
               Left            =   330
               TabIndex        =   25
               Top             =   15
               Width           =   11370
               Begin VSFlex7Ctl.VSFlexGrid Fg 
                  Height          =   5745
                  Index           =   0
                  Left            =   105
                  TabIndex        =   26
                  Top             =   135
                  Width           =   11190
                  _cx             =   19738
                  _cy             =   10134
                  _ConvInfo       =   1
                  Appearance      =   2
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
                  ForeColorSel    =   16777215
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
                  Rows            =   1
                  Cols            =   1
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsAsistencia.frx":33C3
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
               BorderStyle     =   0  'None
               Height          =   6015
               Left            =   12645
               TabIndex        =   23
               Top             =   15
               Width           =   11370
               Begin VSFlex7Ctl.VSFlexGrid Fg 
                  Height          =   5745
                  Index           =   1
                  Left            =   105
                  TabIndex        =   24
                  Top             =   135
                  Width           =   11190
                  _cx             =   19738
                  _cy             =   10134
                  _ConvInfo       =   1
                  Appearance      =   2
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
                  ForeColorSel    =   16777215
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
                  Rows            =   1
                  Cols            =   1
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsAsistencia.frx":34DD
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
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   6105
         Left            =   12555
         TabIndex        =   6
         Top             =   45
         Width           =   11820
         Begin SizerOneLibCtl.TabOne TabOne3 
            Height          =   6045
            Left            =   45
            TabIndex        =   8
            Top             =   15
            Width           =   11715
            _cx             =   20664
            _cy             =   10663
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   0
            MousePointer    =   0
            _ConvInfo       =   1
            Version         =   700
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            FrontTabColor   =   -2147483633
            BackTabColor    =   -2147483633
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483630
            Caption         =   "   Por Dia   | Por Personal "
            Align           =   0
            CurrTab         =   1
            FirstTab        =   0
            Style           =   0
            Position        =   2
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
            Begin VB.Frame Frame6 
               BorderStyle     =   0  'None
               Height          =   6015
               Left            =   330
               TabIndex        =   19
               Top             =   15
               Width           =   11370
               Begin VSFlex7Ctl.VSFlexGrid Fg 
                  Height          =   5745
                  Index           =   4
                  Left            =   105
                  TabIndex        =   21
                  Top             =   135
                  Width           =   11190
                  _cx             =   19738
                  _cy             =   10134
                  _ConvInfo       =   1
                  Appearance      =   2
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
                  ForeColorSel    =   16777215
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
                  Rows            =   1
                  Cols            =   1
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsAsistencia.frx":35C2
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
            Begin VB.Frame Frame4 
               BorderStyle     =   0  'None
               Height          =   6015
               Left            =   -11985
               TabIndex        =   18
               Top             =   15
               Width           =   11370
               Begin VSFlex7Ctl.VSFlexGrid Fg 
                  Height          =   5745
                  Index           =   3
                  Left            =   105
                  TabIndex        =   20
                  Top             =   135
                  Width           =   11190
                  _cx             =   19738
                  _cy             =   10134
                  _ConvInfo       =   1
                  Appearance      =   2
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
                  ForeColorSel    =   16777215
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
                  Rows            =   1
                  Cols            =   1
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsAsistencia.frx":36A7
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
      End
   End
End
Attribute VB_Name = "FrmConsAsistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SeEjecuto As Boolean
Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE

Dim xAcumulado(2, 1) As Double
'--xAcumulado(0,?):: Acumulado por Asiento  ?::0=debe sol; 1::haber sol; 2::debe dol;  3::haber dol
'--xAcumulado(1,?):: Acumulado por libro
'--xAcumulado(2,?):: Acumulado general

Dim nSQLPivot As String
Dim ArrCampos()
Dim ArrAcumular()

Dim mMesActivo As Integer '--indica el mes activo

Private Sub cmd_periodo_Click(Index As Integer)
    mMesActivo = SeleccionaMes(xCon)
    lbl_periodo(0).Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
End Sub

Private Sub pExportar()
    If AnoTra = "" Then
        MsgBox "No Hay Año de trabajo" + vbCr + "No puede continuar", vbExclamation, xTitulo
        Exit Sub
    End If

    If mMesActivo = 0 Then
        MsgBox "Seleccione el Periodo de Consulta", vbExclamation, xTitulo
        cmd_periodo(0).SetFocus
        Exit Sub
    End If
    On Error GoTo error
    
    Dim oExport As New SGI2_funciones.formularios
    Dim mIndex As Integer
    Dim nTitulo As String
    Dim nPeriodo As String
    Dim nTitulo1 As String
    If TabOne1.CurrTab = 0 Then
        If TabOne2.CurrTab = 0 Then '--detalle por marcacion
            mIndex = 0
            nTitulo = "Consulta de Marcación de Asistencia"
            
        ElseIf TabOne2.CurrTab = 1 Then '--detalle por tipo de hora
            mIndex = 1
            nTitulo = "Consulta de Marcación de Asistencia - Tipos de Horas"
        Else '--detalle por personal
            mIndex = 2
            nTitulo = "Consulta de Marcación de Asistencia - Personal"
        End If
    Else
        If TabOne3.CurrTab = 0 Then '--resumen - dia
            mIndex = 3
            nTitulo = "Consulta de Marcación de Asistencia - Resumen por Dia"
        Else '--resumen - personal
            mIndex = 4
            nTitulo = "Consulta de Marcación de Asistencia - Resumen por Personal"
        End If
    End If
    
    nPeriodo = "Periodo: " + lbl_periodo(0).Caption
    If NulosN(lbl_cod(0).Caption) <> 0 Then
         nTitulo1 = "Personal: " & StrConv(lbl_cb(0).Caption, 3)
    End If

    Me.MousePointer = vbHourglass
    oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg(mIndex), nTitulo, nPeriodo, nTitulo1, "Consulta de Asistencia"
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Exportar"
End Sub


Private Sub pImprimir()
    If AnoTra = "" Then
        MsgBox "No Hay Año de trabajo" + vbCr + "No puede continuar", vbExclamation, xTitulo
        Exit Sub
    End If

    If mMesActivo = 0 Then
        MsgBox "Seleccione el Periodo de Consulta", vbExclamation, xTitulo
        cmd_periodo(0).SetFocus
        Exit Sub
    End If
    
    On Error GoTo error

    Dim oPrint  As New SGI2_funciones.formularios
    Dim mIndex As Integer
    Dim nTitulo As String
    Dim nPeriodo As String
    Dim nTitulo1 As String
    If TabOne1.CurrTab = 0 Then
        If TabOne2.CurrTab = 0 Then '--detalle por marcacion
            mIndex = 0
            nTitulo = "Consulta de Marcación de Asistencia"
            
        ElseIf TabOne2.CurrTab = 1 Then '--detalle por tipo de hora
            mIndex = 1
            nTitulo = "Consulta de Marcación de Asistencia - Tipos de Horas"
        Else '--detalle por personal
            mIndex = 2
            nTitulo = "Consulta de Marcación de Asistencia - Personal"
        End If
    Else
        If TabOne3.CurrTab = 0 Then '--resumen - dia
            mIndex = 3
            nTitulo = "Consulta de Marcación de Asistencia - Resumen por Dia"
        Else '--resumen - personal
            mIndex = 4
            nTitulo = "Consulta de Marcación de Asistencia - Resumen por Personal"
        End If
    End If
    
    nPeriodo = "Periodo: " + lbl_periodo(0).Caption
    If NulosN(lbl_cod(0).Caption) <> 0 Then
        nTitulo1 = "Personal: " & StrConv(lbl_cb(0).Caption, 3)
    End If
    
    Me.MousePointer = vbHourglass
    oPrint.Imprimir_x_VSFlexGrid Fg(mIndex), nTitulo, nTitulo1, nPeriodo, False, True
    Set oPrint = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Set oPrint = Nothing
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Exportar"
End Sub

Private Sub pConsultar()
    ''''''''''''
    If AnoTra = "" Then
        MsgBox "No Hay Año de trabajo" + vbCr + "No puede continuar", vbExclamation, xTitulo
        Exit Sub
    End If

    If mMesActivo = 0 Then
        MsgBox "Seleccione el Periodo de Consulta", vbExclamation, xTitulo
        cmd_periodo(0).SetFocus
        Exit Sub
    End If
    
    '''''''''''
    Erase xAcumulado()

    '''''''''''
    BAND_INTERRUMPIR = False
    pConfigurarGrilla
    '----
    fra_barra.Visible = True
    fra_barra.Top = 3195
    fra_barra.Left = 2775
    '----
    Me.TabOne1.CurrTab = 0
    BAND_INTERRUMPIR = False
    pCargarDetalleMarcacion
    If BAND_INTERRUMPIR = True Then GoTo salir:
    pCargarDetalleTipoHora
    If BAND_INTERRUMPIR = True Then GoTo salir:
    pCargarDetallePersonal
    If BAND_INTERRUMPIR = True Then GoTo salir:
    pCargarResumenDia
    If BAND_INTERRUMPIR = True Then GoTo salir:
    pCargarResumenPersonal
    If BAND_INTERRUMPIR = True Then GoTo salir:
    
    Erase xAcumulado()
    '--SI SE NTERRUMPE EL PROCESO => SALIR
    If BAND_INTERRUMPIR = True Then GoTo salir:
    '-----------------------------------------------
salir:
    fra_barra.Visible = False
    Erase xAcumulado()
    If BAND_INTERRUMPIR = True Then
        MsgBox "La consulta fue interrumpida", vbInformation, xTitulo
    End If
        
End Sub


Private Sub Form_Activate()
    If SeEjecuto = False Then
        mMesActivo = xMes
        
        SeEjecuto = True
        txt_cb(0).Text = ""
        LimpiaText lbl_periodo
        lbl_periodo(0).Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
        TabOne1.CurrTab = 0
        pConfigurarGrilla
        cmd_periodo(0).SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        BAND_INTERRUMPIR = True '--interrumpir
    End If
End Sub

Private Sub Form_Load()
    CentrarFrm Me
    SeEjecuto = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    If Button.Index = 3 Then pExportar
    If Button.Index = 4 Then pImprimir
    If Button.Index = 6 Then
        Unload Me
    End If
End Sub

Private Sub pConfigurarGrilla()
    Dim nSQL As String
    Dim mColHoras&
    Dim mColInicio&, mRowArr&
    Dim RstTipoHoras As New ADODB.Recordset

    '-------------------------------------------------
    nSQL = "SELECT mae_tipohora.id, mae_tipohora.descripcion, mae_tipohora.prioridad, mae_tipohora.nomcor " _
        + vbCr + " FROM pla_marcacion INNER JOIN (mae_tipohora INNER JOIN pla_marcacionhora ON mae_tipohora.id = pla_marcacionhora.idhora) ON pla_marcacion.id = pla_marcacionhora.idmarca " _
        + vbCr + " Where (((Month([pla_marcacion].[dia])) = " & mMesActivo & ") And ((Year([pla_marcacion].[dia])) = " & AnoTra & ")) " _
        + vbCr + " GROUP BY mae_tipohora.id, mae_tipohora.descripcion, mae_tipohora.prioridad, mae_tipohora.nomcor " _
        + vbCr + " HAVING (((mae_tipohora.nomcor) Is Not Null)) " _
        + vbCr + " ORDER BY mae_tipohora.prioridad;"
    
    RST_Busq RstTipoHoras, nSQL, xCon
    
    mColHoras = RstTipoHoras.RecordCount
    If mColHoras <> 0 Then RstTipoHoras.MoveFirst
    nSQLPivot = ""
    Erase ArrCampos
    Erase ArrAcumular
    
    ReDim ArrCampos(RstTipoHoras.RecordCount)
    ReDim ArrAcumular(RstTipoHoras.RecordCount + 1)
    
    Do While Not RstTipoHoras.EOF
    
        nSQLPivot = nSQLPivot + "'" + RstTipoHoras.Fields("nomcor") + "',"
        ArrCampos(RstTipoHoras.Bookmark - 1) = Replace(RstTipoHoras.Fields("nomcor"), ".", "_")
        
        RstTipoHoras.MoveNext
    Loop
    If nSQLPivot <> "" Then nSQLPivot = " IN (" + Left(nSQLPivot, Len(nSQLPivot) - 1) + ") "
    Set RstTipoHoras = Nothing
   
    '-------------------------------------------------
    
    With Fg(0) '--detalle - marcacion
        '-----
        .Rows = 1
        .FrozenCols = 0
        .Cols = 8
        .RowHeight(0) = 300
        .ColWidth(0) = 200
        '--DATOS DE FILA
        .TextMatrix(0, 1) = "IdEmp":                .ColWidth(1) = 0:
        .TextMatrix(0, 2) = "Dia":                  .ColWidth(2) = 900:   .ColAlignment(2) = flexAlignCenterCenter:     .Row = 0: .Col = 2: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 3) = "Nombre Dia":           .ColWidth(3) = 1000:  .ColAlignment(3) = flexAlignLeftCenter:       .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 4) = "Apellidos y Nombres":  .ColWidth(4) = 3500:  .ColAlignment(4) = flexAlignLeftCenter:       .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 5) = "Origen":               .ColWidth(5) = 2000:  .ColAlignment(5) = flexAlignLeftCenter:       .Row = 0: .Col = 5: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 6) = "Hora Ingreso":          .ColWidth(6) = 1200:  .ColAlignment(6) = flexAlignCenterCenter:     .Row = 0: .Col = 6: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 7) = "Hora Salida":           .ColWidth(7) = 1200:  .ColAlignment(7) = flexAlignCenterCenter:     .Row = 0: .Col = 7: .CellAlignment = flexAlignCenterCenter
        
        '--muestra columna
        If NulosN(lbl_cod(0).Caption) <> 0 Then .ColWidth(4) = 0
        '-----------------------------------------------------------
        
        .ColFormat(1) = FORMAT_DATE

        .ColFormat(6) = FORMAT_HORA_AL_SEGUNDO
        .ColFormat(7) = FORMAT_HORA_AL_SEGUNDO
        
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDNone
    End With
    
    With Fg(1) '--detalle -tipo hora
        '-----
        .Rows = 1
        .FrozenCols = 0
        .Cols = 7
        .RowHeight(0) = 300
        .ColWidth(0) = 200
        '--DATOS DE FILA
        .TextMatrix(0, 1) = "IdEmp":                  .ColWidth(1) = 0:
        .TextMatrix(0, 2) = "Dia":                  .ColWidth(2) = 900:   .ColAlignment(2) = flexAlignCenterCenter:     .Row = 0: .Col = 2: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 3) = "Nombre Dia":           .ColWidth(3) = 1000:   .ColAlignment(3) = flexAlignLeftCenter:       .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 4) = "Apellidos y Nombres":  .ColWidth(4) = 3500:  .ColAlignment(4) = flexAlignLeftCenter:       .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 5) = "Tipo de Hora":         .ColWidth(5) = 2000:  .ColAlignment(5) = flexAlignLeftCenter:       .Row = 0: .Col = 5: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 6) = "Total":                .ColWidth(6) = 1200:  .ColAlignment(6) = flexAlignCenterCenter:     .Row = 0: .Col = 6: .CellAlignment = flexAlignCenterCenter
        '--muestra columna
        If NulosN(lbl_cod(0).Caption) <> 0 Then .ColWidth(4) = 0
        '-----------------------------------------------------------

        .ColFormat(1) = FORMAT_DATE
        .ColFormat(6) = FORMAT_HORA_LARGO
        
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDNone
    End With
    
    With Fg(2) '--Detalle por personal
        .Rows = 2
        .FixedRows = 2
        .Cols = 5 + mColHoras
        .FrozenCols = 3
        .RowHeight(0) = 300
        .ColWidth(0) = 200
        '--------------
        UNIR_CELDAS Fg(2), 0, 1, 0, 3, " "
        UNIR_CELDAS Fg(2), 0, 4, 0, Fg(2).Cols - 2, "Tipos de Horas"
        UNIR_CELDAS Fg(2), 0, Fg(2).Cols - 1, 0, Fg(2).Cols - 1, "Total Horas"
        '--------------
        .TextMatrix(1, 1) = "Dia":                  .ColWidth(1) = 800:   .ColAlignment(1) = flexAlignCenterCenter:  .Row = 1: .Col = 1: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 2) = "Nombre Dia":           .ColWidth(2) = 1000:  .ColAlignment(2) = flexAlignLeftCenter:    .Row = 1: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 3) = "Apellidos y Nombres":  .ColWidth(3) = 3400:  .ColAlignment(3) = flexAlignLeftCenter:    .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftCenter
        '----------------------------------------
        '--encabezado de los tipos de horas
        mColInicio = 4
        For mRowArr = 0 To UBound(ArrCampos())
            .TextMatrix(1, mColInicio) = ArrCampos(mRowArr):   .ColWidth(mColInicio) = 900:  .ColAlignment(mColInicio) = flexAlignRightCenter:     .Row = 1: .Col = mColInicio: .CellAlignment = flexAlignCenterCenter
            .ColFormat(mColInicio) = FORMAT_HORA_LARGO
            mColInicio = mColInicio + 1
        Next
        '--total
        .TextMatrix(1, mColHoras + 4) = "Total Horas":       .ColWidth(mColHoras + 4) = 1100:  .ColAlignment(mColHoras + 4) = flexAlignCenterCenter:     .Row = 1: .Col = mColHoras + 4: .CellAlignment = flexAlignRightCenter
        '----------------------------------------
       '--muestra columna
        If NulosN(lbl_cod(0).Caption) <> 0 Then .ColWidth(3) = 0
        '-----------------------------------------------------------
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDNone
    End With
    
    With Fg(3) '--resumen por dia
        .Rows = 2
        .FixedRows = 2
        .Cols = 4 + mColHoras
        .FrozenCols = 2
        .RowHeight(0) = 300
        .ColWidth(0) = 200
        '--------------
        UNIR_CELDAS Fg(3), 0, 1, 0, 2, " "
        UNIR_CELDAS Fg(3), 0, 3, 0, Fg(3).Cols - 2, "Tipos de Horas"
        UNIR_CELDAS Fg(3), 0, Fg(3).Cols - 1, 0, Fg(3).Cols - 1, "Total Horas"
        '--------------
        .TextMatrix(1, 1) = "Dia":               .ColWidth(1) = 900:   .ColAlignment(1) = flexAlignCenterCenter:  .Row = 1: .Col = 1: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 2) = "Nombre Dia":        .ColWidth(2) = 1000:  .ColAlignment(2) = flexAlignLeftCenter:    .Row = 1: .Col = 2: .CellAlignment = flexAlignLeftCenter
        '----------------------------------------
        '--encabezado de los tipos de horas
        mColInicio = 3
        For mRowArr = 0 To UBound(ArrCampos())
            .TextMatrix(1, mColInicio) = ArrCampos(mRowArr):  .ColWidth(mColInicio) = 900:  .ColAlignment(mColInicio) = flexAlignRightCenter:     .Row = 1: .Col = mColInicio: .CellAlignment = flexAlignCenterCenter
            .ColFormat(mColInicio) = FORMAT_HORA_LARGO
            mColInicio = mColInicio + 1
        Next
        '--total
        .TextMatrix(1, mColHoras + 3) = "Total Horas":        .ColWidth(mColHoras + 3) = 1100:  .ColAlignment(mColHoras + 3) = flexAlignCenterCenter:     .Row = 1: .Col = mColHoras + 3: .CellAlignment = flexAlignRightCenter
        '----------------------------------------
                
        .ColFormat(1) = FORMAT_DATE
        
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDNone
    End With
    
    With Fg(4) '--resumen por personal
        .Rows = 2
        .FixedRows = 2
        .Cols = 3 + mColHoras
        .FrozenCols = 1
        .RowHeight(0) = 300
        .ColWidth(0) = 200
        '--------------
        UNIR_CELDAS Fg(4), 0, 1, 0, 1, " "
        UNIR_CELDAS Fg(4), 0, 2, 0, Fg(4).Cols - 2, "Tipos de Horas"
        UNIR_CELDAS Fg(4), 0, Fg(4).Cols - 1, 0, Fg(4).Cols - 1, "Total Horas"
        '--------------
        .TextMatrix(1, 1) = "Apellidos y Nombres":   .ColWidth(1) = 3500:   .ColAlignment(1) = flexAlignLeftCenter: .Row = 1: .Col = 1: .CellAlignment = flexAlignLeftCenter
        '----------------------------------------
        '--encabezado de los tipos de horas
        mColInicio = 2
        For mRowArr = 0 To UBound(ArrCampos())
            .TextMatrix(1, mColInicio) = ArrCampos(mRowArr):  .ColWidth(mColInicio) = 900:  .ColAlignment(mColInicio) = flexAlignRightCenter:     .Row = 1: .Col = mColInicio: .CellAlignment = flexAlignCenterCenter
            .ColFormat(mColInicio) = FORMAT_HORA_LARGO
            mColInicio = mColInicio + 1
        Next
        '--total
        .TextMatrix(1, mColHoras + 2) = "Total Horas":        .ColWidth(mColHoras + 2) = 1100:  .ColAlignment(mColHoras + 2) = flexAlignCenterCenter:     .Row = 1: .Col = mColHoras + 2: .CellAlignment = flexAlignRightCenter
        '----------------------------------------
                
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDNone
    End With
    
    TabOne1.CurrTab = 0
    TabOne2.CurrTab = 0
    
    DoEvents
    
End Sub

'****************************************************************************************
Private Sub cb_Click(Index As Integer)
    On Error GoTo error
    Dim xRs As New ADODB.Recordset
    pBuscarPersonal xRs, False
    If xRs.State = 1 Then
        txt_cb(0) = xRs.Fields("id") & "" '--TEXTO A MOSTRAR
        lbl_cb(0).Caption = xRs.Fields("nombres") & "" '--NOMBRE
        lbl_cod(0).Caption = xRs.Fields("id") & "" '--CODIGO
        lbl_cb(0).ToolTipText = xRs.Fields("nombres") & "" '--NOMBRE
        txt_cb(0).SetFocus
    End If
    Set xRs = Nothing
Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub

Private Sub txt_cb_Change(Index As Integer)
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cod(Index).Caption = ""
    End If
End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If txt_cb(Index).Locked = True Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb_Click Index
        Exit Sub
    End If
End Sub

Private Sub txt_cb_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub txt_cb_Validate(Index As Integer, Cancel As Boolean)
    If txt_cb(Index).Text = "" Then Exit Sub
    
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    Select Case Index
        Case 0 '--TIPO DE TRABAJADOR
            nSQL = "SELECT pla_empleados.id, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombre, pla_empleados.id as cod, pla_empleados.numdoc, mae_dociden.abrev AS tipodoc, mae_sexo.abrev AS sexo, Format([pla_empleados].[fchnac],'dd/mm/yyyy') AS fchnac, pla_empleados.numtel, pla_empleados.email " _
                + vbCr + " FROM mae_sexo RIGHT JOIN (mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) ON mae_sexo.id = pla_empleados.idsex " _
                + vbCr + " WHERE  pla_empleados.id  = " & NulosC(txt_cb(Index).Text) & ";"
    End Select

    If xCon.State = 0 Then GoTo salir
    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then GoTo salir

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    If RstTmp.RecordCount > 0 Then
        txt_cb(Index).Text = RstTmp.Fields(0) & ""  '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = RstTmp.Fields(1) & "" '--NOMBRE
        lbl_cod(Index).Caption = RstTmp.Fields(2) & "" '--CODIGO
        lbl_cb(Index).ToolTipText = RstTmp.Fields(1) & "" '--NOMBRE
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cod(Index).Caption = ""
    End If
    '--------------
    Set RstTmp = Nothing
    Exit Sub
error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "txt_cb_Validate (" + CStr(Index) + ")"
    Exit Sub
salir:
    Set RstTmp = Nothing
    txt_cb(Index).Text = ""
End Sub

'****************************************************************************************
Private Sub pCargarDetalleMarcacion()
    
    Dim nSQL As String
    Dim RstTmp As New ADODB.Recordset
    Dim nSQLIdEmp As String
    
    If NulosN(lbl_cod(0).Caption) <> 0 Then
        nSQLIdEmp = " and pla_empleados.id = " & NulosN(lbl_cod(0).Caption)
    End If
    '----
    lblbarra(0).Caption = "Procesando Detalle por Marcación"
    '----
    nSQL = "SELECT pla_empleados.id AS idemp, pla_marcacion.dia, Format([pla_marcacion].[dia],'dddd') AS nomdia, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, pla_marcaciondet.idori, pla_origenes.descripcion AS origen, pla_marcaciondet.hingreso AS hini, pla_marcaciondet.hsalida AS hfin " _
        + vbCr + " FROM pla_empleados INNER JOIN (pla_marcacion INNER JOIN (pla_marcaciondet LEFT JOIN pla_origenes ON pla_marcaciondet.idori = pla_origenes.id) ON pla_marcacion.id = pla_marcaciondet.idmarca) ON pla_empleados.id = pla_marcaciondet.idemp " _
        + vbCr + " Where (((Month([pla_marcacion].[dia])) = " & mMesActivo & ") And ((Year([pla_marcacion].[dia])) = " & AnoTra & ")) " & nSQLIdEmp _
        + vbCr + " ORDER BY pla_marcacion.dia, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom], pla_marcaciondet.hingreso;"

    TabOne1.CurrTab = 0
    TabOne2.CurrTab = 0
    RST_Busq RstTmp, nSQL, xCon
    '-------------
    barra.Min = 1
    If RstTmp.RecordCount > 1 Then barra.Max = RstTmp.RecordCount
    '-----------
    With Fg(0)
        Do While Not RstTmp.EOF
            DoEvents
            '--
            If BAND_INTERRUMPIR = True Then GoTo salir
            barra.Value = RstTmp.Bookmark
            '--
            .Rows = .Rows + 1
            
            .TextMatrix(.Rows - 1, 1) = NulosN(RstTmp("idemp"))
            .TextMatrix(.Rows - 1, 2) = Format(NulosC(RstTmp("dia")), FORMAT_DATE)
            .TextMatrix(.Rows - 1, 3) = NulosC(RstTmp("nomdia"))
            .TextMatrix(.Rows - 1, 4) = NulosC(RstTmp("nombres"))
            .TextMatrix(.Rows - 1, 5) = NulosC(RstTmp("origen"))
            .TextMatrix(.Rows - 1, 6) = Format(NulosC(RstTmp("hini")), FORMAT_HORA_AL_SEGUNDO)
            .TextMatrix(.Rows - 1, 7) = Format(NulosC(RstTmp("hfin")), FORMAT_HORA_AL_SEGUNDO)
            
            RstTmp.MoveNext
        Loop
    End With
salir:
    Set RstTmp = Nothing
    If nSQLIdEmp = "" Then
        GRID_AGRUPAR Fg(0), 1 '--agrupar por personal
    Else
        GRID_AGRUPAR Fg(0), 2 '--agrupar por dia
    End If
End Sub

Private Sub pCargarDetalleTipoHora()
    
    Dim nSQL As String
    Dim RstTmp As New ADODB.Recordset
    Dim nSQLIdEmp As String
    '----
    lblbarra(0).Caption = "Procesando Detalle por Tipo de Horas"
    '----
    If NulosN(lbl_cod(0).Caption) <> 0 Then
        nSQLIdEmp = " and pla_empleados.id =" & NulosN(lbl_cod(0).Caption)
    End If
    
    nSQL = "SELECT pla_empleados.id AS idemp,  pla_marcacion.dia, Format([pla_marcacion].[dia],'dddd') AS nomdia, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, pla_marcacionhora.idhora, mae_tipohora.descripcion as tipohora, pla_marcacionhora.tothor, pla_marcacionhora.totseg " _
        + vbCr + " FROM mae_tipohora INNER JOIN (pla_empleados INNER JOIN (pla_marcacion INNER JOIN pla_marcacionhora ON pla_marcacion.id = pla_marcacionhora.idmarca) ON pla_empleados.id = pla_marcacionhora.idemp) ON mae_tipohora.id = pla_marcacionhora.idhora " _
        + vbCr + " WHERE (((Month([pla_marcacion].[dia])) = " & mMesActivo & ") And ((Year([pla_marcacion].[dia])) = " & AnoTra & " )) " & nSQLIdEmp _
        + vbCr + " ORDER BY pla_marcacion.dia, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom], mae_tipohora.prioridad;"

    RST_Busq RstTmp, nSQL, xCon
    TabOne1.CurrTab = 0
    TabOne2.CurrTab = 1
    '-------------
    barra.Min = 1
    If RstTmp.RecordCount > 1 Then barra.Max = RstTmp.RecordCount
    '-----------
    With Fg(1)
        Do While Not RstTmp.EOF
            DoEvents
            '--
            If BAND_INTERRUMPIR = True Then GoTo salir
            barra.Value = RstTmp.Bookmark
            '--
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = NulosN(RstTmp("idemp"))
            .TextMatrix(.Rows - 1, 2) = Format(NulosC(RstTmp("dia")), FORMAT_DATE)
            .TextMatrix(.Rows - 1, 3) = NulosC(RstTmp("nomdia"))
            .TextMatrix(.Rows - 1, 4) = NulosC(RstTmp("nombres"))
            .TextMatrix(.Rows - 1, 5) = NulosC(RstTmp("tipohora"))
            .TextMatrix(.Rows - 1, 6) = NulosC(RstTmp("tothor"))
            
            RstTmp.MoveNext
        Loop
    End With
salir:
    Set RstTmp = Nothing
    If nSQLIdEmp = "" Then
        GRID_AGRUPAR Fg(1), 1 '--agrupar por personal
    Else
        GRID_AGRUPAR Fg(1), 2 '--agrupar por dia
    End If
End Sub

Private Sub pCargarDetallePersonal()

    Dim nSQL As String
    Dim RstTmp As New ADODB.Recordset
    Dim nSQLIdEmp As String
    Dim mColInicio&, mRowArr&
    
    '----
    lblbarra(0).Caption = "Procesando Detalle por Personal"
    '----
    If NulosN(lbl_cod(0).Caption) <> 0 Then
        nSQLIdEmp = " and pla_empleados.id =" & NulosN(lbl_cod(0).Caption)
    End If

    nSQL = "TRANSFORM Sum(pla_marcacionhora.totseg) AS SumaDetotseg " _
        + vbCr + " SELECT pla_marcacion.dia, Format([dia],'dddd') AS nomdia, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, Sum(pla_marcacionhora.totseg) AS tothor " _
        + vbCr + " FROM pla_empleados INNER JOIN (pla_marcacion INNER JOIN (mae_tipohora INNER JOIN pla_marcacionhora ON mae_tipohora.id = pla_marcacionhora.idhora) ON pla_marcacion.id = pla_marcacionhora.idmarca) ON pla_empleados.id = pla_marcacionhora.idemp " _
        + vbCr + " WHERE (((Year([dia])) = " & AnoTra & ") And ((Month([dia])) = " & mMesActivo & ")) " & nSQLIdEmp _
        + vbCr + " GROUP BY pla_marcacion.dia, Format([dia],'dddd'), [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] " _
        + vbCr + " ORDER BY pla_marcacion.dia, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] " _
        + vbCr + " PIVOT mae_tipohora.nomcor " & nSQLPivot

    RST_Busq RstTmp, nSQL, xCon
    
    TabOne1.CurrTab = 0
    TabOne2.CurrTab = 2
    '-------------
    barra.Min = 1
    If RstTmp.RecordCount > 1 Then barra.Max = RstTmp.RecordCount

    '-----------
    pLimpiarArr ArrAcumular()
    With Fg(2)
        Do While Not RstTmp.EOF
            DoEvents
            '--
            If BAND_INTERRUMPIR = True Then GoTo salir
            barra.Value = RstTmp.Bookmark
            '--
            .Rows = .Rows + 1
            
            .TextMatrix(.Rows - 1, 1) = Format(NulosC(RstTmp("dia")), FORMAT_DATE)
            .TextMatrix(.Rows - 1, 2) = NulosC(RstTmp("nomdia"))
            .TextMatrix(.Rows - 1, 3) = NulosC(RstTmp("nombres"))
            '----------------------------------------
            '--de los tipos de horas
            mColInicio = 4
            For mRowArr = 0 To UBound(ArrCampos) - 1
                If NulosN(RstTmp(ArrCampos(mRowArr))) <> 0 Then
                    .TextMatrix(.Rows - 1, mColInicio) = ConvertHora(NulosN(RstTmp(ArrCampos(mRowArr))))
                    ArrAcumular(mRowArr) = ArrAcumular(mRowArr) + NulosN(RstTmp(ArrCampos(mRowArr)))
                End If
                mColInicio = mColInicio + 1
            Next
            '--total
            .TextMatrix(.Rows - 1, UBound(ArrCampos) + 4) = ConvertHora(NulosN(RstTmp("tothor")))
            ArrAcumular(UBound(ArrCampos)) = ArrAcumular(UBound(ArrCampos)) + NulosN(RstTmp("tothor"))
            '----------------------------------------
            RstTmp.MoveNext
        Loop
        
        '------agregar los totales
        mColInicio = 4
        .Rows = .Rows + 1
        '--muestra columna
        If NulosN(lbl_cod(0).Caption) <> 0 Then
            .TextMatrix(.Rows - 1, 2) = "Totales"
        Else
            .TextMatrix(.Rows - 1, 3) = "Totales"
        End If
        '-----------------------------------------------------------
        For mRowArr = 0 To UBound(ArrAcumular) - 1
            .TextMatrix(.Rows - 1, mColInicio) = ConvertHora(ArrAcumular(mRowArr))
            mColInicio = mColInicio + 1
        Next
    End With
salir:
    Set RstTmp = Nothing
    If nSQLIdEmp = "" Then
        GRID_AGRUPAR Fg(2), 3 '--agrupar por personal
    Else
        GRID_AGRUPAR Fg(2), 1 '--agrupar por dia
    End If
    If BAND_INTERRUMPIR = False Then
        GRID_COLOR_FONDO Fg(2), Fg(2).Rows - 1, 1, Fg(2).Rows - 1, Fg(2).Cols - 1, &HC0C0FF
        GRID_COLOR_FONDO Fg(2), Fg(2).FixedRows, Fg(2).Cols - 1, Fg(2).Rows - 1, Fg(2).Cols - 1, &HC0C0FF
    End If
    '---------------------------------
End Sub

Private Sub pCargarResumenDia()
    Dim nSQL As String
    Dim RstTmp As New ADODB.Recordset
    Dim nSQLIdEmp As String
    Dim mColInicio&, mRowArr&

    '----
    lblbarra(0).Caption = "Procesando Resumen por Dia"
    '----
    If NulosN(lbl_cod(0).Caption) <> 0 Then
        nSQLIdEmp = " and pla_empleados.id=" & NulosN(lbl_cod(0).Caption)
    End If
        
   
    nSQL = "TRANSFORM Sum(pla_marcacionhora.totseg) AS SumaDetotseg " _
        + vbCr + " SELECT pla_marcacion.dia, Format([dia],'dddd') AS nomdia, Sum(pla_marcacionhora.totseg) AS tothor " _
        + vbCr + " FROM pla_empleados INNER JOIN (pla_marcacion INNER JOIN (mae_tipohora INNER JOIN pla_marcacionhora ON mae_tipohora.id = pla_marcacionhora.idhora) ON pla_marcacion.id = pla_marcacionhora.idmarca) ON pla_empleados.id = pla_marcacionhora.idemp " _
        + vbCr + " Where (((Year([dia])) = " & AnoTra & ") And ((Month([dia])) = " & mMesActivo & "))" & nSQLIdEmp _
        + vbCr + " GROUP BY pla_marcacion.dia, Format([dia],'dddd') " _
        + vbCr + " ORDER BY pla_marcacion.dia " _
        + vbCr + " PIVOT mae_tipohora.nomcor " & nSQLPivot


    RST_Busq RstTmp, nSQL, xCon
    
    TabOne1.CurrTab = 1
    TabOne2.CurrTab = 1
    '-------------
    barra.Min = 1
    If RstTmp.RecordCount > 1 Then barra.Max = RstTmp.RecordCount
    '------
    pLimpiarArr ArrAcumular()
    '-----------
    With Fg(3)
        Do While Not RstTmp.EOF
            DoEvents
            '--
            If BAND_INTERRUMPIR = True Then GoTo salir
            barra.Value = RstTmp.Bookmark
            '--
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = Format(NulosC(RstTmp("dia")), FORMAT_DATE)
            .TextMatrix(.Rows - 1, 2) = NulosC(RstTmp("nomdia"))
            '----------------------------------------
            '--de los tipos de horas
            mColInicio = 3
            For mRowArr = 0 To UBound(ArrCampos) - 1
                If NulosN(RstTmp(ArrCampos(mRowArr))) <> 0 Then
                    .TextMatrix(.Rows - 1, mColInicio) = ConvertHora(NulosN(RstTmp(ArrCampos(mRowArr))))
                    ArrAcumular(mRowArr) = ArrAcumular(mRowArr) + NulosN(RstTmp(ArrCampos(mRowArr)))
                End If
                mColInicio = mColInicio + 1
            Next
            '--total
            .TextMatrix(.Rows - 1, UBound(ArrCampos) + 3) = ConvertHora(NulosN(RstTmp("tothor")))
            ArrAcumular(UBound(ArrCampos)) = ArrAcumular(UBound(ArrCampos)) + NulosN(RstTmp("tothor"))
            '----------------------------------------

            RstTmp.MoveNext
        Loop
        '------agregar los totales
        mColInicio = 3
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 2) = "Totales"
        '-----------------------------------------------------------
        
        For mRowArr = 0 To UBound(ArrAcumular) - 1
            .TextMatrix(.Rows - 1, mColInicio) = ConvertHora(ArrAcumular(mRowArr))
            mColInicio = mColInicio + 1
        Next
        '---------------------------------
        
    End With
salir:
    Set RstTmp = Nothing
    GRID_AGRUPAR Fg(3), 1 '--agrupar por dia
    If BAND_INTERRUMPIR = False Then
        GRID_COLOR_FONDO Fg(3), Fg(3).Rows - 1, 1, Fg(3).Rows - 1, Fg(3).Cols - 1, &HC0C0FF
        GRID_COLOR_FONDO Fg(3), Fg(3).FixedRows, Fg(3).Cols - 1, Fg(3).Rows - 1, Fg(3).Cols - 1, &HC0C0FF
    End If

End Sub

Private Sub pCargarResumenPersonal()

    Dim nSQL As String
    Dim RstTmp As New ADODB.Recordset
    Dim nSQLIdEmp As String
    Dim nSQLPeriodo As String
    Dim mColInicio&, mRowArr&
    
    '----
    lblbarra(0).Caption = "Procesando Resumen por Personal"
    '----
    If NulosN(lbl_cod(0).Caption) <> 0 Then
        nSQLIdEmp = vbCr + " and pla_empleados.id =" & NulosN(lbl_cod(0).Caption)
    End If
    
    nSQL = "TRANSFORM Sum(pla_marcacionhora.totseg) AS SumaDetotseg " _
        + vbCr + " SELECT [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, Sum(pla_marcacionhora.totseg) AS tothor " _
        + vbCr + " FROM pla_empleados INNER JOIN (pla_marcacion INNER JOIN (mae_tipohora INNER JOIN pla_marcacionhora ON mae_tipohora.id = pla_marcacionhora.idhora) ON pla_marcacion.id = pla_marcacionhora.idmarca) ON pla_empleados.id = pla_marcacionhora.idemp " _
        + vbCr + " Where (((Year([dia])) = " & AnoTra & ") And ((Month([dia])) = " & mMesActivo & ")) " & nSQLIdEmp _
        + vbCr + " GROUP BY [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] " _
        + vbCr + " ORDER BY [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] " _
        + vbCr + " PIVOT mae_tipohora.nomcor " & nSQLPivot

    
    RST_Busq RstTmp, nSQL, xCon
    
    TabOne1.CurrTab = 1
    TabOne2.CurrTab = 1
    '-------------
    barra.Min = 1
    If RstTmp.RecordCount > 1 Then
        barra.Max = RstTmp.RecordCount
    End If
    '-----------
    pLimpiarArr ArrAcumular()
    '-----------
    With Fg(4)
        Do While Not RstTmp.EOF
            DoEvents
            '--
            If BAND_INTERRUMPIR = True Then GoTo salir
            barra.Value = RstTmp.Bookmark
            '--
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = NulosC(RstTmp("nombres"))
            '----------------------------------------
            '--de los tipos de horas
            mColInicio = 2
            For mRowArr = 0 To UBound(ArrCampos) - 1
                If NulosN(RstTmp(ArrCampos(mRowArr))) <> 0 Then
                    .TextMatrix(.Rows - 1, mColInicio) = ConvertHora(NulosN(RstTmp(ArrCampos(mRowArr))))
                    ArrAcumular(mRowArr) = ArrAcumular(mRowArr) + NulosN(RstTmp(ArrCampos(mRowArr)))
                End If
                mColInicio = mColInicio + 1
            Next
            '--total
            .TextMatrix(.Rows - 1, UBound(ArrCampos) + 2) = ConvertHora(NulosN(RstTmp("tothor")))
            ArrAcumular(UBound(ArrCampos)) = ArrAcumular(UBound(ArrCampos)) + NulosN(RstTmp("tothor"))
            '----------------------------------------

            RstTmp.MoveNext
        Loop
        '------agregar los totales
        mColInicio = 2
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 1) = "Totales"
        For mRowArr = 0 To UBound(ArrAcumular) - 1
            .TextMatrix(.Rows - 1, mColInicio) = ConvertHora(ArrAcumular(mRowArr))
            mColInicio = mColInicio + 1
        Next
        '---------------------------------
        
    End With
salir:
    Set RstTmp = Nothing
    If nSQLIdEmp = "" Then
        GRID_AGRUPAR Fg(4), 3 '--agrupar por personal
    Else
        GRID_AGRUPAR Fg(4), 1 '--agrupar por dia
    End If
    If BAND_INTERRUMPIR = False Then
        GRID_COLOR_FONDO Fg(4), Fg(4).Rows - 1, 1, Fg(4).Rows - 1, Fg(4).Cols - 1, &HC0C0FF
        GRID_COLOR_FONDO Fg(4), Fg(4).FixedRows, Fg(4).Cols - 1, Fg(4).Rows - 1, Fg(4).Cols - 1, &HC0C0FF
    End If
    
End Sub

Private Sub pLimpiarArr(xArr())
    Dim mPos&
    For mPos = 0 To UBound(xArr())
        xArr(mPos) = 0
    Next mPos
End Sub


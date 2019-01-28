VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmEvaluarCosto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Producción - Costo de Personal"
   ClientHeight    =   7575
   ClientLeft      =   105
   ClientTop       =   600
   ClientWidth     =   11820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11820
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4860
         Top             =   90
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
               Picture         =   "FrmEvaluarCosto.frx":0000
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto.frx":0544
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto.frx":08D6
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto.frx":0A5A
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto.frx":0EAE
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto.frx":0FC6
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto.frx":150A
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto.frx":1A4E
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto.frx":1B62
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto.frx":1C76
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto.frx":20CA
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto.frx":2236
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto.frx":277E
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto.frx":2A98
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraProgreso 
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   3030
      TabIndex        =   6
      Top             =   3435
      Visible         =   0   'False
      Width           =   5760
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   90
         TabIndex        =   7
         Top             =   345
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando: Registros"
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
         Left            =   90
         TabIndex        =   9
         Top             =   75
         Width           =   1890
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
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
         Left            =   4140
         TabIndex        =   8
         Top             =   75
         Width           =   1530
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   5745
         X2              =   5745
         Y1              =   -90
         Y2              =   4800
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -150
         X2              =   5895
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   2
         X1              =   -60
         X2              =   6360
         Y1              =   690
         Y2              =   690
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   3
         X1              =   0
         X2              =   15
         Y1              =   15
         Y2              =   5070
      End
   End
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   0
      TabIndex        =   0
      Top             =   255
      Width           =   11850
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   510
         Left            =   10410
         TabIndex        =   19
         Top             =   120
         Width           =   1380
      End
      Begin VB.Frame Frame3 
         Caption         =   "Seleccionar"
         Height          =   525
         Left            =   3990
         TabIndex        =   14
         Top             =   90
         Width           =   6015
         Begin VB.OptionButton OptSeleccion 
            Caption         =   "x Personal"
            Height          =   225
            Index           =   1
            Left            =   900
            TabIndex        =   27
            Top             =   240
            Width           =   1065
         End
         Begin VB.OptionButton OptSeleccion 
            Caption         =   "x Area"
            Height          =   225
            Index           =   0
            Left            =   90
            TabIndex        =   26
            Top             =   240
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.CommandButton cb 
            Height          =   240
            Index           =   0
            Left            =   3225
            Picture         =   "FrmEvaluarCosto.frx":2E2A
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   180
            Width           =   195
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   0
            Left            =   2820
            MaxLength       =   12
            TabIndex        =   16
            Text            =   "txt_cb(0)"
            Top             =   150
            Width           =   615
         End
         Begin VB.Line Line3 
            X1              =   1980
            X2              =   1980
            Y1              =   180
            Y2              =   450
         End
         Begin VB.Label lbl_cb_capt 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Area"
            Height          =   195
            Index           =   0
            Left            =   2130
            TabIndex        =   25
            Top             =   240
            Width           =   585
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
            Height          =   300
            Index           =   0
            Left            =   4770
            TabIndex        =   18
            Top             =   150
            Visible         =   0   'False
            Width           =   1005
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
            Height          =   300
            Index           =   0
            Left            =   3435
            TabIndex        =   17
            Top             =   150
            Width           =   2475
         End
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
         Height          =   300
         Index           =   0
         Left            =   645
         TabIndex        =   2
         Top             =   225
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Valor           =   "25/09/2007"
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
         Height          =   300
         Index           =   1
         Left            =   2550
         TabIndex        =   3
         Top             =   225
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Valor           =   "25/09/2007"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   330
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   2040
         TabIndex        =   4
         Top             =   330
         Width           =   465
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6600
      Left            =   0
      TabIndex        =   10
      Top             =   945
      Width           =   11850
      _cx             =   20902
      _cy             =   11642
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
      FrontTabColor   =   -2147483644
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "     Horas  |     Destajo    "
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
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6180
         Left            =   45
         TabIndex        =   12
         Top             =   45
         Width           =   11760
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   6060
            Left            =   60
            TabIndex        =   13
            Top             =   60
            Width           =   11655
            _cx             =   20558
            _cy             =   10689
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
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmEvaluarCosto.frx":2F5C
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
         Caption         =   "Frame3"
         Height          =   6180
         Left            =   12495
         TabIndex        =   11
         Top             =   45
         Width           =   11760
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   6165
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   11715
            _cx             =   20664
            _cy             =   10874
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
            Caption         =   "     Detalle    |    Resumen    "
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
            Begin VB.Frame Frame5 
               BorderStyle     =   0  'None
               Height          =   6135
               Left            =   330
               TabIndex        =   23
               Top             =   15
               Width           =   11370
               Begin VSFlex7Ctl.VSFlexGrid Fg2 
                  Height          =   6045
                  Left            =   60
                  TabIndex        =   24
                  Top             =   60
                  Width           =   11265
                  _cx             =   19870
                  _cy             =   10663
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
                  SelectionMode   =   1
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
                  FormatString    =   $"FrmEvaluarCosto.frx":316D
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
            Begin VB.Frame Frame1 
               BorderStyle     =   0  'None
               Height          =   6135
               Left            =   12645
               TabIndex        =   21
               Top             =   15
               Width           =   11370
               Begin VSFlex7Ctl.VSFlexGrid Fg3 
                  Height          =   6045
                  Left            =   60
                  TabIndex        =   22
                  Top             =   60
                  Width           =   11250
                  _cx             =   19844
                  _cy             =   10663
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
                  FormatString    =   $"FrmEvaluarCosto.frx":337E
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
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Menu1_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "Eliminar"
      End
   End
   Begin VB.Menu Menu2 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu Menu2_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu2_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu2_3 
         Caption         =   "Eliminar"
      End
   End
End
Attribute VB_Name = "FrmEvaluarCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'--HISTORIA
'--modificado: 18/11/09 Johan Castro
'              cambiar la presentacion de la planilla por horas, antes se mostraba en formato hh:mm:ss AM/PM
'              ahora se muestra segun formato HH:MM:SS (Total Horas, Total HN, Total HE)
'              adicionalmente se cambia la presentacion a la planilla destajo(diferencia de horas)
'--modificado: 07/12/09 Johan Castro
'              Cálculo del turno noche, HE salia duplicado cuando se detallaba las tareas del personal, esto originaba
'              el pago exesivo al personal.
'--modificado: 12/12/09 Johan Castro
'              Mostrar el campo incentivos en pago por horas luego de haber modificado y grabado.




Option Explicit

Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE
'--DE LA IMPRESION
Dim T_RPT_PERIODO As String '--PERIODO DEL REPORTE
Dim T_RPT_TITULO As String  '--TITULO DE REPORTE
'------------

Dim SeEjecuto  As Boolean
Dim Agregando  As Boolean
'------------

Private Sub pConsultar()
''    On Error GoTo error
    Dim rst_select As New ADODB.Recordset
    Dim nSQLSelect As String '--RECIBIR LA CONSULTA
    Dim mTipoConsulta As Integer '--valor para configurar el tipo de consulta y obtener el script sql
        
    If fValidarConsulta() = False Then Exit Sub
    BAND_INTERRUMPIR = False
    
    Me.MousePointer = vbHourglass
    DoEvents
    PosicionarProgBar
    
    lbl(0).Caption = "Procesando: Registros"
    lbl(1).Caption = "Interrumpir = ESC"
    
    If TabOne1.CurrTab = 0 Then
        '--cargar las horas
        pCargarHoras
    Else
        '--cargar los destajos
        pCargarDestajo
    End If
    '------------------------------------------------
    '*********************************************************************
    '*********************************************************************
   '
SALIR:
    FraProgreso.Visible = False
    Set rst_select = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    FraProgreso.Visible = False
    Set rst_select = Nothing
    SHOW_ERROR Me.Name, "pConsultar"
    
End Sub

Private Sub pImprimir()
    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    X_PRINT.Imprimir_x_VSFlexGrid IIf(TabOne1.CurrTab = 0, Fg1, Fg2), T_RPT_TITULO + " ", "", T_RPT_PERIODO, True, True
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimir"

End Sub

Private Sub CmdGrabar_Click()
    
    pGrabar TabOne1.CurrTab
    
End Sub

Private Sub Form_Activate()
    On Error GoTo error
    Dim mTipoConsulta As Integer
    If SeEjecuto = True Then Exit Sub
    SeEjecuto = True
    
    
    TxtFecha(0).valor = Date
    TxtFecha(1).valor = Date
    txt_cb(0).Text = ""
    lbl_cb(0).Caption = ""
    lbl_cod(0).Caption = ""
    '----------------------------
    pConfigurarGrilla
    
    Exit Sub
error:
    SHOW_ERROR Me.Name, "Form_Activate"
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        BAND_INTERRUMPIR = True '--interrumpir
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo error
    SeEjecuto = False
    centrarFrm Me
    
    Exit Sub
error:
    SHOW_ERROR
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    BAND_INTERRUMPIR = True
End Sub

'------
Private Function fValidarConsulta() As Boolean
    '--FUNCION QUE VALIDARA LA CONSULTA
    '--DE LA FECHA ES NULL
    If TxtFecha(0).valor = "" Or TxtFecha(1).valor = "" Then
        MsgBox "Ingrese una fecha", vbExclamation, xTitulo
        If TxtFecha(0).valor = "" Then TxtFecha(0).SetFocus Else TxtFecha(1).SetFocus
        Exit Function
    End If
    If CDate(TxtFecha(0).valor) > CDate(TxtFecha(1).valor) Then
        MsgBox "La fecha inicial es superior al Final", vbExclamation, xTitulo
        TxtFecha(0).SetFocus
        Exit Function
    End If
    fValidarConsulta = True
End Function

Private Sub PosicionarProgBar()
    '--POSICIONAR EL PROGRESO DENTRO DEL FORMULARIO
    '    FraProgreso.Width = 5820
    FraProgreso.Left = (Me.Width - FraProgreso.Width) \ 2
    FraProgreso.Top = (Me.Height - FraProgreso.Height) \ 2
    Me.PgBar.Value = 1
    FraProgreso.Visible = True
End Sub


'--------
Private Sub pExportarExcel()
    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    X_PRINT.VSFlexGrid_Exportar_MSExcel xCon, IIf(TabOne1.CurrTab = 0, Fg1, IIf(TabOne2.CurrTab = 0, Fg2, Fg3)), IIf(TabOne1.CurrTab = 0, "Costo de Personal - Horas", IIf(TabOne2.CurrTab = 0, "Costo de Personal - Destajo Detalle", "Costo de Personal - Destajo Resumen")), "De " & TxtFecha(0).valor & " Al " & TxtFecha(1).valor, , IIf(TabOne1.CurrTab = 0, "Costo Horas", IIf(TabOne2.CurrTab = 0, "Costo Destajo Detalle", "Costo Destajo Resumen"))
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pExportarExcel"
End Sub


Private Sub OptSeleccion_Click(Index As Integer)
    lbl_cb(0).Caption = ""
    txt_cb(0).Text = ""
    lbl_cod(0).Caption = ""
    If OptSeleccion(0).Value = True Then
        lbl_cb_capt(0).Caption = "Area"
    Else
        lbl_cb_capt(0).Caption = "Personal"
    End If
End Sub

'************************************************

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    If Button.Index = 5 Then pExportarExcel
    If Button.Index = 6 Then pImprimir
    If Button.Index = 8 Then
        BAND_INTERRUMPIR = True
        Unload Me
    End If
End Sub

'************************************************


'*******************************************************************************************

Private Sub cb_Click(Index As Integer)
    Dim xCampos() As String
    Dim nCampoBusca As String
    Dim nSQL As String
    Dim nTitulo As String
    On Error GoTo error
    Select Case Index
        Case 0 '--area
            If OptSeleccion(0).Value = True Then
                nTitulo = "Buscando Area"
                nSQL = "SELECT mae_area.id, mae_area.descripcion AS nombre, mae_area.id AS cod, mae_area.id AS idarea " _
                    + vbCr + " FROM pro_area INNER JOIN mae_area ON pro_area.idarea = mae_area.id; "
            Else
                nTitulo = "Buscando Personal"
                nSQL = "SELECT pla_empleados.id, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombre, pla_empleados.id AS cod " _
                    + vbCr + " FROM pla_empleados INNER JOIN pro_pagos ON pla_empleados.id = pro_pagos.idemp " _
                    + vbCr + " GROUP BY pla_empleados.id, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom], pla_empleados.id " _
                    + vbCr + " ORDER BY [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom];"
            
            End If
            
            
    End Select
    
    ReDim xCampos(2, 3) As String
    xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id":           xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"
    
    Dim RstTmp As New ADODB.Recordset
    CARGAR_DLL_EPSBUSCAR xCon, RstTmp, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio

    If RstTmp.State = 0 Then GoTo SALIR
    If RstTmp.EOF = True Or RstTmp.BOF = True Or RstTmp.RecordCount = 0 Then GoTo SALIR

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    txt_cb(Index).Text = NulosC(RstTmp.Fields(0))  '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = NulosC(RstTmp.Fields(1)) '--NOMBRE
    lbl_cod(Index).Caption = NulosN(RstTmp.Fields(2)) '--CODIGO
    lbl_cb(Index).ToolTipText = NulosC(RstTmp.Fields(1))  '--NOMBRE
      

SALIR:
    Set RstTmp = Nothing
Exit Sub
error:
    Set RstTmp = Nothing
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
        If Index <> 1 Then
            SendKeys vbTab
        Else
            If Fg1.Rows >= 2 Then
                Fg1.Row = 1: Fg1.Col = 1
            Else
                Fg1.Row = Fg1.Rows - 1: Fg1.Col = 1
            End If
            Fg1.SetFocus
        End If
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
        Case 0 '--area
            nSQL = "SELECT mae_area.id, mae_area.descripcion AS nombre, mae_area.id AS cod, mae_area.id AS idarea " _
                + vbCr + " FROM pro_area INNER JOIN mae_area ON pro_area.idarea = mae_area.id; "
        
        Case Else
            Exit Sub
            
    End Select

    If xCon.State = 0 Then GoTo SALIR
    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then GoTo SALIR

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    If RstTmp.RecordCount > 0 Then
        txt_cb(Index).Text = NulosC(RstTmp.Fields(0))   '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = NulosC(RstTmp.Fields(1)) '--NOMBRE
        lbl_cod(Index).Caption = NulosN(RstTmp.Fields(2)) '--CODIGO
        lbl_cb(Index).ToolTipText = NulosC(RstTmp.Fields(1)) '--NOMBRE
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cod(Index).Caption = ""
    End If
    
    Set RstTmp = Nothing
    Exit Sub
error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "txt_cb_Validate (" + CStr(Index) + ")"
    Exit Sub
SALIR:
    Set RstTmp = Nothing
    txt_cb(Index).Text = ""
End Sub


Private Sub pCargarHoras()
    Dim Rst As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLFiltro As String
    
    If NulosN(lbl_cod(0).Caption) <> 0 Then
        If OptSeleccion(0).Value = True Then
            nSQLFiltro = " and pro_controltar.idarea=" & NulosN(lbl_cod(0).Caption) & " "
        Else
            nSQLFiltro = " and pla_empleados.id=" & NulosN(lbl_cod(0).Caption) & " "
        End If
    End If
    
    Fg1.Rows = Fg1.FixedRows
    
    DoEvents
    
    '--consulta para determinar la lista de pagos por hora segun filtro seleccionado
    nSQL = "SELECT pla_empleados.id as idemp,mae_area.id as idarea, pro_controltar.fchtra, mae_area.descripcion AS area, pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom AS personal, pla_empleados.paghornor, pla_empleados.paghorext " _
        + vbCr + " FROM (pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN ((pro_controltardet LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id) ON pro_controltar.id = pro_controltardet.idctr " _
        + vbCr + " WHERE (((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltar.tipo)=1)) " & nSQLFiltro _
        + vbCr + " GROUP BY pla_empleados.id, mae_area.id ,pro_controltar.fchtra, mae_area.descripcion, pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom, pla_empleados.paghornor, pla_empleados.paghorext " _
        + vbCr + " ORDER BY pro_controltar.fchtra, mae_area.descripcion, pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom; "
    
    RST_Busq Rst, nSQL, xCon
    
    Dim Difhora As String
    Dim HoraIniNoche As String
    Dim HoraFinNoche As String
    Dim rstTmp1 As New ADODB.Recordset
    Dim rstBonif As New ADODB.Recordset '--registro de los incentivos que se le aplican
    
    If Rst.RecordCount = 0 Then Exit Sub
    Agregando = True
    PgBar.Min = 0
    PgBar.Value = 0
    PgBar.Max = Rst.RecordCount
    
    '---cargando listado de incentivos
    With Fg1
        Do While Not Rst.EOF
            DoEvents
            If BAND_INTERRUMPIR = True Then GoTo SALIR
            PgBar.Value = PgBar.Value + 1
            '------------------------------------------------
            Fg1.Rows = Fg1.Rows + 1
            .TextMatrix(Fg1.Rows - 1, 1) = Rst.Bookmark
            .TextMatrix(Fg1.Rows - 1, 2) = Format(Rst("fchtra"), FORMAT_DATE)
            .TextMatrix(Fg1.Rows - 1, 3) = NulosC(Rst("area"))
            .TextMatrix(Fg1.Rows - 1, 4) = NulosC(Rst("personal"))
            '--obteniendo la hora de inicio y de termino
                
            nSQL = "SELECT ini.fchtra, ini.idref, ini.hinipri, ini.hfinpri, fin.hiniult, fin.hfinult " _
                + vbCr + " From " _
                + vbCr + " (SELECT DISTINCT TOP 1 pro_controltar.fchtra, pro_controltardet.idref, pro_controltardet.horini AS hinipri, pro_controltardet.horfin AS hfinpri " _
                + vbCr + " FROM pro_controltar INNER JOIN pro_controltardet ON pro_controltar.id = pro_controltardet.idctr " _
                + vbCr + " WHERE (((pro_controltar.fchtra)=CDate('" & Rst("fchtra") & "')) AND ((pro_controltardet.idref)=" & NulosN(Rst("idemp")) & ") AND ((pro_controltar.tipo)=1)) and pro_controltardet.horini Is Not Null AND pro_controltardet.horfin Is Not Null" _
                + vbCr + " ORDER BY pro_controltardet.horini ) AS ini "
                
            nSQL = nSQL & vbCr + " Left Join " _
                + vbCr + " (SELECT DISTINCT TOP 1 pro_controltar.fchtra, pro_controltardet.idref, pro_controltardet.horini AS hiniult, pro_controltardet.horfin AS hfinult " _
                + vbCr + " FROM pro_controltar INNER JOIN pro_controltardet ON pro_controltar.id = pro_controltardet.idctr " _
                + vbCr + " WHERE (((pro_controltar.fchtra)=CDate('" & Rst("fchtra") & "')) AND ((pro_controltardet.idref)=" & NulosN(Rst("idemp")) & ") AND ((pro_controltar.tipo)=1)) and pro_controltardet.horini Is Not Null AND pro_controltardet.horfin Is Not Null " _
                + vbCr + " ORDER BY pro_controltardet.horini desc " _
                + vbCr + " ) as fin " _
                + vbCr + " ON (ini.idref = fin.idref) AND (ini.fchtra = fin.fchtra);"
            
            
            RST_Busq RstTmp, nSQL, xCon
            If RstTmp.RecordCount <> 0 Then
                '--evaluando las horas
                'Si la hora de inicio es mayor a la hora de termino de la tarea
                If (CDate(RstTmp("hinipri")) > CDate(RstTmp("hfinult"))) Then
                    Difhora = DiferenciaHoras(RstTmp("hinipri"), CDate("00:00"))
                    Difhora = Format(CDate(Difhora) + CDate(DiferenciaHoras(CDate("00:00"), RstTmp("hfinult"))), "HH:mm")
                    
                    .TextMatrix(Fg1.Rows - 1, 5) = Format(RstTmp("hinipri"), FORMAT_HORA_SIN_SEGUNDO)
                    .TextMatrix(Fg1.Rows - 1, 6) = Format(RstTmp("hfinult"), FORMAT_HORA_SIN_SEGUNDO)
                    .TextMatrix(Fg1.Rows - 1, 7) = "Otro"
                Else
                    'Si los intervalos de tiempo estan en turno noche
                    If CDate(RstTmp("hfinult")) < CDate("10:00") Or CDate(RstTmp("hfinult")) > CDate("23:50") Or (CDate(RstTmp("hinipri")) > CDate("19:00")) Then
                    
                        '-----------------------------------------------------
                        '--obteniendo la ultima hora cuando una persona labora de noche
                        nSQL = "SELECT TOP 1 pro_controltardet.horfin " _
                            + vbCr + " FROM pro_controltar INNER JOIN pro_controltardet ON pro_controltar.id = pro_controltardet.idctr " _
                            + vbCr + " WHERE (((pro_controltardet.horfin) <= CDate('20:00')) AND ((pro_controltar.tipo)=1) AND ((pro_controltar.fchtra)=CDate('" & Rst("fchtra") & "')) AND ((pro_controltardet.idref)=" & Rst("idemp") & ")) " _
                            + vbCr + " ORDER BY pro_controltardet.horfin DESC;"
                            
                        RST_Busq rstTmp1, nSQL, xCon
                        If rstTmp1.RecordCount <> 0 Then
                            HoraFinNoche = rstTmp1("horfin")
                        End If
                        Set rstTmp1 = Nothing
                        
                       '--obteniendo la primera hora cuando una persona labora de noche
                        nSQL = "SELECT TOP 1 pro_controltardet.horini " _
                            + vbCr + " FROM pro_controltar INNER JOIN pro_controltardet ON pro_controltar.id = pro_controltardet.idctr " _
                            + vbCr + " WHERE pro_controltardet.horini >= CDate('10:00') and ((pro_controltar.tipo)=1) AND ((pro_controltar.fchtra)=CDate('" & Rst("fchtra") & "')) AND ((pro_controltardet.idref)=" & Rst("idemp") & ") " _
                            + vbCr + " ORDER BY pro_controltardet.horini asc;"
                            
                        RST_Busq rstTmp1, nSQL, xCon
                        If rstTmp1.RecordCount <> 0 Then
                            HoraIniNoche = rstTmp1("horini")
                        End If
                        Set rstTmp1 = Nothing
                        '-----------------------------------------------------
                        Difhora = DiferenciaHoras(HoraIniNoche, HoraFinNoche)
                        
                        .TextMatrix(Fg1.Rows - 1, 5) = Format(HoraIniNoche, FORMAT_HORA_SIN_SEGUNDO)
                        .TextMatrix(Fg1.Rows - 1, 6) = Format(HoraFinNoche, FORMAT_HORA_SIN_SEGUNDO)
                        .TextMatrix(Fg1.Rows - 1, 7) = "Noche"
                        
                    Else 'Si los intervalos de tiempo estan en turno dia
                    
                        'Se calcula manualmente la diferencia de horas
                        Difhora = Format(CDate(RstTmp("hinipri")) - CDate(RstTmp("hfinult")), "HH:mm")
                        
                        .TextMatrix(Fg1.Rows - 1, 5) = Format(RstTmp("hinipri"), FORMAT_HORA_SIN_SEGUNDO)
                        .TextMatrix(Fg1.Rows - 1, 6) = Format(RstTmp("hfinult"), FORMAT_HORA_SIN_SEGUNDO)
                        .TextMatrix(Fg1.Rows - 1, 7) = "Dia"
                    End If
                End If
                
                .TextMatrix(Fg1.Rows - 1, 8) = Format(Difhora, FORMAT_HORA_LARGO) '--tot horas
                
                Dim h() As String
                Dim tiempo As Double
                '--si es de dia
                If .TextMatrix(Fg1.Rows - 1, 7) = "Dia" Then
                    If CDate(Difhora) > CDate("10:00") Then
                    
                        .TextMatrix(Fg1.Rows - 1, 9) = Format(CDate("10:00"), FORMAT_HORA_LARGO) ' H.Normal
                        .TextMatrix(Fg1.Rows - 1, 10) = NulosN(Rst("paghornor")) 'Costo HN
                        .TextMatrix(Fg1.Rows - 1, 11) = 10 * NulosN(Rst("paghornor")) ' Total HN
                        
                        .TextMatrix(Fg1.Rows - 1, 12) = Format(CDate(CDate(Difhora) - CDate("10:00")), FORMAT_HORA_LARGO)  ' H.Extra
                        
                        .TextMatrix(Fg1.Rows - 1, 13) = Format(NulosN(Rst("paghorext")), FORMAT_MONTO) 'Costo HE
                        
                        'Se calcula el tiempo en horas formato decimales
                        h = Split(Format(.TextMatrix(Fg1.Rows - 1, 12), "HH:mm"), ":")
                        tiempo = Val(h(0)) + (Val(h(1)) / 60)
                        .TextMatrix(Fg1.Rows - 1, 14) = tiempo * NulosN(Rst("paghorext")) 'Total HE
                    Else
                        .TextMatrix(Fg1.Rows - 1, 9) = Format(CDate(Difhora), FORMAT_HORA_LARGO) ' H.Normal
                        .TextMatrix(Fg1.Rows - 1, 10) = Format(NulosN(Rst("paghornor")), FORMAT_MONTO) 'Costo HN
                        .TextMatrix(Fg1.Rows - 1, 11) = Convert1HoraFaccion(Difhora) * NulosN(Rst("paghornor")) ' Total HN
                                                
                        .TextMatrix(Fg1.Rows - 1, 12) = "" 'Format(CDate("00:00"), FORMAT_HORA_SIN_SEGUNDO) ' H.Extra
                        .TextMatrix(Fg1.Rows - 1, 13) = NulosN(Rst("paghorext")) 'Costo HE
                        .TextMatrix(Fg1.Rows - 1, 14) = 0 'Total HE
                    End If
                End If
                '--si es de noche
                If .TextMatrix(Fg1.Rows - 1, 7) = "Noche" Then '--si es de dia
                    If Difhora <> "" Then
                        If CDate(Difhora) > CDate("08:00") Then
                        
                            .TextMatrix(Fg1.Rows - 1, 9) = Format(CDate("08:00"), FORMAT_HORA_LARGO) ' H.Normal
                            .TextMatrix(Fg1.Rows - 1, 10) = NulosN(Rst("paghornor")) 'Costo HN
                            .TextMatrix(Fg1.Rows - 1, 11) = 10 * NulosN(Rst("paghornor")) ' Total HN
                            
                            If CDate(Difhora) > CDate("10:00") Then
                                .TextMatrix(Fg1.Rows - 1, 12) = Format(CDate(CDate(Difhora) - CDate("10:00")), FORMAT_HORA_LARGO)  ' H.Extra
                                .TextMatrix(Fg1.Rows - 1, 13) = Format(NulosN(Rst("paghorext")), FORMAT_MONTO) 'Costo HE
                                .TextMatrix(Fg1.Rows - 1, 14) = Convert1HoraFaccion(CDate(.TextMatrix(Fg1.Rows - 1, 12))) * NulosN(Rst("paghorext")) 'Total HE
                            Else
                                .TextMatrix(Fg1.Rows - 1, 12) = Format(CDate(CDate(Difhora) - CDate("08:00")), FORMAT_HORA_LARGO)  ' H.Extra
                                .TextMatrix(Fg1.Rows - 1, 13) = Format(NulosN(Rst("paghorext")), FORMAT_MONTO) 'Costo HE
                                .TextMatrix(Fg1.Rows - 1, 14) = Convert1HoraFaccion(CDate(.TextMatrix(Fg1.Rows - 1, 12))) * NulosN(Rst("paghorext")) 'Total HE
                            End If
                            
                        Else
                            .TextMatrix(Fg1.Rows - 1, 9) = Format(CDate(Difhora), FORMAT_HORA_LARGO) ' H.Normal
                            .TextMatrix(Fg1.Rows - 1, 10) = Format(NulosN(Rst("paghornor")), FORMAT_MONTO) 'Costo HN
                            .TextMatrix(Fg1.Rows - 1, 11) = Convert1HoraFaccion(Difhora) * NulosN(Rst("paghornor")) ' Total HN
                            
                            .TextMatrix(Fg1.Rows - 1, 12) = "" 'Format(CDate("00:00"), FORMAT_HORA_SIN_SEGUNDO) ' H.Extra
                            .TextMatrix(Fg1.Rows - 1, 13) = NulosN(Rst("paghorext")) 'Costo HE
                            .TextMatrix(Fg1.Rows - 1, 14) = 0 'Total HE
                        End If
                    End If
                End If
                
                '--si es otro caso distinto
                If .TextMatrix(Fg1.Rows - 1, 7) = "Otro" Then
                    If CDate(Difhora) > CDate("10:00") Then
                    
                        .TextMatrix(Fg1.Rows - 1, 9) = Format(CDate("10:00"), FORMAT_HORA_LARGO) ' H.Normal
                        .TextMatrix(Fg1.Rows - 1, 10) = NulosN(Rst("paghornor")) 'Costo HN
                        .TextMatrix(Fg1.Rows - 1, 11) = 10 * NulosN(Rst("paghornor")) ' Total HN
                        
                        .TextMatrix(Fg1.Rows - 1, 12) = Format(CDate(CDate(Difhora) - CDate("10:00")), FORMAT_HORA_LARGO)  ' H.Extra
                        .TextMatrix(Fg1.Rows - 1, 13) = Format(NulosN(Rst("paghorext")), FORMAT_MONTO) 'Costo HE
                        .TextMatrix(Fg1.Rows - 1, 14) = Convert1HoraFaccion(CDate(.TextMatrix(Fg1.Rows - 1, 12))) * NulosN(Rst("paghorext")) 'Total HE
                        
                        
                    Else
                        .TextMatrix(Fg1.Rows - 1, 9) = Format(CDate(Difhora), FORMAT_HORA_LARGO) ' H.Normal
                        .TextMatrix(Fg1.Rows - 1, 10) = Format(NulosN(Rst("paghornor")), FORMAT_MONTO) 'Costo HN
                        .TextMatrix(Fg1.Rows - 1, 11) = Convert1HoraFaccion(Difhora) * NulosN(Rst("paghornor")) ' Total HN
                                                
                        .TextMatrix(Fg1.Rows - 1, 12) = "" 'Format(CDate("00:00"), FORMAT_HORA_SIN_SEGUNDO) ' H.Extra
                        .TextMatrix(Fg1.Rows - 1, 13) = NulosN(Rst("paghorext")) 'Costo HE
                        .TextMatrix(Fg1.Rows - 1, 14) = 0 'Total HE
                    End If
                End If
                
            End If
            
            .TextMatrix(.Rows - 1, 11) = Format(.TextMatrix(Fg1.Rows - 1, 11), FORMAT_MONTO)
            .TextMatrix(.Rows - 1, 14) = Format(.TextMatrix(Fg1.Rows - 1, 14), FORMAT_MONTO)
            
            .TextMatrix(.Rows - 1, 15) = NulosN(.TextMatrix(Fg1.Rows - 1, 11)) + NulosN(.TextMatrix(Fg1.Rows - 1, 14)) 'Tot Pagar
            .TextMatrix(.Rows - 1, 15) = Format(.TextMatrix(Fg1.Rows - 1, 15), FORMAT_MONTO)
            
            '--copiando los datos del pago total
            '-------------------------------------------'-------------------------------------------'-------------------------------------------
            '--incentivos
            nSQL = "SELECT pro_pagos.imptot, pro_pagos.impbon " _
                & " From pro_pagos " _
                & " WHERE (((pro_pagos.idemp)=" & Rst("idemp") & ") AND ((pro_pagos.idarea)=" & Rst("idarea") & ") AND ((pro_pagos.fchtra)=cdate('" & Rst("fchtra") & "')) AND ((pro_pagos.tipo)=1));"
            
            RST_Busq rstBonif, nSQL, xCon
            
            If rstBonif.RecordCount <> 0 Then
                .TextMatrix(.Rows - 1, 16) = NulosN(rstBonif("impbon"))
            Else
                .TextMatrix(.Rows - 1, 16) = 0
            End If
            
            Set rstBonif = Nothing
            '-------------------------------------------
            .TextMatrix(.Rows - 1, 17) = Format(NulosN(.TextMatrix(.Rows - 1, 15)) + NulosN(.TextMatrix(.Rows - 1, 16)), FORMAT_MONTO)
            
            .TextMatrix(.Rows - 1, 18) = NulosN(Rst("idarea"))
            .TextMatrix(.Rows - 1, 19) = NulosN(Rst("idemp"))
            
            Set RstTmp = Nothing
            
            Rst.MoveNext
        Loop
    End With
    Set Rst = Nothing
    Set RstTmp = Nothing
    '----------
    GRID_AGRUPAR Fg1, 3
    
    Dim mRow As Long
    
    For mRow = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(mRow, 17)) = 0 Then
            GRID_COLOR_FONDO Fg1, mRow, 17, mRow, 17, vbRed
        End If
    Next
    
    '--colocando los totales
    Fg1.Rows = Fg1.Rows + 1
    Fg1.TextMatrix(Fg1.Rows - 1, 4) = "Totales"
    Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(GRID_SUMAR_COL(Fg1, 11), FORMAT_MONTO)
    Fg1.TextMatrix(Fg1.Rows - 1, 14) = Format(GRID_SUMAR_COL(Fg1, 14), FORMAT_MONTO)
    Fg1.TextMatrix(Fg1.Rows - 1, 15) = Format(GRID_SUMAR_COL(Fg1, 15), FORMAT_MONTO)
    Fg1.TextMatrix(Fg1.Rows - 1, 17) = Format(GRID_SUMAR_COL(Fg1, 17), FORMAT_MONTO)
    
    GRID_COLOR_FONDO Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1, vbGreen
    
    
SALIR:
Agregando = False
    
End Sub



Private Sub pConfigurarGrilla()
    '===================================================================================================
    'Propósito: Establecer los encabezados del grid
    '
    'Entradas:  Ninguno
    '
    'Resultados: Grilla con Encabezado
    '===================================================================================================
    Dim k As Integer
    
    Agregando = True
    
    With Fg1
        '-----Pago por Horas
        
        .Cols = 20
        .Rows = 1
        
        .ColWidth(0) = 200
        .ColWidth(1) = 0
        .FrozenCols = 4
        .TextMatrix(0, 1) = "Nº":           .ColWidth(1) = 450:         .ColAlignment(1) = flexAlignRightBottom:       .Row = 0: .Col = 1: .CellAlignment = flexAlignRightBottom
        .TextMatrix(0, 2) = "Fecha":        .ColWidth(2) = 800:      .ColAlignment(2) = flexAlignCenterCenter:         .Row = 0: .Col = 2: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 3) = "Area":         .ColWidth(3) = 900:       .ColAlignment(3) = flexAlignLeftBottom:       .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(0, 4) = "Personal":     .ColWidth(4) = 1500:       .ColAlignment(4) = flexAlignLeftBottom:         .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(0, 5) = "Hor. Inicio":  .ColWidth(5) = 900:      .ColAlignment(5) = flexAlignCenterCenter:         .Row = 0: .Col = 5: .CellAlignment = flexAlignCenterCenter
        '--------
        .TextMatrix(0, 6) = "Hor. Fin":     .ColWidth(6) = 900:      .ColAlignment(6) = flexAlignCenterCenter:         .Row = 0: .Col = 6: .CellAlignment = flexAlignCenterCenter
        
        .TextMatrix(0, 7) = "Horario":       .ColWidth(7) = 700:     .ColAlignment(7) = flexAlignLeftBottom:        .Row = 0: .Col = 7: .CellAlignment = flexAlignLeftBottom
        
        .TextMatrix(0, 8) = "Tot. Horas":   .ColWidth(8) = 900:      .ColAlignment(8) = flexAlignCenterCenter:         .Row = 0: .Col = 8: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 9) = "H.Normal":     .ColWidth(9) = 900:       .ColAlignment(9) = flexAlignCenterCenter:       .Row = 0: .Col = 9: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 10) = "Costo HN":     .ColWidth(10) = 800:     .ColAlignment(10) = flexAlignRightCenter:       .Row = 0: .Col = 10: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 11) = "Total HN":    .ColWidth(11) = 800:      .ColAlignment(11) = flexAlignRightBottom:       .Row = 0: .Col = 11: .CellAlignment = flexAlignRightBottom
        .TextMatrix(0, 12) = "H.Extra":     .ColWidth(12) = 900:      .ColAlignment(12) = flexAlignCenterCenter:      .Row = 0: .Col = 12: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 13) = "Costo HE":    .ColWidth(13) = 800:      .ColAlignment(13) = flexAlignRightCenter:        .Row = 0: .Col = 13: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 14) = "Total HE":    .ColWidth(14) = 800:      .ColAlignment(14) = flexAlignRightCenter:        .Row = 0: .Col = 14: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 15) = "Tot Pagar":   .ColWidth(15) = 900:      .ColAlignment(15) = flexAlignRightCenter:        .Row = 0: .Col = 15: .CellAlignment = flexAlignRightCenter
        
        
        .TextMatrix(0, 16) = "Incentivos":   .ColWidth(16) = 900:     .ColAlignment(16) = flexAlignRightCenter:        .Row = 0: .Col = 16: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 17) = "Neto Pagar":   .ColWidth(17) = 900:     .ColAlignment(17) = flexAlignRightCenter:        .Row = 0: .Col = 17: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 18) = "IdArea":   .ColWidth(18) = 0:
        
        .TextMatrix(0, 19) = "IdEmp":   .ColWidth(19) = 0:
        
    End With
            
    With Fg2
        '-----Pago por Destajo
        .Cols = 23
        .Rows = 2
        .FixedRows = 2
        
        .ColWidth(0) = 200
        .ColWidth(1) = 0
        .FrozenCols = 6
        
        
        GRID_COMBINAR Fg2, 0, 1, 0, 11, "Información de Trabajo", flexAlignLeftCenter, True, , , &HD8E9EC, True
        GRID_COMBINAR Fg2, 0, 12, 1, 17, "Eficiencia", flexAlignLeftCenter, False, , , &HD8E9EC, True
        GRID_COMBINAR Fg2, 0, 19, 0, 22, "Pago", flexAlignLeftCenter, True, , , &HD8E9EC, True
        
        
        .TextMatrix(1, 2) = "Fecha":        .ColWidth(2) = 800:      .ColAlignment(2) = flexAlignCenterCenter:         .Row = 1: .Col = 2: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 3) = "Area":         .ColWidth(3) = 450:       .ColAlignment(3) = flexAlignLeftBottom:       .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(1, 4) = "Personal":     .ColWidth(4) = 1200:       .ColAlignment(4) = flexAlignLeftBottom:         .Row = 1: .Col = 4: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(1, 5) = "Tarea":  .ColWidth(5) = 1500:             .ColAlignment(5) = flexAlignLeftBottom:         .Row = 1: .Col = 5: .CellAlignment = flexAlignLeftBottom
        '--------
        .TextMatrix(1, 6) = "Producto":    .ColWidth(6) = 1800:       .ColAlignment(6) = flexAlignLeftBottom:         .Row = 1: .Col = 6: .CellAlignment = flexAlignLeftBottom
        
        
        .TextMatrix(1, 7) = "Observación":    .ColWidth(7) = 1200:       .ColAlignment(7) = flexAlignLeftBottom:    .Row = 1: .Col = 7: .CellAlignment = flexAlignLeftBottom
        
        .TextMatrix(1, 8) = "H.Inicio":    .ColWidth(8) = 800:    .ColAlignment(8) = flexAlignLeftBottom:       .Row = 1: .Col = 8: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(1, 9) = "H.Final":     .ColWidth(9) = 800:        .ColAlignment(9) = flexAlignLeftBottom:         .Row = 1: .Col = 9: .CellAlignment = flexAlignLeftBottom
        
        
        .TextMatrix(1, 10) = "Cant.":     .ColWidth(10) = 700:            .ColAlignment(10) = flexAlignRightBottom:         .Row = 1: .Col = 10: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 11) = "U.M.":     .ColWidth(11) = 450:            .ColAlignment(11) = flexAlignCenterCenter:       .Row = 1: .Col = 11: .CellAlignment = flexAlignCenterCenter
        
        
        .TextMatrix(1, 12) = "Dif.Hora":        .ColWidth(12) = 800:       .ColAlignment(12) = flexAlignRightCenter:      .Row = 1: .Col = 12: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 13) = "Tot.Min":         .ColWidth(13) = 600:     .ColAlignment(13) = flexAlignRightCenter:        .Row = 1: .Col = 13: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 14) = "Unid x Min":      .ColWidth(14) = 0:       .ColAlignment(14) = flexAlignRightCenter:        .Row = 1: .Col = 14: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 15) = "Unid x Hor":      .ColWidth(15) = 950:    .ColAlignment(15) = flexAlignRightCenter:         .Row = 1: .Col = 15: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 16) = "Cant Teo":        .ColWidth(16) = 730:     .ColAlignment(16) = flexAlignRightCenter:        .Row = 1: .Col = 16: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 17) = "%":      .ColWidth(17) = 800:     .ColAlignment(17) = flexAlignRightCenter:        .Row = 1: .Col = 17: .CellAlignment = flexAlignRightCenter
        
        .TextMatrix(1, 18) = " ":           .ColWidth(18) = 0:       .ColAlignment(18) = flexAlignRightCenter:       .Row = 1: .Col = 18: .CellAlignment = flexAlignRightCenter
        
        .TextMatrix(1, 19) = "Cant":        .ColWidth(19) = 800:       .ColAlignment(19) = flexAlignRightCenter:       .Row = 1: .Col = 19: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 20) = "U.M.":        .ColWidth(20) = 450:       .ColAlignment(20) = flexAlignCenterCenter:      .Row = 1: .Col = 20: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 21) = "Pre.Uni":     .ColWidth(21) = 800:       .ColAlignment(21) = flexAlignRightCenter:       .Row = 1: .Col = 21: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 22) = "Total":       .ColWidth(22) = 750:       .ColAlignment(22) = flexAlignRightBottom:       .Row = 1: .Col = 22: .CellAlignment = flexAlignRightBottom
                                                
    End With
            
            
    With Fg3
        '-----
        .Cols = 11
        .Rows = 1
        
        .ColWidth(0) = 200
        .ColWidth(1) = 0
        .FrozenCols = 4
        .TextMatrix(0, 1) = "Nº":           .ColWidth(1) = 450:         .ColAlignment(1) = flexAlignRightBottom:       .Row = 0: .Col = 1: .CellAlignment = flexAlignRightBottom
        .TextMatrix(0, 2) = "Fecha":        .ColWidth(2) = 800:      .ColAlignment(2) = flexAlignCenterCenter:         .Row = 0: .Col = 2: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 3) = "Area":         .ColWidth(3) = 900:       .ColAlignment(3) = flexAlignLeftBottom:       .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(0, 4) = "Personal":     .ColWidth(4) = 3500:       .ColAlignment(4) = flexAlignLeftBottom:         .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftBottom
        
        .TextMatrix(0, 5) = "Horario":      .ColWidth(5) = 700:       .ColAlignment(5) = flexAlignLeftBottom:         .Row = 0: .Col = 5: .CellAlignment = flexAlignLeftBottom
        
        .TextMatrix(0, 6) = "Total":        .ColWidth(6) = 900:      .ColAlignment(6) = flexAlignRightBottom:       .Row = 0: .Col = 6: .CellAlignment = flexAlignRightBottom
        .TextMatrix(0, 7) = "Incentivos":   .ColWidth(7) = 900:     .ColAlignment(7) = flexAlignRightBottom:      .Row = 0: .Col = 7: .CellAlignment = flexAlignRightBottom
        .TextMatrix(0, 8) = "Neto Pagar":   .ColWidth(8) = 900:     .ColAlignment(8) = flexAlignRightCenter:        .Row = 0: .Col = 8: .CellAlignment = flexAlignRightCenter
        
        
        
        .TextMatrix(0, 9) = "IdEmp":   .ColWidth(9) = 0:
        .TextMatrix(0, 10) = "IdArea":   .ColWidth(10) = 0:
                                                
    End With
    
    Agregando = False
    
    DoEvents
End Sub




Private Sub Fg1_EnterCell()

    If Fg1.Col = 16 Then
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If Col <> 16 Then Exit Sub
    
    If IsNumeric(Fg1.TextMatrix(Row, 16)) = False Then
        Fg1.TextMatrix(Row, 17) = Fg1.TextMatrix(Row, 15)
        Fg1.TextMatrix(Row, 16) = 0
        Fg1.SetFocus
        Exit Sub
    End If
    Fg1.TextMatrix(Row, 16) = Format(Fg1.TextMatrix(Row, 16), FORMAT_MONTO)
    
    Fg1.TextMatrix(Row, 17) = NulosN(Fg1.TextMatrix(Row, 16) + NulosN(Fg1.TextMatrix(Row, 15)))
    
    '--totalizar
    Fg1.TextMatrix(Fg1.Rows - 1, 16) = Format(GRID_SUMAR_COL(Fg1, 16) - NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 16)), FORMAT_MONTO)
    Fg1.TextMatrix(Fg1.Rows - 1, 17) = Format(GRID_SUMAR_COL(Fg1, 17) - NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 17)), FORMAT_MONTO)
    
End Sub


Private Sub Fg3_EnterCell()

    If Fg3.Col = 7 Then
        Fg3.Editable = flexEDKbdMouse
    Else
        Fg3.Editable = flexEDNone
    End If
End Sub

Private Sub Fg3_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If Col <> 7 Then Exit Sub
    
    If IsNumeric(Fg3.TextMatrix(Row, 7)) = False Then
        Fg3.TextMatrix(Row, 8) = Fg3.TextMatrix(Row, 6)
        Fg3.TextMatrix(Row, 7) = 0
        Fg3.SetFocus
        Exit Sub
    End If
    Fg3.TextMatrix(Row, 7) = Format(Fg3.TextMatrix(Row, 7), FORMAT_MONTO)
    
    Fg3.TextMatrix(Row, 8) = NulosN(Fg3.TextMatrix(Row, 7) + NulosN(Fg3.TextMatrix(Row, 6)))
    
    '--totalizar
    Fg3.TextMatrix(Fg3.Rows - 1, 6) = Format(GRID_SUMAR_COL(Fg3, 6) - NulosN(Fg3.TextMatrix(Fg3.Rows - 1, 6)), FORMAT_MONTO)
    Fg3.TextMatrix(Fg3.Rows - 1, 7) = Format(GRID_SUMAR_COL(Fg3, 7) - NulosN(Fg3.TextMatrix(Fg3.Rows - 1, 7)), FORMAT_MONTO)
    Fg3.TextMatrix(Fg3.Rows - 1, 8) = Format(GRID_SUMAR_COL(Fg3, 8) - NulosN(Fg3.TextMatrix(Fg3.Rows - 1, 8)), FORMAT_MONTO)
    
End Sub


Private Sub pGrabar(Tipo As Integer)
    '--tipo 0=horas, 1=destajo
    
    If Tipo = 0 Then
        If Fg1.Rows = 1 Then
            MsgBox "No hay Registros para grabar", vbInformation
            Exit Sub
        End If
    Else
        If Fg3.Rows = 1 Then
            MsgBox "No hay Registros para grabar", vbInformation
            Exit Sub
        End If
    
    
    End If
    
    If MsgBox("Seguro desea continuar", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
    habilitar TxtFecha, False
    
    Dim xFil&
    Dim xCod&
    Dim RstDet As New ADODB.Recordset
    
    On Error GoTo error
    xCon.BeginTrans
    RST_Busq RstDet, "SELECT top 1 * FROM pro_pagos ", xCon
            
    '--registro por horas
    If Tipo = 0 Then
    
        '--eliminar si se eligio una area en especial, ej Acabado
        '--se eliminara los registro contenidos en el intervalo de fechas del area seleccionada
        If NulosN(lbl_cod(0).Caption) <> 0 Then
            xCon.Execute "DELETE pro_pagos.*  FROM pro_pagos " _
                & " WHERE (((CDate([pro_pagos].[fchtra])) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_pagos.idarea)=" & NulosN(lbl_cod(0).Caption) & ") AND ((pro_pagos.tipo)=1)); "
        End If
        
        With Fg1
            '--eliminar los registros
            
            For xFil = 1 To Fg1.Rows - 2
            
                DoEvents
                
                '--pro_pagos.tipo:: 1 =hora, 2=destajo
                '--solo se eliminara cuando se consulte sin filtro de area
                If NulosN(lbl_cod(0).Caption) = 0 Then
                    xCon.Execute "DELETE * FROM pro_pagos where pro_pagos.tipo=1 and cdate(pro_pagos.fchtra) = '" & CDate(.TextMatrix(xFil, 2)) & "' and idemp = " & NulosN(.TextMatrix(xFil, 19))
                End If
'                xCod = HallaCodigoTabla("pro_pagos", xCon, "id")
                RstDet.AddNew
'                RstDet("id") = xCod
                RstDet("tipo") = 1
                RstDet("fchtra") = CDate(.TextMatrix(xFil, 2))
                RstDet("imptot") = NulosN(.TextMatrix(xFil, 15))
                RstDet("impbon") = NulosN(.TextMatrix(xFil, 16))
                RstDet("impbrut") = NulosN(.TextMatrix(xFil, 17))
                
                RstDet("idemp") = NulosN(.TextMatrix(xFil, 19))
                RstDet("idarea") = NulosN(.TextMatrix(xFil, 18))
                
                If UCase(.TextMatrix(xFil, 7)) = "NOCHE" Then
                    RstDet("turno") = 2
                Else
                    RstDet("turno") = 1
                End If
                
                RstDet.Update
            Next
        End With
        
    '--registro por destajo
    Else
    
        '--eliminando registros del area seleccionada
        If NulosN(lbl_cod(0).Caption) <> 0 Then
        
            xCon.Execute "DELETE pro_pagos.*  FROM pro_pagos " _
                & " WHERE (((CDate([pro_pagos].[fchtra])) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_pagos.idarea)=" & NulosN(lbl_cod(0).Caption) & ") AND ((pro_pagos.tipo)=2)); "
                        
        End If
                
        With Fg3
            For xFil = 1 To .Rows - 2
                DoEvents
                
                If NulosN(lbl_cod(0).Caption) = 0 Then
                    xCon.Execute "DELETE * FROM pro_pagos where pro_pagos.tipo=2 and cdate(pro_pagos.fchtra) = '" & CDate(.TextMatrix(xFil, 2)) & "' and idemp = " & NulosN(.TextMatrix(xFil, 9))
                End If
'                xCod = HallaCodigoTabla("pro_pagos", xCon, "id")
                RstDet.AddNew
'                RstDet("id") = xCod
                RstDet("tipo") = 2
                
                RstDet("fchtra") = CDate(.TextMatrix(xFil, 2))
                
                RstDet("imptot") = NulosN(.TextMatrix(xFil, 6))
                RstDet("impbon") = NulosN(.TextMatrix(xFil, 7))
                RstDet("impbrut") = NulosN(.TextMatrix(xFil, 8))
                
                RstDet("idemp") = NulosN(.TextMatrix(xFil, 9))
                RstDet("idarea") = NulosN(.TextMatrix(xFil, 10))
                
                If UCase(.TextMatrix(xFil, 5)) = "NOCHE" Then
                    RstDet("turno") = 2
                Else
                    RstDet("turno") = 1
                End If
                
                RstDet.Update
            Next
        End With
    End If

    xCon.CommitTrans
    Set RstDet = Nothing
    
    habilitar TxtFecha, True
    
    MsgBox "Información se grabó con éxito", vbInformation, xTitulo
    Exit Sub
error:
    xCon.RollbackTrans
    habilitar TxtFecha, True
    SHOW_ERROR Me.Name, "pGrabar"
    Set RstDet = Nothing
End Sub

Private Sub pCargarDestajo()

    Dim Rst As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLFiltro As String
    
    Dim mRow&
    
'    On Error GoTo error
    
    Fg2.Rows = Fg2.FixedRows
    Fg3.Rows = Fg3.FixedRows
    DoEvents

    If NulosN(lbl_cod(0).Caption) <> 0 Then
        If OptSeleccion(0).Value = True Then
            nSQLFiltro = " and pro_controltar.idarea=" & NulosN(lbl_cod(0).Caption) & " "
        Else
            nSQLFiltro = " and pla_empleados.id=" & NulosN(lbl_cod(0).Caption) & " "
        End If

    End If
    
    
    '--actualizar el costo en la base de datos
    pActualizarCostoDestajo1
    
    DoEvents
    lbl(0).Caption = "Procesando: Registros"
    lbl(1).Caption = "Interrumpir = ESC"

    '--generar la consulta para presentar el informe
    nSQL = "SELECT vwtarea.*, vwcosto.canteo, iif(vwtarea.unidxhor = 0 or vwcosto.canteo = 0 or vwcosto.canteo is null ,0, (vwtarea.unidxhor/vwcosto.canteo)*100 ) as Eficiencia  " _
        + vbCr + " FROM ( "
    nSQL = nSQL _
        + vbCr + " SELECT IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk, pro_controltardet.idctr, pro_controltardet.corr, pla_empleados.id AS idemp, pro_controltar.idarea, pro_controltardet.idtar, pro_controltardet.idrec, pro_controltardet.idunimed, pro_controltar.fchtra, mae_area.descripcion AS area, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS personal, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, pro_controltardet.cant AS CantReal, mae_unidades.abrev, pro_controltardet.observacion, " _
        + vbCr + " pro_controltardet.horini, pro_controltardet.horfin, IIf([pro_controltardet].[horini] Is Null Or [pro_controltardet].[horfin] Is Null,'',IIf([pro_controltardet].[horini]<CDate('13:20:00') And [pro_controltardet].[horfin]>CDate('14:00:00'),Format(CDate(Format([pro_controltardet].[horfin]-[pro_controltardet].[horini],'hh:mm:ss'))-CDate('01:00:00'),'hh:mm:ss'),Format([pro_controltardet].[horfin]-[pro_controltardet].[horini],'hh:mm:ss'))) AS difhora, IIf([difhora] Is Null Or [difhora]='',0,Hour([difhora])*60+Minute([difhora])) AS totmin, IIf([CantReal]=0 Or [totmin]=0,0,[CantReal]/[totmin]) AS UnidXMin, [UnidXMin]*60 AS UnidXHor, pro_controltardet.canpro, mae_unidades_1.abrev AS abrev1, pro_controltardet.preuni, pro_controltardet.imptot " _
        + vbCr + " FROM ((pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN (alm_inventario RIGHT JOIN ((((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr) LEFT JOIN mae_unidades AS mae_unidades_1 ON pro_controltardet.idunid = mae_unidades_1.id " _
        + vbCr + " WHERE ((pro_controltar.tipo =2 AND pro_controltardet.tipo =1 ) AND ((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) ) " & nSQLFiltro
    nSQL = nSQL _
        + vbCr + " UNION "
    nSQL = nSQL _
        + vbCr + " SELECT IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk, pro_controltardet.idctr, pro_controltardet.corr, pla_empleados.id AS idemp, pro_controltar.idarea, pro_controltardet.idtar, pro_controltardet.idrec, pro_controltardet.idunimed, pro_controltar.fchtra, mae_area.descripcion AS area, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS personal, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, pro_controltardetgr.cant AS CantReal, mae_unidades.abrev, pro_controltardet.observacion,  " _
        + vbCr + "  pro_controltardetgr.horini, pro_controltardetgr.horfin, IIf([pro_controltardetgr].[horini] Is Null Or [pro_controltardetgr].[horfin] Is Null,'',IIf([pro_controltardetgr].[horini]<CDate('13:20:00') And [pro_controltardetgr].[horfin]>CDate('14:00:00'),Format(CDate(Format([pro_controltardetgr].[horfin]-[pro_controltardetgr].[horini],'hh:mm:ss'))-CDate('01:00:00'),'hh:mm:ss'),Format([pro_controltardetgr].[horfin]-[pro_controltardetgr].[horini],'hh:mm:ss'))) AS difhora, IIf([difhora] Is Null Or [difhora]='',0,Hour([difhora])*60+Minute([difhora])) AS totmin, IIf([CantReal]=0 Or [totmin]=0,0,[CantReal]/[totmin]) AS UnidXMin, [UnidXMin]*60 AS UnidXHor, pro_controltardetgr.canpro, mae_unidades_1.abrev AS abrev1, pro_controltardetgr.preuni, pro_controltardetgr.imptot " _
        + vbCr + " FROM ((pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN ((alm_inventario RIGHT JOIN (((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (pro_controltardetgr LEFT JOIN pla_empleados ON pro_controltardetgr.idper = pla_empleados.id) ON (pro_controltardet.corr = pro_controltardetgr.corr) AND (pro_controltardet.idctr = pro_controltardetgr.idctr)) ON pro_controltar.id = pro_controltardet.idctr) LEFT JOIN mae_unidades AS mae_unidades_1 ON pro_controltardetgr.idunid = mae_unidades_1.id " _
        + vbCr + " WHERE pro_controltar.tipo =2 AND (((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltardet.idtar)<>0) AND ((pro_controltardet.tipo)=2) AND ((pro_controltardetgr.activo)=-1)) " & nSQLFiltro
    nSQL = nSQL _
        + vbCr + " ) AS vwtarea "
    nSQL = nSQL _
        + vbCr + " Left Join " _
        + vbCr + " (SELECT IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] & '*' & [pro_costodet].[idunimed] AS codigopk, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa') AS origen, IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas_1].[descripcion]) AS referencia, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio, pro_costodet.jornal, pro_costodet.cant AS canteo, pro_costodet.costo, pro_costodet.orden " _
        + vbCr + " FROM pro_tareas INNER JOIN ((alm_inventario RIGHT JOIN ((pro_costo LEFT JOIN pro_receta ON pro_costo.idref = pro_receta.id) LEFT JOIN pro_tareas AS pro_tareas_1 ON pro_costo.idref = pro_tareas_1.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costodet.idtar " _
        + vbCr + " ) AS vwcosto"
    nSQL = nSQL _
        + vbCr + " ON vwtarea.codigopk = vwcosto.codigopk "
            
    nSQL = nSQL _
     + vbCr + " ORDER BY vwtarea.fchtra, vwtarea.area,vwtarea.personal,vwtarea.horini; "
    
    
    
    RST_Busq Rst, nSQL, xCon
    
    PgBar.Min = 0
    PgBar.Value = 0
    If Rst.RecordCount <> 0 Then PgBar.Max = Rst.RecordCount
    With Fg2
        
        Do While Not Rst.EOF
            DoEvents
            If BAND_INTERRUMPIR = True Then GoTo SALIR
            PgBar.Value = PgBar.Value + 1
            '------------------------------------------------
            .Rows = .Rows + 1
            .TextMatrix(Fg2.Rows - 1, 1) = NulosN(Rst("idemp"))
            .TextMatrix(Fg2.Rows - 1, 2) = Format(Rst("fchtra"), FORMAT_DATE)
            .TextMatrix(Fg2.Rows - 1, 3) = NulosC(Rst("area"))
            .TextMatrix(Fg2.Rows - 1, 4) = NulosC(Rst("personal"))

            .TextMatrix(Fg2.Rows - 1, 5) = NulosC(Rst("tarea")) 'Tarea
            '--------
            .TextMatrix(Fg2.Rows - 1, 6) = NulosC(Rst("Producto")) 'Producto
            
            .TextMatrix(Fg2.Rows - 1, 7) = NulosC(Rst("observacion")) 'observacion
            
            
            .TextMatrix(Fg2.Rows - 1, 8) = Format(Rst("horini"), FORMAT_HORA_SIN_SEGUNDO)   'hini
            .TextMatrix(Fg2.Rows - 1, 9) = Format(Rst("horfin"), FORMAT_HORA_SIN_SEGUNDO)   'hfin
            
            
            .TextMatrix(Fg2.Rows - 1, 10) = Format(NulosN(Rst("cantreal")), FORMAT_MONTO) 'Cantidad
            .TextMatrix(Fg2.Rows - 1, 11) = Rst("abrev") 'U.M.
            
            
            .TextMatrix(Fg2.Rows - 1, 12) = Format(Rst("difhora"), FORMAT_HORA_LARGO)    'difhora
            .TextMatrix(Fg2.Rows - 1, 13) = NulosN(Rst("totmin")) 'Tot.Min"
            .TextMatrix(Fg2.Rows - 1, 14) = NulosN(Rst("unidxmin")) 'Unid x Min
            .TextMatrix(Fg2.Rows - 1, 15) = Format(NulosN(Rst("unidxhor")), "#,##0.00000") 'Unid x Hor
            .TextMatrix(Fg2.Rows - 1, 16) = Format(Rst("canteo"), FORMAT_MONTO) 'Cant Teo
            
            If NulosN(Rst.Fields("eficiencia")) = 100 Then  '--negro
                .TextMatrix(Fg2.Rows - 1, 17) = Format(Rst.Fields("eficiencia"), FORMAT_PORCENTAJE) & "%"
            ElseIf NulosN(Rst.Fields("eficiencia")) = 0 Then  '--no mostrar datos
                
            ElseIf NulosN(Rst.Fields("eficiencia")) > 100 Then '--azul (supera la eficiencia)
                FORMATO_CELDA Fg2, .Rows - 1, 17, &HFF0000, False, &HFFFFFF, Format(NulosN(Rst.Fields("eficiencia")), FORMAT_PORCENTAJE) + "%"
            ElseIf NulosN(Rst.Fields("eficiencia")) < 100 Then '--rojo (menos eficiente)
                FORMATO_CELDA Fg2, .Rows - 1, 17, &HFF, False, &HFFFFFF, Format(NulosN(Rst.Fields("eficiencia")), FORMAT_PORCENTAJE) + "%"
            End If
                
            
            .TextMatrix(Fg2.Rows - 1, 19) = Format(NulosN(Rst("canpro")), FORMAT_MONTO)
            .TextMatrix(Fg2.Rows - 1, 20) = NulosC(Rst("abrev1"))
            .TextMatrix(Fg2.Rows - 1, 21) = Format(NulosN(Rst("preuni")), "0.000000") 'Pre.Uni
            .TextMatrix(Fg2.Rows - 1, 22) = Format(NulosN(Rst("imptot")), FORMAT_MONTO) 'Total
                

            Rst.MoveNext
        Loop
    End With
    Set Rst = Nothing
    '----------
    GRID_AGRUPAR Fg2, 4
    
    '--pintar los montos =0
    
    For mRow = Fg2.FixedRows To Fg2.Rows - 1
        If NulosN(Fg2.TextMatrix(mRow, 22)) = 0 Then
            GRID_COLOR_FONDO Fg2, mRow, 22, mRow, 22, vbRed
        End If
    Next
    '--colocando los totales
    Fg2.Rows = Fg2.Rows + 1
    Fg2.TextMatrix(Fg2.Rows - 1, 4) = "Totales"
    Fg2.TextMatrix(Fg2.Rows - 1, 22) = Format(GRID_SUMAR_COL(Fg2, 22), FORMAT_MONTO)
    
    GRID_COLOR_FONDO Fg2, Fg2.Rows - 1, 1, Fg2.Rows - 1, Fg2.Cols - 1, vbGreen
    
    '----------------------------------------------------------------------------------------
    '-- colocando los datos en el resumen
    
    nSQL = "SELECT vwtarea.idemp, vwtarea.idarea, vwtarea.fchtra, vwtarea.area, vwtarea.personal, Sum(vwtarea.total) AS toting, IIf([vwbono].[impbon] Is Null,0,[vwbono].[impbon]) AS totbono, [toting]+[totbono] AS totneto, First(vwtarea.hinipri) AS hinipri1, First(vwtarea.hfinpri) AS hfinpri1, Last(vwtarea.hiniult) AS hiniult1, Last(vwtarea.hfinult) AS hfinult1 " _
        + vbCr + " FROM ( "
    nSQL = nSQL _
        + vbCr + " SELECT Format([pro_controltar].[fchtra],'dd/mm/yy') & '-' & [pla_empleados].[id] AS codigopk, pla_empleados.id AS idemp, pro_controltar.idarea, pro_controltar.fchtra, mae_area.descripcion AS area, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS personal, pro_controltardet.imptot AS total, pro_controltardet.horini AS hinipri, pro_controltardet.horfin AS hfinpri, pro_controltardet.horini AS hiniult, pro_controltardet.horfin AS hfinult " _
        + vbCr + " FROM (pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN (pro_controltardet LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id) ON pro_controltar.id = pro_controltardet.idctr " _
        + vbCr + " WHERE (((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltar.tipo)=2) AND ((pro_controltardet.tipo)=1)) " & nSQLFiltro
    nSQL = nSQL _
        + vbCr + " UNION "
    nSQL = nSQL _
        + vbCr + " SELECT Format([pro_controltar].[fchtra],'dd/mm/yy') & '-' & [pla_empleados].[id] AS codigopk, pla_empleados.id AS idemp, pro_controltar.idarea, pro_controltar.fchtra, mae_area.descripcion AS area, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS personal, pro_controltardetgr.imptot AS total, pro_controltardetgr.horini AS hinipri, pro_controltardetgr.horfin AS hfinpri, pro_controltardetgr.horini AS hiniult, pro_controltardetgr.horfin AS hfinult " _
        + vbCr + " FROM (pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN (pro_controltardet INNER JOIN (pro_controltardetgr LEFT JOIN pla_empleados ON pro_controltardetgr.idper = pla_empleados.id) ON (pro_controltardet.corr = pro_controltardetgr.corr) AND (pro_controltardet.idctr = pro_controltardetgr.idctr)) ON pro_controltar.id = pro_controltardet.idctr " _
        + vbCr + " WHERE (((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltar.tipo)=2) AND ((pro_controltardet.tipo)=2) AND ((pro_controltardetgr.activo)=-1)) " & nSQLFiltro
    nSQL = nSQL _
        + vbCr + " ) AS vwtarea "
    nSQL = nSQL _
        + vbCr + " Left Join " _
        + vbCr + " (SELECT Format([pro_pagos].[fchtra],'dd/mm/yy') & '-' & [pro_pagos].[idemp] AS codigopk, pro_pagos.idemp, pro_pagos.fchtra, pro_pagos.impbon " _
        + vbCr + " FROM pro_pagos WHERE (((pro_pagos.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "'))) " _
        + vbCr + " GROUP BY Format([pro_pagos].[fchtra],'dd/mm/yy') & '-' & [pro_pagos].[idemp], pro_pagos.idemp, pro_pagos.fchtra, pro_pagos.impbon  " _
        + vbCr + " ) AS vwbono"
    nSQL = nSQL _
        + vbCr + " ON vwtarea.codigopk = vwbono.codigopk "
            
    nSQL = nSQL _
     + vbCr + " GROUP BY vwtarea.idemp, vwtarea.idarea, vwtarea.fchtra, vwtarea.area, vwtarea.personal, IIf([vwbono].[impbon] Is Null,0,[vwbono].[impbon]) " _
     + vbCr + " ORDER BY vwtarea.fchtra, vwtarea.area, vwtarea.personal, First(vwtarea.hinipri); "
    
    
    
    RST_Busq Rst, nSQL, xCon
    
    If Rst.RecordCount = 0 Then Exit Sub
    Agregando = True
    PgBar.Min = 0
    PgBar.Value = 0
    PgBar.Max = Rst.RecordCount
    
    
    With Fg3
        Do While Not Rst.EOF
            DoEvents
            If BAND_INTERRUMPIR = True Then GoTo SALIR
            PgBar.Value = PgBar.Value + 1
            '------------------------------------------------
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = Rst.Bookmark
            .TextMatrix(.Rows - 1, 2) = Format(Rst("fchtra"), FORMAT_DATE)
            .TextMatrix(.Rows - 1, 3) = NulosC(Rst("area"))
            .TextMatrix(.Rows - 1, 4) = NulosC(Rst("personal"))

            '--determinar el horario
            
            If IsDate(Rst("hfinult1")) = True Then
                '--evaluando los las horas
                If CDate(Rst("hfinult1")) < CDate("10:00") Then
                    .TextMatrix(.Rows - 1, 5) = "Noche"
                Else
                    .TextMatrix(.Rows - 1, 5) = "Dia"
                End If
            End If
            
            '--------
            .TextMatrix(.Rows - 1, 6) = Format(NulosN(Rst("toting")), FORMAT_MONTO) 'totingreso
            .TextMatrix(.Rows - 1, 7) = Format(NulosN(Rst("totbono")), FORMAT_MONTO) 'totbono
            .TextMatrix(.Rows - 1, 8) = Format(NulosN(Rst("totneto")), FORMAT_MONTO) 'totneto
            
'            '--si es de noche el costo de la tarea incrementar en un porentaje ejm 30%
'            If .TextMatrix(.Rows - 1, 5) = "Noche" Then
'                .TextMatrix(.Rows - 1, 6) = Format(NulosN(.TextMatrix(.Rows - 1, 6)) * 1.3, FORMAT_MONTO)
'                .TextMatrix(.Rows - 1, 8) = Format(NulosN(.TextMatrix(.Rows - 1, 6)) + NulosN(.TextMatrix(.Rows - 1, 7)), FORMAT_MONTO)
'            End If
'
            
            .TextMatrix(.Rows - 1, 9) = NulosN(Rst("idemp"))
            .TextMatrix(.Rows - 1, 10) = NulosN(Rst("idarea"))
            
            Rst.MoveNext
        Loop
    End With
    Set Rst = Nothing
    '----------
    GRID_AGRUPAR Fg3, 3
    
    '--colocando los totales
    Fg3.Rows = Fg3.Rows + 1
    Fg3.TextMatrix(Fg3.Rows - 1, 4) = "Totales"
    Fg3.TextMatrix(Fg3.Rows - 1, 6) = Format(GRID_SUMAR_COL(Fg3, 6), FORMAT_MONTO)
    Fg3.TextMatrix(Fg3.Rows - 1, 7) = Format(GRID_SUMAR_COL(Fg3, 7), FORMAT_MONTO)
    Fg3.TextMatrix(Fg3.Rows - 1, 8) = Format(GRID_SUMAR_COL(Fg3, 8), FORMAT_MONTO)
    
    GRID_COLOR_FONDO Fg3, Fg3.Rows - 1, 1, Fg3.Rows - 1, Fg3.Cols - 1, vbGreen
    
    
    '----------------------------------------------------------------------------------------
    
    
SALIR:
Agregando = False
Exit Sub
error:
    SHOW_ERROR Me.Name, "pCargarDestajo"
    
End Sub

Private Sub pActualizarCostoDestajo()

    Dim Rst As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLFiltro  As String
    
    '--actualizando el destajo individual
    lbl(0).Caption = "Actualizando Costos 1/2"
    lbl(1).Caption = "No Interrumpir"
    DoEvents
    
    If NulosN(lbl_cod(0).Caption) <> 0 Then
        If OptSeleccion(0).Value = True Then
            nSQLFiltro = " and pro_controltar.idarea=" & NulosN(lbl_cod(0).Caption) & " "
        Else
            nSQLFiltro = " and pla_empleados.id=" & NulosN(lbl_cod(0).Caption) & " "
        End If
        
    End If
    
    
    nSQL = "SELECT vwtarea.idctr,vwtarea.corr, vwtarea.idemp, vwtarea.idarea, vwtarea.idtar, vwtarea.idrec, vwtarea.idunimed, vwtarea.fchtra, vwtarea.personal, vwtarea.area, vwtarea.producto, vwtarea.tarea, iif(vwcosto.paghor=0 or vwcosto.paghor is null,vwtarea.CantReal,vwtarea.tothor) AS canpro, vwtarea.abrev,  iif(vwcosto.paghor=0 or vwcosto.paghor is null,vwcosto.costo,0)  AS preuni, [preuni]*[canpro] AS tot ,  vwcosto.paghor,vwtarea.tothor, iif(vwcosto.paghor=0 or vwcosto.paghor is null,vwtarea.idunimed,7)  as idunid " _
        + vbCr + " FROM ( "
    nSQL = nSQL _
        + vbCr + " SELECT IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk, pro_controltardet.idctr, pro_controltardet.corr,pla_empleados.id AS idemp, pro_controltar.idarea, pro_controltardet.idtar, pro_controltardet.idrec, pro_controltardet.idunimed, pro_controltar.fchtra, mae_area.descripcion AS area, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS personal, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, pro_controltardet.cant AS CantReal, mae_unidades.abrev,pro_controltardet.tothor " _
        + vbCr + " FROM (pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN (alm_inventario RIGHT JOIN ((((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr " _
        + vbCr + " WHERE ((pro_controltar.tipo =2 AND pro_controltardet.tipo =1 ) AND ((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltardet.cant)<>0)) " & nSQLFiltro
    nSQL = nSQL _
        + vbCr + " ) AS vwtarea "
    nSQL = nSQL _
        + vbCr + " Left Join " _
        + vbCr + " (SELECT IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] & '*' & [pro_costodet].[idunimed] AS codigopk, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa') AS origen, IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas_1].[descripcion]) AS referencia, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio, pro_costodet.jornal, pro_costodet.cant AS canteo, pro_costodet.costo, pro_costodet.orden,pro_costodet.paghor " _
        + vbCr + " FROM pro_tareas INNER JOIN ((alm_inventario RIGHT JOIN ((pro_costo LEFT JOIN pro_receta ON pro_costo.idref = pro_receta.id) LEFT JOIN pro_tareas AS pro_tareas_1 ON pro_costo.idref = pro_tareas_1.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costodet.idtar " _
        + vbCr + " ) AS vwcosto"
    nSQL = nSQL _
        + vbCr + " ON vwtarea.codigopk = vwcosto.codigopk "
            
    nSQL = nSQL _
     + vbCr + " ORDER BY vwtarea.fchtra, vwtarea.personal, vwtarea.area, vwtarea.producto; "
    
    
    RST_Busq RstTmp, nSQL, xCon
    
    PgBar.Min = 0
    PgBar.Value = 0
    
    Dim xHoras As Double
    If RstTmp.RecordCount <> 0 Then
        PgBar.Max = RstTmp.RecordCount
        RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            PgBar.Value = RstTmp.Bookmark
            xCon.Execute "update pro_controltardet set canpro=" & NulosN(RstTmp("canpro")) & " ,idunid=" & NulosN(RstTmp("idunid")) & ", preuni = " & NulosN(RstTmp("preuni")) & ", imptot =" & NulosN(RstTmp("tot")) & " where idctr = " & NulosN(RstTmp("idctr")) & " and corr = " & NulosN(RstTmp("corr")) & " and tipo=1 and idref = " & NulosN(RstTmp("idemp")) & " and idtar = " & NulosN(RstTmp("idtar")) & " and idrec = " & NulosN(RstTmp("idrec")) & " and idunimed =" & NulosN(RstTmp("idunimed"))
            RstTmp.MoveNext
        Loop
    End If
    Set RstTmp = Nothing
    
    
    '-----------
    '--actualizando el destajo grupal
    lbl(0).Caption = "Actualizando Costos 2/2"
    DoEvents
    
    nSQL = "SELECT vwtarea.idctr,vwtarea.corr, vwtarea.idemp, vwtarea.idarea, vwtarea.idtar, vwtarea.idrec, vwtarea.idunimed, vwtarea.fchtra, vwtarea.personal, vwtarea.area, vwtarea.producto, vwtarea.tarea, iif(vwcosto.paghor=0 or vwcosto.paghor is null,vwtarea.CantReal,vwtarea.tothor) AS canpro, vwtarea.abrev,  iif(vwcosto.paghor=0 or vwcosto.paghor is null,vwcosto.costo,0)  AS preuni, [preuni]*[canpro] AS tot ,  vwcosto.paghor,vwtarea.tothor, iif(vwcosto.paghor=0 or vwcosto.paghor is null,vwtarea.idunimed,7)  as idunid " _
        + vbCr + " FROM ( "
    nSQL = nSQL _
        + vbCr + " SELECT IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk,pro_controltardet.idctr, pro_controltardet.corr, pla_empleados.id AS idemp, pro_controltar.idarea, pro_controltardet.idtar, pro_controltardet.idrec, pro_controltardet.idunimed, pro_controltar.fchtra, mae_area.descripcion AS area, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS personal, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, pro_controltardetgr.cant AS CantReal, mae_unidades.abrev,pro_controltardetgr.tothor " _
        + vbCr + " FROM (pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN (alm_inventario RIGHT JOIN ((((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) INNER JOIN (pro_controltardetgr LEFT JOIN pla_empleados ON pro_controltardetgr.idper = pla_empleados.id) ON (pro_controltardet.corr = pro_controltardetgr.corr) AND (pro_controltardet.idctr = pro_controltardetgr.idctr)) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr " _
        + vbCr + " WHERE pro_controltar.tipo =2 AND (((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltardet.idtar)<>0) AND ((pro_controltardet.tipo)=2) AND ((pro_controltardetgr.cant)<>0) AND ((pro_controltardetgr.activo)=-1)) " & nSQLFiltro
    nSQL = nSQL _
        + vbCr + " ) AS vwtarea "
    nSQL = nSQL _
        + vbCr + " Left Join " _
        + vbCr + " (SELECT IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] & '*' & [pro_costodet].[idunimed] AS codigopk, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa') AS origen, IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas_1].[descripcion]) AS referencia, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio, pro_costodet.jornal, pro_costodet.cant AS canteo, pro_costodet.costo, pro_costodet.orden,pro_costodet.paghor " _
        + vbCr + " FROM pro_tareas INNER JOIN ((alm_inventario RIGHT JOIN ((pro_costo LEFT JOIN pro_receta ON pro_costo.idref = pro_receta.id) LEFT JOIN pro_tareas AS pro_tareas_1 ON pro_costo.idref = pro_tareas_1.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costodet.idtar " _
        + vbCr + " ) AS vwcosto"
    nSQL = nSQL _
        + vbCr + " ON vwtarea.codigopk = vwcosto.codigopk "
            
    nSQL = nSQL _
     + vbCr + " ORDER BY vwtarea.fchtra, vwtarea.personal, vwtarea.area, vwtarea.producto; "
    
    
    RST_Busq RstTmp, nSQL, xCon
    
    If RstTmp.RecordCount <> 0 Then
        PgBar.Min = 0
        PgBar.Value = 0
        PgBar.Max = RstTmp.RecordCount
        
        RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            
            PgBar.Value = RstTmp.Bookmark
            
            xCon.Execute "UPDATE pro_controltardet INNER JOIN pro_controltardetgr ON (pro_controltardet.idctr = pro_controltardetgr.idctr) AND (pro_controltardet.corr = pro_controltardetgr.corr) SET pro_controltardetgr.canpro=" & NulosN(RstTmp("canpro")) & ",pro_controltardetgr.idunid=" & NulosN(RstTmp("idunid")) & ", pro_controltardetgr.preuni = " & NulosN(RstTmp("preuni")) & ", pro_controltardetgr.imptot = " & NulosN(RstTmp("tot")) _
                & " WHERE (((pro_controltardet.idctr)=" & NulosN(RstTmp("idctr")) & ") AND ((pro_controltardet.corr)=" & NulosN(RstTmp("corr")) & ") AND " _
                & " ((pro_controltardet.idtar)=" & NulosN(RstTmp("idtar")) & ") AND ((pro_controltardet.idrec)=" & NulosN(RstTmp("idrec")) & " ) AND " _
                & " ((pro_controltardetgr.idper)=" & NulosN(RstTmp("idemp")) & ") AND ((pro_controltardet.idunimed)=" & NulosN(RstTmp("idunimed")) & " ) AND " _
                & " ((pro_controltardet.tipo)=2)); "

            
            RstTmp.MoveNext
        Loop
    End If
    Set RstTmp = Nothing
        
SALIR:
Agregando = False

End Sub







Private Sub pActualizarCostoDestajo1()

    Dim Rst As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLFiltro  As String
    
    '--actualizando el destajo individual
    lbl(0).Caption = "Actualizando Costos 1/2"
    lbl(1).Caption = "No Interrumpir"
    DoEvents
    
    If NulosN(lbl_cod(0).Caption) <> 0 Then
        If OptSeleccion(0).Value = True Then
            nSQLFiltro = " and pro_controltar.idarea=" & NulosN(lbl_cod(0).Caption) & " "
        Else
            nSQLFiltro = " and pla_empleados.id=" & NulosN(lbl_cod(0).Caption) & " "
        End If
        
    End If
    
    
    nSQL = "SELECT vwtarea.idctr,vwtarea.corr, vwtarea.idemp, vwtarea.idarea, vwtarea.idtar, vwtarea.idrec, vwtarea.idunimed, vwtarea.fchtra, vwtarea.personal, vwtarea.area, vwtarea.producto, vwtarea.tarea, iif(vwcosto.paghor=0 or vwcosto.paghor is null,vwtarea.CantReal,vwtarea.tothor) AS canpro, vwtarea.abrev,  IIf([vwcosto].[paghor]=0,[vwcosto].[costo],[vwcostoh].[costo]) AS preuni, [preuni]*[canpro] AS tot ,  vwcosto.paghor,vwtarea.tothor, iif(vwcosto.paghor=0,vwtarea.idunimed,7)  as idunid " _
        + vbCr + " FROM (( "
    nSQL = nSQL _
        + vbCr + " SELECT IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk,IIf(pro_controltardet.idrec=0 Or pro_controltardet.idrec Is Null,'-',pro_controltardet.idrec) & '*' & pro_controltardet.idtar AS codigopk1, pro_controltardet.idctr, pro_controltardet.corr,pla_empleados.id AS idemp, pro_controltar.idarea, pro_controltardet.idtar, pro_controltardet.idrec, pro_controltardet.idunimed, pro_controltar.fchtra, mae_area.descripcion AS area, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS personal, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, pro_controltardet.cant AS CantReal, mae_unidades.abrev,pro_controltardet.tothor " _
        + vbCr + " FROM (pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN (alm_inventario RIGHT JOIN ((((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr " _
        + vbCr + " WHERE ((pro_controltar.tipo =2 AND pro_controltardet.tipo =1 ) AND ((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltardet.cant)<>0)) " & nSQLFiltro
    nSQL = nSQL _
        + vbCr + " ) AS vwtarea "
    nSQL = nSQL _
        + vbCr + " Left Join " _
        + vbCr + " (SELECT IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] & '*' & [pro_costodet].[idunimed] AS codigopk, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa') AS origen, IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas_1].[descripcion]) AS referencia, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio, pro_costodet.jornal, pro_costodet.cant AS canteo, pro_costodet.costo, pro_costodet.orden,pro_costodet.paghor " _
        + vbCr + " FROM pro_tareas INNER JOIN ((alm_inventario RIGHT JOIN ((pro_costo LEFT JOIN pro_receta ON pro_costo.idref = pro_receta.id) LEFT JOIN pro_tareas AS pro_tareas_1 ON pro_costo.idref = pro_tareas_1.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costodet.idtar " _
        + vbCr + " ) AS vwcosto "
    nSQL = nSQL _
        + vbCr + " ON vwtarea.codigopk = vwcosto.codigopk ) "
            
    nSQL = nSQL _
        + vbCr + " Left Join " _
        + vbCr + " (SELECT IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] & '*' & [pro_costodet].[idunimed] AS codigopk, IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] AS codigopk1, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa') AS origen, IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas_1].[descripcion]) AS referencia, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio, pro_costodet.jornal, pro_costodet.cant AS canteo, pro_costodet.costo, pro_costodet.orden, pro_costodet.paghor " _
        + vbCr + " FROM (alm_inventario RIGHT JOIN ((pro_costo LEFT JOIN pro_receta ON pro_costo.idref = pro_receta.id) LEFT JOIN pro_tareas AS pro_tareas_1 ON pro_costo.idref = pro_tareas_1.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (pro_tareas INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_tareas.id = pro_costodet.idtar) ON pro_costo.id = pro_costodet.idcos " _
        + vbCr + " WHERE (((pro_costodet.idunimed)=7)) " _
        + vbCr + " ) AS vwcostoh "
    nSQL = nSQL _
        + vbCr + "  ON vwtarea.codigopk1 = vwcostoh.codigopk1 "
            
            
    nSQL = nSQL _
     + vbCr + " ORDER BY vwtarea.fchtra, vwtarea.personal, vwtarea.area, vwtarea.producto; "
    
    
    RST_Busq RstTmp, nSQL, xCon
    
    PgBar.Min = 0
    PgBar.Value = 0
    
    Dim xHoras As Double
    If RstTmp.RecordCount <> 0 Then
        PgBar.Max = RstTmp.RecordCount
        RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            PgBar.Value = RstTmp.Bookmark
            xCon.Execute "update pro_controltardet set canpro=" & NulosN(RstTmp("canpro")) & " ,idunid=" & NulosN(RstTmp("idunid")) & ", preuni = " & NulosN(RstTmp("preuni")) & ", imptot =" & NulosN(RstTmp("tot")) & " where idctr = " & NulosN(RstTmp("idctr")) & " and corr = " & NulosN(RstTmp("corr")) & " and tipo=1 and idref = " & NulosN(RstTmp("idemp")) & " and idtar = " & NulosN(RstTmp("idtar")) & " and idrec = " & NulosN(RstTmp("idrec")) & " and idunimed =" & NulosN(RstTmp("idunimed"))
            RstTmp.MoveNext
        Loop
    End If
    Set RstTmp = Nothing
    
    
    '-----------
    '--actualizando el destajo grupal
    lbl(0).Caption = "Actualizando Costos 2/2"
    DoEvents
    
    nSQL = "SELECT vwtarea.idctr,vwtarea.corr, vwtarea.idemp, vwtarea.idarea, vwtarea.idtar, vwtarea.idrec, vwtarea.idunimed, vwtarea.fchtra, vwtarea.personal, vwtarea.area, vwtarea.producto, vwtarea.tarea, iif(vwcosto.paghor=0 or vwcosto.paghor is null,vwtarea.CantReal,vwtarea.tothor) AS canpro, vwtarea.abrev, IIf([vwcosto].[paghor]=0,[vwcosto].[costo],[vwcostoh].[costo]) AS preuni, [preuni]*[canpro] AS tot ,  vwcosto.paghor,vwtarea.tothor, iif(vwcosto.paghor=0,vwtarea.idunimed,7)  as idunid " _
        + vbCr + " FROM (( "
    nSQL = nSQL _
        + vbCr + " SELECT IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk,IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] AS codigopk1, pro_controltardet.idctr, pro_controltardet.corr, pla_empleados.id AS idemp, pro_controltar.idarea, pro_controltardet.idtar, pro_controltardet.idrec, pro_controltardet.idunimed, pro_controltar.fchtra, mae_area.descripcion AS area, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS personal, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, pro_controltardetgr.cant AS CantReal, mae_unidades.abrev,pro_controltardetgr.tothor " _
        + vbCr + " FROM (pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN (alm_inventario RIGHT JOIN ((((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) INNER JOIN (pro_controltardetgr LEFT JOIN pla_empleados ON pro_controltardetgr.idper = pla_empleados.id) ON (pro_controltardet.corr = pro_controltardetgr.corr) AND (pro_controltardet.idctr = pro_controltardetgr.idctr)) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr " _
        + vbCr + " WHERE pro_controltar.tipo =2 AND (((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltardet.idtar)<>0) AND ((pro_controltardet.tipo)=2) AND ((pro_controltardetgr.cant)<>0) AND ((pro_controltardetgr.activo)=-1)) " & nSQLFiltro
    nSQL = nSQL _
        + vbCr + " ) AS vwtarea "
    nSQL = nSQL _
        + vbCr + " Left Join " _
        + vbCr + " (SELECT IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] & '*' & [pro_costodet].[idunimed] AS codigopk, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa') AS origen, IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas_1].[descripcion]) AS referencia, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio, pro_costodet.jornal, pro_costodet.cant AS canteo, pro_costodet.costo, pro_costodet.orden,pro_costodet.paghor " _
        + vbCr + " FROM pro_tareas INNER JOIN ((alm_inventario RIGHT JOIN ((pro_costo LEFT JOIN pro_receta ON pro_costo.idref = pro_receta.id) LEFT JOIN pro_tareas AS pro_tareas_1 ON pro_costo.idref = pro_tareas_1.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costodet.idtar " _
        + vbCr + " ) AS vwcosto"
    nSQL = nSQL _
        + vbCr + " ON vwtarea.codigopk = vwcosto.codigopk ) "
            
    nSQL = nSQL _
        + vbCr + " Left Join " _
        + vbCr + " (SELECT IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] & '*' & [pro_costodet].[idunimed] AS codigopk, IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] AS codigopk1, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa') AS origen, IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas_1].[descripcion]) AS referencia, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio, pro_costodet.jornal, pro_costodet.cant AS canteo, pro_costodet.costo, pro_costodet.orden, pro_costodet.paghor " _
        + vbCr + " FROM (alm_inventario RIGHT JOIN ((pro_costo LEFT JOIN pro_receta ON pro_costo.idref = pro_receta.id) LEFT JOIN pro_tareas AS pro_tareas_1 ON pro_costo.idref = pro_tareas_1.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (pro_tareas INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_tareas.id = pro_costodet.idtar) ON pro_costo.id = pro_costodet.idcos " _
        + vbCr + " WHERE (((pro_costodet.idunimed)=7)) " _
        + vbCr + " ) AS vwcostoh "
    nSQL = nSQL _
        + vbCr + "  ON vwtarea.codigopk1 = vwcostoh.codigopk1 "
            
    nSQL = nSQL _
     + vbCr + " ORDER BY vwtarea.fchtra, vwtarea.personal, vwtarea.area, vwtarea.producto; "
    
    
    RST_Busq RstTmp, nSQL, xCon
    
    If RstTmp.RecordCount <> 0 Then
        PgBar.Min = 0
        PgBar.Value = 0
        PgBar.Max = RstTmp.RecordCount
        
        RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            
            PgBar.Value = RstTmp.Bookmark
            
            xCon.Execute "UPDATE pro_controltardet INNER JOIN pro_controltardetgr ON (pro_controltardet.idctr = pro_controltardetgr.idctr) AND (pro_controltardet.corr = pro_controltardetgr.corr) SET pro_controltardetgr.canpro=" & NulosN(RstTmp("canpro")) & ",pro_controltardetgr.idunid=" & NulosN(RstTmp("idunid")) & ", pro_controltardetgr.preuni = " & NulosN(RstTmp("preuni")) & ", pro_controltardetgr.imptot = " & NulosN(RstTmp("tot")) _
                & " WHERE (((pro_controltardet.idctr)=" & NulosN(RstTmp("idctr")) & ") AND ((pro_controltardet.corr)=" & NulosN(RstTmp("corr")) & ") AND " _
                & " ((pro_controltardet.idtar)=" & NulosN(RstTmp("idtar")) & ") AND ((pro_controltardet.idrec)=" & NulosN(RstTmp("idrec")) & " ) AND " _
                & " ((pro_controltardetgr.idper)=" & NulosN(RstTmp("idemp")) & ") AND ((pro_controltardet.idunimed)=" & NulosN(RstTmp("idunimed")) & " ) AND " _
                & " ((pro_controltardet.tipo)=2)); "

            
            RstTmp.MoveNext
        Loop
    End If
    Set RstTmp = Nothing
        
SALIR:
Agregando = False

End Sub








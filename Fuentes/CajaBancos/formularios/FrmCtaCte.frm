VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmCtaCte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caja y Bancos - Cuenta Corriente (Cliente, Proveedor)"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraBarra 
      BorderStyle     =   0  'None
      Caption         =   "FrmConsultaDiario"
      Height          =   780
      Left            =   2760
      TabIndex        =   29
      Top             =   3525
      Visible         =   0   'False
      Width           =   6285
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   330
         Left            =   150
         TabIndex        =   30
         Top             =   315
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   582
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   -15
         Y2              =   900
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   6270
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   1
         X1              =   -75
         X2              =   6500
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   6270
         X2              =   6270
         Y1              =   -30
         Y2              =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Procesando Documentos"
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
         Height          =   195
         Left            =   195
         TabIndex        =   32
         Top             =   75
         Width           =   2130
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   2
         Left            =   4605
         TabIndex        =   31
         Top             =   75
         Width           =   1530
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
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
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3285
         Top             =   0
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
               Picture         =   "FrmCtaCte.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte.frx":0544
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte.frx":08D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte.frx":0A30
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte.frx":0DC2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte.frx":0F46
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte.frx":139A
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte.frx":14B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte.frx":19F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte.frx":1F3A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte.frx":204E
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte.frx":2162
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte.frx":25B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte.frx":2722
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1185
      Left            =   30
      TabIndex        =   8
      Top             =   285
      Width           =   11835
      Begin VB.CheckBox chk_descuadrado 
         Caption         =   "Descuadrados"
         Enabled         =   0   'False
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
         Height          =   195
         Left            =   6180
         TabIndex        =   16
         ToolTipText     =   "Mostrará solo los documentos cuyo saldo final es negativo"
         Top             =   840
         Width           =   1545
      End
      Begin VB.Frame Frame3 
         Caption         =   "[  Seleccionar  ]"
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
         Height          =   975
         Left            =   9900
         TabIndex        =   14
         Top             =   150
         Width           =   1830
         Begin VB.OptionButton OptSel1 
            Caption         =   "Todos"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   150
            TabIndex        =   7
            Top             =   270
            Value           =   -1  'True
            Width           =   1560
         End
         Begin VB.OptionButton OptSel2 
            Caption         =   "Seleccionar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   150
            TabIndex        =   15
            Top             =   525
            Width           =   1560
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "[  Seleccionar  ]"
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
         Height          =   975
         Left            =   7920
         TabIndex        =   12
         Top             =   150
         Width           =   1830
         Begin VB.OptionButton OptTodos 
            Caption         =   "Todos"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   225
            TabIndex        =   23
            Top             =   720
            Width           =   1350
         End
         Begin VB.OptionButton OptCan 
            Caption         =   "Cancelados"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   225
            TabIndex        =   6
            Top             =   480
            Width           =   1350
         End
         Begin VB.OptionButton OptPen 
            Caption         =   "Pendientes"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   225
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   1350
         End
      End
      Begin VB.OptionButton OptProvee 
         Caption         =   "Proveedor"
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
         Height          =   195
         Left            =   6180
         TabIndex        =   4
         Top             =   495
         Width           =   1230
      End
      Begin VB.OptionButton OptCliente 
         Caption         =   "Cliente"
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
         Height          =   195
         Left            =   6180
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.CommandButton CmdBusCliPro 
         Enabled         =   0   'False
         Height          =   240
         Left            =   5580
         Picture         =   "FrmCtaCte.frx":2C6A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   210
         Width           =   240
      End
      Begin VB.TextBox TxtCliPro 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1155
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "TxtCliPro"
         Top             =   180
         Width           =   4695
      End
      Begin VB.Frame Frame5 
         Caption         =   "[ Hasta el dia ]"
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
         Height          =   660
         Left            =   75
         TabIndex        =   21
         Top             =   495
         Width           =   2385
         Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
            Height          =   315
            Left            =   930
            TabIndex        =   0
            Top             =   270
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   315
            Width           =   450
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "[ Expresado en ]"
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
         Height          =   660
         Left            =   2505
         TabIndex        =   17
         Top             =   495
         Width           =   3345
         Begin VB.CommandButton CmdBusMon 
            Height          =   240
            Left            =   1275
            Picture         =   "FrmCtaCte.frx":2D9C
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   300
            Width           =   210
         End
         Begin VB.TextBox TxtIdMon 
            Height          =   300
            Left            =   810
            MaxLength       =   1
            TabIndex        =   1
            Text            =   "TxtIdMon"
            Top             =   270
            Width           =   705
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda"
            Height          =   195
            Index           =   4
            Left            =   105
            TabIndex        =   20
            Top             =   315
            Width           =   585
         End
         Begin VB.Label LblMoneda 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblMoneda"
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
            Left            =   1515
            TabIndex        =   19
            Top             =   270
            Width           =   1710
         End
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   6105
         X2              =   7665
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         Index           =   3
         X1              =   6045
         X2              =   6045
         Y1              =   165
         Y2              =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         Index           =   0
         X1              =   6060
         X2              =   6060
         Y1              =   165
         Y2              =   1095
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   6105
         X2              =   7665
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         Index           =   2
         X1              =   7740
         X2              =   7740
         Y1              =   165
         Y2              =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         Index           =   1
         X1              =   7725
         X2              =   7725
         Y1              =   165
         Y2              =   1095
      End
      Begin VB.Label LblIdCliPro 
         Caption         =   "LblIdCliPro"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   4605
         TabIndex        =   11
         Top             =   135
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   9
         Top             =   270
         Width           =   840
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6075
      Left            =   30
      TabIndex        =   24
      Top             =   1470
      Width           =   11850
      _cx             =   20902
      _cy             =   10716
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
      FrontTabColor   =   14215660
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "      Detalle     |      Resumen     "
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
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   5655
         Left            =   45
         TabIndex        =   27
         Top             =   45
         Width           =   11760
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   5475
            Left            =   30
            TabIndex        =   28
            Top             =   90
            Width           =   11685
            _cx             =   20611
            _cy             =   9657
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
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   128
            ForeColorSel    =   16777215
            BackColorBkg    =   -2147483636
            BackColorAlternate=   16777215
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
            Rows            =   2
            Cols            =   13
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmCtaCte.frx":2ECE
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
         Height          =   5655
         Left            =   12495
         TabIndex        =   25
         Top             =   45
         Width           =   11760
         Begin VSFlex7Ctl.VSFlexGrid Fg2 
            Height          =   5580
            Left            =   15
            TabIndex        =   26
            Top             =   30
            Width           =   11850
            _cx             =   20902
            _cy             =   9842
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
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmCtaCte.frx":3055
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
Attribute VB_Name = "FrmCtaCte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstCta As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE


Private Sub CmdBusCliPro_Click()
    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    
    If OptCliente.Value = True Then
        xForm.Titulo = "Buscando Clientes"
        xForm.SQLCad = "SELECT mae_cliente.numruc, mae_cliente.nombre, mae_cliente.id From mae_cliente ORDER BY mae_cliente.nombre"
        xCampos(0, 0) = "Cliente":       xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
    Else
        xForm.Titulo = "Buscando Proveedores"
        xForm.SQLCad = "SELECT mae_prov.numruc, mae_prov.nombre, mae_prov.id From mae_prov ORDER BY mae_prov.nombre"
        xCampos(0, 0) = "Proveedor":     xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
    End If

    xCampos(1, 0) = "Nº R.U.C.":     xCampos(1, 1) = "numruc":        xCampos(1, 2) = "1400":         xCampos(1, 3) = "C"

    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "nombre"
    xForm.CampoBusca = "nombre"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtCliPro.Text = xRs("nombre")
        LblIdCliPro.Caption = xRs("id")
        TxtFecha.SetFocus
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        pConfigurarGrilla
        TxtIdMon.Text = 1
        TxtIdMon_Validate False
        SeEjecuto = True
    End If
End Sub

Sub CargarCli(IdCliPro)
    Dim Rst As New ADODB.Recordset
    Dim Rstabo As New ADODB.Recordset
    Dim A, B, xFila As Integer
    Dim TotDebe, TotHaber As Double
    Dim TotGralDebe, TotGralHaber As Double
    Dim xNomPro As String
    Dim Cambio As Boolean
    Dim nSQL As String
'    On Error GoTo error
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "Seleccione una Moneda", vbExclamation, xTitulo
        TxtIdMon.SetFocus
        Exit Sub
    End If
    BAND_INTERRUMPIR = False
    pConfigurarGrilla
    '--------------------------
    fraBarra.Left = 2798
    fraBarra.Top = 2925
    
    ProgressBar1.Max = 1
    ProgressBar1.Value = 0
    fraBarra.Visible = True
    fraBarra.Refresh
    
    
    Dim nSQLWhere As String
    Dim nCampoMuestra As String '--indica el campo que se mostrara esta en funcion de la moneda seleccionada
    nSQLWhere = ""
    If OptCliente.Value = True Then '--ventas
        If IdCliPro <> 0 Then nSQLWhere = " and vta_ventas.idcli = " & IdCliPro & " "
        nSQL = "SELECT vta_ventas.id,IIf([vta_ventas]![anulado]=-1,' ',[mae_cliente]![numruc]) AS numruc, IIf([vta_ventas]![anulado]=-1,'Anulado',[mae_cliente]![nombre]) AS nombre, IIf([vta_ventas].[numreg] Is Null Or [vta_ventas].[numreg]='','',Left([vta_ventas].[numreg],2) & [mae_libros].[codsun] & Right([vta_ventas].[numreg],4)) AS registro, 'Ventas' AS libro, mae_documento.codsun,mae_documento.abrev, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc2, vta_ventas.fchdoc, vta_ventas.fchven, mae_moneda.simbolo, con_tc.impven AS tipcam, " _
            + vbCr + " vta_ventas.idmon,vta_ventas.imptotdoc AS imptotal,vta_ventas.impsal, " _
            + vbCr + " IIf([vta_ventas].[imptotdoc] Is Null,0,IIf([vta_ventas].[idmon]=1,[vta_ventas].[imptotdoc],IIf([con_tc].[impven] Is Null,0,[vta_ventas].[imptotdoc]*[con_tc].[impven]))) AS imptotsol, " _
            + vbCr + " IIf([vta_ventas].[imptotdoc] Is Null,0,IIf([vta_ventas].[idmon]=2,[vta_ventas].[imptotdoc],IIf([con_tc].[impven] Is Null,0,[vta_ventas].[imptotdoc]/[con_tc].[impven]))) AS imptotdol " _
            + vbCr + " FROM ((((vta_ventas LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id " _
            + vbCr + " WHERE vta_ventas.fchdoc <= CDate('" & TxtFecha.Valor & "') " & nSQLWhere _
            + vbCr + " ORDER BY IIf([vta_ventas]![anulado]=-1,'Anulado',[mae_cliente]![nombre]), [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc];"
    
    Else '--compras
        If IdCliPro <> 0 Then nSQLWhere = " and com_compras.idpro = " & IdCliPro & " "
        
        nSQL = "SELECT  com_compras.id,mae_prov.numruc, mae_prov!nombre AS nombre, IIf(com_compras.numreg Is Null Or com_compras.numreg='','',Left(com_compras.numreg,2) & mae_libros.codsun & Right(com_compras.numreg,4)) AS registro, 'Compras' AS libro, mae_documento.codsun,mae_documento.abrev, com_compras!numser+'-'+com_compras!numdoc AS numdoc2, com_compras.fchdoc, com_compras.fchven, mae_moneda.simbolo, con_tc.impven AS tipcam, " _
            + vbCr + " com_compras.idmon,com_compras.imptot AS imptotal,com_compras.impsal, " _
            + vbCr + " IIf([com_compras].[imptot] Is Null,0,IIf([com_compras].[idmon]=1,[com_compras].[imptot],IIf([con_tc].[impven] Is Null,0,[com_compras].[imptot]*[con_tc].[impven]))) AS imptotsol, " _
            + vbCr + " IIf([com_compras].[imptot] Is Null,0,IIf([com_compras].[idmon]=2,[com_compras].[imptot],IIf([con_tc].[impven] Is Null,0,[com_compras].[imptot]/[con_tc].[impven]))) AS imptotdol " _
            + vbCr + " FROM (mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN (mae_documento RIGHT JOIN (com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON mae_documento.id = com_compras.tipdoc) ON mae_prov.id = com_compras.idpro) ON mae_moneda.id = com_compras.idmon) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha " _
            + vbCr + " WHERE (com_compras.fchdoc <=CDate('" & TxtFecha.Valor & "') )" & nSQLWhere & "AND ( com_compras.tipdoc <> 7) " _
            + vbCr + " ORDER BY mae_prov!nombre, com_compras.fchdoc;"
        
'WHERE (((com_compras.fchdoc)<=CDate('17/12/2008')) AND ((com_compras.idpro)=1176) AND ((com_compras.tipdoc)<>7))
    End If
    
    If NulosN(TxtIdMon.Text) = 1 Then
        nCampoMuestra = "imptotsol"
    ElseIf NulosN(TxtIdMon.Text) = 2 Then
        nCampoMuestra = "imptotdol"
    Else
        fraBarra.Visible = False
        MsgBox "Por el momento no se puede expresar en " & LblMoneda.Caption, vbInformation, xTitulo
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    '--ejecutar la conulta
    RST_Busq Rst, nSQL, xCon
    '--filtrar lo que se va mostrar
    If chk_descuadrado.Value = 0 Then
        '--obs. si selecciona la opcion todos no hace el fintro
        If OptPen.Value = True Then Rst.Filter = "impsal > 0" ' FILTRAMOS LOS PENDIENTE
        If OptCan.Value = True Then Rst.Filter = "impsal <= 0" ' FILTRAMOS LOS CANCELADOS
    End If
    If Rst.RecordCount = 0 Then
        MsgBox "No hay documentos del cliente seleccionado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fraBarra.Visible = False
        Set Rst = Nothing
        Exit Sub
    End If
    ProgressBar1.Max = Rst.RecordCount
    
    Dim xSaldoDoc As Double
    Dim xFilaIni&
    Dim xColor&
    
    Me.MousePointer = vbHourglass
     
    xColor = 0
    If Rst.RecordCount <> 0 Then
        DoEvents
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then GoTo Salir:

        Rst.MoveFirst
        xSaldoDoc = 0
        xNomPro = NulosC(Rst("nombre"))
        xFila = Fg1.FixedRows
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(xFila, 1) = "Nº R.U.C. :"
        Fg1.TextMatrix(xFila, 2) = NulosC(Rst("numruc"))
        Fg1.TextMatrix(xFila, 4) = NulosC(Rst("nombre"))
        '*****resumen
        Fg2.Rows = Fg2.Rows + 1
        Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosC(Rst("numruc"))
        Fg2.TextMatrix(Fg2.Rows - 1, 2) = NulosC(Rst("nombre"))
        '******
        xFilaIni = xFila
        With Fg1
            .Cell(flexcpForeColor, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1) = &H800000
            .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
            .FillStyle = flexFillRepeat
            .CellFontBold = True
        End With
    
        TotDebe = 0
        TotHaber = 0
        
        Cambio = False
        
        Dim mRowIni As Integer

        For A = 1 To Rst.RecordCount    '--GRUPO DE CLIENTE/PROVEEDOR
            DoEvents
            '--SI SE NTERRUMPE EL PROCESO => SALIR
            If BAND_INTERRUMPIR = True Then GoTo Salir:
            ProgressBar1.Value = A
            
            xSaldoDoc = 0
            
            If NulosC(Rst("nombre")) <> xNomPro Then
                DoEvents
                Cambio = True
                xNomPro = NulosC(Rst("nombre"))
                Fg1.Rows = Fg1.Rows + 1
                xFila = xFila + 1
                Fg1.TextMatrix(xFila, 4) = "TOTAL -->"
                Fg1.TextMatrix(xFila, 10) = Format(TotDebe, FORMAT_MONTO)
                Fg1.TextMatrix(xFila, 11) = Format(TotHaber, FORMAT_MONTO)
                Fg1.TextMatrix(xFila, 12) = Format(TotDebe - TotHaber, FORMAT_MONTO)
                
                '*****resumen
                Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(TotDebe, FORMAT_MONTO)
                Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(TotHaber, FORMAT_MONTO)
                Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotDebe - TotHaber, FORMAT_MONTO)
                '******

                
                With Fg1
                    .Cell(flexcpForeColor, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1) = &H80000008
                    .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
                    .FillStyle = flexFillRepeat
                    .CellFontBold = True
                End With
                
                '----MOSTRAR SOLO DESCUADRADOS ---------
                If chk_descuadrado.Value = 1 Then
                    If NulosN(Fg1.TextMatrix(xFila, 12)) = 0 Then
                        GRID_DELETE Fg1, Fg1.Rows - 2, Fg1.Rows - 1, e_Fila
                        Fg1.Rows = Fg1.Rows + 1
                        xFila = Fg1.Rows - 1
                    Else
                        Fg1.Rows = Fg1.Rows + 2
                        xFila = xFila + 2
                    End If
                    '---del resumen
                    If NulosN(Fg2.TextMatrix(Fg2.Rows - 1, 5)) = 0 Then
                        GRID_DELETE Fg2, Fg2.Rows - 1, Fg2.Rows - 1, e_Fila
                    End If
                    '---------------
                Else
                    Fg1.Rows = Fg1.Rows + 2
                    xFila = xFila + 2
                End If
                '---------------------------------------------------------
                TotGralHaber = TotGralHaber + TotHaber
                TotGralDebe = TotGralDebe + TotDebe
                
                TotHaber = 0
                TotDebe = 0
                '---------------------------------------------------------
                Fg1.TextMatrix(xFila, 1) = "Nº R.U.C. :"
                Fg1.TextMatrix(xFila, 2) = NulosC(Rst("numruc"))
                Fg1.TextMatrix(xFila, 4) = xNomPro
                
                '*****resumen
                Fg2.Rows = Fg2.Rows + 1
                Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosC(Rst("numruc"))
                Fg2.TextMatrix(Fg2.Rows - 1, 2) = xNomPro
                '******

                
                With Fg1
                    .Cell(flexcpForeColor, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1) = &H800000
                    .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
                    .FillStyle = flexFillRepeat
                    .CellFontBold = True
                End With
            Else
                Cambio = False
            End If

            Fg1.Rows = Fg1.Rows + 1
            xFila = xFila + 1
            xFilaIni = xFila
            
            Fg1.TextMatrix(xFila, 1) = NulosC(Rst("registro"))
            
            Fg1.TextMatrix(xFila, 2) = NulosC(Rst("libro"))
            Fg1.TextMatrix(xFila, 3) = NulosC(Rst("codsun"))
            Fg1.TextMatrix(xFila, 4) = NulosC(Rst("numdoc2"))
            Fg1.TextMatrix(xFila, 5) = Format(Rst("fchdoc"), FORMAT_DATE)
            Fg1.TextMatrix(xFila, 6) = Format(Rst("fchven"), FORMAT_DATE)
            Fg1.TextMatrix(xFila, 7) = NulosC(Rst("simbolo"))
            Fg1.TextMatrix(xFila, 8) = Format(NulosN(Rst("imptotal")), FORMAT_MONTO)
            Fg1.TextMatrix(xFila, 9) = Format(NulosN(Rst("tipcam")), "###0.##0") & ""
            
            Fg1.TextMatrix(xFila, 10) = Format(NulosN(Rst(nCampoMuestra)), FORMAT_MONTO)
            '--saldo
            Fg1.TextMatrix(xFila, 12) = Format(NulosN(Rst(nCampoMuestra)), FORMAT_MONTO)
            
            xSaldoDoc = NulosN(Rst("impsal"))
            TotDebe = TotDebe + NulosN(Rst(nCampoMuestra))
            
            
            If OptCliente.Value = True Then
                'Buscamos los abonos del cliente
                'Retenciones UNION Caja y Bancos UNION Canje de documentos UNION Canje de Letra
                nSQL = "SELECT DISTINCT Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, 'Retenciones' AS libro, mae_documento.codsun,mae_documento.abrev, [con_retencion]![numser]+'-'+[con_retencion]![numdoc] AS numdoc, con_retencion.fchemi, mae_moneda.simbolo, con_tc.impven AS tipcam, con_retenciondet.impret AS imptotal, IIf([con_retenciondet].[impret] Is Null,0,IIf([con_retencion].[idmon]=1,[con_retenciondet].[impret],IIf([con_tc].[impven] Is Null,0,[con_retenciondet].[impret]*[con_tc].[impven]))) AS imptotsol, IIf([con_retenciondet].[impret] Is Null,0,IIf([con_retencion].[idmon]=2,[con_retenciondet].[impret],IIf([con_tc].[impven] Is Null OR [con_tc].[impven]=0,0,[con_retenciondet].[impret]/[con_tc].[impven]))) AS imptotdol " _
                    + vbCr + " FROM mae_moneda RIGHT JOIN ((((con_diario RIGHT JOIN (con_retencion LEFT JOIN mae_documento ON con_retencion.iddoc = mae_documento.id) ON con_diario.idmov = con_retencion.id) LEFT JOIN con_tc ON con_retencion.fchemi = con_tc.fecha) LEFT JOIN mae_libros ON con_retencion.idlib = mae_libros.id) LEFT JOIN con_retenciondet ON con_retencion.id = con_retenciondet.id) ON mae_moneda.id = con_retencion.idmon " _
                    + vbCr + " WHERE (((con_diario.idlib) = 5) And ((con_retencion.tipo) = 2) And ((con_retenciondet.iddoc) = " & Rst("id") & ")); " _
                    + vbCr + " UNION " _
                    + vbCr + " SELECT Mid([numreg],1,2)+'01'+Mid([numreg],3,4) AS registro, 'Caja Bancos' AS libro, '' AS codsun, tes_documentos.abrev, " _
                    + vbCr + " IIf([tes_cajaorigendet]![numser]<>'',[tes_cajaorigendet]![numser]+'-'+[tes_cajaorigendet]![numdoc],[tes_cajaorigendet]![numdoc]) AS numdoc, " _
                    + vbCr + " tes_caja.fchope AS fchemi, mae_moneda.simbolo, con_tc.impven AS tipcam, tes_cajadestinodet.acuenta AS imptotal, " _
                    + vbCr + " IIf([tes_caja]![idmon]=1,[tes_cajadestinodet]![acuenta],[tes_cajadestinodet]![acuenta]*[con_tc]![impven]) AS imptotsol, " _
                    + vbCr + " IIf([tes_caja]![idmon]=2,[tes_cajadestinodet]![acuenta],[tes_cajadestinodet]![acuenta]/[con_tc]![impven]) AS imptotdol " _
                    + vbCr + " FROM (((((tes_caja LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) " _
                    + vbCr + " INNER JOIN (tes_cajadestino INNER JOIN tes_cajadestinodet ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND " _
                    + vbCr + " (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON tes_caja.id = tes_cajadestino.idtes) INNER JOIN tes_cajaori " _
                    + vbCr + " ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN tes_cajaorigendet ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND " _
                    + vbCr + " (tes_cajaori.idtes = tes_cajaorigendet.idtes)) LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id " _
                    + vbCr + " WHERE (((tes_cajadestinodet.idmod)=2) AND ((tes_cajadestinodet.iddoc)=" & Rst("id") & ") AND ((tes_caja.tipmov)=1))" _
                    + vbCr + " UNION" _
                    + vbCr + " SELECT DISTINCT Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, 'Canje de documentos' AS libro, '99' AS codsun, 'CAN' AS abrev, [con_canjes].[numser] & '-' & [con_canjes].[numdoc] AS numdoc, con_canjes.fchemi, mae_moneda.simbolo, con_tc.impven AS tipcam, con_canjesdet.impcan AS imptotal, IIf([con_canjesdet].[impcan] Is Null,0,IIf([con_canjes].[idmon]=1,[con_canjesdet].[impcan],IIf([con_tc].[impven] Is Null,0,[con_canjesdet].[impcan]*[con_tc].[impven]))) AS imptotsol, IIf([con_canjesdet].[impcan] Is Null,0,IIf([con_canjes].[idmon]=2,[con_canjesdet].[impcan],IIf([con_tc].[impven] Is Null Or [con_tc].[impven]=0,0,[con_canjesdet].[impcan]/[con_tc].[impven]))) AS imptotdol " _
                    + vbCr + " FROM ((((con_canjes LEFT JOIN con_diario ON con_canjes.id = con_diario.idmov) LEFT JOIN con_tc ON con_canjes.fchemi = con_tc.fecha) LEFT JOIN mae_libros ON con_canjes.idlib = mae_libros.id) LEFT JOIN con_canjesdet ON con_canjes.id = con_canjesdet.idcan) LEFT JOIN mae_moneda ON con_canjes.idmon = mae_moneda.id " _
                    + vbCr + " WHERE (((con_diario.idlib)=8) AND ((con_canjesdet.iddoc)=" & Rst("id") & " and con_canjesdet.tipo=1)); " _
                    + vbCr + " UNION " _
                    + vbCr + " SELECT DISTINCT Format(con_diario.idmes,'00') & IIf(mae_libros.codsun Is Null Or mae_libros.codsun='','FF',mae_libros.codsun) & con_diario.numasi AS registro, 'Canje de Letra' AS libro, '100' AS codsun, 'LE' AS abrev, con_letradet.numlet AS numdoc, con_letradet.fchemi, mae_moneda.simbolo, con_tc.impven AS tipcam, con_diario.imphabsol AS imptotal, IIf([con_letra].[idmon]=1,[con_diario].[imphabsol],IIf([con_tc].[impven] Is Null Or [con_tc].[impven]=0 Or [con_diario].[imphabdol] Is Null Or [con_diario].[imphabdol]=0,0,[con_diario].[imphabdol]*[con_tc].[impven])) AS imptotsol, IIf([con_letra].[idmon]=2,[con_diario].[imphabdol],IIf([con_tc].[impven] Is Null Or [con_tc].[impven]=0 Or [con_diario].[imphabsol] Is Null Or [con_diario].[imphabsol]=0,0,[con_diario].[imphabsol]/[con_tc].[impven])) AS imptotdol " _
                    + vbCr + " FROM (((con_letra LEFT JOIN con_tc ON con_letra.fchemi = con_tc.fecha) LEFT JOIN mae_libros ON con_letra.idlib = mae_libros.id) LEFT JOIN mae_moneda ON con_letra.idmon = mae_moneda.id) RIGHT JOIN ((con_letradet LEFT JOIN con_diario ON (con_letradet.corr = con_diario.correlativo) AND (con_letradet.idlet = con_diario.idmov)) LEFT JOIN con_letradoc ON con_diario.iddocpro = con_letradoc.iddoc) ON con_letra.id = con_letradet.idlet " _
                    + vbCr + " WHERE con_letra.tiplet=2 AND (((con_letradoc.iddoc)=" & Rst("id") & " ) AND ((con_diario.idlib)=37));"
                    
            Else
            
                'Buscamos los abonos al proveedor
                'Caja y Bancos UNION Canje de documentos UNION Canje de Letra UNION Rendición de Cuenta
                'nSQL = "SELECT Mid([numreg],1,2)+'01'+Mid([numreg],3,4) AS registro, 'Caja y Bancos' AS libro, '' AS codsun, tes_documentos.abrev, " _
                    & " IIf([tes_cajaorigendet]![numser]<>'',[tes_cajaorigendet]![numser]+'-'+[tes_cajaorigendet]![numdoc],[tes_cajaorigendet]![numdoc]) AS numdoc, " _
                    & " tes_caja.fchope AS fchemi, mae_moneda.simbolo, con_tc.impven AS tipcam, tes_cajadestinodet.acuenta AS imptotal, " _
                    & " IIf([tes_caja]![idmon]=1,[tes_cajadestinodet]![acuenta],[tes_cajadestinodet]![acuenta]*[con_tc]![impven]) AS imptotsol, " _
                    & " IIf([tes_caja]![idmon]=2,[tes_cajadestinodet]![acuenta],[tes_cajadestinodet]![acuenta]/[con_tc]![impven]) AS imptotdol " _
                    & " FROM (((((tes_caja LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) " _
                    & " INNER JOIN (tes_cajadestino INNER JOIN tes_cajadestinodet ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) " _
                    & " AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON tes_caja.id = tes_cajadestino.idtes) LEFT JOIN tes_cajaori " _
                    & " ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN tes_cajaorigendet ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND " _
                    & " (tes_cajaori.idtes = tes_cajaorigendet.idtes)) LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id  " _
                    & " WHERE (((tes_cajadestinodet.idmod)=1) AND ((tes_caja.tipmov)=2) AND ((tes_cajadestinodet.iddoc)=" & Rst("id") & ")) " _
                    + vbCr + " UNION " _
                    + vbCr + " SELECT Format(con_diario.idmes,'00') & IIf(mae_libros.codsun Is Null Or mae_libros.codsun='','FF',mae_libros.codsun) & con_diario.numasi AS registro, 'Canje de documentos' AS libro, '99' AS codsun, 'CAN' AS abrev,con_canjes.numser & '-' & con_canjes.numdoc AS numdoc, con_canjes.fchemi, mae_moneda.simbolo, con_tc.impven AS tipcam, con_diario.imphabsol AS imptotal, IIf(con_canjes.idmon=1,con_diario.imphabsol,IIf(con_tc.impven Is Null Or con_tc.impven=0 Or con_diario.imphabdol Is Null Or con_diario.imphabdol=0,0,con_diario.imphabdol*con_tc.impven)) AS imptotsol, IIf(con_canjes.idmon=2,con_diario.imphabdol,IIf(con_tc.impven Is Null Or con_tc.impven=0 Or con_diario.imphabsol Is Null Or con_diario.imphabsol=0,0,con_diario.imphabsol/con_tc.impven)) AS imptotdol " _
                    + vbCr + " FROM ((((con_canjes LEFT JOIN con_diario ON con_canjes.id = con_diario.idmov) LEFT JOIN con_tc ON con_canjes.fchemi = con_tc.fecha) LEFT JOIN mae_libros ON con_canjes.idlib = mae_libros.id) LEFT JOIN mae_moneda ON con_canjes.idmon = mae_moneda.id) LEFT JOIN con_canjesdet ON (con_diario.iddocpro = con_canjesdet.iddoc) AND (con_diario.idmov = con_canjesdet.idcan) " _
                    + vbCr + " WHERE (((con_diario.idlib) =8) And ((con_canjesdet.iddoc) = " & Rst("id") & ") And con_canjesdet.tipo = 2); " _
                    + vbCr + " UNION " _
                    + vbCr + " SELECT DISTINCT Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, 'Canje de Letra' AS libro, '100' AS codsun,'LE' AS abrev, con_letradet.numlet AS numdoc, con_letradet.fchemi, mae_moneda.simbolo, con_tc.impven AS tipcam, con_diario.impdebsol AS imptotal, IIf([con_letra].[idmon]=1,[con_diario].[impdebsol],IIf([con_tc].[impven] Is Null Or [con_tc].[impven]=0 Or [con_diario].[impdebdol] Is Null Or [con_diario].[impdebdol]=0,0,[con_diario].[impdebdol]*[con_tc].[impven])) AS imptotsol, IIf([con_letra].[idmon]=2,[con_diario].[impdebdol],IIf([con_tc].[impven] Is Null Or [con_tc].[impven]=0 Or [con_diario].[impdebsol] Is Null Or [con_diario].[impdebsol]=0,0,[con_diario].[impdebsol]/[con_tc].[impven])) AS imptotdol " _
                    + vbCr + " FROM (((con_letra LEFT JOIN con_tc ON con_letra.fchemi = con_tc.fecha) LEFT JOIN mae_libros ON con_letra.idlib = mae_libros.id) LEFT JOIN mae_moneda ON con_letra.idmon = mae_moneda.id) RIGHT JOIN ((con_letradet LEFT JOIN con_diario ON (con_letradet.corr = con_diario.correlativo) AND (con_letradet.idlet = con_diario.idmov)) LEFT JOIN con_letradoc ON con_diario.iddocpro = con_letradoc.iddoc) ON con_letra.id = con_letradet.idlet " _
                    + vbCr + " WHERE con_letra.tiplet=1 AND (((con_letradoc.iddoc)=" & Rst("id") & " ) AND ((con_diario.idlib)=37));" _
                    + vbCr + " UNION " _
                    + vbCr + " SELECT Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, 'Rendición de Cuenta' AS libro, '101' AS codsun,'REN' AS abrev, con_devoluciones.numdoc, con_devoluciones.fchemi, mae_moneda.simbolo, con_tc.impven AS tipcam, con_diario.impdebsol AS imptotal, IIf([con_devoluciones].[idmon]=1,[con_diario].[impdebsol],IIf([con_tc].[impven] Is Null Or [con_tc].[impven]=0 Or [con_diario].[impdebdol] Is Null Or [con_diario].[impdebdol]=0,0,[con_diario].[impdebdol]*[con_tc].[impven])) AS imptotsol, IIf([con_devoluciones].[idmon]=2,[con_diario].[impdebdol],IIf([con_tc].[impven] Is Null Or [con_tc].[impven]=0 Or [con_diario].[impdebsol] Is Null Or [con_diario].[impdebsol]=0,0,[con_diario].[impdebsol]/[con_tc].[impven])) AS imptotdol " _
                    + vbCr + " FROM ((con_devoluciones LEFT JOIN con_tc ON con_devoluciones.fchemi = con_tc.fecha) LEFT JOIN mae_moneda ON con_devoluciones.idmon = mae_moneda.id) INNER JOIN (mae_libros INNER JOIN (con_devolucionesdet INNER JOIN con_diario ON (con_devolucionesdet.idcom = con_diario.iddocpro) AND (con_devolucionesdet.id = con_diario.idmov)) ON mae_libros.id = con_diario.idlib) ON con_devoluciones.id = con_devolucionesdet.id " _
                    + vbCr + " WHERE (((con_devolucionesdet.idcom)=" & Rst("id") & " ) AND ((con_diario.idlib)=38));"
                    
                    
                nSQL = "SELECT Mid([numreg],1,2)+'01'+Mid([numreg],3,4) AS registro, 'Caja y Bancos' AS libro, '' AS codsun, tes_documentos.abrev, " _
                    & " IIf([tes_cajaorigendet]![numser]<>'',[tes_cajaorigendet]![numser]+'-'+[tes_cajaorigendet]![numdoc],[tes_cajaorigendet]![numdoc]) AS numdoc,  " _
                    & " tes_caja.fchope AS fchemi, mae_moneda.simbolo, con_tc.impven AS tipcam, tes_cajadestinodet.acuenta AS imptotal,  " _
                    & " IIf([tes_caja]![idmon]=1,[tes_cajadestinodet]![acuenta],[tes_cajadestinodet]![acuenta]*[con_tc]![impven]) AS imptotsol,  " _
                    & " IIf([tes_caja]![idmon]=2,[tes_cajadestinodet]![acuenta],[tes_cajadestinodet]![acuenta]/[con_tc]![impven]) AS imptotdol  " _
                    & " FROM (((((tes_caja LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha)  " _
                    & " INNER JOIN (tes_cajadestino INNER JOIN tes_cajadestinodet ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) " _
                    & " AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON tes_caja.id = tes_cajadestino.idtes) LEFT JOIN tes_cajaori " _
                    & " ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN tes_cajaorigendet ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND " _
                    & " (tes_cajaori.idtes = tes_cajaorigendet.idtes)) LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id  " _
                    & " WHERE (((tes_cajadestinodet.idmod)=1) AND ((tes_caja.tipmov)=2) AND ((tes_cajadestinodet.iddoc)=" & Rst("id") & ")) " _
                    & " Union " _
                    & " SELECT Format(con_diario.idmes,'00') & IIf(mae_libros.codsun Is Null Or mae_libros.codsun='','FF',mae_libros.codsun) & con_diario.numasi AS registro, " _
                    & " 'Canje de documentos' AS libro, '99' AS codsun, 'CAN' AS abrev,con_canjes.numser & '-' & con_canjes.numdoc AS numdoc, con_canjes.fchemi, " _
                    & " mae_moneda.simbolo, con_tc.impven AS tipcam, con_diario.imphabsol AS imptotal, IIf(con_canjes.idmon=1,con_diario.imphabsol, " _
                    & " IIf(con_tc.impven Is Null Or con_tc.impven=0 Or con_diario.imphabdol Is Null Or con_diario.imphabdol=0,0,con_diario.imphabdol*con_tc.impven)) AS imptotsol, " _
                    & " IIf(con_canjes.idmon=2,con_diario.imphabdol,IIf(con_tc.impven Is Null Or con_tc.impven=0 Or con_diario.imphabsol Is Null Or " _
                    & " con_diario.imphabsol=0,0,con_diario.imphabsol/con_tc.impven)) AS imptotdol FROM ((((con_canjes LEFT JOIN con_diario ON con_canjes.id = con_diario.idmov) " _
                    & " LEFT JOIN con_tc ON con_canjes.fchemi = con_tc.fecha) LEFT JOIN mae_libros ON con_canjes.idlib = mae_libros.id) LEFT JOIN mae_moneda " _
                    & " ON con_canjes.idmon = mae_moneda.id) LEFT JOIN con_canjesdet ON (con_diario.iddocpro = con_canjesdet.iddoc) AND (con_diario.idmov = con_canjesdet.idcan) " _
                    & " WHERE (((con_diario.idlib) =8) And ((con_canjesdet.iddoc) = " & Rst("id") & ") And con_canjesdet.tipo = 2) "
                nSQL = nSQL & " Union " _
                    & " SELECT DISTINCT Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, " _
                    & " 'Canje de Letra' AS libro, '100' AS codsun,'LE' AS abrev, con_letradet.numlet AS numdoc, con_letradet.fchemi, mae_moneda.simbolo, " _
                    & " con_tc.impven AS tipcam, con_diario.impdebsol AS imptotal, IIf([con_letra].[idmon]=1,[con_diario].[impdebsol],IIf([con_tc].[impven] Is Null " _
                    & " Or [con_tc].[impven]=0 Or [con_diario].[impdebdol] Is Null Or [con_diario].[impdebdol]=0,0,[con_diario].[impdebdol]*[con_tc].[impven])) AS imptotsol, " _
                    & " IIf([con_letra].[idmon]=2,[con_diario].[impdebdol],IIf([con_tc].[impven] Is Null Or [con_tc].[impven]=0 Or [con_diario].[impdebsol] Is Null " _
                    & " Or [con_diario].[impdebsol]=0,0,[con_diario].[impdebsol]/[con_tc].[impven])) AS imptotdol  FROM (((con_letra LEFT JOIN con_tc ON con_letra.fchemi = con_tc.fecha) " _
                    & " LEFT JOIN mae_libros ON con_letra.idlib = mae_libros.id) LEFT JOIN mae_moneda ON con_letra.idmon = mae_moneda.id) RIGHT JOIN ((con_letradet " _
                    & " LEFT JOIN con_diario ON (con_letradet.corr = con_diario.correlativo) AND (con_letradet.idlet = con_diario.idmov)) LEFT JOIN con_letradoc " _
                    & " ON con_diario.iddocpro = con_letradoc.iddoc) ON con_letra.id = con_letradet.idlet  WHERE con_letra.tiplet=1 AND (((con_letradoc.iddoc)=" & Rst("id") & " ) " _
                    & " AND ((con_diario.idlib)=37)) " _
                    & " Union " _
                    & " SELECT Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, " _
                    & " 'Rendición de Cuenta' AS libro, '101' AS codsun,'REN' AS abrev, con_devoluciones.numdoc, con_devoluciones.fchemi, mae_moneda.simbolo, " _
                    & " con_tc.impven AS tipcam, con_diario.impdebsol AS imptotal, IIf([con_devoluciones].[idmon]=1,[con_diario].[impdebsol],IIf([con_tc].[impven] Is Null " _
                    & " Or [con_tc].[impven]=0 Or [con_diario].[impdebdol] Is Null Or [con_diario].[impdebdol]=0,0,[con_diario].[impdebdol]*[con_tc].[impven])) AS imptotsol, " _
                    & " IIf([con_devoluciones].[idmon]=2,[con_diario].[impdebdol],IIf([con_tc].[impven] Is Null Or [con_tc].[impven]=0 Or [con_diario].[impdebsol] Is Null " _
                    & " Or [con_diario].[impdebsol]=0,0,[con_diario].[impdebsol]/[con_tc].[impven])) AS imptotdol FROM ((con_devoluciones LEFT JOIN con_tc " _
                    & " ON con_devoluciones.fchemi = con_tc.fecha) LEFT JOIN mae_moneda ON con_devoluciones.idmon = mae_moneda.id) INNER JOIN (mae_libros " _
                    & " INNER JOIN (con_devolucionesdet INNER JOIN con_diario ON (con_devolucionesdet.idcom = con_diario.iddocpro) AND (con_devolucionesdet.id = con_diario.idmov)) " _
                    & " ON mae_libros.id = con_diario.idlib) ON con_devoluciones.id = con_devolucionesdet.id  WHERE (((con_devolucionesdet.idcom)=" & Rst("id") & " ) " _
                    & " AND ((con_diario.idlib)=38))"
                nSQL = nSQL & " Union " _
                    & " SELECT Mid([com_compras]![numreg],1,2) & [mae_libros]![codsun] & Mid([com_compras]![numreg],3,4) AS registro, mae_libros.descripcion AS libro, " _
                    & " mae_libros.codsun, '' AS abrev, [com_compras]![numser] & '-' & [com_compras]![numdoc] AS numdoc, com_compras.fchdoc AS fchemi, mae_moneda.simbolo, " _
                    & " con_tc.impven AS timpcam, com_compras.imptot AS imptotal, IIf([com_compras]![idmon]=1,[com_compras]![imptot],[com_compras]![imptot]*[con_tc]![impven]) AS imptotsol, " _
                    & " IIf([com_compras]![idmon]=2,[com_compras]![imptot],[com_compras]![imptot]/[con_tc]![impven]) AS imptotdol FROM mae_moneda RIGHT JOIN " _
                    & " ((com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) " _
                    & " ON mae_moneda.id = com_compras.idmon WHERE (((com_compras.iddocref)=" & Rst("id") & "))"

                '    & " SELECT Mid([com_compras]![numreg],1,2) & [mae_libros]![codsun] & Mid([com_compras]![numreg],3,4) AS registro, " _
                    & " mae_libros.descripcion, mae_libros.codsun, [com_compras]![numser] & '-' & [com_compras]![numdoc] AS numdoc, com_compras.fchdoc AS fchemi, " _
                    & " mae_moneda.simbolo, con_tc.impven AS timpcam, com_compras.imptot AS imptotal, IIf([com_compras]![idmon]=1,[com_compras]![imptot], " _
                    & " [com_compras]![imptot]*[con_tc]![impven]) AS imptotsol, IIf([com_compras]![idmon]=2,[com_compras]![imptot],[com_compras]![imptot]/[con_tc]![impven]) AS imptotdol " _
                    & " FROM (mae_moneda RIGHT JOIN (com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON mae_moneda.id = com_compras.idmon) " _
                    & " LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha WHERE (((com_compras.iddocref)=" & Rst("id") & "))"
            End If
            
            Set Rstabo = Nothing
            RST_Busq Rstabo, nSQL, xCon
            If Rstabo.RecordCount <> 0 Then
                Rstabo.MoveFirst
                Rstabo.Sort = "fchemi ASC"
                Do While Not Rstabo.EOF
                    '--SI SE NTERRUMPE EL PROCESO => SALIR
'                    DoEvents
                    If BAND_INTERRUMPIR = True Then GoTo Salir:
                    Fg1.Rows = Fg1.Rows + 1
                    xFila = xFila + 1
                    
                    Fg1.TextMatrix(xFila, 1) = NulosC(Rstabo("registro"))
                    
                    Fg1.TextMatrix(xFila, 2) = NulosC(Rstabo("libro"))
                    Fg1.TextMatrix(xFila, 3) = NulosC(Rstabo("codsun"))
                    Fg1.TextMatrix(xFila, 4) = NulosC(Rstabo("numdoc"))
                    Fg1.TextMatrix(xFila, 5) = Format(Rstabo("fchemi"), FORMAT_DATE)
                    Fg1.TextMatrix(xFila, 7) = NulosC(Rstabo("simbolo"))
                    Fg1.TextMatrix(xFila, 8) = Format(NulosN(Rstabo("imptotal")), FORMAT_MONTO)
                    Fg1.TextMatrix(xFila, 9) = Format(NulosN(Rstabo("tipcam")), "####.###")
                    
                    Fg1.TextMatrix(xFila, 11) = Format(NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                    Fg1.TextMatrix(xFila, 12) = Format(NulosN(Fg1.TextMatrix(xFila - 1, 12)) - NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                    TotHaber = TotHaber + NulosN(Rstabo(nCampoMuestra))
                    Rstabo.MoveNext
                Loop
            End If
            
            '---ACTUALIZANDO EL SALDO AL DOCUMENTO
            If xSaldoDoc <> NulosN(Fg1.TextMatrix(xFila, 12)) And NulosN(Rst("idmon")) = NulosN(TxtIdMon.Text) Then
                If OptCliente.Value = True Then     '--VENTAS
                    xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = " & NulosN(Fg1.TextMatrix(xFila, 12)) & " WHERE (((vta_ventas.id)=" & Rst("id") & "))"
                Else                                '--COMPRAS
                    xCon.Execute "UPDATE com_compras SET com_compras.impsal = " & NulosN(Fg1.TextMatrix(xFila, 12)) & " WHERE (((com_compras.id)=" & Rst("id") & "))"
                End If
            End If
            '----MOSTRAR SOLO DESCUADRADOS ---------
            If chk_descuadrado.Value = 1 Then
                If NulosN(Fg1.TextMatrix(xFila, 12)) >= 0 Then
                    GRID_DELETE Fg1, Fg1.Rows - 1 - Rstabo.RecordCount, Fg1.Rows - 1, e_Fila
                    '*********************************************
                    If Rstabo.RecordCount <> 0 Then
                        Rstabo.MoveFirst
                        Do While Not Rstabo.EOF
                            TotHaber = TotHaber - NulosN(Rstabo(nCampoMuestra))
                            Rstabo.MoveNext
                        Loop
                    End If
                    TotDebe = TotDebe - NulosN(Rst(nCampoMuestra))
                    '*********************************************
                    xFila = Fg1.Rows - 1
                    mRowIni = -1
                Else
                    mRowIni = 0
                End If
            End If
            '---------------------------------------------------------
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
            If mRowIni = 0 Then
                If xColor = 0 Then
                    GRID_COLOR_FONDO Fg1, xFilaIni, 1, Fg1.Rows - 1, Fg1.Cols - 1, &H80000005
                    xColor = 1
                Else
                    GRID_COLOR_FONDO Fg1, xFilaIni, 1, Fg1.Rows - 1, Fg1.Cols - 1, &HE0DCDA
                    xColor = 0
                End If
            End If
        Next A
        
        Fg1.Rows = Fg1.Rows + 1
        xFila = xFila + 1
        Fg1.TextMatrix(xFila, 4) = "TOTAL -->"

        Fg1.TextMatrix(xFila, 10) = Format(TotDebe, FORMAT_MONTO)
        Fg1.TextMatrix(xFila, 11) = Format(TotHaber, FORMAT_MONTO)
        Fg1.TextMatrix(xFila, 12) = Format(TotDebe - TotHaber, FORMAT_MONTO)
        '*****resumen
        Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(TotDebe, FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(TotHaber, FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotDebe - TotHaber, FORMAT_MONTO)
        '******

        With Fg1
            .Cell(flexcpForeColor, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1) = &H80000008
            .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
            .FillStyle = flexFillRepeat
            .CellFontBold = True
        End With
        '----MOSTRAR SOLO DESCUADRADOS ---------
        If chk_descuadrado.Value = 1 Then
            If NulosN(Fg1.TextMatrix(xFila, 12)) = 0 Then
                GRID_DELETE Fg1, Fg1.Rows - 2, Fg1.Rows - 1, e_Fila
                xFila = Fg1.Rows - 1
            End If
            '--del resumen
            If NulosN(Fg2.TextMatrix(Fg2.Rows - 1, 5)) = 0 Then
                GRID_DELETE Fg2, Fg2.Rows - 1, Fg2.Rows - 1, e_Fila
            End If
            '---------------
        End If
        
        '---------------------------------------------------------

        If TotGralDebe <> 0 Or TotGralHaber <> 0 Then
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = "TOTAL GRAL -->"
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(TotGralDebe, FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(TotGralHaber, FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(TotGralDebe - TotGralHaber, FORMAT_MONTO)
            With Fg1
                .Cell(flexcpForeColor, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1) = &H80000008
                .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
                .FillStyle = flexFillRepeat
                .CellFontBold = True
            End With
            '*****resumen
            Fg2.Rows = Fg2.Rows + 2
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = "TOTAL GRAL -->"
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(TotGralDebe, FORMAT_MONTO)
            Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(TotGralHaber, FORMAT_MONTO)
            Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotGralDebe - TotGralHaber, FORMAT_MONTO)
            With Fg2
                .Cell(flexcpForeColor, Fg2.Rows - 1, 1, Fg2.Rows - 1, Fg2.Cols - 1) = &H80000008
                .Select Fg2.Rows - 1, 1, Fg2.Rows - 1, Fg2.Cols - 1
                .FillStyle = flexFillRepeat
                .CellFontBold = True
            End With
            '******
    End If
    
    End If
    If mRowIni = 0 Then
        If xColor = 0 Then
            GRID_COLOR_FONDO Fg1, xFilaIni, 1, Fg1.Rows - 2, Fg1.Cols - 1, &H80000005
            xColor = 1
        Else
            GRID_COLOR_FONDO Fg1, xFilaIni, 1, Fg1.Rows - 2, Fg1.Cols - 1, &HE0DCDA
            xColor = 0
        End If
    End If
    Set Rst = Nothing
    Set Rstabo = Nothing
    fraBarra.Visible = False
    Me.MousePointer = vbDefault
    MsgBox "La Consulta fue se realizó Correctamente", vbInformation, xTitulo
    Exit Sub
Salir:
    Me.MousePointer = vbDefault
    fraBarra.Visible = False
    Set Rst = Nothing
    Set Rstabo = Nothing
    MsgBox "La Consulta fue Interrumpida", vbInformation, xTitulo
error:
    Me.MousePointer = vbDefault
    fraBarra.Visible = False
    Set Rst = Nothing
    Set Rstabo = Nothing
    SHOW_ERROR Me.Name, "CargarCli"
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        BAND_INTERRUMPIR = True '--interrumpir
    ElseIf KeyCode = vbKeyF3 And Shift = 0 Then
        BuscarVSFlexGrid
    End If

End Sub

Private Sub Form_Load()
    TxtCliPro.Text = ""
    TxtFecha.Valor = ""
    TxtFecha.Valor = Date
    LblMoneda.Caption = ""
    TxtIdMon.Text = ""
    SeEjecuto = False

End Sub

Private Sub OptCan_Click()
    chk_descuadrado.Enabled = True
End Sub

Private Sub OptCliente_Click()
    TxtCliPro.Text = ""
    LblIdCliPro.Caption = ""
    Label1(0).Caption = "Cliente"
End Sub

Private Sub OptPen_Click()
    chk_descuadrado.Value = 0
    chk_descuadrado.Enabled = False
End Sub

Private Sub OptProvee_Click()
    TxtCliPro.Text = ""
    LblIdCliPro.Caption = ""
    Label1(0).Caption = "Proveedor"
End Sub

Private Sub OptSel1_Click()
    TxtCliPro.Text = ""
    LblIdCliPro.Caption = ""
    TxtCliPro.Enabled = False
    CmdBusCliPro.Enabled = False
End Sub

Private Sub OptSel2_Click()
    TxtCliPro.Enabled = True
    CmdBusCliPro.Enabled = True
End Sub

Private Sub OptTodos_Click()
    chk_descuadrado.Enabled = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        
        If OptSel1.Value = True Then CargarCli 0
        If OptSel2.Value = True Then
            If NulosC(TxtCliPro.Text) = "" Then
                If OptCliente.Value = True Then
                    MsgBox "No ha especificado el cliente a consultar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                End If
                If OptProvee.Value = True Then
                    MsgBox "No ha especificado el proveedor a consultar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                End If
                TxtCliPro.SetFocus
                Exit Sub
            End If
            CargarCli NulosN(LblIdCliPro.Caption)
        End If
    End If
    
    If Button.Index = 3 Then pExportar
    If Button.Index = 4 Then pImprimir
    If Button.Index = 6 Then
        Set RstCta = Nothing
        Unload Me
    End If
End Sub

Sub pExportar()
    On Error GoTo error
    Dim oExport As New SGI2_funciones.formularios
    Dim nPeriodo As String
    Dim nTitulo1 As String
    
    nPeriodo = "Al  " + CStr(TxtFecha.Valor)
    If NulosN(TxtIdMon.Text) = 1 Then
        nTitulo1 = "(Expresado en Nuevos Soles)"
    ElseIf NulosN(TxtIdMon.Text) = 2 Then
        nTitulo1 = "(Expresado en Dolares Americanos)"
    End If
    If TabOne1.CurrTab = 0 Then '--detalle
        oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "Cuenta Corriente - " + IIf(OptCliente.Value = True, "Cliente", "Proveedor"), nPeriodo, nTitulo1, "Cuenta Corriente Análisis"
    Else
        oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg2, "Resumen de Cuenta Corriente - " + IIf(OptCliente.Value = True, "Cliente", "Proveedor"), nPeriodo, nTitulo1, "Cuenta Corriente Análisis"
    End If
    
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pExportar"
End Sub

Private Sub TxtCliPro_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCliPro_Click
    End If
End Sub

'***********************************************************************************************
'------------CAMBIOS AL 020108

Private Sub CmdBusMon_Click()
    
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Nombre":      xCampos(0, 1) = "descripcion":     xCampos(0, 2) = "4500":      xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id":   xCampos(1, 1) = "id":              xCampos(1, 2) = "500":      xCampos(1, 3) = "N"

    CARGAR_DLL_EPSBUSCAR xCon, xRs, "SELECT * FROM mae_moneda ORDER BY descripcion ;", xCampos(), "Buscando Moneda", "descripcion", "descripcion", Principio
    
    If xRs.State = 0 Then GoTo Salir
    If xRs.RecordCount = 0 Then GoTo Salir
    TxtIdMon.Text = xRs("id") & ""
    LblMoneda.Caption = xRs("descripcion") & ""
    
Salir:
    Set xRs = Nothing
End Sub

Private Sub TxtIdMon_Change()
    If Trim(TxtIdMon.Text) = "" Then LblMoneda.Caption = ""
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtIdMon_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMon_Click
    End If
End Sub

Private Sub TxtIdMon_Validate(Cancel As Boolean)
    If NulosC(TxtIdMon.Text) <> "" Then
        LblMoneda.Caption = Busca_Codigo(TxtIdMon.Text, "id", "descripcion", "mae_moneda", "N", xCon)
        If NulosC(LblMoneda.Caption) = "" Then
            TxtIdMon.Text = ""
        End If
    End If
End Sub

Private Sub BuscarVSFlexGrid()
    On Error GoTo error
    
    Dim oExport As New SGI2_funciones.formularios
    Dim xCampos(4, 3) As String
    'campo     'columna del grid    'tipo(N:Numerico, C:caracter, F:fecha)      campo predeterminado(0:no se muestra, -1:se muestra al iniciar el formulario)
    xCampos(0, 0) = "Nº.Registro":      xCampos(0, 1) = "1":    xCampos(0, 2) = "C":    xCampos(0, 3) = "-1"
    xCampos(1, 0) = "Origen":           xCampos(1, 1) = "2":    xCampos(1, 2) = "C":    xCampos(1, 3) = "0"
    xCampos(2, 0) = "Nº Documento":     xCampos(2, 1) = "4":    xCampos(2, 2) = "C":    xCampos(2, 3) = "0"
    xCampos(3, 0) = Label1(0):          xCampos(3, 1) = "4":    xCampos(3, 2) = "C":    xCampos(3, 3) = "0"
    xCampos(4, 0) = "Fch.Emi.":         xCampos(4, 1) = "5":    xCampos(4, 2) = "F":    xCampos(4, 3) = "0"
    
    oExport.VSFlexGrid_Buscar Me.hWnd, Fg1, xCampos(), Fg1.Row
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "BuscarVSFlexGrid"
End Sub

Private Sub pConfigurarGrilla()
    Dim A As Integer
    '--PERMITIRA CONFIGURAR EL FORMATO DE LA CONSULTA
    With Fg1
        '-----
        .Rows = 2
        .Cols = 13
        .FixedRows = 2
        .FrozenCols = 0
        .RowHeight(0) = 250
        .ColWidth(0) = 200
        UNIR_CELDAS Fg1, 0, 1, 0, 9, "DATOS DEL DOCUMENTO", flexAlignCenterCenter
        FORMATO_CELDA Fg1, 0, 1, vbBlack, True, &HD8E9EC
        If Trim(LblMoneda.Caption) = "" Then
            UNIR_CELDAS Fg1, 0, 10, 0, 12, "IMPORTES", flexAlignCenterCenter
        Else
            UNIR_CELDAS Fg1, 0, 10, 0, 12, "IMPORTES EN " & UCase(LblMoneda.Caption), flexAlignCenterCenter
        End If
        FORMATO_CELDA Fg1, 0, 10, vbBlack, True, &HD8E9EC
        .ColWidth(1) = 350
'
        .TextMatrix(1, 1) = "N° Registro":  .ColWidth(1) = 900:   .ColAlignment(1) = flexAlignLeftCenter:     .Row = 1: .Col = 1: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 2) = "Origen":       .ColWidth(2) = 1200:   .ColAlignment(2) = flexAlignLeftCenter:     .Row = 1: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 3) = "T.D.":         .ColWidth(3) = 450:    .ColAlignment(3) = flexAlignLeftCenter:     .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 4) = "N°.Documento": .ColWidth(4) = 1600:   .ColAlignment(4) = flexAlignLeftCenter:     .Row = 1: .Col = 4: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 5) = "Fch.Emi.":     .ColWidth(5) = 800:    .ColAlignment(5) = flexAlignCenterBottom:   .Row = 1: .Col = 5: .CellAlignment = flexAlignCenterBottom
        .TextMatrix(1, 6) = "Fch.Ven.":     .ColWidth(6) = 800:    .ColAlignment(6) = flexAlignCenterBottom:   .Row = 1: .Col = 6: .CellAlignment = flexAlignCenterBottom
        .TextMatrix(1, 7) = "M":            .ColWidth(7) = 450:    .ColAlignment(7) = flexAlignLeftCenter:    .Row = 1: .Col = 7: .CellAlignment = flexAlignCenterBottom
        
        .TextMatrix(1, 8) = "Imp":          .ColWidth(8) = 900:    .ColAlignment(8) = flexAlignRightCenter:   .Row = 1: .Col = 8: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 9) = "T.C.":         .ColWidth(9) = 500:    .ColAlignment(9) = flexAlignRightCenter:    .Row = 1: .Col = 9: .CellAlignment = flexAlignRightCenter
        '----------------------
        .TextMatrix(1, 10) = "Cargo":       .ColWidth(10) = 1150:  .ColAlignment(10) = flexAlignRightCenter:   .Row = 1: .Col = 10: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 11) = "Abono":       .ColWidth(11) = 1150:  .ColAlignment(11) = flexAlignRightCenter:   .Row = 1: .Col = 11: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 12) = "Saldo":       .ColWidth(12) = 1150:  .ColAlignment(12) = flexAlignRightCenter:   .Row = 1: .Col = 12: .CellAlignment = flexAlignRightCenter
        .SelectionMode = flexSelectionByRow
    End With
    
    With Fg2
        '-----
        .Rows = 1
        .Cols = 6
        .FixedRows = 1
        .FrozenCols = 0
        .RowHeight(0) = 250
        .ColWidth(0) = 200:
        .TextMatrix(0, 1) = "R.U.C.":   .ColWidth(1) = 1200:  .ColAlignment(1) = flexAlignCenterCenter: .Row = 0: .Col = 1: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 2) = "Nombres":  .ColWidth(2) = 5500:  .ColAlignment(2) = flexAlignLeftCenter:   .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftCenter
        '----------------------
        .TextMatrix(0, 3) = "Cargo":    .ColWidth(3) = 1300:  .ColAlignment(3) = flexAlignRightCenter:  .Row = 0: .Col = 3: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 4) = "Abono":    .ColWidth(4) = 1300:  .ColAlignment(4) = flexAlignRightCenter:  .Row = 0: .Col = 4: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 5) = "Saldo":    .ColWidth(5) = 1300:  .ColAlignment(5) = flexAlignRightCenter:  .Row = 0: .Col = 5: .CellAlignment = flexAlignRightCenter
        For A = 1 To .Cols - 1
            FORMATO_CELDA Fg2, 0, A, vbBlack, True, &HD8E9EC
        Next
        .SelectionMode = flexSelectionByRow
    End With
    TabOne1.CurrTab = 0
    DoEvents
End Sub


Private Sub pImprimir()

    On Error GoTo error
    
    If TabOne1.CurrTab = 0 Then
        FrmPrinCtaCtaCli.Show
        FrmPrinCtaCtaCli.SetFocus
    Else
        Dim oPrint As New SGI2_funciones.formularios
        Dim nPeriodo As String
        Dim nTitulo As String
        Dim nTitulo1 As String
        nPeriodo = "Al  " + CStr(TxtFecha.Valor)
        If NulosN(TxtIdMon.Text) = 1 Then
            nTitulo1 = "(Expresado en Nuevos Soles)"
        ElseIf NulosN(TxtIdMon.Text) = 2 Then
            nTitulo1 = "(Expresado en Dolares Americanos)"
        End If
        nTitulo = "Resumen de Cuenta Corriente - " + IIf(OptCliente.Value = True, "Cliente", "Proveedor")
        Me.MousePointer = vbHourglass
        oPrint.Imprimir_x_VSFlexGrid Fg2, nTitulo, nTitulo1, nPeriodo, False, True
        Set oPrint = Nothing
    
    End If
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimir"
End Sub

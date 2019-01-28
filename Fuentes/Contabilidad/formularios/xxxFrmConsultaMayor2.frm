VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SIZERONE.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsultaMayor2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Mayor Auxiliar"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   11925
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   1005
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
         Top             =   45
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
               Picture         =   "FrmConsultaMayor2.frx":0000
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsultaMayor2.frx":0544
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsultaMayor2.frx":08D6
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsultaMayor2.frx":0A5A
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsultaMayor2.frx":0EAE
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsultaMayor2.frx":0FC6
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsultaMayor2.frx":150A
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsultaMayor2.frx":1A4E
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsultaMayor2.frx":1B62
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsultaMayor2.frx":1C76
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsultaMayor2.frx":20CA
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsultaMayor2.frx":2236
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsultaMayor2.frx":277E
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsultaMayor2.frx":2A98
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1290
      Left            =   2055
      TabIndex        =   7
      Top             =   360
      Width           =   8085
      Begin VB.CommandButton cmd_periodo 
         Height          =   255
         Index           =   1
         Left            =   1875
         Picture         =   "FrmConsultaMayor2.frx":2E2A
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1650
         Width           =   285
      End
      Begin VB.CommandButton cmd_periodo 
         Height          =   255
         Index           =   0
         Left            =   1875
         Picture         =   "FrmConsultaMayor2.frx":31AC
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1350
         Width           =   285
      End
      Begin VB.CommandButton CmdDel 
         Caption         =   "Eliminar"
         Height          =   435
         Left            =   7155
         TabIndex        =   6
         Top             =   750
         Width           =   855
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "Agregar "
         Height          =   435
         Left            =   7155
         TabIndex        =   5
         Top             =   195
         Width           =   855
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg2 
         Height          =   990
         Left            =   2355
         TabIndex        =   4
         Top             =   195
         Width           =   4740
         _cx             =   8361
         _cy             =   1746
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmConsultaMayor2.frx":352E
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
      Begin VB.OptionButton OptDolares 
         Caption         =   "Dolares"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1185
         TabIndex        =   3
         Top             =   990
         Width           =   900
      End
      Begin VB.OptionButton OptSoles 
         Caption         =   "Soles"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   990
         Value           =   -1  'True
         Width           =   900
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   945
         TabIndex        =   0
         Top             =   225
         Width           =   1245
         _ExtentX        =   2196
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
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
         Height          =   300
         Left            =   945
         TabIndex        =   1
         Top             =   540
         Width           =   1245
         _ExtentX        =   2196
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
      End
      Begin VB.Label lbl_periodo 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_periodo(1)"
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
         Index           =   1
         Left            =   525
         TabIndex        =   33
         Top             =   1620
         Width           =   1680
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
         Left            =   525
         TabIndex        =   31
         Top             =   1320
         Width           =   1680
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Inicio"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   9
         Top             =   270
         Width           =   735
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Final"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   8
         Top             =   585
         Width           =   690
      End
   End
   Begin VB.CheckBox chk 
      Caption         =   "Procesar Todas las Cuentas"
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   60
      TabIndex        =   29
      Top             =   1200
      Width           =   1770
   End
   Begin VB.Frame Frame6 
      Caption         =   "[ Tipo de Consulta ]"
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
      Height          =   750
      Left            =   15
      TabIndex        =   26
      Top             =   360
      Width           =   1980
      Begin VB.OptionButton opt_fecha 
         Caption         =   "Por Fecha"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   28
         Top             =   225
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton opt_fecha 
         Caption         =   "Por Periodo"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   27
         Top             =   465
         Width           =   1125
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "[ Ordenado por ]"
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
      Height          =   1290
      Left            =   10200
      TabIndex        =   22
      Top             =   360
      Width           =   1725
      Begin VB.OptionButton opt 
         Caption         =   "Nº Documento"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   25
         Top             =   615
         Width           =   1440
      End
      Begin VB.OptionButton opt 
         Caption         =   "Fecha de Emisión"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   24
         Top             =   315
         Width           =   1560
      End
      Begin VB.OptionButton opt 
         Caption         =   "Nº Registro"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   23
         Top             =   915
         Value           =   -1  'True
         Width           =   1425
      End
   End
   Begin VB.Frame fra_msg 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   270
      Left            =   6015
      TabIndex        =   19
      Top             =   7350
      Visible         =   0   'False
      Width           =   5730
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Obs: Se recomienda minimizar la ventana para agilizar el proceso"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   30
         TabIndex        =   20
         Top             =   15
         Width           =   5535
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   795
      Left            =   3045
      TabIndex        =   15
      Top             =   2985
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   105
         TabIndex        =   16
         Top             =   330
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   5925
         Y1              =   780
         Y2              =   765
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   5925
         X2              =   5925
         Y1              =   15
         Y2              =   945
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   -30
         X2              =   5940
         Y1              =   0
         Y2              =   15
      End
      Begin VB.Label lbl 
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
         Index           =   2
         Left            =   4365
         TabIndex        =   18
         Top             =   90
         Width           =   1530
      End
      Begin VB.Label Label3 
         Caption         =   "Procesando Asientos"
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
         Left            =   165
         TabIndex        =   17
         Top             =   90
         Width           =   4020
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
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   5985
      Left            =   0
      TabIndex        =   10
      Top             =   1665
      Width           =   11910
      _cx             =   21008
      _cy             =   10557
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
      Caption         =   "   Detalle   |   Resumen   "
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
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   5565
         Left            =   12555
         TabIndex        =   12
         Top             =   45
         Width           =   11820
         Begin VSFlex7Ctl.VSFlexGrid Fg3 
            Height          =   5550
            Left            =   15
            TabIndex        =   13
            Top             =   0
            Width           =   11790
            _cx             =   20796
            _cy             =   9790
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
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmConsultaMayor2.frx":35B3
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
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   5565
         Left            =   45
         TabIndex        =   11
         Top             =   45
         Width           =   11820
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   5550
            Left            =   15
            TabIndex        =   14
            Top             =   0
            Width           =   11790
            _cx             =   20796
            _cy             =   9790
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
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmConsultaMayor2.frx":36B9
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
Attribute VB_Name = "FrmConsultaMayor2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstTmp As New ADODB.Recordset
Dim RstTmp2 As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE
Dim mMesIni As Integer
Dim mMesFin As Integer



Private Sub chk_Click()
    If Fg2.Rows = 1 Then Exit Sub
    Fg2.Rows = 1
End Sub

Private Sub CmdAdd_Click()
    On Error GoTo error
    Dim xfrm As New SGI2_funciones.formularios
    Dim Rst As New ADODB.Recordset
    Dim k As Integer
    Dim MSG_CUENTA As String    '--MUSTRA EL MENSAJE SI DESEA AGREGAR UNA CUENTA, CUANDO YA EXISTE UNA CUENTA DE NIVEL SUPERIOR O NIVEL INFERIOR
                                '--NO MOSTRAR MENSAJE SOLO CUANDO LAS CUENTAS SEA DEL MISMO NIVEL
    
    If chk.Value = 1 Then chk.Value = 0
    
    Set Rst = xfrm.SelePlanCuentas(xCon)
    If Rst.State = 1 Then
        If Rst.RecordCount <> 0 Then
            If GRID_BUSCAR_VALOR(Fg2, 3, CStr(Rst("id") & ""), False) <> "-1" Then
                MsgBox "La cuenta contable Nº " + Trim(Rst("cuenta") & "") + " ya fue seleccionada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
            
            For k = 1 To Fg2.Rows - 1
                If Len(Trim(Rst.Fields("cuenta") & "")) < Len(Trim(Fg2.TextMatrix(k, 1))) Then
                    
                    If Trim(Rst.Fields("cuenta") & "") = Mid(Trim(Fg2.TextMatrix(k, 1)), 1, Len(Trim(Rst.Fields("cuenta") & ""))) Then
                        MSG_CUENTA = "Ya agregó la cuenta Nª: " + Trim(Fg2.TextMatrix(k, 1)) + " cuyo nivel es Inferior a la cuenta Nº: " + Trim(Rst.Fields("cuenta") & "") + " que desea agregar" _
                                    + vbCr + "Sólo puede agregar Cuentas del mismo nivel " _
                                    + vbCr + "Si desea continuar elimine la fila que contenga la Cuenta Nº: " + Trim(Fg2.TextMatrix(k, 1))
                        Exit For
                    End If
                    
                Else
                    If Trim(Fg2.TextMatrix(k, 1)) = Mid(Trim(Rst.Fields("cuenta") & ""), 1, Len(Trim(Fg2.TextMatrix(k, 1)))) Then
                        MSG_CUENTA = "Ya agregó la cuenta Nª: " + Trim(Fg2.TextMatrix(k, 1)) + " cuyo nivel es Superior a la cuenta Nº: " + Trim(Rst.Fields("cuenta") & "") + " que desea agregar" _
                                    + vbCr + "Sólo puede agregar Cuentas del mismo nivel " _
                                    + vbCr + "Si desea continuar elimine la fila que contenga la Cuenta Nº: " + Trim(Fg2.TextMatrix(k, 1))
                        Exit For
                    End If
                    
                End If
            Next k
            If MSG_CUENTA <> "" Then
                MsgBox MSG_CUENTA, vbExclamation, xTitulo
                GoTo SALIR
            End If
            
            
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosC(Rst("cuenta"))
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = NulosC(Rst("descripcion"))
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = NulosC(Rst("id"))
        End If
    End If
SALIR:
    Set xfrm = Nothing
    Set Rst = Nothing
    Exit Sub
error:
    Set xfrm = Nothing
    Set Rst = Nothing
    SHOW_ERROR Me.Name, "CmdAdd_Click"
End Sub

Private Sub CmdDel_Click()
    If Fg2.Row <= 0 Then Exit Sub
    If Fg2.Rows <= 1 Then
        MsgBox "No hay cuentas seleccionadas para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    Else
        Fg2.RemoveItem Fg2.Row
        Fg2.Refresh
    End If
End Sub

Private Sub pImprimir()

    On Error GoTo error
    
    If fValidarConsulta() = False Then Exit Sub
    
    
    Me.MousePointer = vbHourglass
    If Me.TabOne1.CurrTab = 0 Then
        FrmPrintMayor.Show
    Else
        Dim X_PRINT As New SGI2_funciones.formularios
        Dim xMoneda As String
        Dim nPeriodo  As String
        If opt_fecha(0).Value = True Then  '--por fecha
            If CDate(Me.TxtFchIni.Valor) <> CDate(Me.TxtFchFin.Valor) Then
                nPeriodo = "Del: " + CStr(TxtFchIni.Valor) + " Al: " + CStr(TxtFchFin.Valor)
            Else
                nPeriodo = "Al: " + CStr(TxtFchIni.Valor)
            End If
        Else '--por periodo
            If mMesIni = mMesFin Then
                nPeriodo = "Periodo : " + lbl_periodo(0).Caption
            Else
                nPeriodo = "Periodo : De " + lbl_periodo(0).Caption & " A " & lbl_periodo(1).Caption
            End If
        End If
        If OptSoles.Value = True Then
            xMoneda = "Nuevos Soles"
        Else
            xMoneda = "Dolares Americanos"
        End If
        X_PRINT.Imprimir_x_VSFlexGrid Fg3, "LIBRO MAYOR ", "(Expresado en " + xMoneda + ")", nPeriodo, False, True
        Set X_PRINT = Nothing
        
    End If
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "CmdImprimir_Click"
End Sub

Private Sub pConsultar()
    If fValidarConsulta() = False Then Exit Sub

    BAND_INTERRUMPIR = False
    Me.ProgressBar1.Value = 1
    Me.TabOne1.CurrTab = 0
    pConfigurarGrilla True
    If MuestraMayor() = False Then Exit Sub
    If BAND_INTERRUMPIR = True Then Exit Sub
    Me.TabOne1.CurrTab = 1
    DoEvents
    CargarResumen
    
End Sub

Sub CargarResumen()
    On Error GoTo error
    Dim RstRes As New ADODB.Recordset
    Dim A&
    Dim xTotal1, xTotal2, xTotal3, xTotal4 As Double
    Dim xAcumulado(7) As Double
   
    Frame5.Left = 3413
    Frame5.Top = 2685
    Me.ProgressBar1.Value = 1
    Frame5.Visible = True
    Label3.Caption = "Procesando Resumen"
    DoEvents
    
    Dim SQL_CUENTA As String
    Dim nSQL As String
    '--------------------------
    SQL_CUENTA = ""
    '--SI AGREGA CUENTAS AS GRID, GENERAR EL FILTRO A CONCATENAR A LA CONSULTA
    For A = 1 To Fg2.Rows - 1
        If Trim(Fg2.TextMatrix(A, 1)) <> "" Then
            SQL_CUENTA = SQL_CUENTA + " con_planctas.cuenta Like '" & Trim(Fg2.TextMatrix(A, 1)) & "%' OR "
        End If
    Next A
    If SQL_CUENTA <> "" Then SQL_CUENTA = " WHERE (" + Left(SQL_CUENTA, Len(SQL_CUENTA) - 3) + ") "
    '--------------------------

    nSQL = "SELECT con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion AS descri , con_planctas.tipsal , "
    
    If Me.OptSoles = True Then
        nSQL = nSQL _
            + vbCr + " (SELECT Sum(IIf(con_diario1.impdebdol=0,con_diario1.impdebsol,IIf(con_tc1.impven Is Null,0,(con_tc1.impven*con_diario1.impdebdol)))) AS saldebesol  FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue  WHERE con_diario1.fchasi < CDate('01/01/07') Or con_diario1.fchasi Is Null GROUP BY con_diario1.idcue  HAVING con_diario1.idcue = con_diario.idcue ) AS saldebesol, " _
            + vbCr + " (SELECT Sum(IIf(con_diario1.imphabdol=0,con_diario1.imphabsol,IIf(con_tc1.impven Is Null,0,(con_tc1.impven*con_diario1.imphabdol)))) AS salhabersol FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue WHERE con_diario1.fchasi < CDate('01/01/07') Or con_diario1.fchasi Is Null GROUP BY con_diario1.idcue  HAVING con_diario1.idcue = con_diario.idcue ) AS salhabersol, " _
            + vbCr + " (SELECT Sum(IIf(con_diario1.impdebdol=0,con_diario1.impdebsol,IIf(con_tc1.impven Is Null,0,(con_tc1.impven*con_diario1.impdebdol))))  AS debesol FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue WHERE  con_diario1.fchasi >= CDate('01/01/07') and con_diario1.fchasi <= cdate('31/12/07') GROUP BY con_diario1.idcue  HAVING con_diario1.idcue=con_diario.idcue  ) AS  debesol, " _
            + vbCr + " (SELECT Sum(IIf(con_diario1.imphabdol=0,con_diario1.imphabsol,IIf(con_tc1.impven Is Null,0,(con_tc1.impven*con_diario1.imphabdol))))  AS habersol FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue WHERE  con_diario1.fchasi >= CDate('01/01/07') and con_diario1.fchasi <= cdate('31/12/07')  GROUP BY con_diario1.idcue  HAVING con_diario1.idcue=con_diario.idcue  ) AS  habersol, " _
            + vbCr + " IIf(debesol Is Null,0+IIf(saldebesol Is Null,0,saldebesol),debesol+IIf(saldebesol Is Null,0,saldebesol)) AS maydebesol, " _
            + vbCr + " IIf(habersol Is Null,0+IIf(salhabersol Is Null,0,salhabersol),habersol+IIf(salhabersol Is Null,0,salhabersol)) AS mayhabersol, " _
            + vbCr + " (IIF (con_planctas.tipsal='D' OR con_planctas.tipsal IS NULL OR con_planctas.tipsal ='', (maydebesol -  mayhabersol), (mayhabersol - maydebesol))) as saldosol, " _
            + vbCr + " IIf(maydebesol>mayhabersol,(maydebesol-mayhabersol),0) AS deudorsol, " _
            + vbCr + " IIf(mayhabersol>maydebesol,(mayhabersol-maydebesol),0) AS acreedorsol "
    End If
                        
    If OptDolares.Value = True Then
        nSQL = nSQL _
            + vbCr + " (SELECT Sum(IIf(con_diario1.impdebdol<>0,con_diario1.impdebdol,IIf(con_tc1.impven Is Null Or con_diario1.impdebsol=0,0,(con_diario1.impdebsol/con_tc1.impven)))) AS saldebedol FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue WHERE con_diario1.fchasi < CDate('01/01/07') Or con_diario1.fchasi Is Null GROUP BY con_diario1.idcue  HAVING (((con_diario1.idcue)=con_diario.idcue))) AS saldebedol, " _
            + vbCr + " (SELECT Sum(IIf(con_diario1.imphabdol<>0,con_diario1.imphabdol,IIf(con_tc1.impven Is Null Or con_diario1.imphabsol=0,0,(con_diario1.imphabsol/con_tc1.impven)))) AS salhaberdol FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue WHERE con_diario1.fchasi < CDate('01/01/07') Or con_diario1.fchasi Is Null GROUP BY con_diario1.idcue  HAVING (((con_diario1.idcue)=con_diario.idcue))) AS salhaberdol, " _
            + vbCr + " (SELECT Sum(IIf(con_diario1.impdebdol<>0,con_diario1.impdebdol,IIf(con_tc1.impven Is Null Or con_diario1.impdebsol=0,0,(con_diario1.impdebsol/con_tc1.impven)))) AS debedol FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue WHERE  con_diario1.fchasi >= CDate('01/01/07') and con_diario1.fchasi <= cdate('31/12/07')  GROUP BY con_diario1.idcue HAVING con_diario1.idcue=con_diario.idcue ) AS  debedol, " _
            + vbCr + " (SELECT Sum(IIf(con_diario1.imphabdol<>0,con_diario1.imphabdol,IIf(con_tc1.impven Is Null Or con_diario1.imphabsol=0,0,(con_diario1.imphabsol/con_tc1.impven)))) AS haberdol FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue WHERE  con_diario1.fchasi >= CDate('01/01/07') and con_diario1.fchasi <= cdate('31/12/07')  GROUP BY con_diario1.idcue  HAVING con_diario1.idcue=con_diario.idcue ) AS  haberdol, " _
            + vbCr + " IIf(debedol Is Null,0+IIf(saldebedol Is Null,0,saldebedol),debedol+IIf(saldebedol Is Null,0,saldebedol)) AS maydebedol, " _
            + vbCr + " IIf(haberdol Is Null,0+IIf(salhaberdol Is Null,0,salhaberdol),haberdol+IIf(salhaberdol Is Null,0,salhaberdol)) AS mayhaberdol, " _
            + vbCr + " (IIF (con_planctas.tipsal='D' OR con_planctas.tipsal IS NULL OR con_planctas.tipsal ='', (maydebedol -  mayhaberdol), (mayhaberdol - maydebedol))) as saldodol, " _
            + vbCr + " IIf(maydebedol>mayhaberdol,(maydebedol-mayhaberdol),0) AS deudordol, " _
            + vbCr + " IIf(mayhaberdol > maydebedol, (mayhaberdol - maydebedol), 0) As acreedordol "
    End If
        
    nSQL = nSQL _
        + vbCr + " FROM con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue " _
        + SQL_CUENTA _
        + vbCr + " GROUP BY con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion,con_planctas.tipsal " _
        + vbCr + " ORDER BY con_planctas.cuenta, con_planctas.descripcion;"

    '--UNIFICANDO LOS NOMBRES DE LOS CAMPOS TANTO PARA DOLARES Y SOLES
    nSQL = Replace(nSQL, "saldebesol", "saldeb")
    nSQL = Replace(nSQL, "salhabersol", "salhab")
    nSQL = Replace(nSQL, "maydebesol", "maydeb")
    nSQL = Replace(nSQL, "mayhabersol", "mayhab")
    nSQL = Replace(nSQL, "debesol", "movdeb")
    nSQL = Replace(nSQL, "habersol", "movhab")
    nSQL = Replace(nSQL, "deudorsol", "deudor")
    nSQL = Replace(nSQL, "acreedorsol", "acreedor")
    

    nSQL = Replace(nSQL, "saldebedol", "saldeb")
    nSQL = Replace(nSQL, "salhaberdol", "salhab")
    nSQL = Replace(nSQL, "maydebedol", "maydeb")
    nSQL = Replace(nSQL, "mayhaberdol", "mayhab")
    nSQL = Replace(nSQL, "debedol", "movdeb")
    nSQL = Replace(nSQL, "haberdol", "movhab")
    nSQL = Replace(nSQL, "deudordol", "deudor")
    nSQL = Replace(nSQL, "acreedordol", "acreedor")
    
    If opt_fecha(0).Value = True Then '--por fecha
        '--REEMPLAZANDO EL INTERVALO DE FECHA
        nSQL = Replace(nSQL, "con_diario1.fchasi >= CDate('01/01/07') and con_diario1.fchasi <= cdate('31/12/07')", " ( con_diario1.fchasi >=CDate('" + Me.TxtFchIni.Valor + "') And con_diario1.fchasi <= CDate('" + Me.TxtFchFin.Valor + "') ) ")
        '--REEMPLAZANDO LA FECHA DE INICIO PARA OBTENER LOS SALDOS
        nSQL = Replace(nSQL, "con_diario1.fchasi < CDate('01/01/07')", " ( con_diario1.fchasi < CDate('" + Me.TxtFchIni.Valor + "') ) ")
    Else '--por periodo
        '--REEMPLAZANDO EL INTERVALO DE FECHA
        nSQL = Replace(nSQL, "con_diario1.fchasi >= CDate('01/01/07') and con_diario1.fchasi <= cdate('31/12/07')", " ( con_diario1.año = " & AnoTra & " and con_diario1.idmes >= " & mMesIni & " and con_diario1.idmes <= " & mMesFin & " ) ")
        '--REEMPLAZANDO LA FECHA DE INICIO PARA OBTENER LOS SALDOS
        nSQL = Replace(nSQL, "con_diario1.fchasi < CDate('01/01/07')", " ( con_diario1.año = " & AnoTra & " and con_diario1.idmes < " & mMesIni & " ) ")
    End If
    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then GoTo SALIR:
    RstTmp.Filter = adFilterNone
    RstTmp.Sort = "cuenta"
    fra_msg.Visible = True
    If RstTmp.BOF = False Or RstTmp.EOF = False Or RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
    If RstTmp.RecordCount > 0 Then ProgressBar1.Max = RstTmp.RecordCount
    Label3.Caption = "Procesando Resumen"
    For A = 1 To RstTmp.RecordCount
        DoEvents
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then GoTo SALIR:
        '-----------------------------------------------
        ProgressBar1.Value = A
        Fg3.Rows = Fg3.Rows + 1
        Fg3.TextMatrix(A + 1, 1) = RstTmp("cuenta") & ""
        
        Fg3.TextMatrix(A + 1, 2) = RstTmp("descri") & ""
        
        Fg3.TextMatrix(A + 1, 3) = Format(NulosN(RstTmp("saldeb")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 4) = Format(NulosN(RstTmp("salhab")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 5) = Format(NulosN(RstTmp("movdeb")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 6) = Format(NulosN(RstTmp("movhab")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 7) = Format(NulosN(RstTmp("maydeb")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 8) = Format(NulosN(RstTmp("mayhab")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 9) = Format(NulosN(RstTmp("deudor")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 10) = Format(NulosN(RstTmp("acreedor")), FORMAT_MONTO)
        
        
        xAcumulado(0) = xAcumulado(0) + NulosN(Fg3.TextMatrix(A + 1, 3)) '--saldeb
        xAcumulado(1) = xAcumulado(1) + NulosN(Fg3.TextMatrix(A + 1, 4)) '--salhab
        xAcumulado(2) = xAcumulado(2) + NulosN(Fg3.TextMatrix(A + 1, 5)) '--movdeb
        xAcumulado(3) = xAcumulado(3) + NulosN(Fg3.TextMatrix(A + 1, 6)) '--movhab
        xAcumulado(4) = xAcumulado(4) + NulosN(Fg3.TextMatrix(A + 1, 7)) '--maydeb
        xAcumulado(5) = xAcumulado(5) + NulosN(Fg3.TextMatrix(A + 1, 8)) '--mayhab
        xAcumulado(6) = xAcumulado(6) + NulosN(Fg3.TextMatrix(A + 1, 9)) '--deudor
        xAcumulado(7) = xAcumulado(7) + NulosN(Fg3.TextMatrix(A + 1, 10)) '--acreedor
        
        RstTmp.MoveNext
        If RstTmp.EOF = True Then Exit For
    Next A
    
    Fg3.Rows = Fg3.Rows + 2
    
    Fg3.TextMatrix(Fg3.Rows - 1, 2) = "TOTAL =>"
    '-------------------------------
    Dim Col&
    For A = 0 To UBound(xAcumulado())
        Fg3.TextMatrix(Fg3.Rows - 1, 3 + A) = Format(xAcumulado(A), FORMAT_MONTO)
        FORMATO_CELDA Fg3, Fg3.Rows - 1, 3 + A, , True
    Next A
    Erase xAcumulado()
    '-------------------------------
    If RstTmp.RecordCount > 0 Then
        GRID_COLOR_FONDO Fg3, 2, 3, Fg3.Rows - 3, 4, RGB(255, 255, 236)
        GRID_COLOR_FONDO Fg3, 2, 7, Fg3.Rows - 3, 8, RGB(255, 255, 236)
        
    End If
        GRID_COLOR_FONDO Fg3, Fg3.Rows - 2, 1, Fg3.Rows - 1, Fg3.Cols - 1, RGB(231, 254, 224)
    
SALIR:
    Set RstRes = Nothing
    Frame5.Visible = False
    fra_msg.Visible = False
    
    MsgBox "El Mayor se terminó de procesar con éxito", vbInformation, xTitulo
    
    Exit Sub
error:
    Frame5.Visible = False
    Set RstRes = Nothing
    fra_msg.Visible = False
    SHOW_ERROR Me.Name, "CargarResumen"
End Sub

Private Sub pExportar()
    If TabOne1.CurrTab = 0 Then
        If Fg1.Rows = 1 Then
            MsgBox "No hay registros para exportar", vbExclamation, xTitulo
            Exit Sub
        End If
    ElseIf TabOne1.CurrTab = 1 Then
        If Fg3.Rows = 1 Then
            MsgBox "No hay registros para exportar", vbExclamation, xTitulo
            Exit Sub
        End If
    End If
    If fValidarConsulta() = False Then Exit Sub
    
    If TabOne1.CurrTab = 0 Then ExportarExcelDetalle
    If TabOne1.CurrTab = 1 Then ExportarExcelResumen
    
End Sub

Private Sub Fg2_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 45 Then
        CmdAdd_Click
    End If
    If KeyCode = 46 Then
        CmdDel_Click
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
    
        lbl_periodo(0).Caption = Busca_Codigo(xMes, "id", "descripcion", "con_meses", "N", xCon)
        lbl_periodo(1).Caption = lbl_periodo(0).Caption
        mMesIni = xMes
        mMesFin = xMes

        TabOne1.CurrTab = 0

        TxtFchIni.SetFocus
        SeEjecuto = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        BAND_INTERRUMPIR = True '--interrumpir
    ElseIf KeyCode = vbKeyF3 And Shift = 0 Then
        BuscarVSFlexGrid
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    Fg2.Rows = 1
    Fg1.Rows = 1
    Fg3.Rows = 1
    Fg1.SelectionMode = flexSelectionByRow
    Fg2.SelectionMode = flexSelectionByRow
    Fg3.SelectionMode = flexSelectionByRow
    Fg1.Editable = flexEDNone
    Fg2.Editable = flexEDNone
    Fg3.Editable = flexEDNone
    
    Fg2.Tag = Fg2.FormatString
    
    Fg2.ColWidth(3) = 0
    Frame2.BackColor = &H8000000F
    Frame3.BackColor = &H8000000F
    TxtFchIni.Valor = Date
    TxtFchFin.Valor = Date
    
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    
    pConfigurarGrilla True
    
    TabOne1.CurrTab = 0
  
End Sub


Private Function MuestraMayor() As Boolean
    Dim RstMay As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstSal As New ADODB.Recordset
    
    Dim A&, B&, C&
    Dim nSQL As String
    On Error GoTo error

    Frame5.Left = 3413
    Frame5.Top = 2685
    Frame5.Visible = True
    ProgressBar1.Min = 1
    DoEvents
    
    Dim SQL_CUENTA As String
    SQL_CUENTA = ""
    '--SI AGREGA CUENTAS AS GRID, GENERAR EL FILTRO A CONCATENAR A LA CONSULTA
    For A = 1 To Fg2.Rows - 1
        If Trim(Fg2.TextMatrix(A, 1)) <> "" Then
            SQL_CUENTA = SQL_CUENTA + " con_planctas.cuenta Like '" & Trim(Fg2.TextMatrix(A, 1)) & "%' OR "
        End If
    Next A
    If SQL_CUENTA <> "" Then SQL_CUENTA = " AND (" + Left(SQL_CUENTA, Len(SQL_CUENTA) - 3) + ") "
    '---------
    '--ESTABLECER EL CAMPO A TOTALIZAR EN FUNCION DEL RECORDSET TMP (RstTmp2) , TANTO A SOLES Y DOLARES
    Dim CAMPO_DEBE, CAMPO_HABER  As String
    If OptSoles.Value = True Then
        CAMPO_DEBE = "impdebsol":  CAMPO_HABER = "imphabsol"
    End If
    If OptDolares.Value = True Then
        CAMPO_DEBE = "impdebdol":  CAMPO_HABER = "imphabdol"
    End If
    '-----------------------------------------------------------------------
    '--SI SE NTERRUMPE EL PROCESO => SALIR
     If BAND_INTERRUMPIR = True Then GoTo SALIR:
     '-----------------------------------------------
    Set RstTmp2 = Nothing
    'nSQL = "SELECT con_diario.idcue AS id,  con_diario.idcue as idcuenta, con_planctas.cuenta, con_planctas.descripcion AS descri, con_planctas.tipsal, con_diario.idlib, con_diario.idmov, Format([con_diario]![idmes],'00') & IIf(mae_libros.codsun Is Null OR mae_libros.codsun ='','FF',Format([mae_libros].[codsun],'00')) & Trim([con_diario]![numasi]) AS numreg, mae_libros.descripcion AS nomlib, con_tc.impven, " _
        + vbCr + " IIf(con_diario.idlib=0 Or con_diario.idlib Is Null,' ',IIf(con_diario.idlib=1, mae_documento.abrev,IIf(con_diario.idlib=2, mae_documento_1.abrev,IIf(con_diario.idlib=3, mae_documento_2.abrev,IIf(con_diario.idlib=4, mae_documento_3.abrev,IIf(con_diario.idlib=5, mae_documento_4.abrev,IIf(con_diario.idlib=6, mae_doccajaban.abrev,IIf(con_diario.idlib=8,'CAN',IIf(con_diario.idlib=9,' ',IIf(con_diario.idlib=37,'CAN',IIf(con_diario.idlib=38,mae_doccajaban_1.abrev,IIf(con_diario.idlib=39,mae_documento_5.abrev,'OTROS LIBROS')))))))))))) AS tipdoc, " _
        + vbCr + " IIf(con_diario.idlib=0 Or con_diario.idlib Is Null,' ',IIf(con_diario.idlib=1,com_compras.fchdoc,IIf(con_diario.idlib=2,vta_ventas.fchdoc,IIf(con_diario.idlib=3,con_proviciones.fchdoc,IIf(con_diario.idlib=4,con_percepcion.fchdoc,IIf(con_diario.idlib=5,con_retencion.fchemi,IIf(con_diario.idlib=6,con_cajabanco.fchope,IIf(con_diario.idlib=8,con_canjes.fchemi,IIf(con_diario.idlib=9,' ',IIf(con_diario.idlib=37,con_letra.fchemi,IIf(con_diario.idlib=38,con_ctasrendir.fchemi,IIf(con_diario.idlib=39,con_devoluciones.fchemi,'OTROS LIBROS')))))))))))) AS fchemi, " _
        + vbCr + " IIf(con_diario.idlib=0 Or con_diario.idlib Is Null,' ',IIf(con_diario.idlib=1,com_compras.numser & '-' & com_compras.numdoc,IIf(con_diario.idlib=2,vta_ventas.numser & '-' & vta_ventas.numdoc,IIf(con_diario.idlib=3,con_proviciones.numser & '-' & con_proviciones.numdoc,IIf(con_diario.idlib=4,con_percepcion.numser & '-' & con_percepcion.numdoc,IIf(con_diario.idlib=5,con_retencion.numser & '-' & con_retencion.numdoc,IIf(con_diario.idlib=6,con_cajabanco.numdoc,IIf(con_diario.idlib=8,[con_canjes].[numser] & '-' & [con_canjes].[numdoc],IIf(con_diario.idlib=9,' ',IIf(con_diario.idlib=37,' ',IIf(con_diario.idlib=38,con_ctasrendir.numdoc,IIf(con_diario.idlib=39,con_devoluciones.numdoc,'OTROS LIBROS')))))))))))) AS numdoc, " _
        + vbCr + " IIf(con_diario.impdebdol<>0,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol) AS impdebsol, " _
        + vbCr + " IIf(con_diario.imphabdol<>0,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol) AS imphabsol, " _
        + vbCr + " IIf([con_diario].[impdebdol]<>0,[con_diario].[impdebdol],IIf([con_tc].[impven] Is Null Or [con_diario].[impdebsol]=0,0,([con_diario].[impdebsol]/[con_tc].[impven]))) AS impdebdol, " _
        + vbCr + " IIf([con_diario].[imphabdol]<>0,[con_diario].[imphabdol],IIf([con_tc].[impven] Is Null Or [con_diario].[imphabsol]=0,0,([con_diario].[imphabsol]/[con_tc].[impven]))) AS imphabdol, " _
        + vbCr + " (IIf(con_planctas.tipsal='D' Or con_planctas.tipsal Is Null Or con_planctas.tipsal='',(impdebsol-imphabsol),(imphabsol-impdebsol))) AS saldosol, " _
        + vbCr + " (IIf(con_planctas.tipsal='D' Or con_planctas.tipsal Is Null Or con_planctas.tipsal='',(impdebdol-imphabdol),(imphabdol-impdebdol))) AS saldodol " _
        + vbCr + " FROM (((((mae_libros RIGHT JOIN (con_planctas RIGHT JOIN ((((((((((((((con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN com_compras ON con_diario.idmov = com_compras.id) LEFT JOIN vta_ventas ON con_diario.idmov = vta_ventas.id) LEFT JOIN con_retencion ON con_diario.idmov = con_retencion.id) LEFT JOIN con_percepcion ON con_diario.idmov = con_percepcion.id) LEFT JOIN con_proviciones ON con_diario.idmov = con_proviciones.id) LEFT JOIN con_cajabanco ON con_diario.idmov = con_cajabanco.id) LEFT JOIN con_canjes ON con_diario.idmov = con_canjes.id) LEFT JOIN mae_doccajaban ON con_cajabanco.iddoc = mae_doccajaban.id) LEFT JOIN mae_documento ON com_compras.tipdoc = mae_documento.id) LEFT JOIN mae_documento AS mae_documento_1 ON vta_ventas.tipdoc = mae_documento_1.id) LEFT JOIN mae_documento AS mae_documento_2 ON con_proviciones.tipdoc = mae_documento_2.id)  " _
        + vbCr + "      LEFT JOIN mae_documento AS mae_documento_3 ON con_percepcion.tipdoc = mae_documento_3.id) LEFT JOIN mae_documento AS mae_documento_4 ON con_retencion.iddoc = mae_documento_4.id) ON con_planctas.id = con_diario.idcue) ON mae_libros.id = con_diario.idlib) LEFT JOIN con_letra ON con_diario.idmov = con_letra.id) LEFT JOIN con_ctasrendir ON con_diario.idmov = con_ctasrendir.id) LEFT JOIN con_devoluciones ON con_diario.idmov = con_devoluciones.id) LEFT JOIN mae_doccajaban AS mae_doccajaban_1 ON con_ctasrendir.tipdoc = mae_doccajaban_1.id) LEFT JOIN mae_documento AS mae_documento_5 ON con_devoluciones.iddoc = mae_documento_5.id "
        
    nSQL = "SELECT con_diario.idcue AS id, con_diario.idcue AS idcuenta, con_planctas.cuenta, con_planctas.descripcion AS descri, con_planctas.tipsal, con_diario.idlib, con_diario.idmov, Format([con_diario]![idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',Format([mae_libros].[codsun],'00')) & Trim([con_diario]![numasi]) AS numreg, IIf([con_diario].[idlib]<>3,[mae_libros].[descripcion],[mae_librossub].[descripcion]) AS nomlib, con_tc.impven, " _
        & vbCr & " IIf(con_diario.idlib=0 Or con_diario.idlib Is Null,' ',IIf(con_diario.idlib=1,mae_documento.abrev,IIf(con_diario.idlib=2,mae_documento_1.abrev,IIf(con_diario.idlib=3,mae_documento_2.abrev,IIf(con_diario.idlib=4,mae_documento_3.abrev,IIf(con_diario.idlib=5,mae_documento_4.abrev,IIf(con_diario.idlib=6,mae_doccajaban.abrev,IIf(con_diario.idlib=8,'CAN',IIf(con_diario.idlib=9,' ',IIf(con_diario.idlib=37,'CAN',IIf(con_diario.idlib=38,mae_doccajaban_1.abrev,IIf(con_diario.idlib=39,mae_documento_5.abrev,'OTROS LIBROS')))))))))))) AS tipdoc, " _
        & vbCr & " IIf(con_diario.idlib=0 Or con_diario.idlib Is Null,' ',IIf(con_diario.idlib=1,com_compras.fchdoc,IIf(con_diario.idlib=2,vta_ventas.fchdoc,IIf(con_diario.idlib=3,con_proviciones.fchdoc,IIf(con_diario.idlib=4,con_percepcion.fchdoc,IIf(con_diario.idlib=5,con_retencion.fchemi,IIf(con_diario.idlib=6,con_cajabanco.fchope,IIf(con_diario.idlib=8,con_canjes.fchemi,IIf(con_diario.idlib=9,' ',IIf(con_diario.idlib=37,con_letra.fchemi,IIf(con_diario.idlib=38,con_ctasrendir.fchemi,IIf(con_diario.idlib=39,con_devoluciones.fchemi,'OTROS LIBROS')))))))))))) AS fchemi, " _
        & vbCr & " IIf(con_diario.idlib=0 Or con_diario.idlib Is Null,' ',IIf(con_diario.idlib=1,com_compras.numser & '-' & com_compras.numdoc,IIf(con_diario.idlib=2,vta_ventas.numser & '-' & vta_ventas.numdoc,IIf(con_diario.idlib=3,con_proviciones.numser & '-' & con_proviciones.numdoc,IIf(con_diario.idlib=4,con_percepcion.numser & '-' & con_percepcion.numdoc,IIf(con_diario.idlib=5,con_retencion.numser & '-' & con_retencion.numdoc,IIf(con_diario.idlib=6,con_cajabanco.numdoc,IIf(con_diario.idlib=8,[con_canjes].[numser] & '-' & [con_canjes].[numdoc],IIf(con_diario.idlib=9,' ',IIf(con_diario.idlib=37,' ',IIf(con_diario.idlib=38,con_ctasrendir.numdoc,IIf(con_diario.idlib=39,con_devoluciones.numdoc,'OTROS LIBROS')))))))))))) AS numdoc, " _
        & vbCr & " IIf(con_diario.impdebdol<>0,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol) AS impdebsol, " _
        & vbCr & " IIf(con_diario.imphabdol<>0,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol) AS imphabsol, " _
        & vbCr & " IIf([con_diario].[impdebdol]<>0,[con_diario].[impdebdol],IIf([con_tc].[impven] Is Null Or [con_diario].[impdebsol]=0,0,([con_diario].[impdebsol]/[con_tc].[impven]))) AS impdebdol, " _
        & vbCr & " IIf([con_diario].[imphabdol]<>0,[con_diario].[imphabdol],IIf([con_tc].[impven] Is Null Or [con_diario].[imphabsol]=0,0,([con_diario].[imphabsol]/[con_tc].[impven]))) AS imphabdol, " _
        & vbCr & " (IIf(con_planctas.tipsal='D' Or con_planctas.tipsal Is Null Or con_planctas.tipsal='',(impdebsol-imphabsol),(imphabsol-impdebsol))) AS saldosol, " _
        & vbCr & " (IIf(con_planctas.tipsal='D' Or con_planctas.tipsal Is Null Or con_planctas.tipsal='',(impdebdol-imphabdol),(imphabdol-impdebdol))) AS saldodol, " _
        & vbCr & " IIf([con_diario].[idlib]=2,[mae_cliente]![nombre],IIf([con_diario].[idlib]=1,[mae_prov]![nombre],'')) AS nomclipro " _
        & vbCr & " FROM (mae_prov RIGHT JOIN (mae_libros RIGHT JOIN (mae_cliente RIGHT JOIN ((((((con_planctas RIGHT JOIN ((((((((((((((con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN com_compras ON con_diario.idmov = com_compras.id) LEFT JOIN vta_ventas ON con_diario.idmov = vta_ventas.id) LEFT JOIN con_retencion ON con_diario.idmov = con_retencion.id) LEFT JOIN con_percepcion ON con_diario.idmov = con_percepcion.id) LEFT JOIN con_proviciones ON con_diario.idmov = con_proviciones.id) LEFT JOIN con_cajabanco ON con_diario.idmov = con_cajabanco.id) LEFT JOIN con_canjes ON con_diario.idmov = con_canjes.id) LEFT JOIN mae_doccajaban ON con_cajabanco.iddoc = mae_doccajaban.id) LEFT JOIN mae_documento ON com_compras.tipdoc = mae_documento.id) LEFT JOIN mae_documento AS mae_documento_1 ON vta_ventas.tipdoc = mae_documento_1.id) LEFT JOIN mae_documento AS mae_documento_2 ON con_proviciones.tipdoc = mae_documento_2.id)  " _
        & vbCr & " LEFT JOIN mae_documento AS mae_documento_3 ON con_percepcion.tipdoc = mae_documento_3.id) LEFT JOIN mae_documento AS mae_documento_4 ON con_retencion.iddoc = mae_documento_4.id) ON con_planctas.id = con_diario.idcue) LEFT JOIN con_letra ON con_diario.idmov = con_letra.id) LEFT JOIN con_ctasrendir ON con_diario.idmov = con_ctasrendir.id) LEFT JOIN con_devoluciones ON con_diario.idmov = con_devoluciones.id) LEFT JOIN mae_doccajaban AS mae_doccajaban_1 ON con_ctasrendir.tipdoc = mae_doccajaban_1.id) LEFT JOIN mae_documento AS mae_documento_5 ON con_devoluciones.iddoc = mae_documento_5.id) ON mae_cliente.id = vta_ventas.idcli) ON mae_libros.id = con_diario.idlib) ON mae_prov.id = com_compras.idpro) LEFT JOIN mae_librossub ON (con_proviciones.idlib = mae_librossub.idlib) AND (con_proviciones.idsublib = mae_librossub.id) "

'WHERE (((con_planctas.cuenta) Like '42*') AND ((con_diario.fchasi)>=CDate('01/01/2007') And (con_diario.fchasi)<=CDate('31/12/2007') And (con_diario.fchasi)>=CDate('01/01/2007') And (con_diario.fchasi)<=CDate('31/12/2007')))
'ORDER BY con_planctas.cuenta;

   
    If opt_fecha(0).Value = True Then
        nSQL = nSQL + vbCr + " WHERE (con_diario.fchasi >=CDate('" + TxtFchIni.Valor + "') And con_diario.fchasi<=CDate('" + TxtFchFin.Valor + "')) " _
            + vbCr + " AND ( con_diario.fchasi >=CDate('01/01/" + AnoTra + "') And con_diario.fchasi <= CDate('31/12/" + AnoTra + "') ) "
    Else
        nSQL = nSQL + vbCr + " WHERE ( con_diario.idmes >= " & mMesIni & " and con_diario.idmes <= " & mMesFin & " ) and con_diario.año = " & AnoTra & " "
    End If
        
    nSQL = nSQL + SQL_CUENTA + vbCr + " ORDER BY con_planctas.cuenta ASC "


    RST_Busq RstTmp2, nSQL, xCon

    'HACEMOS UNA CONSULTA DE LOS REGISTROS UNICOS DE LA CONSULTA ANTERIOR, PARA PODER TOTALIZARLA

    Fg1.Rows = 1
    Dim xFila&
    Dim xSaldo As Double
    Dim xTotal1, xTotal2 As Double
    Dim xTotal1_1, xTotal2_1 As Double
    
    xFila = 1

    DoEvents
    '--SI SE NTERRUMPE EL PROCESO => SALIR
     If BAND_INTERRUMPIR = True Then GoTo SALIR:
    '-----------------------------------------------
    nSQL = "SELECT con_diario.idcue as idcuenta, con_planctas.cuenta, con_planctas.descripcion, Sum(con_diario.impdebsol) AS SumaDeimpdebsol, " _
         + vbCr + " Sum(con_diario.imphabsol) AS SumaDeimphabsol, Sum(con_diario.impdebdol) AS SumaDeimpdebdol, Sum(con_diario.imphabdol) AS SumaDeimphabdol " _
         + vbCr + " FROM con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue "
                  
    If opt_fecha(0).Value = True Then
         nSQL = nSQL + vbCr + " WHERE (con_diario.fchasi >=CDate('" & TxtFchIni.Valor & "') And con_diario.fchasi <=CDate('" & TxtFchFin.Valor & "')) " _
         + vbCr + " AND ( con_diario.fchasi >=CDate('01/01/" + AnoTra + "') And con_diario.fchasi <= CDate('31/12/" + AnoTra + "') ) "
    Else
        nSQL = nSQL + vbCr + " WHERE ( con_diario.idmes >= " & mMesIni & " and con_diario.idmes <= " & mMesFin & " )and con_diario.año = " & AnoTra & " "
    End If
         
    nSQL = nSQL + vbCr + SQL_CUENTA _
         + vbCr + " GROUP BY con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion ORDER BY con_planctas.cuenta ASC "
    
    RST_Busq RstMay, nSQL, xCon

    xSaldo = 0
    '---
    If RstMay.RecordCount <> 0 Then
    
        fra_msg.Visible = True
        
        RstMay.MoveFirst
        
        If RstMay.RecordCount > 1 Then ProgressBar1.Max = RstMay.RecordCount
        
        DoEvents
        Do While Not RstMay.EOF
            DoEvents
            '--SI SE NTERRUMPE EL PROCESO => SALIR
            If BAND_INTERRUMPIR = True Then GoTo SALIR:
            '-----------------------------------------------
            ProgressBar1.Value = RstMay.Bookmark
    
            xSaldo = 0
            Fg1.Rows = Fg1.Rows + 1
            
            UNIR_CELDAS Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 7, "Cta Nº  :  " + NulosC(RstMay("cuenta")) & "   - " + NulosC(RstMay.Fields("descripcion")), flexAlignLeftCenter
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 1, , True
                        
            xTotal1 = 0
            xTotal2 = 0
            
           'hallamos el saldo anterior de la cuenta
            Set RstSal = Nothing
                            
            nSQL = "SELECT con_diario.idcue as idcuenta, con_planctas.cuenta,con_planctas.tipsal, " _
                + vbCr + " Sum(IIf([con_diario].[impdebdol]<>0,IIf([con_tc].[impven] Is Null,0,[con_diario].[impdebdol]*[con_tc].[impven]),[con_diario].[impdebsol])) AS impdebsol, " _
                + vbCr + " Sum(IIf([con_diario].[imphabdol]<>0,IIf([con_tc].[impven] Is Null,0,[con_diario].[imphabdol]*[con_tc].[impven]),[con_diario].[imphabsol])) AS imphabsol, " _
                + vbCr + " Sum(IIf([con_diario].[impdebdol]<>0,[con_diario].[impdebdol],IIf([con_diario].[impdebsol]=0 Or [con_tc].[impven] Is Null,0,[con_diario].[impdebsol]/[con_tc].[impven]))) AS impdebdol, " _
                + vbCr + " Sum(IIf([con_diario].[imphabdol]<>0,[con_diario].[imphabdol],IIf([con_diario].[imphabsol]=0 Or [con_diario]![imphabsol] Is Null Or [con_tc].[impven] Is Null, 0, [con_diario].[imphabsol] / [con_tc].[impven]))) As imphabdol " _
                + vbCr + " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue "
            If opt_fecha(0).Value = True Then
                nSQL = nSQL + vbCr + " WHERE (con_diario.fchasi Is Null and con_diario.año = " & AnoTra & " ) Or (con_diario.fchasi < CDate('" & TxtFchIni.Valor & "')) "
            Else
                nSQL = nSQL + vbCr + " WHERE (con_diario.idmes < " & mMesIni & " AND con_diario.año = " & AnoTra & " )  or (con_diario.año < " & AnoTra & " ) "
            End If
            nSQL = nSQL + vbCr + " GROUP BY con_diario.idcue, con_planctas.cuenta, con_planctas.tipsal  " _
                + vbCr + " HAVING con_planctas.cuenta ='" & RstMay("cuenta") & "' "
                            
            RST_Busq RstSal, nSQL, xCon
            
            If RstSal.RecordCount <> 0 Then
                    If UCase(RstSal.Fields("tipsal") & "") = "D" Or NulosC(RstSal.Fields("tipsal")) = "" Then
                        xSaldo = (NulosN(RstSal(CAMPO_DEBE)) - NulosN(RstSal(CAMPO_HABER)))
                    Else
                        xSaldo = (NulosN(RstSal(CAMPO_HABER)) - NulosN(RstSal(CAMPO_DEBE)))
                    End If
                    Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(xSaldo, FORMAT_MONTO)
            Else
                xSaldo = 0
                Fg1.TextMatrix(Fg1.Rows - 1, 10) = "0.00"
            End If
                        
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, , True
            
            RstTmp2.Filter = adFilterNone
            RstTmp2.Filter = "idcuenta = '" & NulosC(RstMay("idcuenta")) & "'"
            Label3.Caption = "Procesando Cta Nº  :  " + NulosC(RstMay("cuenta"))
            DoEvents
            xFila = xFila + 1
            If RstTmp2.RecordCount <> 0 Then
            
                RstTmp2.MoveFirst
                If opt(0).Value = True Then
                    RstTmp2.Sort = "fchemi"
                ElseIf opt(1).Value = True Then
                    RstTmp2.Sort = "numdoc"
                Else
                    RstTmp2.Sort = "numreg"
                End If
                Do While Not RstTmp2.EOF
                    DoEvents
                    '--SI SE NTERRUMPE EL PROCESO => SALIR
                    If BAND_INTERRUMPIR = True Then GoTo SALIR
                    '-----------------------------------------------
                    Fg1.Rows = Fg1.Rows + 1
                    Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(RstMay("cuenta"))
                    Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstTmp2("numreg"))
                    Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(RstTmp2("nomlib"))
                    
                    Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(RstTmp2("tipdoc"))
                    
                    If IsDate(RstTmp2("fchemi")) = True Then
                        Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(RstTmp2("fchemi"), FORMAT_DATE)
                    End If
                    Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosC(RstTmp2("numdoc"))
                    Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosC(RstTmp2("impven"))
                    Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(NulosN(RstTmp2(CAMPO_DEBE)), FORMAT_MONTO)
                    Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(NulosN(RstTmp2(CAMPO_HABER)), FORMAT_MONTO)
                    
                    If UCase(RstTmp2.Fields("tipsal") & "") = "D" Or NulosC(RstTmp2.Fields("tipsal")) = "" Then
                        xSaldo = xSaldo + (NulosN(RstTmp2(CAMPO_DEBE)) - NulosN(RstTmp2(CAMPO_HABER)))
                    Else
                        xSaldo = xSaldo + (NulosN(RstTmp2(CAMPO_HABER)) - NulosN(RstTmp2(CAMPO_DEBE)))
                    End If
                    
                    Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(xSaldo, FORMAT_MONTO)
                    
                    Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosC(RstTmp2("nomclipro"))
                    
                    xTotal1 = xTotal1 + NulosN(RstTmp2(CAMPO_DEBE))
                    xTotal2 = xTotal2 + NulosN(RstTmp2(CAMPO_HABER))
                    RstTmp2.MoveNext
                    If RstTmp2.EOF = True Then
                        Fg1.Rows = Fg1.Rows + 1
                        Fg1.TextMatrix(Fg1.Rows - 1, 7) = "Total =>"
                        Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(xTotal1, FORMAT_MONTO)
                        Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(xTotal2, FORMAT_MONTO)
                        
                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 7, , True
                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 8, , True
                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 9, , True
                        
                        Fg1.Rows = Fg1.Rows + 1
                        Exit Do
                    End If
                    xFila = xFila + 1
                Loop
            End If
            
            xTotal1_1 = xTotal1_1 + xTotal1
            xTotal2_1 = xTotal2_1 + xTotal2

            RstMay.MoveNext
            
        Loop
        
    Else
        xSaldo = 0

        'hallamos el saldo anterior de la cuenta

        nSQL = "SELECT con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion ,con_planctas.tipsal, " _
            + vbCr + " Sum(IIf([con_diario].[impdebdol]<>0,IIf([con_tc].[impven] Is Null,0,[con_diario].[impdebdol]*[con_tc].[impven]),[con_diario].[impdebsol])) AS impdebsol, " _
            + vbCr + " Sum(IIf([con_diario].[imphabdol]<>0,IIf([con_tc].[impven] Is Null,0,[con_diario].[imphabdol]*[con_tc].[impven]),[con_diario].[imphabsol])) AS imphabsol, " _
            + vbCr + " Sum(IIf([con_diario].[impdebdol]<>0,[con_diario].[impdebdol],IIf([con_diario].[impdebsol]=0 Or [con_tc].[impven] Is Null,0,[con_diario].[impdebsol]/[con_tc].[impven]))) AS impdebdol, " _
            + vbCr + " Sum(IIf([con_diario].[imphabdol]<>0,[con_diario].[imphabdol], IIf([con_diario].[imphabsol] = 0 Or [con_diario].[imphabsol] Is Null Or [con_tc].[impven] Is Null, 0, [con_diario].[imphabsol] / [con_tc].[impven]))) As imphabdol " _
            + vbCr + " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue "

        If opt_fecha(0).Value = True Then
            nSQL = nSQL + vbCr + " WHERE (con_diario.fchasi Is Null and con_diario.año = " & AnoTra & " ) Or (con_diario.fchasi < CDate('" & TxtFchIni.Valor & "')) "
        Else
            nSQL = nSQL + vbCr + " WHERE (con_diario.idmes < " & mMesIni & " AND con_diario.año = " & AnoTra & " )  or (con_diario.año < " & AnoTra & " ) "
        End If
        
        nSQL = nSQL + SQL_CUENTA _
            + vbCr + " GROUP BY con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion,con_planctas.tipsal ; " _
            + vbCr + " "
            
        RST_Busq RstSal, nSQL, xCon
        
        If RstSal.EOF = False Or RstSal.BOF = False Or RstSal.RecordCount <> 0 Then
            Fg1.Rows = Fg1.Rows + 1
            RstSal.MoveFirst
        End If
        
        Do While Not RstSal.EOF
            UNIR_CELDAS Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 5, "Cta Nº  :  " + NulosC(RstSal("cuenta")) & "   - " + NulosC(RstSal.Fields("descripcion")), flexAlignLeftCenter
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 1, , True
            
            If UCase(RstSal.Fields("tipsal") & "") = "D" Or RstSal.Fields("tipsal") = "" Then
                xSaldo = (NulosN(RstSal(CAMPO_DEBE)) - NulosN(RstSal(CAMPO_HABER)))
            Else
                xSaldo = (NulosN(RstSal(CAMPO_HABER)) - NulosN(RstSal(CAMPO_DEBE)))
            End If
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(xSaldo, FORMAT_MONTO)
            Fg1.Rows = Fg1.Rows + 1
            RstSal.MoveNext
        Loop
        Set RstSal = Nothing
        '----------------------------
    End If

    If xTotal1_1 <> 0 Or xTotal2_1 <> 0 Then
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 7) = "Total Gral =>"
        Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(xTotal1_1, FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(xTotal2_1, FORMAT_MONTO)
        
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 7, , True
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 8, , True
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 9, , True
    End If

SALIR:
    Set RstMay = Nothing:     Set RstDet = Nothing:     Set RstSal = Nothing
    Frame5.Visible = False
    fra_msg.Visible = False
    MuestraMayor = True
    Exit Function
error:
    Set RstMay = Nothing:     Set RstDet = Nothing:     Set RstSal = Nothing
    Frame5.Visible = False
    fra_msg.Visible = False
    SHOW_ERROR Me.Name, "MuestraMayor"
End Function


Sub ExportarExcelDetalle()
    Dim A&, B&, xFilas&
    
    On Error GoTo error

    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    
    'determina el numero de hojas que se mostrara en el Excel
    objExcel.SheetsInNewWorkbook = 1
    
    'abre el Libro
    objExcel.Workbooks.Add  'Trim(App.Path) + "\RegCompras.xls"
    
    objExcel.WindowState = 1
    
    objExcel.Visible = True
    With objExcel.ActiveSheet
        
        .Cells(1, 2) = NomEmp
        .Cells(1, 11) = Date
        .Cells(2, 2) = "Nº R.U.C. : " + NumRUC
        
        '**********************************
        Dim nPeriodo As String
        Dim nTitulo1 As String
        If opt_fecha(0).Value = True Then
            If CDate(Me.TxtFchIni.Valor) <> CDate(Me.TxtFchFin.Valor) Then
                nPeriodo = " Del: " + CStr(TxtFchIni.Valor) + " Al: " + CStr(TxtFchFin.Valor)
            Else
                nPeriodo = "Al: " + CStr(TxtFchIni.Valor)
            End If
        Else
            If mMesIni = mMesFin Then
                nPeriodo = "Periodo: " + lbl_periodo(0).Caption
            Else
                nPeriodo = "Periodo: De " + lbl_periodo(0).Caption & " A " & lbl_periodo(1).Caption
            End If
            
        End If
        If Me.OptSoles.Value = True Then
            nTitulo1 = "(Expresado en Nuevos Soles)"
        Else
            nTitulo1 = "(Expresado en Dolares Americanos)"
        End If
        .Cells(4, 2) = "Libro Mayor"
        .Cells(5, 2) = nPeriodo
        .Cells(6, 2) = nTitulo1
        '**********************************
        '--ancho de columna
        For B = 1 To Fg1.Cols - 1
            .Columns(B + 1).ColumnWidth = Fg1.ColWidth(B) / 100
        Next B
        '--encabezado
        xFilas = 8
        For B = 1 To Fg1.Cols - 1
            .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(0, B)
        Next B
        
        .Range("B6:K7").Font.Bold = True
    
        xFilas = xFilas + 1
        For A = 1 To Fg1.Rows - 1
            DoEvents
            For B = 1 To Fg1.Cols - 1
                
                If B <= 7 Then
                    If B = 1 Then
                        .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(A, B)
                        If InStr(Fg1.TextMatrix(A, B), "Cta Nº  :") <> 0 Then
                            .Cells(xFilas, 2) = "'" + Fg1.TextMatrix(A, B)
                            GoTo SIG_FIL
                        End If
                    Else
                        .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(A, B)
                    End If
                    
                    
                Else
                    If IsNumeric(Fg1.TextMatrix(A, B)) = True Then
                        .Cells(xFilas, B + 1) = NulosN(Fg1.TextMatrix(A, B))
                    Else
                        .Cells(xFilas, B + 1) = NulosC(Fg1.TextMatrix(A, B))
                    End If
                End If

            Next B
SIG_FIL:
            xFilas = xFilas + 1
        Next A
    End With
    
    MsgBox "El proceso de exportación terminó con éxito", vbInformation, xTitulo
    objExcel.WindowState = 1
    Set objExcel = Nothing
    Exit Sub
error:
    Set objExcel = Nothing
    SHOW_ERROR Me.Name, "ExportarExcelDetalle", , IIf(Err.Number <> 50290, "", "No manipule el archivo hasta que termine de exportar!!!!")
    
End Sub


Private Sub pConfigurarGrilla(Optional F_CONSERVAR_FORMATO As Boolean = False)
    '--PERMITIRA CONFIGURAR EL FORMATO DE LA CONSULTA
    '--DE ACUERDO A LO QUE SE SELECCIONA
                                   
    Dim k&, j&
    With Fg1
        '-----
        If F_CONSERVAR_FORMATO = True Then LimpiarGrid Fg1, , 1
        .FrozenCols = 0
        Fg1.Cols = 12
                 
        .ColWidth(0) = 200
        '--DATOS DE FILA
        .TextMatrix(0, 1) = "Nº.Cuenta":  .ColWidth(1) = 1000:       .ColAlignment(1) = flexAlignLeftBottom:     .FixedAlignment(1) = flexAlignCenterTop
        .TextMatrix(0, 2) = "Num.Reg.":     .ColWidth(2) = 850:    .ColAlignment(2) = flexAlignLeftCenter:     .FixedAlignment(2) = flexAlignCenterTop
        .TextMatrix(0, 3) = "Libro":        .ColWidth(3) = 1500:    .ColAlignment(3) = flexAlignLeftBottom:     .FixedAlignment(3) = flexAlignCenterTop
        .TextMatrix(0, 4) = "T.D.":         .ColWidth(4) = 450:     .ColAlignment(4) = flexAlignLeftCenter:     .FixedAlignment(4) = flexAlignCenterTop
        .TextMatrix(0, 5) = "Fch. Doc":     .ColWidth(5) = 900:     .ColAlignment(5) = flexAlignCenterBottom:   .FixedAlignment(5) = flexAlignCenterTop
        .TextMatrix(0, 6) = "Nº.Documento": .ColWidth(6) = 1500:    .ColAlignment(6) = flexAlignLeftBottom:     .FixedAlignment(6) = flexAlignCenterTop
        .TextMatrix(0, 7) = "T.C.":         .ColWidth(7) = 800:     .ColAlignment(7) = flexAlignRightBottom:   .FixedAlignment(7) = flexAlignCenterTop
        .TextMatrix(0, 8) = "Debe":         .ColWidth(8) = 1300:    .ColAlignment(8) = flexAlignRightBottom:    .FixedAlignment(8) = flexAlignCenterTop
        .TextMatrix(0, 9) = "Haber":        .ColWidth(9) = 1300:    .ColAlignment(9) = flexAlignRightBottom:    .FixedAlignment(9) = flexAlignCenterTop
        .TextMatrix(0, 10) = "Saldo":       .ColWidth(10) = 1300:   .ColAlignment(10) = flexAlignRightBottom:   .FixedAlignment(10) = flexAlignCenterTop
        .TextMatrix(0, 11) = "Cliente / Proveedor":       .ColWidth(11) = 3000:   .ColAlignment(11) = flexAlignLeftTop:  .FixedAlignment(11) = flexAlignCenterTop
    End With
    
    With Fg3
        '-----
        If F_CONSERVAR_FORMATO = True Then LimpiarGrid Fg3, , 2
        
        .Cols = 11
        .FixedRows = 2
        .FrozenCols = 2
        .RowHeight(0) = 500
        
        UNIR_CELDAS Fg3, 0, 1, 0, 2, "Datos de la Cuenta", flexAlignCenterCenter
        UNIR_CELDAS Fg3, 0, 3, 0, 4, "Saldos Iniciales", flexAlignCenterCenter
        UNIR_CELDAS Fg3, 0, 5, 0, 6, "Movimiento del Periodo", flexAlignCenterCenter
        UNIR_CELDAS Fg3, 0, 7, 0, 8, "Sumas del Mayor", flexAlignCenterCenter
        UNIR_CELDAS Fg3, 0, 9, 0, 10, "Saldos Finales", flexAlignCenterCenter
        
'        '--DATOS DE FILA
        .TextMatrix(1, 1) = "Nº. Cuenta":       .ColWidth(1) = 1100:       .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(1, 2) = "Descripción":      .ColWidth(2) = 3000:       .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(1, 3) = "Debe":       .ColWidth(3) = 1300:       .ColAlignment(3) = flexAlignRightCenter
        .TextMatrix(1, 4) = "Haber":      .ColWidth(4) = 1300:       .ColAlignment(4) = flexAlignRightCenter
        .TextMatrix(1, 5) = "Debe":       .ColWidth(5) = 1320:       .ColAlignment(5) = flexAlignRightCenter
        .TextMatrix(1, 6) = "Haber":      .ColWidth(6) = 1320:       .ColAlignment(6) = flexAlignRightCenter
        .TextMatrix(1, 7) = "Debe":       .ColWidth(7) = 1320:       .ColAlignment(7) = flexAlignRightCenter
        .TextMatrix(1, 8) = "Haber":      .ColWidth(8) = 1320:       .ColAlignment(8) = flexAlignRightCenter
        .TextMatrix(1, 9) = "Debe":       .ColWidth(9) = 1200:       .ColAlignment(9) = flexAlignRightCenter
        .TextMatrix(1, 10) = "Haber":     .ColWidth(10) = 1200:      .ColAlignment(10) = flexAlignRightCenter
        
        '--AGREGANDO LAS FECHAS EN LA CABECERA
        If opt_fecha(0).Value = True Then
            If IsDate(TxtFchIni.Valor) = True Then UNIR_CELDAS Fg3, 0, 3, 0, 4, "Saldos Iniciales" + vbCr + " Al " + CStr(CDate(TxtFchIni.Valor) - 1), flexAlignCenterCenter
            If IsDate(TxtFchFin.Valor) = True Then UNIR_CELDAS Fg3, 0, 9, 0, 10, "Saldos Finales" + vbCr + " Al " + CStr(CDate(TxtFchFin.Valor) - 1), flexAlignCenterCenter
        Else
            If IsDate(TxtFchIni.Valor) = True Then UNIR_CELDAS Fg3, 0, 3, 0, 4, "Saldos Iniciales" + vbCr + " A " + lbl_periodo(0).Caption, flexAlignCenterCenter
            If IsDate(TxtFchFin.Valor) = True Then UNIR_CELDAS Fg3, 0, 9, 0, 10, "Saldos Finales" + vbCr + " A " + lbl_periodo(1).Caption, flexAlignCenterCenter
        End If
    End With
    
    DoEvents
End Sub


Private Function fValidarConsulta() As Boolean
    '--FUNCION QUE VALIDARA LA CONSULTA
    If AnoTra = "" Then
        MsgBox "No Hay Año de trabajo" + vbCr + "No puede continuar", vbExclamation, xTitulo
        Exit Function
    End If
    If opt_fecha(0).Value = True Then '--por fecha
        If NulosC(TxtFchIni.Valor) = "" Then
            MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchIni.SetFocus
            Exit Function
        End If
        
        If NulosC(TxtFchFin.Valor) = "" Then
            MsgBox "No ha especificado la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchFin.SetFocus
            Exit Function
        End If
        
        If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
            MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchIni.SetFocus
            Exit Function
        End If
        
        If (Year(TxtFchIni.Valor) <> Year(TxtFchFin.Valor)) Then
            MsgBox "El rango de fechas debe estar en el Año de Trabajo : " + CStr(AnoTra), vbExclamation, xTitulo
            TxtFchIni.SetFocus
            Exit Function
        ElseIf Year(TxtFchIni.Valor) <> CStr(AnoTra) Then
            MsgBox "El rango de fechas debe estar en el Año de Trabajo : " + CStr(AnoTra), vbExclamation, xTitulo
            TxtFchIni.SetFocus
            Exit Function
        End If
    Else '--por periodo
        If mMesIni > mMesFin Then
            MsgBox "El periodo de inicio debe ser inferior o igual al periodo final", vbExclamation, xTitulo
            cmd_periodo(0).SetFocus
            Exit Function
        End If
    End If
    
'
    If Fg2.Rows = 1 And chk.Value = 0 Then
        MsgBox "No ha especificado una cuenta contable a mayorizar" + vbCr + "Si desea ver todas las cuentas, Active la opción: Procesar Todas las Cuentas...", vbExclamation, xTitulo
        CmdAdd.SetFocus
        Exit Function
    End If
    
    fValidarConsulta = True
End Function



Private Sub ExportarExcelResumen()
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios
    Dim nPeriodo As String
    Dim nTitulo1 As String
    If opt_fecha(0).Value = True Then  '--por fecha
        If CDate(Me.TxtFchIni.Valor) <> CDate(Me.TxtFchFin.Valor) Then
            nPeriodo = " Del: " + CStr(TxtFchIni.Valor) + " Al: " + CStr(TxtFchFin.Valor)
        Else
            nPeriodo = "Al: " + CStr(TxtFchIni.Valor)
        End If
    Else '--por periodo
        If mMesIni = mMesFin Then
            nPeriodo = "Periodo:  " + lbl_periodo(0).Caption
        Else
            nPeriodo = "Periodo:  De " + lbl_periodo(0).Caption & " A " & lbl_periodo(1).Caption
        End If
        
    End If
    If Me.OptSoles.Value = True Then
        nTitulo1 = "(Expresado en Nuevos Soles)"
    Else
        nTitulo1 = "(Expresado en Dolares Americanos)"
    End If
    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, Fg3, "RESUMEN DEL MAYOR", nPeriodo, nTitulo1, "Resumen del Mayor"
    Set X_EXPORT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Exportar"
End Sub


Private Sub BuscarVSFlexGrid()
    On Error GoTo error
    
    If Me.TabOne1.CurrTab <> 0 Then Exit Sub
    Dim X_EXPORT As New SGI2_funciones.formularios
    Dim xCampos(3, 3) As String
    'campo     'columna del grid    'tipo(N:Numerico, C:caracter, F:fecha)      campo predeterminado(0:no se muestra, -1:se muestra al iniciar el formulario)
    xCampos(0, 0) = "Num.Reg.":     xCampos(0, 1) = "2":    xCampos(0, 2) = "C":    xCampos(0, 3) = "-1"
    xCampos(1, 0) = "Libro":        xCampos(1, 1) = "3":    xCampos(1, 2) = "C":    xCampos(1, 3) = "0"
    xCampos(2, 0) = "Fch. Doc   ":  xCampos(2, 1) = "4":    xCampos(2, 2) = "F":    xCampos(2, 3) = "0"
    xCampos(3, 0) = "Nº Documento": xCampos(3, 1) = "5":    xCampos(3, 2) = "C":    xCampos(3, 3) = "0"
    
    X_EXPORT.VSFlexGrid_Buscar Me.hWnd, Fg1, xCampos(), Fg1.Row
    Set X_EXPORT = Nothing
    Me.MousePointer = vbDefault
    
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "BuscarVSFlexGrid"
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    If Button.Index = 3 Then pExportar
    If Button.Index = 4 Then pImprimir
    If Button.Index = 6 Then
        Unload Me
    End If
End Sub

Private Sub opt_fecha_Click(Index As Integer)
    If Index = 0 Then '--por fecha
        TxtFchFin.Visible = True
        TxtFchIni.Visible = True
        lbl(0).Caption = "Fch. Inicio"
        lbl(1).Caption = "Fch. Final"
        
        Ocultar cmd_periodo, False
        Ocultar lbl_periodo, False
    Else '--por periodo
        TxtFchFin.Visible = False
        TxtFchIni.Visible = False
        lbl(0).Caption = "De"
        lbl(1).Caption = "A"
        cmd_periodo(0).Top = 240
        lbl_periodo(0).Top = 210
        cmd_periodo(1).Top = 540
        lbl_periodo(1).Top = 510
        Ocultar cmd_periodo, True
        Ocultar lbl_periodo, True
    End If
End Sub

Private Sub cmd_periodo_Click(Index As Integer)
    If Index = 0 Then
        mMesIni = SeleccionaMes(xCon)
        lbl_periodo(0).Caption = Busca_Codigo(mMesIni, "id", "descripcion", "con_meses", "N", xCon)
    Else
        mMesFin = SeleccionaMes(xCon)
        lbl_periodo(1).Caption = Busca_Codigo(mMesFin, "id", "descripcion", "con_meses", "N", xCon)
    End If
End Sub

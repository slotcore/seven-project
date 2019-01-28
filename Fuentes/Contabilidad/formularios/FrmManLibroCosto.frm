VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmManLibroCosto 
   Caption         =   "Contabilidad - Libro de Costos de Producción"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11625
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   11625
      _ExtentX        =   20505
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
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   14
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
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11040
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto.frx":277E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto.frx":2B10
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLibroCosto.frx":2E2A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Gastos de Fabrica ]"
      Height          =   1065
      Left            =   9120
      TabIndex        =   30
      Top             =   330
      Width           =   2475
      Begin VB.OptionButton optGacFab 
         Caption         =   "Aplicar Todos"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   34
         Top             =   800
         Width           =   1305
      End
      Begin VB.OptionButton optGacFab 
         Caption         =   "Aplicar Prod. Ventas"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   550
         Width           =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblGasFab 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblGasFab"
         Height          =   285
         Left            =   780
         TabIndex        =   32
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "[ Tipo de Operac. ]"
      Height          =   1050
      Left            =   2040
      TabIndex        =   25
      Top             =   350
      Width           =   1665
      Begin VB.CheckBox ckoptCon 
         Caption         =   "Mostrar Ventas"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   29
         Top             =   600
         Width           =   1365
      End
      Begin VB.CheckBox ckoptCon 
         Caption         =   "Mostrar Todos"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   28
         Top             =   410
         Width           =   1335
      End
      Begin VB.OptionButton opttipop 
         Caption         =   "Procesamiento"
         Height          =   195
         Index           =   1
         Left            =   30
         TabIndex        =   27
         Top             =   810
         Width           =   1365
      End
      Begin VB.OptionButton opttipop 
         Caption         =   "Consulta"
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   26
         Top             =   210
         Width           =   975
      End
   End
   Begin VB.CheckBox ckOpcion 
      Caption         =   "[ Proceso ]"
      Height          =   285
      Index           =   0
      Left            =   3810
      TabIndex        =   20
      Top             =   360
      Width           =   1155
   End
   Begin VB.Frame frmOpcion 
      Enabled         =   0   'False
      Height          =   1050
      Index           =   0
      Left            =   3720
      TabIndex        =   19
      Top             =   350
      Width           =   1485
      Begin VB.OptionButton optProceso 
         Caption         =   "Seleccionar"
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   23
         Top             =   450
         Width           =   1185
      End
      Begin VB.OptionButton optProceso 
         Caption         =   "Todos"
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   250
         Width           =   855
      End
      Begin VB.ComboBox ComboSemanas 
         Height          =   315
         ItemData        =   "FrmManLibroCosto.frx":3144
         Left            =   420
         List            =   "FrmManLibroCosto.frx":3146
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   700
         Width           =   1000
      End
   End
   Begin VB.CheckBox ckOpcion 
      Caption         =   "[ Producto ]"
      Height          =   285
      Index           =   1
      Left            =   5310
      TabIndex        =   18
      Top             =   360
      Width           =   1155
   End
   Begin VB.Frame frmOpcion 
      Enabled         =   0   'False
      Height          =   1050
      Index           =   1
      Left            =   5220
      TabIndex        =   14
      Top             =   350
      Width           =   3885
      Begin VB.CommandButton cmd 
         Height          =   240
         Index           =   0
         Left            =   810
         Picture         =   "FrmManLibroCosto.frx":3148
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   240
      End
      Begin VB.TextBox txtIdItem 
         Height          =   300
         Left            =   150
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   16
         Text            =   "txtIdItem"
         Top             =   330
         Width           =   915
      End
      Begin VB.Label lblItem 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblItem"
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   150
         TabIndex        =   17
         Top             =   660
         Width           =   3675
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "[ Mes ]"
      ForeColor       =   &H00800000&
      Height          =   1050
      Left            =   0
      TabIndex        =   5
      Top             =   350
      Width           =   1995
      Begin VB.ListBox LbMes 
         Height          =   735
         Left            =   60
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   210
         Width           =   1860
      End
   End
   Begin VB.Frame FraProgreso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   4080
      TabIndex        =   0
      Top             =   2670
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   225
         TabIndex        =   1
         Top             =   420
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Shape Shape1 
         Height          =   885
         Index           =   0
         Left            =   60
         Top             =   90
         Width           =   5805
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
         TabIndex        =   4
         Top             =   180
         Width           =   525
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
         TabIndex        =   3
         Top             =   180
         Width           =   1035
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
         TabIndex        =   2
         Top             =   720
         Width           =   1260
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   4365
      Left            =   0
      TabIndex        =   8
      Top             =   4740
      Width           =   11655
      _cx             =   20558
      _cy             =   7699
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
      Caption         =   "    &MP     |  &M. Obra  | &G. Fabrica "
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
      Begin VB.Frame Frame5 
         Caption         =   "[ Costo de Gastos de Fabrica ]"
         ForeColor       =   &H00800000&
         Height          =   3990
         Left            =   45
         TabIndex        =   12
         Top             =   45
         Width           =   11565
         Begin VSFlex7Ctl.VSFlexGrid fg 
            Height          =   3645
            Index           =   4
            Left            =   60
            TabIndex        =   13
            Top             =   270
            Width           =   11415
            _cx             =   20135
            _cy             =   6429
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
            FormatString    =   $"FrmManLibroCosto.frx":327A
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
      Begin VB.Frame Frame6 
         Caption         =   "[ Costo de Mano de Obra ]"
         Height          =   3990
         Left            =   -12210
         TabIndex        =   10
         Top             =   45
         Width           =   11565
         Begin VSFlex7Ctl.VSFlexGrid fg 
            Height          =   3630
            Index           =   3
            Left            =   60
            TabIndex        =   11
            Top             =   270
            Width           =   11415
            _cx             =   20135
            _cy             =   6403
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
            FormatString    =   $"FrmManLibroCosto.frx":3347
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
         Caption         =   "[Costo Materia Prima ]"
         ForeColor       =   &H00800000&
         Height          =   3990
         Left            =   -12510
         TabIndex        =   9
         Top             =   45
         Width           =   11565
         Begin VSFlex7Ctl.VSFlexGrid fg 
            Height          =   3630
            Index           =   2
            Left            =   60
            TabIndex        =   24
            Top             =   270
            Width           =   11415
            _cx             =   20135
            _cy             =   6403
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
            FormatString    =   $"FrmManLibroCosto.frx":341A
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
      Height          =   3315
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   1410
      Width           =   11625
      _cx             =   20505
      _cy             =   5847
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
      FormatString    =   $"FrmManLibroCosto.frx":3510
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
   Begin VB.Menu menu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu00 
         Caption         =   "Aplicar solo a Productos en Venta"
      End
      Begin VB.Menu separador 
         Caption         =   "-"
      End
      Begin VB.Menu menu01 
         Caption         =   "Aplicar a Todos"
      End
   End
End
Attribute VB_Name = "FrmManLibroCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMPLANEPRODUCCION.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : MUESTRA LOS PRODUCTOS Y LAS CANTIDADES A PRODUCIR SEGUN EL MES, ADEMAS HACE LA PROGRAMACION MENSUAL
'* DISEÑADO POR     : jOSE CHACON MANRIQUE
'*****************************************************************************************************
Option Explicit

Dim Agregando As Boolean        ' INFORMA QUE SE ESTA AGREGANDO UNA FILA AL CONTROL FLEXGRID
Dim SeEjecuto As Boolean        ' VERIFICA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim QueHace As Integer          ' ESPECIFICA EN QUE MODO SE ENCUENTRA EL FORMULARIO
Dim IdMenuActivo As Integer     ' INDICA EL CODIGO DEL MENU ACTIVO
Dim xHorIni As Date             ' ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO

Dim cSQL As String
Dim BANDERA_ As Boolean
Dim RECORDSETPREUNI_ As New ADODB.Recordset
Dim RECORDSETMOBRA_ As New ADODB.Recordset
Dim RECORDSETGFABRICA_ As New ADODB.Recordset
Dim RECORDSETERRORES_ As New ADODB.Recordset

Dim RSTCABECERA As New ADODB.Recordset
Dim RSTDETALLEMATPRI As New ADODB.Recordset
Dim RSTDETALLEMANOBR As New ADODB.Recordset
Dim RSTDETALLEGASFAB As New ADODB.Recordset
Dim CORRELATIVO_ As Double

Private Enum COLUMNA_
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

Dim OrigFX As Long
Dim OrigFY As Long

Sub llenarDefinirRST(ByRef RST_ As ADODB.Recordset, Optional TIPO_ As Integer = 0, _
                                        Optional CARGAR_ As Boolean = True, Optional LIMITEFECHA_ As String, _
                                        Optional LIMITEPROCESO_ As Integer, Optional IDMES_ As Integer, _
                                        Optional REROCESO_ As Boolean = False)
                                            
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
    Select Case TIPO_
        Case 0 ' -------------------COSTO MP
            If NulosC(LIMITEFECHA_) <> "" Then
                nSQLId = nSQLId & " AND ((con_centrocostopreuni.fecha)<CDate('" & LIMITEFECHA_ & "'))"
            End If
            
            If NulosN(LIMITEPROCESO_) <> 0 Then
                nSQLId = nSQLId & " AND ((con_centrocostopreuni.proceso)<" & LIMITEPROCESO_ & ") "
            End If
            
            cSQL = "SELECT con_centrocostopreuni.iditem, con_centrocostopreuni.fecha, con_centrocostopreuni.premprima AS preuni, con_centrocostopreuni.proceso, con_centrocostopreuni.horini, con_centrocostopreuni.horfin, con_centrocostopreuni.idprod " _
                + vbCr + "FROM con_centrocostopreuni " _
                + vbCr + "WHERE (((con_centrocostopreuni.premprima)>0)) " & nSQLId
            
            Set xRs = Nothing
            RST_Busq xRs, cSQL, xCon
            
            If xRs.State = 0 Then Exit Sub
            DEFINIR_RST_TMP RST_, xRs
            
            If Not CARGAR_ Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            CARGAR_RST_TMP RST_, xRs
            
        Case 1 ' ------------------------COSTO DE MANO DE OBRA
            If NulosC(LIMITEFECHA_) <> "" Then
                nSQLId = nSQLId & " AND ((con_centrocostopreuni.fecha)<=CDate('" & LIMITEFECHA_ & "'))"
            End If
            
            If NulosN(LIMITEPROCESO_) <> 0 Then
                nSQLId = nSQLId & " AND ((con_centrocostopreuni.proceso)<=" & LIMITEPROCESO_ & ") "
            End If
            
            cSQL = "SELECT con_centrocostopreuni.iditem, con_centrocostopreuni.fecha, con_centrocostopreuni.premobra AS preuni, con_centrocostopreuni.proceso, con_centrocostopreuni.horini, con_centrocostopreuni.horfin, con_centrocostopreuni.idprod " _
                + vbCr + "FROM con_centrocostopreuni " _
                + vbCr + "WHERE (((con_centrocostopreuni.premobra)>0)) " & nSQLId
            
            Set xRs = Nothing
            RST_Busq xRs, cSQL, xCon
            
            If xRs.State = 0 Then Exit Sub
            DEFINIR_RST_TMP RST_, xRs
            
            If Not CARGAR_ Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            CARGAR_RST_TMP RST_, xRs
        
        Case 2 ' -------------------ERRORES
            ReDim xCampos(6, 3) As String
            
            xCampos(0, 0) = "numdoc":           xCampos(0, 1) = "C":      xCampos(0, 2) = "20"
            xCampos(1, 0) = "item":             xCampos(1, 1) = "C":      xCampos(1, 2) = "60"
            xCampos(2, 0) = "preuni":           xCampos(2, 1) = "D":      xCampos(2, 2) = ""
            xCampos(3, 0) = "detalle":          xCampos(3, 1) = "C":      xCampos(3, 2) = "40"
            xCampos(4, 0) = "fecha":            xCampos(4, 1) = "F":      xCampos(4, 2) = ""
            xCampos(5, 0) = "insumo":           xCampos(5, 1) = "C":      xCampos(5, 2) = "60"
            
            Set RST_ = xFun.CrearRstTMP(xCampos)
            RST_.Open
            
        Case 3 ' NUEVA REFORMULACION
'            Dim IDLIBRO_ As Integer
'
'            If NulosN(LIMITEPROCESO_) <> 0 Then
'                nSQLId = nSQLId & " AND ((con_librocosto.proceso)<=" & LIMITEPROCESO_ & ") "
'            End If
'
'            ' ----------------------CABECERA
'            cSQL = "SELECT * FROM con_librocosto WHERE ((con_librocosto.idmes)=" & IDMES_ & ")" & nSQLId
'            Set xRs = Nothing
'            RST_Busq xRs, cSQL, xCon
'            If xRs.State = 0 Then Exit Sub
'            'If xRs.RecordCount = 0 Then Exit Sub
'            'IDLIBRO_ = NulosN(xRs("id"))
'            If RSTCABECERA.State = 0 Then DEFINIR_RST_TMP RSTCABECERA, xRs
'            If Not REROCESO_ Then CARGAR_RST_TMP RSTCABECERA, xRs
'
'            nSQLId = GENERAR_SQL_ID_RST(xRs, "id", "idlibro")
'            If nSQLId = "" Then nSQLId = "idlibro=0"
'            ' ---------------------DETALLE MATERIA PRIMA
'            cSQL = "SELECT * FROM con_librocostomatpri WHERE " & nSQLId
'            Set xRs = Nothing
'            RST_Busq xRs, cSQL, xCon
'            If RSTDETALLEMATPRI.State = 0 Then DEFINIR_RST_TMP RSTDETALLEMATPRI, xRs
'            If Not REROCESO_ Then CARGAR_RST_TMP RSTDETALLEMATPRI, xRs
'
'            ' --------------------DETALLE MANO DE OBRA
'            cSQL = "SELECT * FROM con_librocostomanobr WHERE " & nSQLId
'            Set xRs = Nothing
'            RST_Busq xRs, cSQL, xCon
'            If RSTDETALLEMANOBR.State = 0 Then DEFINIR_RST_TMP RSTDETALLEMANOBR, xRs
'            If Not REROCESO_ Then CARGAR_RST_TMP RSTDETALLEMANOBR, xRs
'
'            ' --------------------DETALLE GASTOS DE FABRICA
'            cSQL = "SELECT * FROM con_librocostogasfab WHERE " & nSQLId
'            Set xRs = Nothing
'            RST_Busq xRs, cSQL, xCon
'            If RSTDETALLEGASFAB.State = 0 Then DEFINIR_RST_TMP RSTDETALLEGASFAB, xRs
'            If Not REROCESO_ Then CARGAR_RST_TMP RSTDETALLEGASFAB, xRs
        
        Case 4 ' GASTOS DE FABRICA
            If NulosC(LIMITEFECHA_) <> "" Then
                nSQLId = nSQLId & " AND ((con_centrocostopreuni.fecha)<=CDate('" & LIMITEFECHA_ & "'))"
            End If
            
            If NulosN(LIMITEPROCESO_) <> 0 Then
                nSQLId = nSQLId & " AND ((con_centrocostopreuni.proceso)<=" & LIMITEPROCESO_ & ") "
            End If
            
            cSQL = "SELECT con_centrocostopreuni.iditem, con_centrocostopreuni.fecha, con_centrocostopreuni.pregfabrica AS preuni, con_centrocostopreuni.proceso, con_centrocostopreuni.horini, con_centrocostopreuni.horfin, con_centrocostopreuni.idprod " _
                + vbCr + "FROM con_centrocostopreuni " _
                + vbCr + "WHERE (((con_centrocostopreuni.pregfabrica)>0)) " & nSQLId
            
            Set xRs = Nothing
            RST_Busq xRs, cSQL, xCon
            
            If xRs.State = 0 Then Exit Sub
            DEFINIR_RST_TMP RST_, xRs
            
            If Not CARGAR_ Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            CARGAR_RST_TMP RST_, xRs
        
    End Select
End Sub

Private Sub ckOpcion_Click(Index As Integer)
    Select Case Index
        Case 0 ' PROCESO
            If ckOpcion(Index).Value = 1 Then
                ckOpcion(1).Value = 0
                frmOpcion(Index).Enabled = True
            Else
                frmOpcion(Index).Enabled = False
            End If
            
        Case 1 ' PRODUCTO
            If ckOpcion(Index).Value = 1 Then
                ckOpcion(0).Value = 0
                frmOpcion(Index).Enabled = True
            Else
                frmOpcion(Index).Enabled = False
            End If
    End Select
End Sub

Private Sub ckoptCon_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case Index
        Case 0
            ckoptCon(1).Value = 0
        Case 1
            ckoptCon(0).Value = 0
    End Select
End Sub

Private Sub Cmd_Click(Index As Integer)
    Dim nTitulo As String
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    Dim MES_ As Integer
    
    Select Case Index
        Case 0 ' AGREGAR MATERIA PRIMA
            ReDim xCampos(2, 4) As String
            'descripcion                     'campo                       'tamaño                         'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "desitem":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":            xCampos(1, 1) = "iditem":     xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
            
            ' SE HALLA EL MES SELECCIONADO
            For A = 0 To LbMes.ListCount - 1
                LbMes.ListIndex = A
                MES_ = A + 1
                If LbMes.Selected(A) = False Then GoTo SIGUIENTE
                A = LbMes.ListCount - 1
SIGUIENTE:
            Next A
            
            cSQL = "SELECT pro_producciondet.iditem, alm_inventario.descripcion AS desitem " _
                + vbCr + "FROM pro_produccion INNER JOIN (pro_producciondet LEFT JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id) ON pro_produccion.id = pro_producciondet.idpro " _
                + vbCr + "Where (((Month([pro_produccion].[dia])) = " & MES_ & ") And ((pro_producciondet.estado) = 2)) " _
                + vbCr + "GROUP BY pro_producciondet.iditem, alm_inventario.descripcion;"
                
            nTitulo = "Buscando Productos"
            
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "desitem", "desitem", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            ' SE CARGA EL ITEM
            txtIdItem.Text = NulosN(xRs("iditem"))
            lblItem.Caption = NulosC(xRs("desitem"))
        
        Case 1
            PopupMenu menu
        
    End Select
End Sub

Private Sub llenarDetallePersonal()
    Dim RECORDSET_ As New ADODB.Recordset
    Dim TOTALPRODUCCION_ As Double
    Dim TOTALPLANILLA_ As Double
    Dim FECHA_ As String
    
    FECHA_ = NulosC(fg(0).TextMatrix(fg(0).Row, COLUMNAFECHA_))
    ' ---------------COSTO DE LA PRODUCCION
    If NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACOSTOMOBRA_)) = 0 Then Exit Sub
    TOTALPRODUCCION_ = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNACOSTOMOBRA_))
    
    ' ---------------COSTO TOTAL DE LA PLANILLA
    ' SE FILTRAN AREAS RELACIONADAS CON LA PRODUCCION
    cSQL = "SELECT Sum(pro_pagos.imptot) AS montotot " _
        + vbCr + "FROM pro_pagos " _
        + vbCr + "WHERE (((pro_pagos.fchtra)=CDate('" & FECHA_ & "')) AND ((pro_pagos.idarea) In (3,4,8,23)));"
    
    Set RECORDSET_ = Nothing
    RST_Busq RECORDSET_, cSQL, xCon
    
    If RECORDSET_.State = 0 Then Exit Sub
    If RECORDSET_.RecordCount = 0 Then Exit Sub
    TOTALPLANILLA_ = NulosN(RECORDSET_("montotot"))
    
    ' -------------PERSONAS INVOLUCRDAS EN LA RODUCCION
    cSQL = "SELECT Sum(pro_pagos.imptot) AS montotot, pro_pagos.idemp, pla_empleados.nombre, pro_pagos.idarea " _
        + vbCr + "FROM pro_pagos INNER JOIN pla_empleados ON pro_pagos.idemp = pla_empleados.id " _
        + vbCr + "WHERE (((pro_pagos.fchtra)=CDate('" & FECHA_ & "')) AND ((pro_pagos.idarea) In (3,4,8,23))) " _
        + vbCr + "GROUP BY pro_pagos.idemp, pla_empleados.nombre, pro_pagos.idarea;"
    
    Set RECORDSET_ = Nothing
    RST_Busq RECORDSET_, cSQL, xCon
    
    If RECORDSET_.State = 0 Then Exit Sub
    If RECORDSET_.RecordCount = 0 Then Exit Sub
    
    fg(3).Rows = fg(3).FixedRows
    RECORDSET_.MoveFirst
    While Not RECORDSET_.EOF
        fg(3).Rows = fg(3).Rows + 1
        fg(3).TextMatrix(fg(3).Rows - 1, 1) = Busca_Codigo(NulosN(RECORDSET_("idemp")), "id", "numdoc", "pla_empleados", "N", xCon)
        fg(3).TextMatrix(fg(3).Rows - 1, 2) = NulosC(RECORDSET_("nombre"))
        fg(3).TextMatrix(fg(3).Rows - 1, 3) = Busca_Codigo(NulosN(RECORDSET_("idarea")), "id", "descripcion", "mae_area", "N", xCon)
        fg(3).TextMatrix(fg(3).Rows - 1, 4) = NulosN(RECORDSET_("montotot")) / (TOTALPLANILLA_ / TOTALPRODUCCION_)
        fg(3).TextMatrix(fg(3).Rows - 1, 4) = Format(fg(3).TextMatrix(fg(3).Rows - 1, 4), "0.0000")
        fg(3).TextMatrix(fg(3).Rows - 1, 5) = NulosN(RECORDSET_("idemp"))
        fg(3).TextMatrix(fg(3).Rows - 1, 6) = NulosN(RECORDSET_("idarea"))
        
        RECORDSET_.MoveNext
    Wend
    
    fg(3).Rows = fg(3).Rows + 1
    fg(3).TextMatrix(fg(3).Rows - 1, 3) = "TOTAL"
    fg(3).TextMatrix(fg(3).Rows - 1, 4) = Format(GRID_SUMAR_COL(fg(3), 4), "0.0000")
End Sub

Private Sub llenarDerivados(IDMATPRIMA_ As Integer, MES_ As Integer)
    Dim xRs As New ADODB.Recordset
    Dim RECORDSETDERIVADOS_ As New ADODB.Recordset
    Dim CANTIDADREGISTROS_ As Integer
    Dim NUMEROPROCESO_ As Integer
    Dim nSQLId As String

    fg(1).Rows = fg(1).FixedRows
    NUMEROPROCESO_ = 0
    ' PRIMER PROCESO
    cSQL = "SELECT pro_producciondet.iditem, alm_inventario.descripcion AS item, pro_producciondetins.iditem AS idins " _
        + vbCr + "FROM ((pro_produccion INNER JOIN (((pro_producciondet INNER JOIN pro_emp ON pro_producciondet.idres = pro_emp.id) INNER JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id) INNER JOIN mae_unidades ON pro_producciondet.idunimed = mae_unidades.id) ON pro_produccion.id = pro_producciondet.idpro) INNER JOIN pro_producciondetins ON (pro_producciondet.idrec = pro_producciondetins.idrec) AND (pro_producciondet.numparte = pro_producciondetins.numparte) AND (pro_producciondet.idpro = pro_producciondetins.idpro)) INNER JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id " _
        + vbCr + "WHERE (((pro_producciondetins.iditem)=" & IDMATPRIMA_ & ") AND ((Month([pro_produccion].[dia])) = " & MES_ & ") And ((pro_producciondet.estado) = 2)) " & nSQLId _
        + vbCr + "GROUP BY pro_producciondet.iditem, alm_inventario.descripcion, pro_producciondetins.iditem " _
        + vbCr + "ORDER BY alm_inventario.descripcion;"
        
    Set RECORDSETDERIVADOS_ = Nothing
    RST_Busq RECORDSETDERIVADOS_, cSQL, xCon
    
    If RECORDSETDERIVADOS_.State = 0 Then Exit Sub
    If RECORDSETDERIVADOS_.RecordCount = 0 Then Exit Sub
    CANTIDADREGISTROS_ = RECORDSETDERIVADOS_.RecordCount
    
    ' PROCESOS SIGUIENTES
    While CANTIDADREGISTROS_ > 0
        NUMEROPROCESO_ = NUMEROPROCESO_ + 1
        nSQLId = GENERAR_SQL_ID_RST(RECORDSETDERIVADOS_, "iditem", "AND pro_producciondetins.iditem")
        
        RECORDSETDERIVADOS_.MoveFirst
        Agregando = True
        fg(1).Rows = fg(1).Rows + 1
        fg(1).TextMatrix(fg(1).Rows - 1, 2) = "PROCESO: " & NUMEROPROCESO_
        fg(1).Select fg(1).Rows - 1, 2
        fg(1).CellFontBold = True
        
        While Not RECORDSETDERIVADOS_.EOF
            fg(1).Rows = fg(1).Rows + 1
            fg(1).TextMatrix(fg(1).Rows - 1, 1) = NulosN(RECORDSETDERIVADOS_("iditem"))
            fg(1).TextMatrix(fg(1).Rows - 1, 2) = NulosC(RECORDSETDERIVADOS_("item"))
            fg(1).TextMatrix(fg(1).Rows - 1, 3) = NulosC(Busca_Codigo(NulosN(RECORDSETDERIVADOS_("idins")), "id", "descripcion", "alm_inventario", "N", xCon))
            RECORDSETDERIVADOS_.MoveNext
        Wend
        Agregando = False
        
        cSQL = "SELECT pro_producciondet.iditem, alm_inventario.descripcion AS item, pro_producciondetins.iditem AS idins " _
            + vbCr + "FROM ((pro_produccion INNER JOIN (((pro_producciondet INNER JOIN pro_emp ON pro_producciondet.idres = pro_emp.id) INNER JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id) INNER JOIN mae_unidades ON pro_producciondet.idunimed = mae_unidades.id) ON pro_produccion.id = pro_producciondet.idpro) INNER JOIN pro_producciondetins ON (pro_producciondet.idrec = pro_producciondetins.idrec) AND (pro_producciondet.numparte = pro_producciondetins.numparte) AND (pro_producciondet.idpro = pro_producciondetins.idpro)) INNER JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id " _
            + vbCr + "WHERE (((Month(pro_produccion.dia)) = " & MES_ & ") And ((pro_producciondet.estado) = 2)) " & nSQLId _
            + vbCr + "GROUP BY pro_producciondet.iditem, alm_inventario.descripcion, pro_producciondetins.iditem " _
            + vbCr + "ORDER BY alm_inventario.descripcion;"
        
        Set RECORDSETDERIVADOS_ = Nothing
        RST_Busq RECORDSETDERIVADOS_, cSQL, xCon
        
        CANTIDADREGISTROS_ = RECORDSETDERIVADOS_.RecordCount
    Wend
    
End Sub

Private Function GENERAR_SQL_ID_RST(Rst As ADODB.Recordset, nDesc As String, _
                            nCampo As String, Optional nTipoIn As String = "IN", _
                            Optional fEsNumero As Boolean = True) As String
    Dim nSQL As String
    Dim k&
    nSQL = ""
    
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

Private Sub fg_DblClick(Index As Integer)
    If Index <> 0 Then Exit Sub
    If Agregando Then Exit Sub
    If fg(0).Row = fg(0).Rows - 1 Then Exit Sub
    
    Me.MousePointer = vbHourglass
    llenarDetalleInsumos
    llenarDetallePersonal
    Me.MousePointer = vbDefault
End Sub

Private Sub llenarDetalleInsumos(Optional RESUMEN_ As Boolean = True)
    Dim IDDOCUMENTO_ As Integer
    Dim IDITEM_ As Integer
    Dim FECHA_ As String
    Dim RECORDSET_ As New ADODB.Recordset
    
    If Agregando Then Exit Sub
    fg(2).Rows = fg(2).FixedRows
    
    IDDOCUMENTO_ = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNAIDPROD_))
    IDITEM_ = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNAIDITEM_))
    FECHA_ = NulosC(fg(0).TextMatrix(fg(0).Row, COLUMNAFECHA_))
    
    cSQL = "SELECT pro_producciondetins.iditem AS idins, pro_producciondetins.canutil AS cantidad " _
        + vbCr + "FROM (pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro) INNER JOIN pro_producciondetins ON (pro_producciondet.idrec = pro_producciondetins.idrec) AND (pro_producciondet.numparte = pro_producciondetins.numparte) AND (pro_producciondet.idpro = pro_producciondetins.idpro) " _
        + vbCr + "WHERE (((pro_producciondetins.canutil)>0) AND ((pro_produccion.id)=" & IDDOCUMENTO_ & ") AND ((pro_producciondet.iditem)=" & IDITEM_ & "));"
    
    Set RECORDSET_ = Nothing
    RST_Busq RECORDSET_, cSQL, xCon
    If RECORDSET_.State = 0 Then Me.MousePointer = vbDefault: Exit Sub
    If RECORDSET_.RecordCount = 0 Then Me.MousePointer = vbDefault: Exit Sub
    
    RECORDSETPREUNI_.Filter = adFilterNone
    With fg(2)
        RECORDSET_.MoveFirst
        While Not RECORDSET_.EOF
            RECORDSETPREUNI_.Filter = "iditem=" & NulosN(RECORDSET_("idins")) & " AND fecha=" & FECHA_
            If RECORDSETPREUNI_.RecordCount = 0 Then GoTo SIGUIENTE_
            fg(2).Rows = fg(2).Rows + 1
            .TextMatrix(.Rows - 1, 6) = NulosN(RECORDSET_("idins"))
            .TextMatrix(.Rows - 1, 2) = UCase(Busca_Codigo(NulosN(RECORDSET_("idins")), "id", "descripcion", "alm_inventario", "N", xCon))
            .TextMatrix(.Rows - 1, 7) = Busca_Codigo(NulosN(RECORDSET_("idins")), "id", "tippro", "alm_inventario", "N", xCon)
            .TextMatrix(.Rows - 1, 1) = UCase(Busca_Codigo(NulosN(.TextMatrix(.Rows - 1, 7)), "id", "descripcion", "mae_tipoproducto", "N", xCon))
            .TextMatrix(.Rows - 1, 3) = Format(NulosN(RECORDSET_("cantidad")), "0.0000")
            .TextMatrix(.Rows - 1, 4) = Format(NulosN(RECORDSETPREUNI_("preuni")), "0.0000")
            .TextMatrix(.Rows - 1, 5) = Format(NulosN(RECORDSET_("cantidad")) * NulosN(RECORDSETPREUNI_("preuni")), "0.0000")
SIGUIENTE_:
            RECORDSET_.MoveNext
        Wend
        .Rows = .Rows + 1
        .Select .Rows - 1, 2
        .CellFontBold = True
        .TextMatrix(.Rows - 1, 2) = "TOTAL"
        .TextMatrix(.Rows - 1, 5) = GRID_SUMAR_COL(fg(2), 5)
        
    End With
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE LE FORMULARIO
    If SeEjecuto = False Then
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
    
        'OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
        SeEjecuto = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        '--interrumpir
        BANDERA_ = True
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    SeEjecuto = False
    iniciarCampos
End Sub

Private Sub Form_Resize()
    ' Si esta minimizado no se hace nada
    If Me.WindowState = 1 Then Exit Sub
    
    If Me.Width <= 12000 Then Me.Width = 12000
    If Me.Height <= 8100 Then Me.Height = 8100
        
    ' Se dimensiona el contenido
    fg(0).Width = Me.Width - 105
    fg(0).Height = Me.Height - 6195
    
    TabOne1.Top = Me.Height - 4755
    TabOne1.Width = Me.Width - 75
    fg(2).Width = TabOne1.Width - 240
    fg(3).Width = TabOne1.Width - 240
    fg(4).Width = TabOne1.Width - 240
    Frame8.Width = TabOne1.Width - 90
End Sub

Private Sub iniciarCampos()
    Dim A As Integer
    
    opttipop(0).Value = True
    ckOpcion(0).Value = 1
    optProceso(0).Value = True
    
    fg(0).FrozenCols = COLUMNAITEM_
    
    fg(2).AllowUserResizing = flexResizeColumns
    fg(2).AutoSearch = flexSearchFromTop
    fg(2).ExplorerBar = flexExSortShow
    fg(2).ForeColorSel = &H80000005
    fg(2).BackColorSel = &H80&
    fg(2).Editable = flexEDKbdMouse
    fg(2).Rows = fg(2).FixedRows
    fg(2).ColWidth(6) = 0
    fg(2).ColWidth(7) = 0
    
    fg(3).AllowUserResizing = flexResizeColumns
    fg(3).AutoSearch = flexSearchFromTop
    fg(3).ExplorerBar = flexExSortShow
    fg(3).ForeColorSel = &H80000005
    fg(3).BackColorSel = &H80&
    fg(3).Editable = flexEDKbdMouse
    fg(3).ColWidth(1) = 0
    fg(3).Rows = fg(3).FixedRows
    fg(3).ColWidth(5) = 0
    fg(3).ColWidth(6) = 0
    
    fg(4).AllowUserResizing = flexResizeColumns
    fg(4).AutoSearch = flexSearchFromTop
    fg(4).ExplorerBar = flexExSortShow
    fg(4).ForeColorSel = &H80000005
    fg(4).BackColorSel = &H80&
    fg(4).Editable = flexEDKbdMouse
    fg(4).ColWidth(1) = 0
    fg(4).Rows = fg(4).FixedRows
    
    fg(0).AllowUserResizing = flexResizeColumns
    fg(0).AutoSearch = flexSearchFromTop
    fg(0).ExplorerBar = flexExSortShow
    fg(0).SelectionMode = flexSelectionByRow
    fg(0).ForeColorSel = &H80000005
    fg(0).BackColorSel = &H80&
    fg(0).ColWidth(COLUMNAIDITEM_) = 0
    fg(0).ColWidth(COLUMNAIDPROD_) = 0
    fg(0).ColWidth(COLUMNACORRELATIVO_) = 0
    fg(0).Rows = 1
    
    lblGasFab.Caption = ""
    optGacFab(1).Value = True
    
    ' ---SE CARGAN MESES
    Llenar_Mes LbMes
    
    ' ---SE CARGA MES ACTUAL
    LbMes.Selected(Month(Date) - 1) = True
    pHallarGastoFabrica Month(Date) - 1
    '----SE CARGAN PROCESOS
    For A = 1 To 5
        ComboSemanas.AddItem A
    Next A
    txtIdItem.Text = ""
    lblItem.Caption = ""
    BANDERA_ = False
    CORRELATIVO_ = -9999
End Sub

Private Sub pintarGrid(GRID_ As VSFlexGrid, COLUMNA_ As Integer, COLOR1_ As String, COLOR2_ As String)
    Dim A As Integer
    
    With GRID_
        For A = GRID_.FixedRows To .Rows - 1
            .Select A, COLUMNA_
            If NulosN(.TextMatrix(A, COLUMNA_)) >= 0 Then
                .CellForeColor = COLOR1_
            Else
                .CellForeColor = COLOR2_
            End If
        Next A
    End With
End Sub

Private Sub ConfigurarGrid()
End Sub

Private Sub aplicarFiltrado()
    Dim A As Integer
    Dim INDICE_ As Integer
    Dim INDICETOPE_ As Integer
    Dim MES_ As Integer
        
    ' Se encuentran las caracteristicas del indice seleccionado
    INDICE_ = LbMes.ListIndex
    INDICETOPE_ = LbMes.TopIndex
        
    For A = 0 To LbMes.ListCount - 1
        LbMes.ListIndex = A
        MES_ = A + 1
        If LbMes.Selected(A) = False Then GoTo SIGUIENTE
        LblProg.Caption = "Cargando Productos"
        If opttipop(0).Value = True Then
            pLlenarDatos MES_
        Else
            pProcesarDatos MES_
        End If
        
        A = LbMes.ListCount - 1
SIGUIENTE:
    Next A
    
    LbMes.TopIndex = INDICETOPE_
    LbMes.ListIndex = INDICE_
End Sub

Private Function pHallarGastoFabrica(MESATRABAJAR_ As Integer) As Double
    Dim ULTIMODIAMES_ As Date
    Dim PRIMERDIAMES_ As Date
    Dim ANIOACTUAL_ As Double
    Dim MESACTUAL_ As Double
    Dim IMPGASFAB_ As Double
    Dim xRs As New ADODB.Recordset

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

    cSQL = "SELECT con_planctas.cuenta, con_planctas.descripcion, IIf(MovPeriodo.DebSol Is Null,0,MovPeriodo.DebSol) AS MPDebSol, con_planctas.id AS IdCta " _
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
        + vbCr + "WHERE ((Left([con_planctas].[cuenta],2)='93') AND ((con_planctas.id) In (SELECT con_diario.idcue FROM con_diario WHERE  (con_diario.ajuste in (0, 1) )  AND (  (((con_diario.fchasi) Between CDate('01/01/2012') And CDate('31/01/2012')))  OR  (con_diario.fchasi)<CDate('01/01/2012')  OR  (con_diario.fchasi) is null  )   ))) " _
        + vbCr + "ORDER BY con_planctas.cuenta;"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then pHallarGastoFabrica = 0: Exit Function
    If xRs.RecordCount = 0 Then pHallarGastoFabrica = 0: Exit Function
    xRs.MoveFirst
    IMPGASFAB_ = 0
    While Not xRs.EOF
        IMPGASFAB_ = IMPGASFAB_ + NulosN(xRs("MPDebSol"))
        xRs.MoveNext
    Wend
    
    pHallarGastoFabrica = IMPGASFAB_
End Function

Private Sub pLlenarDatos(MESATRABAJAR_ As Integer)
    Dim xRs As New ADODB.Recordset
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
        
    ' MATERIAS PRIMAS
    cSQL = "SELECT pro_producciondetins.iditem, alm_inventario.descripcion AS desitem " _
        + vbCr + "FROM ((pro_produccion INNER JOIN (pro_producciondet INNER JOIN mae_unidades ON pro_producciondet.idunimed = mae_unidades.id) ON pro_produccion.id = pro_producciondet.idpro) INNER JOIN pro_producciondetins ON (pro_producciondet.idrec = pro_producciondetins.idrec) AND (pro_producciondet.numparte = pro_producciondetins.numparte) AND (pro_producciondet.idpro = pro_producciondetins.idpro)) INNER JOIN alm_inventario ON pro_producciondetins.iditem = alm_inventario.id " _
        + vbCr + "WHERE (((alm_inventario.tippro) = 1) And ((pro_producciondet.estado) = 2) And ((Month([pro_produccion].[dia])) = " & MESATRABAJAR_ & ")) " _
        + vbCr + "GROUP BY pro_producciondetins.iditem, alm_inventario.descripcion;"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
      
    If xRs.State = 0 Then GoTo SALIR_
    If xRs.RecordCount = 0 Then GoTo SALIR_
    
    ' INICIALIZAMOS PROCESO Y NUMERO DE REGISTROS
    PROCESO_ = 0
    NUMEROREGISTROS_ = 1
        
    ' SE DEFINE COSTOS DE INSUMOS
    llenarDefinirRST RECORDSETPREUNI_, , , NulosC(ULTIMODIAMES_ + 1)
    llenarDefinirRST RECORDSETMOBRA_, 1, , NulosC(ULTIMODIAMES_ + 1)
    llenarDefinirRST RECORDSETERRORES_, 2, , NulosC(ULTIMODIAMES_ + 1)
    llenarDefinirRST RECORDSETGFABRICA_, 4, , NulosC(ULTIMODIAMES_ + 1)
        
    fg(2).Rows = fg(2).FixedRows
    fg(0).Rows = fg(0).FixedRows
    While NUMEROREGISTROS_ > 0
        PROCESO_ = PROCESO_ + 1
        
        nSQLId = GENERAR_SQL_ID_RST(xRs, "iditem", " AND pro_recetains.iditem")
        nSQLIdNot = GENERAR_SQL_ID_RST(xRs, "iditem", " AND pro_producciondet.iditem", "NOT IN")
        
        ' HALLAMOS PRODUCTOS DEL PROCESO
        cSQL = "SELECT pro_receta.iditem " _
            + vbCr + "FROM pro_receta INNER JOIN pro_recetains ON pro_receta.id = pro_recetains.idrec " _
            + vbCr + "WHERE ((pro_recetains.canpro)<>0) " & nSQLId _
            + vbCr + "GROUP BY pro_receta.iditem;"
        
        Set xRs = Nothing
        RST_Busq xRs, cSQL, xCon
        
        If xRs.State = 0 Then GoTo SALIR_
        If xRs.RecordCount = 0 Then GoTo SALIR_
        nSQLId = GENERAR_SQL_ID_RST(xRs, "iditem", " AND pro_producciondet.iditem")
    
        ' BUSCAMOS PRODUCCION DEL PROCESO
        cSQL = "SELECT pro_produccion.id, pro_produccion.dia AS fchdoc, pro_producciondet.numparte, pro_producciondet.iditem, alm_inventario.descripcion AS item, pro_receta.codrec, pro_producciondet.idres AS idresp, pla_empleados.nombre AS desresp, pro_producciondet.cantidad, mae_unidades.abrev, pro_producciondet.horini, pro_producciondet.horfin, IIf([cPREVEN].[preven]<>0,'V','P') AS tipo, cPREVEN.preven " _
            + vbCr + "FROM (pro_produccion INNER JOIN (((((pro_producciondet INNER JOIN pro_emp ON pro_producciondet.idres = pro_emp.id) INNER JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id) INNER JOIN mae_unidades ON pro_producciondet.idunimed = mae_unidades.id) INNER JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id) ON pro_produccion.id = pro_producciondet.idpro) LEFT JOIN " _
            + vbCr + "( " _
            + vbCr + "SELECT vta_ventasdet.iditem, Avg(vta_ventasdet.preuni) AS preven " _
            + vbCr + "FROM vta_ventas INNER JOIN vta_ventasdet ON vta_ventas.id = vta_ventasdet.idvta " _
            + vbCr + "WHERE (((vta_ventas.fchdoc)>=CDate('" & PRIMERDIAMES_ & "') And (vta_ventas.fchdoc)<=CDate('" & ULTIMODIAMES_ & "'))) " _
            + vbCr + "GROUP BY vta_ventasdet.iditem " _
            + vbCr + ") AS cPREVEN ON pro_producciondet.iditem = cPREVEN.iditem " _
            + vbCr + "WHERE (((pro_producciondet.cantidad)>0) AND ((Month([pro_produccion].[dia]))=" & MESATRABAJAR_ & ") AND ((pro_producciondet.estado)=2)) " & nSQLId & nSQLIdNot _
            + vbCr + "ORDER BY pro_produccion.dia, pro_producciondet.iditem;"
        
        Set xRs = Nothing
        RST_Busq xRs, cSQL, xCon
        
        If xRs.State = 0 Then GoTo SALIR_
        If xRs.RecordCount = 0 Then GoTo SALIR_
        
        ' HALLAMOS NUMERO DE REGISTROS
        NUMEROREGISTROS_ = xRs.RecordCount
        
        If ckOpcion(0).Value = 1 Then ' PROCESO
            If optProceso(1).Value = True Then
                If NulosN(ComboSemanas.Text) <> PROCESO_ Then
                    GoTo SIGUIENTEPROCESO_
                End If
            End If
        End If
        
        If xRs.State = 0 Then Exit Sub
        
        VALOR_ = 0
        TOTALHORAS_ = 0
        
        If ckoptCon(0).Value = 0 And ckoptCon(1).Value = 1 Then
            xRs.Filter = "tipo='V'"
        End If
        
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
                
                If ckOpcion(1).Value = 1 Then ' PRODUCTO
                    If IDITEM_ <> NulosN(txtIdItem.Text) Then
                        GoTo SIGUIENTEITEM_
                    End If
                End If
                
                .Rows = .Rows + 1
                .TopRow = .Rows - 1
                FraProgreso.Refresh
                LblProg.Caption = NulosC(xRs("item"))
                PgBar.Value = PgBar.Value + 1
                
                IDPROD_ = NulosN(xRs("id"))
                FECHA_ = NulosC(xRs("fchdoc"))
                .TextMatrix(.Rows - 1, COLUMNAFECHA_) = Format(NulosC(xRs("fchdoc")), FORMAT_DATE)
                .TextMatrix(.Rows - 1, COLUMNAREGPROD_) = NulosC(xRs("numparte"))
                .TextMatrix(.Rows - 1, COLUMNATIPO_) = NulosC(xRs("tipo"))
                .TextMatrix(.Rows - 1, COLUMNAPROCESO_) = PROCESO_
                .TextMatrix(.Rows - 1, COLUMNAITEM_) = NulosC(xRs("item"))
                .TextMatrix(.Rows - 1, COLUMNARECETA_) = NulosC(xRs("codrec"))
                .TextMatrix(.Rows - 1, COLUMNARESPONSABLE_) = NulosC(xRs("desresp"))
                .TextMatrix(.Rows - 1, COLUMNAUNIMED_) = NulosC(xRs("abrev"))
                .TextMatrix(.Rows - 1, COLUMNACANTIDAD_) = Format(NulosN(xRs("cantidad")), "0.0000")
                .TextMatrix(.Rows - 1, COLUMNAHORINI_) = Format(NulosC(xRs("horini")), FORMAT_HORA_SIN_SEGUNDO)
                .TextMatrix(.Rows - 1, COLUMNAHORFIN_) = Format(NulosC(xRs("horfin")), FORMAT_HORA_SIN_SEGUNDO)
                                
                ' ---------------COSTO DE MP
                RECORDSETPREUNI_.Filter = adFilterNone
                RECORDSETPREUNI_.Filter = "idprod=" & IDPROD_ & " AND iditem=" & IDITEM_
                If RECORDSETPREUNI_.RecordCount = 0 Then
                    .TextMatrix(.Rows - 1, COLUMNACOSTOMP_) = Format(0, "0.0000")
                Else
                    .TextMatrix(.Rows - 1, COLUMNACOSTOMP_) = Format(NulosN(xRs("cantidad")) * NulosN(RECORDSETPREUNI_("preuni")), "0.0000")
                End If
                ' ---------------COSTO DE MANO DE OBRA
                RECORDSETMOBRA_.Filter = "idprod=" & IDPROD_ & " AND iditem=" & IDITEM_
                If RECORDSETMOBRA_.RecordCount = 0 Then
                    .TextMatrix(.Rows - 1, COLUMNACOSTOMOBRA_) = Format(0, "0.0000")
                Else
                    .TextMatrix(.Rows - 1, COLUMNACOSTOMOBRA_) = Format(NulosN(xRs("cantidad")) * NulosN(RECORDSETMOBRA_("preuni")), "0.0000")
                End If
                
                .TextMatrix(.Rows - 1, COLUMNACOSTOPRIMO_) = Format(NulosN(.TextMatrix(.Rows - 1, COLUMNACOSTOMP_)) + NulosN(.TextMatrix(.Rows - 1, COLUMNACOSTOMOBRA_)), "0.0000")
                
                ' ---------------GASTOS DE FABRICA
                RECORDSETGFABRICA_.Filter = "idprod=" & IDPROD_ & " AND iditem=" & IDITEM_
                If RECORDSETGFABRICA_.RecordCount = 0 Then
                    .TextMatrix(.Rows - 1, COLUMNACOSTOFABRICA_) = Format(0, "0.0000")
                Else
                    .TextMatrix(.Rows - 1, COLUMNACOSTOFABRICA_) = Format(NulosN(xRs("cantidad")) * NulosN(RECORDSETGFABRICA_("preuni")), FORMAT_IMPORTEKARDEX)
                End If
                .TextMatrix(.Rows - 1, COLUMNACOSTOTOTAL_) = Format(NulosN(.TextMatrix(.Rows - 1, COLUMNACOSTOMP_)) + NulosN(.TextMatrix(.Rows - 1, COLUMNACOSTOMOBRA_)) + NulosN(.TextMatrix(.Rows - 1, COLUMNACOSTOFABRICA_)), "0.0000")
                .TextMatrix(.Rows - 1, COLUMNACOSTOUNIPRODUCCION_) = Format(NulosN(.TextMatrix(.Rows - 1, COLUMNACOSTOTOTAL_)) / NulosN(.TextMatrix(.Rows - 1, COLUMNACANTIDAD_)), "0.0000")
                
                If NulosN(xRs("preven")) <> 0 Then
                    .TextMatrix(.Rows - 1, COLUMNAPRECIOVENTA_) = Format(NulosN(xRs("preven")), "0.0000")
                    .TextMatrix(.Rows - 1, COLUMNAIMPORTEVENTA_) = Format(NulosN(xRs("cantidad")) * NulosN(xRs("preven")), "0.0000")
                    .TextMatrix(.Rows - 1, COLUMNADESVIACION_) = Format(NulosN(.TextMatrix(.Rows - 1, COLUMNAIMPORTEVENTA_)) - NulosN(.TextMatrix(.Rows - 1, COLUMNACOSTOTOTAL_)), "0.0000")
                    .TextMatrix(.Rows - 1, COLUMNADESVIACIONPORC_) = Format((NulosN(.TextMatrix(.Rows - 1, COLUMNADESVIACION_)) / NulosN(.TextMatrix(.Rows - 1, COLUMNAIMPORTEVENTA_))) * 100, FORMAT_CANTIDAD)
                End If
                
                .TextMatrix(.Rows - 1, COLUMNAIDPROD_) = IDPROD_
                .TextMatrix(.Rows - 1, COLUMNAIDITEM_) = IDITEM_
SIGUIENTEITEM_:
                xRs.MoveNext
            Wend
            xRs.Filter = adFilterNone
        End With
SIGUIENTEPROCESO_:
    Wend
    
SALIR_:
    pintarGrid fg(0), COLUMNADESVIACION_, &H0&, &HFF&
    pintarGrid fg(0), COLUMNADESVIACIONPORC_, &H0&, &HFF&
    
    fg(0).Rows = fg(0).Rows + 1
    FORMATO_CELDA fg(0), fg(0).Rows - 1, COLUMNAHORFIN_, , True, , "TOTAL"
    fg(0).TextMatrix(fg(0).Rows - 1, COLUMNACOSTOMP_) = Format(GRID_SUMAR_COL(fg(0), COLUMNACOSTOMP_), FORMAT_IMPORTEKARDEX)
    fg(0).TextMatrix(fg(0).Rows - 1, COLUMNACOSTOMOBRA_) = Format(GRID_SUMAR_COL(fg(0), COLUMNACOSTOMOBRA_), FORMAT_IMPORTEKARDEX)
    fg(0).TextMatrix(fg(0).Rows - 1, COLUMNACOSTOPRIMO_) = Format(GRID_SUMAR_COL(fg(0), COLUMNACOSTOPRIMO_), FORMAT_IMPORTEKARDEX)
    fg(0).TextMatrix(fg(0).Rows - 1, COLUMNACOSTOFABRICA_) = Format(GRID_SUMAR_COL(fg(0), COLUMNACOSTOFABRICA_), FORMAT_IMPORTEKARDEX)
    fg(0).TextMatrix(fg(0).Rows - 1, COLUMNACOSTOTOTAL_) = Format(GRID_SUMAR_COL(fg(0), COLUMNACOSTOTOTAL_), FORMAT_IMPORTEKARDEX)
    fg(0).TextMatrix(fg(0).Rows - 1, COLUMNAIMPORTEVENTA_) = Format(GRID_SUMAR_COL(fg(0), COLUMNAIMPORTEVENTA_), FORMAT_IMPORTEKARDEX)
    fg(0).TopRow = fg(0).Rows - 1
    
    FraProgreso.Visible = False
    Agregando = False
    BANDERA_ = False
End Sub

Private Function pCostoManoObraUnitario(IDITEM_ As Integer, FECHA_ As String, XCON_ As ADODB.Connection, _
                                            IDDOCUMENTO_ As Integer, CANTIDAD_ As Double)
    
    Dim RECORDSET_ As New ADODB.Recordset
    Dim PRECIOPROMEDIO_ As Double
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
    
    If RECORDSET_.State = 0 Then pCostoManoObraUnitario = 0: Exit Function
    If RECORDSET_.RecordCount = 0 Then pCostoManoObraUnitario = 0: Exit Function
    DURHORASARREGLO = Split(Format(RECORDSET_("dur"), "HH:mm"), ":")
    DURACPRODUCCION_ = NulosN(DURHORASARREGLO(0)) + (NulosN(DURHORASARREGLO(1)) / 60)
    
    ' ---------------TOTAL PLANILLA DEL DIA
    cSQL = "SELECT Sum(pro_pagos.imptot) AS montotot " _
        + vbCr + "FROM pro_pagos " _
        + vbCr + "WHERE (((pro_pagos.fchtra)=CDate('" & FECHA_ & "')) AND ((pro_pagos.idarea) In (3,4,8,23)));"
    
    Set RECORDSET_ = Nothing
    RST_Busq RECORDSET_, cSQL, XCON_
    
    If RECORDSET_.State = 0 Then pCostoManoObraUnitario = 0: Exit Function
    If RECORDSET_.RecordCount = 0 Then pCostoManoObraUnitario = 0: Exit Function
    TOTALPLANILLA_ = NulosN(RECORDSET_("montotot"))
    
    ' ---------------TOTAL HORAS DE PRODUCCION DEL DIA
    cSQL = "SELECT pro_producciondet.iditem, CDate([pro_producciondet].[horfin]-[pro_producciondet].[horini]) AS dur " _
        + vbCr + "FROM pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro " _
        + vbCr + "WHERE (((pro_produccion.dia)=CDate('" & FECHA_ & "')));"
    
    Set RECORDSET_ = Nothing
    RST_Busq RECORDSET_, cSQL, XCON_
    
    If RECORDSET_.State = 0 Then pCostoManoObraUnitario = 0: Exit Function
    If RECORDSET_.RecordCount = 0 Then pCostoManoObraUnitario = 0: Exit Function
    RECORDSET_.MoveFirst
    While Not RECORDSET_.EOF
        DURHORASARREGLO = Split(Format(RECORDSET_("dur"), "HH:mm"), ":")
        DURHORASNUMERICO_ = NulosN(DURHORASARREGLO(0)) + (NulosN(DURHORASARREGLO(1)) / 60)
        TOTALHORASPRODUCCION_ = TOTALHORASPRODUCCION_ + DURHORASNUMERICO_
        RECORDSET_.MoveNext
    Wend
        
    ' ---------------COSTO PROMEDIO POR HORA
    PRECIOPROMEDIO_ = (TOTALPLANILLA_ / TOTALHORASPRODUCCION_) * DURACPRODUCCION_ / CANTIDAD_
    
    pCostoManoObraUnitario = PRECIOPROMEDIO_
End Function

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

Private Function pCostoPrimoUnitario(IDITEM_ As Integer, FECHA_ As String, HORINI_ As Date, HORFIN_ As Date, XCON_ As ADODB.Connection, _
                                Optional TIPO_ As Integer = 1, Optional TIPODOCUMENTO_ As String, _
                                Optional IDDOCUMENTO_ As Integer, Optional CANTIDAD_ As Double) As Double
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
        
    '---------------DETALLE DE MOVIMIENTOS
    RECORDSETPREUNI_.Filter = "iditem=" & IDITEM_ & " And fecha<=" & FECHA_
         
    If IDITEM_ = 1609 Then
        MsgBox "Entro"
    End If
         
         
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
        PRECIOINICIAL_ = RECORDSETPREUNI_("preuni")
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
                If RECORDSET_.State = 0 Then pCostoPrimoUnitario = 0: Exit Function
                If RECORDSET_.RecordCount = 0 Then pCostoPrimoUnitario = 0: Exit Function
                
                RECORDSET_.MoveFirst
                IMPORTEACUMULADO_ = 0
                While Not RECORDSET_.EOF
                    RECORDSETPREUNI_.Filter = "iditem=" & NulosN(RECORDSET_("idins")) & " AND fecha=" & FECHA_
                    If RECORDSETPREUNI_.RecordCount = 0 Then
                        PRECIOUNITARIO_ = pCostoPrimoUnitario(RECORDSET_("idins"), CDate(FECHA_), HORINI_, HORFIN_, XCON_)
                        
                        ' SI ES PRODUCTO DE AGREGA LA MANO DE OBRA
                        TIPOPRODUCTO_ = Busca_Codigo(NulosN(RECORDSET_("idins")), "id", "tippro", "alm_inventario", "N", XCON_)
                        If TIPOPRODUCTO_ = 3 Then
                            RECORDSETMOBRA_.Filter = "iditem=" & NulosN(RECORDSET_("idins")) & " AND fecha=" & FECHA_
                            If RECORDSETMOBRA_.RecordCount = 0 Then
                                PRECIOMANOOBRA_ = NulosN(pCostoManoObraUnitario(NulosN(RECORDSET_("idins")), FECHA_, XCON_, IDDOCUMENTO_, CANTIDAD_))
                                ' ------------SE AGREGA EL PRECIO DE MANO DE OBRA
                                RECORDSETMOBRA_.AddNew
                                RECORDSETMOBRA_("iditem") = IDITEM_
                                RECORDSETMOBRA_("fecha") = FECHA_
                                RECORDSETMOBRA_("preuni") = PRECIOMANOOBRA_
                                RECORDSETMOBRA_.Update
                            Else
                                PRECIOMANOOBRA_ = NulosN(RECORDSETMOBRA_("preuni"))
                            End If
                        Else
                            PRECIOMANOOBRA_ = 0
                        End If
                        
                        If PRECIOUNITARIO_ < 0 Then
                            MsgBox "Ocurrio un error al procesar el costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                            RECORDSETERRORES_.AddNew
                            RECORDSETERRORES_("numdoc") = Busca_Codigo(IDDOCUMENTO_, "idpro", "numparte", "pro_producciondet", "N", XCON_)
                            RECORDSETERRORES_("item") = Busca_Codigo(IDITEM_, "id", "descripcion", "Alm_inventario", "N", XCON_)
                            RECORDSETERRORES_("insumo") = Busca_Codigo(NulosN(RECORDSET_("idins")), "id", "descripcion", "Alm_inventario", "N", XCON_)
                            RECORDSETERRORES_("preuni") = PRECIOUNITARIO_
                            RECORDSETERRORES_("detalle") = "Costo MP - Precio unitario negativo"
                            RECORDSETERRORES_("fecha") = FECHA_
                            RECORDSETERRORES_.Update
                            BANDERA_ = True
                        ElseIf PRECIOUNITARIO_ = 0 Then
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
                        ' SE AGREGA AL RECORDSET DE PRECIOS UNITARIOS
                        RECORDSETPREUNI_.AddNew
                        RECORDSETPREUNI_("iditem") = RECORDSET_("idins")
                        RECORDSETPREUNI_("fecha") = FECHA_
                        RECORDSETPREUNI_("horini") = HORINI_
                        RECORDSETPREUNI_("horfin") = HORFIN_
                        RECORDSETPREUNI_("preuni") = PRECIOUNITARIO_ + PRECIOMANOOBRA_
                        RECORDSETPREUNI_.Update
                        
                        ' ****************************************************************
                        ' ****************************************************************
'                        RSTDETALLEMATPRI.AddNew
'                        RSTDETALLEMATPRI("idlibro") = CORRELATIVO_
'                        RSTDETALLEMATPRI("iditem") = NulosN(RECORDSET_("idins"))
'                        RSTDETALLEMATPRI("cantidad") = NulosN(RECORDSET_("cantidad"))
'                        RSTDETALLEMATPRI("impmatpri") = (PRECIOUNITARIO_ + PRECIOMANOOBRA_) * NulosN(RECORDSET_("cantidad"))
'                        RSTDETALLEMATPRI.Update
                        ' ****************************************************************
                        ' ****************************************************************
                    Else
                        PRECIOUNITARIO_ = NulosN(RECORDSETPREUNI_("preuni"))
                    End If
                    IMPORTEACUMULADO_ = IMPORTEACUMULADO_ + (PRECIOUNITARIO_ * RECORDSET_("cantidad"))
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
                
                PRECIOPROMEDIO_ = IMPORTEACUMULADO_ / CANTIDAD_
                ' SE AGREGA AL RECORDSET DE PRECIOS UNITARIOS
                RECORDSETPREUNI_.AddNew
                RECORDSETPREUNI_("iditem") = IDITEM_
                RECORDSETPREUNI_("fecha") = FECHA_
                RECORDSETPREUNI_("preuni") = PRECIOPROMEDIO_
                RECORDSETPREUNI_("horini") = HORINI_
                RECORDSETPREUNI_("horfin") = HORFIN_
                RECORDSETPREUNI_("idprod") = IDDOCUMENTO_
                RECORDSETPREUNI_.Update
                
                ' ****************************************************************
                ' ****************************************************************
'                RSTCABECERA.AddNew
'                RSTCABECERA("id") = CORRELATIVO_
'                RSTCABECERA("iditem") = IDITEM_
'                RSTCABECERA("idprod") = IDDOCUMENTO_
'                RSTCABECERA("proceso") = IDITEM_
'                RSTCABECERA("impmprima") = IDITEM_
'                RSTCABECERA("idmes") = IDITEM_
'                RSTCABECERA("cantidad") = CANTIDAD_
'                RSTCABECERA.Update
                ' ****************************************************************
                ' ****************************************************************
            ' ----------------------------------------------------------SALIDAS
            Else
            End If
                
        Case 1
            If RECORDSET_.RecordCount = 0 Then pCostoPrimoUnitario = PRECIOINICIAL_: Exit Function
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
                        PRECIOUNITARIO_ = PrecioUni(RECORDSET_("id"), CDbl(IDITEM_), NulosC(RECORDSET_("tipo")))
                        
                        If PRECIOUNITARIO_ < 0 Then
                            MsgBox "Ocurrio un error al procesar el costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                            RECORDSETERRORES_.AddNew
                            RECORDSETERRORES_("numdoc") = NulosC(RECORDSET_("numdoc"))
                            RECORDSETERRORES_("item") = Busca_Codigo(IDITEM_, "id", "descripcion", "Alm_inventario", "N", XCON_)
                            RECORDSETERRORES_("preuni") = PRECIOUNITARIO_
                            RECORDSETERRORES_("detalle") = "Costo MP - Precio unitario negativo"
                            RECORDSETERRORES_("fecha") = RECORDSET_("fchdoc")
                            RECORDSETERRORES_.Update
                            BANDERA_ = True
                        ElseIf PRECIOUNITARIO_ = 0 Then
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
                        RECORDSETPREUNI_.Filter = "iditem=" & IDITEM_ & " AND fecha=" & RECORDSET_("fchdoc")
                        If RECORDSETPREUNI_.RecordCount = 0 Then
                            ' --------------TIPO DE ITEM
                            TIPOPRODUCTO_ = Busca_Codigo(IDITEM_, "id", "tippro", "alm_inventario", "N", XCON_)
                            Select Case TIPOPRODUCTO_
                                Case 3
                                    PRECIOUNITARIO_ = pCostoPrimoUnitario(IDITEM_, RECORDSET_("fchdoc"), RECORDSET_("horini"), RECORDSET_("horfin"), XCON_, 0, RECORDSET_("tipo"), RECORDSET_("id"), NulosN(RECORDSET_("canpro")))
                                Case Else
                                    PRECIOUNITARIO_ = NulosN(RECORDSET_("preuni"))
                                    ' SE AGREGA AL RECORDSET DE PRECIOS UNITARIOS
                                    RECORDSETPREUNI_.AddNew
                                    RECORDSETPREUNI_("iditem") = IDITEM_
                                    RECORDSETPREUNI_("fecha") = RECORDSET_("fchdoc")
                                    RECORDSETPREUNI_("preuni") = PRECIOUNITARIO_
                                    RECORDSETPREUNI_.Update
                            End Select
                        Else
                            PRECIOUNITARIO_ = NulosN(RECORDSETPREUNI_("preuni"))
                        End If
                    End If
                    ' --------------------------------IMPORTE ACUMULADO
                    If RECORDSET_("descdoc") = "NC" Then
                        IMPORTEACUMULADO_ = IMPORTEACUMULADO_ - (NulosN(RECORDSET_("canpro")) * PRECIOUNITARIO_)
                    Else
                        IMPORTEACUMULADO_ = IMPORTEACUMULADO_ + (NulosN(RECORDSET_("canpro")) * PRECIOUNITARIO_)
                    End If
                    ' --------------------------------PRECIO PROMEDIO
                    If CANTIDADACUMULADA_ > 0 Then
                        PRECIOPROMEDIO_ = PRECIOUNITARIO_ 'IMPORTEACUMULADO_ / CANTIDADACUMULADA_
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
                    PRECIOUNITARIO_ = IMPORTEACUMULADO_ / CANTIDADACUMULADA_
                    
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
                    
                    'PRECIOUNITARIO_ = PRECIOPROMEDIO_
                    ' --------------------------------IMPORTE ACUMULADO
                    If RECORDSET_("descdoc") = "NC" Then
                        IMPORTEACUMULADO_ = IMPORTEACUMULADO_ + (NulosN(RECORDSET_("canpro")) * PRECIOUNITARIO_)
                    Else
                        IMPORTEACUMULADO_ = IMPORTEACUMULADO_ - (NulosN(RECORDSET_("canpro")) * PRECIOUNITARIO_)
                    End If
                End If
SIGUIENTE_:
                RECORDSET_.MoveNext
            Wend
    End Select
    
    pCostoPrimoUnitario = PRECIOPROMEDIO_
End Function

Private Function PrecioUni(IdDocumento, IdItem As Double, DondeBuscar As String) As Double
    '===================================================================================================
    'Creado:     01/07/11 Johan Castro
    'Propósito:  Obtener el Precio unitario del registro de compras vinculado con documentos (de ingreso de almacen, Guia Remision)
    '
    'Entradas:   IdDocumento = Código de Libro
    '            IdItem = Código del Item (Producto, Materia prima, Insumo, etc)
    '            DondeBuscar = Indica el origen del registro
    '
    'Resultados: Precio unitario del item segun el documento ingresado
    '===================================================================================================
    
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
    Dim MESACTUAL_ As Double
    Dim nSQLId As String
    Dim nSQLIdNot As String
    Dim CONSULTA_ As String
    Dim NUMEROREGISTROS_ As Integer
    Dim PROCESO_ As Integer
    Dim IMPORTEPARCIAL_ As Double
    Dim INDICEFAB_ As Double
    Dim A As Integer
    
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
        
    ' MATERIAS PRIMAS
    cSQL = "SELECT pro_producciondetins.iditem, alm_inventario.descripcion AS desitem " _
        + vbCr + "FROM ((pro_produccion INNER JOIN (pro_producciondet INNER JOIN mae_unidades ON pro_producciondet.idunimed = mae_unidades.id) ON pro_produccion.id = pro_producciondet.idpro) INNER JOIN pro_producciondetins ON (pro_producciondet.idrec = pro_producciondetins.idrec) AND (pro_producciondet.numparte = pro_producciondetins.numparte) AND (pro_producciondet.idpro = pro_producciondetins.idpro)) INNER JOIN alm_inventario ON pro_producciondetins.iditem = alm_inventario.id " _
        + vbCr + "WHERE (((alm_inventario.tippro) = 1) And ((pro_producciondet.estado) = 2) And ((Month([pro_produccion].[dia])) = " & MESATRABAJAR_ & ")) " _
        + vbCr + "GROUP BY pro_producciondetins.iditem, alm_inventario.descripcion;"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
      
    If xRs.State = 0 Then GoTo SALIR_
    If xRs.RecordCount = 0 Then GoTo SALIR_
    
    ' INICIALIZAMOS PROCESO Y NUMERO DE REGISTROS
    PROCESO_ = 0
    NUMEROREGISTROS_ = 1
        
    ' SE DEFINE COSTOS DE INSUMOS
    If optProceso(0).Value = True Then
        llenarDefinirRST RECORDSETPREUNI_, , , NulosC(PRIMERDIAMES_)
        llenarDefinirRST RECORDSETMOBRA_, 1, , NulosC(PRIMERDIAMES_)
        llenarDefinirRST RECORDSETERRORES_, 2, , NulosC(PRIMERDIAMES_)
        
        llenarDefinirRST RECORDSETERRORES_, 3, , , , MESATRABAJAR_, True
        llenarDefinirRST RECORDSETGFABRICA_, 4, , NulosC(PRIMERDIAMES_)
    Else
        llenarDefinirRST RECORDSETPREUNI_, , , NulosC(ULTIMODIAMES_), NulosN(ComboSemanas.Text)
        llenarDefinirRST RECORDSETMOBRA_, 1, , NulosC(ULTIMODIAMES_), NulosN(ComboSemanas.Text)
        llenarDefinirRST RECORDSETERRORES_, 2, , NulosC(ULTIMODIAMES_), NulosN(ComboSemanas.Text)
        
        llenarDefinirRST RECORDSETERRORES_, 3, , , NulosN(ComboSemanas.Text), MESATRABAJAR_, False
        llenarDefinirRST RECORDSETGFABRICA_, 4, , NulosC(ULTIMODIAMES_), NulosN(ComboSemanas.Text)
    End If
        
    fg(2).Rows = fg(2).FixedRows
    fg(0).Rows = fg(0).FixedRows
    While NUMEROREGISTROS_ > 0
        PROCESO_ = PROCESO_ + 1
        
        nSQLId = GENERAR_SQL_ID_RST(xRs, "iditem", " AND pro_recetains.iditem")
        nSQLIdNot = GENERAR_SQL_ID_RST(xRs, "iditem", " AND pro_producciondet.iditem", "NOT IN")
        
        ' HALLAMOS PRODUCTOS DEL PROCESO
        cSQL = "SELECT pro_receta.iditem " _
            + vbCr + "FROM pro_receta INNER JOIN pro_recetains ON pro_receta.id = pro_recetains.idrec " _
            + vbCr + "WHERE ((pro_recetains.canpro)<>0) " & nSQLId _
            + vbCr + "GROUP BY pro_receta.iditem;"
        
        Set xRs = Nothing
        RST_Busq xRs, cSQL, xCon
        
        If xRs.State = 0 Then GoTo SALIR_
        If xRs.RecordCount = 0 Then GoTo GASTOSDEFABRICA_
        nSQLId = GENERAR_SQL_ID_RST(xRs, "iditem", " AND pro_producciondet.iditem")
        
        ' BUSCAMOS PRODUCCION DEL PROCESO
        cSQL = "SELECT pro_produccion.id, pro_produccion.dia AS fchdoc, pro_producciondet.numparte, pro_producciondet.iditem, alm_inventario.descripcion AS item, pro_receta.codrec, pro_producciondet.idres AS idresp, pla_empleados.nombre AS desresp, pro_producciondet.cantidad, mae_unidades.abrev, pro_producciondet.horini, pro_producciondet.horfin " _
            + vbCr + "FROM (pro_produccion INNER JOIN ((((pro_producciondet INNER JOIN pro_emp ON pro_producciondet.idres = pro_emp.id) INNER JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id) INNER JOIN mae_unidades ON pro_producciondet.idunimed = mae_unidades.id) INNER JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) ON pro_produccion.id = pro_producciondet.idpro) LEFT JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id " _
            + vbCr + "WHERE (((pro_producciondet.cantidad)>0) AND ((Month([pro_produccion].[dia])) = " & MESATRABAJAR_ & ") And ((pro_producciondet.estado) = 2)) " & nSQLId & nSQLIdNot _
            + vbCr + "ORDER BY pro_produccion.dia, pro_producciondet.iditem;"
        
        Set xRs = Nothing
        RST_Busq xRs, cSQL, xCon
        
        If xRs.State = 0 Then GoTo SALIR_
        If xRs.RecordCount = 0 Then GoTo GASTOSDEFABRICA_
        
        ' HALLAMOS NUMERO DE REGISTROS
        NUMEROREGISTROS_ = xRs.RecordCount
        
        If ckOpcion(0).Value = 1 Then ' PROCESO
            If optProceso(1).Value = True Then
                If NulosN(ComboSemanas.Text) <> PROCESO_ Then
                    GoTo SIGUIENTEPROCESO_
                End If
            End If
        End If
        
        If xRs.State = 0 Then Exit Sub
        
        VALOR_ = 0
        TOTALHORAS_ = 0
        With fg(0)
            If xRs.RecordCount = 0 Then Exit Sub
            
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
                If NUMEROREGISTROS_ = 0 Then GoTo GASTOSDEFABRICA_
                
                IDITEM_ = NulosN(xRs("iditem"))
                
                If ckOpcion(1).Value = 1 Then ' PRODUCTO
                    If IDITEM_ = NulosN(txtIdItem.Text) Then
                        GoTo SIGUIENTEITEM_
                    End If
                End If
                
                .Rows = .Rows + 1
                .TopRow = .Rows - 1
                FraProgreso.Refresh
                LblProg.Caption = NulosC(xRs("item"))
                PgBar.Value = PgBar.Value + 1
                
                IDPROD_ = NulosN(xRs("id"))
                FECHA_ = NulosC(xRs("fchdoc"))
                .TextMatrix(.Rows - 1, COLUMNAFECHA_) = Format(NulosC(xRs("fchdoc")), FORMAT_DATE)
                .TextMatrix(.Rows - 1, COLUMNAREGPROD_) = NulosC(xRs("numparte"))
                .TextMatrix(.Rows - 1, COLUMNAPROCESO_) = PROCESO_
                .TextMatrix(.Rows - 1, COLUMNAITEM_) = NulosC(xRs("item"))
                .TextMatrix(.Rows - 1, COLUMNARECETA_) = NulosC(xRs("codrec"))
                .TextMatrix(.Rows - 1, COLUMNARESPONSABLE_) = NulosC(xRs("desresp"))
                .TextMatrix(.Rows - 1, COLUMNAUNIMED_) = NulosC(xRs("abrev"))
                .TextMatrix(.Rows - 1, COLUMNACANTIDAD_) = Format(NulosN(xRs("cantidad")), "0.0000")
                .TextMatrix(.Rows - 1, COLUMNAHORINI_) = Format(NulosC(xRs("horini")), FORMAT_HORA_SIN_SEGUNDO)
                .TextMatrix(.Rows - 1, COLUMNAHORFIN_) = Format(NulosC(xRs("horfin")), FORMAT_HORA_SIN_SEGUNDO)
                .TextMatrix(.Rows - 1, COLUMNACOSTOMP_) = Format(NulosN(fg(0).TextMatrix(.Rows - 1, COLUMNACANTIDAD_)) * pCostoPrimoUnitario(IDITEM_, FECHA_, xRs("horini"), xRs("horfin"), xCon, 0, "P", IDPROD_, NulosC(xRs("cantidad"))), "0.0000")
                
                RECORDSETMOBRA_.Filter = "idprod=" & IDPROD_
                If RECORDSETMOBRA_.RecordCount = 0 Then
                    .TextMatrix(.Rows - 1, COLUMNACOSTOMOBRA_) = pCostoManoObraUnitario(IDITEM_, FECHA_, xCon, IDPROD_, NulosC(xRs("cantidad")))
                    ' ------------SE AGREGA EL PRECIO DE MANO DE OBRA
                    RECORDSETMOBRA_.AddNew
                    RECORDSETMOBRA_("iditem") = IDITEM_
                    RECORDSETMOBRA_("fecha") = FECHA_
                    RECORDSETMOBRA_("preuni") = NulosN(.TextMatrix(.Rows - 1, COLUMNACOSTOMOBRA_))
                    RECORDSETMOBRA_("idprod") = IDPROD_
                    RECORDSETMOBRA_.Update
                    
                    '**********************************************************
                    '**********************************************************
'                    RSTCABECERA.Filter = "id=" & CORRELATIVO_
'                    RSTCABECERA("impmanobr") = NulosN(.TextMatrix(.Rows - 1, COLUMNACOSTOMOBRA_)) * NulosN(xRs("cantidad"))
'                    RSTCABECERA.Update
                    '**********************************************************
                    '**********************************************************
                    
                    .TextMatrix(.Rows - 1, COLUMNACOSTOMOBRA_) = Format(NulosN(fg(0).TextMatrix(.Rows - 1, COLUMNACANTIDAD_)) * .TextMatrix(.Rows - 1, COLUMNACOSTOMOBRA_), FORMAT_IMPORTEKARDEX)
                Else
                    .TextMatrix(.Rows - 1, COLUMNACOSTOMOBRA_) = Format(NulosN(fg(0).TextMatrix(.Rows - 1, COLUMNACANTIDAD_)) * NulosN(RECORDSETMOBRA_("preuni")), "0.0000")
                End If

                If .TextMatrix(.Rows - 1, COLUMNACOSTOMOBRA_) < 0 Then
                    MsgBox "Ocurrio un error al procesar el costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    RECORDSETERRORES_.AddNew
                    RECORDSETERRORES_("numdoc") = NulosC(xRs("numparte"))
                    RECORDSETERRORES_("item") = NulosC(xRs("item"))
                    RECORDSETERRORES_("preuni") = .TextMatrix(.Rows - 1, COLUMNACOSTOMOBRA_)
                    RECORDSETERRORES_("detalle") = "Mano de Obra - Precio unitario negativo"
                    RECORDSETERRORES_.Update
                    BANDERA_ = True
                ElseIf .TextMatrix(.Rows - 1, COLUMNACOSTOMOBRA_) = 0 Then
                    MsgBox "Ocurrio un error al procesar el costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    RECORDSETERRORES_.AddNew
                    RECORDSETERRORES_("numdoc") = NulosC(xRs("numparte"))
                    RECORDSETERRORES_("item") = NulosC(xRs("item"))
                    RECORDSETERRORES_("preuni") = .TextMatrix(.Rows - 1, COLUMNACOSTOMOBRA_)
                    RECORDSETERRORES_("detalle") = "Mano de Obra - Precio unitario cero"
                    RECORDSETERRORES_.Update
                    BANDERA_ = True
                End If

                .TextMatrix(.Rows - 1, COLUMNACOSTOPRIMO_) = Format(NulosN(.TextMatrix(.Rows - 1, COLUMNACOSTOMP_)) + NulosN(.TextMatrix(.Rows - 1, COLUMNACOSTOMOBRA_)), "0.0000")
                .TextMatrix(.Rows - 1, COLUMNAIDPROD_) = IDPROD_
                .TextMatrix(.Rows - 1, COLUMNAIDITEM_) = IDITEM_
                .TextMatrix(.Rows - 1, COLUMNACORRELATIVO_) = CORRELATIVO_
                CORRELATIVO_ = CORRELATIVO_ + 1
SIGUIENTEITEM_:
                xRs.MoveNext
            Wend
        End With
APLICARCAMBIOS_:
        LblProg.Caption = ""
        lbl(2).Caption = "GRABANDO COSTOS"
        aplicarCambios PROCESO_, MESATRABAJAR_
        lbl(2).Caption = "Cancelar = ESC"
        
SIGUIENTEPROCESO_:
    Wend
    
GASTOSDEFABRICA_:
    
    IMPORTEPARCIAL_ = 0
    IMPORTEPARCIAL_ = GRID_SUMAR_COL(fg(0), COLUMNACOSTOPRIMO_)
    INDICEFAB_ = NulosN(lblGasFab.Caption) / IMPORTEPARCIAL_
    lbl(2).Caption = "APLICANDO GASTOS DE FABRICA"
    For A = fg(0).FixedRows To fg(0).Rows - 1
        DoEvents
        fg(0).TopRow = A
        If optGacFab(0).Value = True Then '----------- SOLO VENTAS
            If NulosC(fg(0).TextMatrix(A, COLUMNATIPO_)) = "V" Then
                ' ------------SE AGREGA EL PRECIO DE FABRICA
'                RECORDSETGFABRICA_.AddNew
'                RECORDSETGFABRICA_("iditem") = NulosN(fg(0).TextMatrix(A, COLUMNAIDITEM_))
'                RECORDSETGFABRICA_("fecha") = NulosC(fg(0).TextMatrix(A, COLUMNAFECHA_))
'                RECORDSETGFABRICA_("preuni") = (NulosN(fg(0).TextMatrix(A, COLUMNACOSTOPRIMO_)) * INDICEFAB_) / NulosN(fg(0).TextMatrix(A, COLUMNACANTIDAD_))
'                RECORDSETGFABRICA_("idprod") = NulosN(fg(0).TextMatrix(A, COLUMNAIDPROD_))
'                RECORDSETGFABRICA_.Update
                
                fg(0).TextMatrix(A, COLUMNACOSTOFABRICA_) = Format(fg(0).TextMatrix(A, COLUMNACOSTOPRIMO_) * INDICEFAB_, FORMAT_IMPORTEKARDEX)
            Else
                INDICEFAB_ = 0
                ' ------------SE AGREGA EL PRECIO DE FABRICA
'                RECORDSETGFABRICA_.AddNew
'                RECORDSETGFABRICA_("iditem") = NulosN(fg(0).TextMatrix(A, COLUMNAIDITEM_))
'                RECORDSETGFABRICA_("fecha") = NulosC(fg(0).TextMatrix(A, COLUMNAFECHA_))
'                RECORDSETGFABRICA_("preuni") = (NulosN(fg(0).TextMatrix(A, COLUMNACOSTOPRIMO_)) * INDICEFAB_) / NulosN(fg(0).TextMatrix(A, COLUMNACANTIDAD_))
'                RECORDSETGFABRICA_("idprod") = NulosN(fg(0).TextMatrix(A, COLUMNAIDPROD_))
'                RECORDSETGFABRICA_.Update
                
                fg(0).TextMatrix(fg(0).Rows - 1, COLUMNACOSTOFABRICA_) = Format(0, FORMAT_IMPORTEKARDEX)
            End If
        Else '------------------------------------------- TODOS
            ' ------------SE AGREGA EL PRECIO DE FABRICA
'            RECORDSETGFABRICA_.AddNew
'            RECORDSETGFABRICA_("iditem") = NulosN(fg(0).TextMatrix(A, COLUMNAIDITEM_))
'            RECORDSETGFABRICA_("fecha") = NulosC(fg(0).TextMatrix(A, COLUMNAFECHA_))
'            RECORDSETGFABRICA_("preuni") = (NulosN(fg(0).TextMatrix(A, COLUMNACOSTOPRIMO_)) * INDICEFAB_) / NulosN(fg(0).TextMatrix(A, COLUMNACANTIDAD_))
'            RECORDSETGFABRICA_("idprod") = NulosN(fg(0).TextMatrix(A, COLUMNAIDPROD_))
'            RECORDSETGFABRICA_.Update
            
            fg(0).TextMatrix(A, COLUMNACOSTOFABRICA_) = Format(INDICEFAB_ * NulosN(fg(0).TextMatrix(A, COLUMNACOSTOPRIMO_)), FORMAT_IMPORTEKARDEX)
        End If
                                
        '**********************************************************
        '**********************************************************
        aplicarCambios 0, MESATRABAJAR_, True, NulosN(fg(0).TextMatrix(A, COLUMNAIDPROD_)), (NulosN(fg(0).TextMatrix(A, COLUMNACOSTOPRIMO_)) * INDICEFAB_) / NulosN(fg(0).TextMatrix(A, COLUMNACANTIDAD_))
        '**********************************************************
        '**********************************************************
        
        fg(0).TextMatrix(A, COLUMNACOSTOTOTAL_) = Format(NulosN(fg(0).TextMatrix(A, COLUMNACOSTOMP_)) + NulosN(fg(0).TextMatrix(A, COLUMNACOSTOMOBRA_)) + NulosN(fg(0).TextMatrix(A, COLUMNACOSTOFABRICA_)), FORMAT_IMPORTEKARDEX)
        fg(0).TextMatrix(A, COLUMNACOSTOUNIPRODUCCION_) = Format(fg(0).TextMatrix(A, COLUMNACOSTOTOTAL_) / NulosN(fg(0).TextMatrix(A, COLUMNACANTIDAD_)), FORMAT_IMPORTEKARDEX)
    Next A

SALIR_:
    pExportar
    FraProgreso.Visible = False
    Agregando = False
    BANDERA_ = False
End Sub

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

Private Sub LbMes_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim A As Integer
    Dim ENCONTRO_ As Integer
    Dim INDICE_ As Double
    Dim INDICETOPE_ As Double
    
    If Button = 1 Then
        ' Se encuentran las caracteristicas del indice selecionado
        INDICE_ = LbMes.ListIndex
        INDICETOPE_ = LbMes.TopIndex
        
        ' Se verifica que indices estan seleccionados
        For A = 0 To LbMes.ListCount - 1
            LbMes.ListIndex = A
            If LbMes.Selected(A) = True Then
                ENCONTRO_ = ENCONTRO_ + 1
                If ENCONTRO_ = 2 Then A = LbMes.ListCount - 1
            End If
        Next
        
        ' Si hay mas de un seleccionado
        If ENCONTRO_ = 2 Then
            LbMes.Selected(INDICE_) = False
            LbMes.TopIndex = INDICETOPE_
            LbMes.ListIndex = INDICE_
            Exit Sub
        ElseIf ENCONTRO_ = 0 Then
            lblGasFab.Caption = ""
            LbMes.TopIndex = INDICETOPE_
            LbMes.ListIndex = INDICE_
            Exit Sub
        End If
        
        ' Se seleccionan los indices del inicio
        LbMes.TopIndex = INDICETOPE_
        LbMes.ListIndex = INDICE_
        
        lblGasFab.Caption = Format(NulosN(pHallarGastoFabrica(INDICE_ + 1)), FORMAT_IMPORTEKARDEX)
        
        txtIdItem.Text = ""
        lblItem.Caption = ""
        fg(0).Rows = fg(0).FixedRows
        'fg(1).Rows = fg(1).FixedRows
        fg(2).Rows = fg(2).FixedRows
        fg(3).Rows = fg(3).FixedRows
        fg(4).Rows = fg(4).FixedRows
    End If
End Sub

Private Sub optProceso_Click(Index As Integer)
    Select Case Index
        Case 0 ' TODOS
            ComboSemanas.Enabled = False
            
        Case 1 ' SELECCIONAR
            ComboSemanas.Enabled = True
            
    End Select
End Sub

Private Sub opttipop_Click(Index As Integer)
    If opttipop(0).Value = True Then
        ckoptCon(0).Enabled = True
        ckoptCon(1).Enabled = True
        ckoptCon(0).Value = 1
        ckoptCon(1).Value = 0
    ElseIf opttipop(1).Value = True Then
        ckoptCon(0).Enabled = False
        ckoptCon(1).Enabled = False
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim MES_ As Integer
    Dim A As Integer
    
    If Button.Index = 1 Then ' Buscar
        aplicarFiltrado
    End If
    
    If Button.Index = 3 Then ' EXPORTAR EXCEL
        ExportarExcel fg(0)
    End If
    
    If Button.Index = 6 Then ' Salir
        Unload Me
    End If
End Sub

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
        RECORDSETPREUNI_.Filter = "fecha>=" & PRIMERDIAMES_ & " AND fecha<=" & ULTIMODIAMES_
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

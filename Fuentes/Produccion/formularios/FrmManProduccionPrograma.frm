VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmManProduccionPrograma 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Producción - Programación de la Producción"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9810
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManProduccionPrograma.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManProduccionPrograma.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManProduccionPrograma.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManProduccionPrograma.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManProduccionPrograma.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManProduccionPrograma.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManProduccionPrograma.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManProduccionPrograma.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManProduccionPrograma.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManProduccionPrograma.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManProduccionPrograma.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManProduccionPrograma.frx":2236
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7245
      Left            =   30
      TabIndex        =   7
      Top             =   390
      Width           =   11865
      _cx             =   20929
      _cy             =   12779
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
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   6825
         Left            =   12510
         TabIndex        =   9
         Top             =   375
         Width           =   11775
         Begin VB.Frame Frame3 
            Caption         =   "( Periodo )"
            Height          =   660
            Left            =   10335
            TabIndex        =   30
            Top             =   570
            Width           =   1410
            Begin VB.Label lblperiodo 
               Alignment       =   2  'Center
               Caption         =   "lblperiodo"
               BeginProperty Font 
                  Name            =   "Courier"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   31
               Top             =   240
               Width           =   1245
            End
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   2
            Left            =   10545
            Locked          =   -1  'True
            TabIndex        =   24
            Tag             =   "null"
            Text            =   "txt(2)"
            Top             =   240
            Width           =   1170
         End
         Begin VB.Frame Frame4 
            Height          =   660
            Left            =   6270
            TabIndex        =   26
            Top             =   570
            Width           =   4065
            Begin VB.CommandButton cmd 
               Caption         =   "Agregar Producto"
               Enabled         =   0   'False
               Height          =   420
               Index           =   0
               Left            =   60
               TabIndex        =   29
               ToolTipText     =   "Agregar Documentos"
               Top             =   180
               Width           =   1275
            End
            Begin VB.CommandButton cmd 
               Caption         =   "&Eliminar"
               Enabled         =   0   'False
               Height          =   420
               Index           =   1
               Left            =   2700
               TabIndex        =   28
               ToolTipText     =   "Eliminar Documentos Seleccionados"
               Top             =   180
               Width           =   1275
            End
            Begin VB.CommandButton cmd 
               Caption         =   "Seleccionar Producto"
               Enabled         =   0   'False
               Height          =   420
               Index           =   2
               Left            =   1380
               TabIndex        =   27
               ToolTipText     =   "Agregar Documentos"
               Top             =   180
               Width           =   1275
            End
         End
         Begin VB.TextBox txt 
            BackColor       =   &H0080FF80&
            Height          =   315
            Index           =   0
            Left            =   7185
            TabIndex        =   20
            Tag             =   "null"
            Text            =   "txt(0)"
            Top             =   345
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.CommandButton cb 
            Enabled         =   0   'False
            Height          =   240
            Index           =   0
            Left            =   2355
            Picture         =   "FrmManProduccionPrograma.frx":277E
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   690
            Width           =   240
         End
         Begin VB.TextBox txt 
            Height          =   1020
            Index           =   1
            Left            =   60
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Tag             =   "null"
            Text            =   "FrmManProduccionPrograma.frx":28B0
            Top             =   5700
            Width           =   11655
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   4035
            Left            =   60
            TabIndex        =   5
            Top             =   1365
            Width           =   11655
            _cx             =   20558
            _cy             =   7117
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
            Rows            =   5
            Cols            =   9
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmManProduccionPrograma.frx":28B9
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
         Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
            Height          =   300
            Index           =   0
            Left            =   1395
            TabIndex        =   0
            Top             =   345
            Width           =   1215
            _ExtentX        =   2143
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
            Locked          =   -1  'True
            Valor           =   "21/11/2007"
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
            Height          =   300
            Index           =   1
            Left            =   1395
            TabIndex        =   3
            Top             =   975
            Width           =   1215
            _ExtentX        =   2143
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
            Locked          =   -1  'True
            Valor           =   "21/11/2007"
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
            Height          =   300
            Index           =   2
            Left            =   3285
            TabIndex        =   4
            Top             =   975
            Width           =   1260
            _ExtentX        =   2223
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
            Locked          =   -1  'True
            Valor           =   "21/11/2007"
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   0
            Left            =   1395
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   1
            Text            =   "txt_cb(0)"
            ToolTipText     =   "Ingrese DNI del Programador"
            Top             =   660
            Width           =   1215
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "N° Programación"
            Height          =   195
            Index           =   2
            Left            =   9165
            TabIndex        =   25
            Top             =   300
            Width           =   1200
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "codigo"
            Height          =   195
            Index           =   0
            Left            =   6540
            TabIndex        =   21
            Top             =   405
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label lbl_cb_cod 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb_cod(0)"
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
            Left            =   4785
            TabIndex        =   19
            Top             =   660
            Visible         =   0   'False
            Width           =   1290
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
            Left            =   2610
            TabIndex        =   18
            Top             =   660
            Width           =   3540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Observación:"
            Height          =   195
            Index           =   1
            Left            =   75
            TabIndex        =   16
            Top             =   5460
            Width           =   945
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Al"
            Height          =   195
            Left            =   2940
            TabIndex        =   14
            Top             =   1020
            Width           =   135
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Periodo    Del"
            Height          =   195
            Left            =   75
            TabIndex        =   13
            Top             =   1020
            Width           =   960
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Left            =   75
            TabIndex        =   12
            Top             =   405
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Programado Por"
            Height          =   195
            Left            =   75
            TabIndex        =   11
            Top             =   720
            Width           =   1140
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Programa de Producción"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   105
            TabIndex        =   10
            Top             =   45
            Width           =   11400
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6825
         Left            =   45
         TabIndex        =   8
         Top             =   375
         Width           =   11775
         Begin TrueOleDBGrid70.TDBGrid Dg3 
            Height          =   6480
            Left            =   15
            TabIndex        =   15
            Top             =   345
            Width           =   11730
            _ExtentX        =   20690
            _ExtentY        =   11430
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
            Columns(1).Caption=   "Programa N°"
            Columns(1).DataField=   "num"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fecha"
            Columns(2).DataField=   "fecha"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Programado Por"
            Columns(3).DataField=   "prog"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fch. Ini."
            Columns(4).DataField=   "fchini"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Fch. Fin"
            Columns(5).DataField=   "fchfin"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1191"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1111"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2355"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2275"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2037"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1958"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=7938"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=7858"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=2037"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1958"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=2090"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2011"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HDBFDFD&,.fgcolor=&H80000002&"
            _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H800000&"
            _StyleDefs(11)  =   ":id=2,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H80&"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.fgcolor=&H0&"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H80&"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80&"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
            _StyleDefs(60)  =   "Named:id=33:Normal"
            _StyleDefs(61)  =   ":id=33,.parent=0,.fgcolor=&H80000008&"
            _StyleDefs(62)  =   "Named:id=34:Heading"
            _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000A&,.fgcolor=&H800000&"
            _StyleDefs(64)  =   ":id=34,.wraptext=-1"
            _StyleDefs(65)  =   "Named:id=35:Footing"
            _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(67)  =   "Named:id=36:Selected"
            _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(69)  =   "Named:id=37:Caption"
            _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(71)  =   "Named:id=38:HighlightRow"
            _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=39:EvenRow"
            _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(75)  =   "Named:id=40:OddRow"
            _StyleDefs(76)  =   ":id=40,.parent=33"
            _StyleDefs(77)  =   "Named:id=41:RecordSelector"
            _StyleDefs(78)  =   ":id=41,.parent=34"
            _StyleDefs(79)  =   "Named:id=42:FilterBar"
            _StyleDefs(80)  =   ":id=42,.parent=33"
         End
         Begin VB.Label lblperiodo 
            Caption         =   "lblperiodo"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Index           =   0
            Left            =   9750
            TabIndex        =   23
            Top             =   45
            Width           =   2010
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Programa de Producción"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   135
            TabIndex        =   22
            Top             =   45
            Width           =   11595
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
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
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Listado"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Registro"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Agregar producto           "
      End
      Begin VB.Menu Menu1_4 
         Caption         =   "Seleccionar Producto"
      End
      Begin VB.Menu menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Eliminar producto"
      End
   End
   Begin VB.Menu Menu3 
      Caption         =   "Menu3"
      Visible         =   0   'False
      Begin VB.Menu Menu3_1 
         Caption         =   "Eliminar Producto"
      End
      Begin VB.Menu Menu3_2 
         Caption         =   "Eliminar Todo"
      End
   End
   Begin VB.Menu Menu4 
      Caption         =   "Menu4"
      Visible         =   0   'False
      Begin VB.Menu Menu4_1 
         Caption         =   "Seleccionar un Producto"
      End
      Begin VB.Menu Menu4_2 
         Caption         =   "Seleccionar Varios Registros"
      End
   End
   Begin VB.Menu menu2 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu mn_insumo 
         Caption         =   "Insumos"
         Begin VB.Menu mn_insumo_prod 
            Caption         =   "x Producto"
            Begin VB.Menu mn_insumo_prod_1 
               Caption         =   "Toda la Programación"
            End
            Begin VB.Menu mn_insumo_prod_2 
               Caption         =   "Dia Actual"
            End
         End
         Begin VB.Menu mn_insumo_tprod 
            Caption         =   "Todos los Productos"
            Begin VB.Menu mn_insumo_tprod_3 
               Caption         =   "Resumen de Toda la Programación"
            End
            Begin VB.Menu mn_insumo_tprod_1 
               Caption         =   "Toda la Programación"
            End
            Begin VB.Menu mn_insumo_tprod_2 
               Caption         =   "Dia Actual"
            End
         End
      End
      Begin VB.Menu mn_tarea 
         Caption         =   "Tarea"
         Begin VB.Menu mn_tarea_prod 
            Caption         =   "x Producto"
            Begin VB.Menu mn_tarea_prod_1 
               Caption         =   "Toda la Programación"
            End
            Begin VB.Menu mn_tarea_prod_2 
               Caption         =   "Dia Actual"
            End
         End
         Begin VB.Menu mn_tarea_tprod 
            Caption         =   "Todos los Productos"
            Begin VB.Menu mn_tarea_tprod_3 
               Caption         =   "Resumen de Toda la Programación"
            End
            Begin VB.Menu mn_tarea_tprod_1 
               Caption         =   "Toda la Programación"
            End
            Begin VB.Menu mn_tarea_tprod_2 
               Caption         =   "Dia Actual"
            End
         End
      End
      Begin VB.Menu mn_equipo 
         Caption         =   "Equipos"
         Visible         =   0   'False
         Begin VB.Menu mn_equipo_prod 
            Caption         =   "x Producto"
            Begin VB.Menu mn_equipo_prod_1 
               Caption         =   "Toda la Programación"
            End
            Begin VB.Menu mn_equipo_prod_2 
               Caption         =   "Dia Actual"
            End
         End
         Begin VB.Menu mn_equipo_tprod 
            Caption         =   "Todos los Productos"
            Begin VB.Menu mn_equipo_tprod_1 
               Caption         =   "Toda la Programación"
            End
            Begin VB.Menu mn_equipo_tprod_2 
               Caption         =   "Dia Actual"
            End
         End
      End
      Begin VB.Menu mn_separador 
         Caption         =   "-"
      End
      Begin VB.Menu mn_exportar 
         Caption         =   "Exportar MSExcel"
      End
      Begin VB.Menu mn_imprimir 
         Caption         =   "Imprimir"
      End
   End
End
Attribute VB_Name = "FrmManProduccionPrograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMMANPRODUCCIONPROGRAMA.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : PERMITE LA PROGRAMACION DE LA PRODUCCION, ASI SE PODRA SIMULAR EL GASTO DE
'*                    MATERIA PRIMA E INSUMOS
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 06/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim QueHace As Integer                                   ' INDICA EN QUE MODO SE ENCUENTRA EL FORMULARIO
Dim SeEjecuto As Boolean                                 ' VERIFICA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim RstFrm As New ADODB.Recordset                        ' RECORDSET QUE ALMACENARA INFORMACION DE LA TABLA pro_programa
Dim Agregando As Boolean                                 ' INDICA QUE SE ESTAN AGREGANDO FILAS AL CONTROL FLEXGRID
Dim ARR_TMP() As String                                  ' DEPENDERA DEL FILTRO DE FECHAS
Dim mMesActivo  As Integer                               ' INDICA EL MES ACTIVO
Private Const FORMAT_NUM_PRODUCCION As String = "000000"
Private Const FORMAT_NUM_PROGRAMA As String = "000000"   ' INDICA EL FORMATO DE LA COLUMNA CON EL NUMERO DE PROGRAMA
Private Const Q_ANCHO_COL_FECHA As Integer = 820         ' ANCHO DE LAS COLUMNA DE FECHA, +300 ANCHO DEL TOTAL
Dim fOrdenLista As Boolean                               ' especfica el orden de la lista de la consulta

'*****************************************************************************************************
'* Nombre           : pRegistroAdd
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : GENERAR EL WHERE DE LOS ID'S RECETA PARA QUE NO SE REPITAN
'* Paranetros       : NOMBRE           |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    fSeleccionVarios |  Boolean   |
'* Devuelve         :
'*****************************************************************************************************
Private Sub pRegistroAdd(Optional fSeleccionVarios As Boolean = True)
    Dim SQL_IDREC As String
    Dim xRs As New ADODB.Recordset
    Dim N_SQL As String
    ReDim xCampos(3, 5) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "descripcion":  xCampos(0, 2) = "5000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
    xCampos(1, 0) = "Familia":          xCampos(1, 1) = "famdesc":      xCampos(1, 2) = "1500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
    xCampos(2, 0) = "U.M.":             xCampos(2, 1) = "abrev":        xCampos(2, 2) = "600":      xCampos(2, 3) = "C":    xCampos(2, 4) = "N"

    N_SQL = "SELECT alm_inventario.id as iditem, alm_inventario.descripcion , mae_familia.descripcion AS famdesc, mae_unidades.abrev " _
        + vbCr + " FROM mae_unidades RIGHT JOIN (alm_inventario LEFT JOIN mae_familia ON alm_inventario.idfam = mae_familia.id) ON mae_unidades.id = alm_inventario.idunimed " _
        + vbCr + " WHERE (((alm_inventario.tippro) = 3)) AND alm_inventario.activo = -1 " _
        + vbCr + " ORDER BY alm_inventario.descripcion, mae_familia.descripcion;"
    
    Me.MousePointer = vbHourglass
    If fSeleccionVarios = True Then
        CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, N_SQL, xCampos(), "Buscando Productos"
    Else
        CARGAR_DLL_EPSBUSCAR xCon, xRs, N_SQL, xCampos(), "Buscando Productos", "descripcion", "descripcion", Principio
    End If
    
    If xRs.State = 0 Then GoTo SALIR:
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    
    Dim RstReceta As New ADODB.Recordset           ' BUSCAR LA RECETA PREDETERMINADA
    Dim xFila As Integer
    
    Agregando = True
    Me.MousePointer = vbHourglass
    If Fg1.Rows > 1 Then Fg1.Rows = Fg1.Rows - 1   ' ELIMINANDO TOTALES DE FILAS
    
    If fSeleccionVarios = True Then xRs.MoveFirst
    Do While Not xRs.EOF
        ' SI YA EXISTE LA RECETA NO AGREGAR
        ADD_REG Fg1, Fila_Ninguno
        
        With Fg1
            .TextMatrix(Fg1.Rows - 1, 1) = xRs.Fields("descripcion") & ""      ' ESCRIPPCION RECETA
            ' CARGAR RECETA PREDETERMINADA
            RST_Busq RstReceta, "SELECT TOP 1 pro_receta.id AS idrec, pro_receta.descripcion, pro_receta.codrec, mae_unidades.abrev , pro_receta.idunimed" _
                + vbCr + " FROM pro_receta INNER JOIN mae_unidades ON pro_receta.idunimed = mae_unidades.id " _
                + vbCr + " Where (((pro_receta.iditem) = " + CStr(xRs.Fields("iditem")) + "))" _
                + vbCr + " ORDER BY pro_receta.prirec;", xCon
                                
            If RstReceta.EOF = False Or RstReceta.BOF = False Or RstReceta.RecordCount <> 0 Then
                .TextMatrix(Fg1.Rows - 1, 2) = RstReceta.Fields("codrec") & "" ' CODIGO RECETA
                .TextMatrix(Fg1.Rows - 1, 3) = RstReceta.Fields("abrev") & ""  ' UNIDAD
                .TextMatrix(Fg1.Rows - 1, 4) = RstReceta.Fields("idrec")       ' ID RECETA
            End If
            .TextMatrix(Fg1.Rows - 1, 5) = xRs.Fields("iditem") ' ID ITEM
            Set RstReceta = Nothing
        End With
        
        If fSeleccionVarios = False Then Exit Do
        xRs.MoveNext
    Loop
    GRID_ADD_TOTAL_ROW False
    SumarTodasFilas

SALIR:
    Me.MousePointer = vbDefault
    Set xRs = Nothing
    Agregando = False
    Exit Sub

error:
    Me.MousePointer = vbDefault
    Set xRs = Nothing
    Set RstReceta = Nothing
    Agregando = False
    SHOW_ERROR Me.Name, "pRegistroAdd"
End Sub

'*****************************************************************************************************
'* Nombre           : GRID_ADD_TOTAL_ROW
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : AÑADE FILAS DE TOTAL AL CONTROL Fg1
'* Paranetros       : NOMBRE        |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    LIMPIAR_TOTAL |  Boolean    |
'* Devuelve         :
'*****************************************************************************************************
Sub GRID_ADD_TOTAL_ROW(Optional LIMPIAR_TOTAL As Boolean = True)
    Dim A As Integer
    Dim xTotal As Double
    
    ' LIMPIANDO TOTAL ANTERIOR
    Agregando = True
    If Fg1.Rows > 1 And LIMPIAR_TOTAL = True Then
        For A = 1 To Fg1.Cols - 1
            Fg1.TextMatrix(Fg1.Rows - 1, A) = ""
        Next A
    End If
    
    Fg1.Rows = Fg1.Rows + 1
    Fg1.TextMatrix(Fg1.Rows - 1, 3) = "Total"
    Agregando = False
End Sub

'*****************************************************************************************************
'* Nombre           : pRegistroDel
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UNA FILA DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pRegistroDel()
    If Fg1.Row < 0 Then Exit Sub
    
    If Fg1.Row = Fg1.Rows - 1 Then
        MsgBox "La fila no se puede eliminar" + vbCr + "Seleccione una correcta", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Seguro desea eliminar el Registro" & vbCr & "Producto: " & Fg1.TextMatrix(Fg1.Row, 1), vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    Fg1.RemoveItem (Fg1.Row)
    SumarTodasFilas
End Sub

Private Sub cmd_Click(Index As Integer)
    If QueHace = 3 Then Exit Sub
    Select Case Index
        Case 0 ' AGREGAR PRODUCTO
            pRegistroAdd False
        
        Case 1 ' ELIMINAR REGISTROS AGREGADOS
            pRegistroDel
        
        Case 2
            pRegistroAdd
    End Select
End Sub

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace <> 3 Then
            If Index = 0 Then PopupMenu Menu4
            If Index = 1 Then PopupMenu menu3
        End If
    End If
End Sub

Private Sub Dg3_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg3_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA EN FORMA ASCENDENTE O DECENDENTE LAS COLUMNAS DEL CONTROL Dg3
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstFrm.Sort = CStr(Dg3.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    ' EJECUTA LA BUSQUEDA DE RECETAS
    If Col <> 2 Then Exit Sub
    Dim xRs As New ADODB.Recordset
    Dim N_SQL As String

    If NulosN(Fg1.TextMatrix(Row, 5)) = 0 Then
        MsgBox "Seleccione un Producto", vbExclamation, xTitulo
        Exit Sub
    End If
    
    ReDim xCampos(3, 4) As String
    xCampos(0, 0) = "Descripción":     xCampos(0, 1) = "nombre":        xCampos(0, 2) = "5000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
    xCampos(1, 0) = "Código":          xCampos(1, 1) = "codrec":        xCampos(1, 2) = "1500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
    xCampos(2, 0) = "U.M.":            xCampos(2, 1) = "abrev":         xCampos(2, 2) = "500":      xCampos(2, 3) = "C":    xCampos(2, 4) = "N"

    N_SQL = "SELECT pro_receta.id AS idrec, pro_receta.descripcion as nombre, pro_receta.codrec, mae_unidades.abrev, pro_receta.idunimed " _
        + vbCr + " FROM alm_inventario INNER JOIN (pro_receta INNER JOIN mae_unidades ON pro_receta.idunimed = mae_unidades.id) ON alm_inventario.id = pro_receta.iditem " _
        + vbCr + " WHERE (((alm_inventario.id) = " + CStr(Fg1.TextMatrix(Row, 5)) + ")) " _
        + vbCr + " ORDER BY pro_receta.descripcion;"

    CARGAR_DLL_EPSBUSCAR xCon, xRs, N_SQL, xCampos(), "Buscando Recetas", "nombre", "nombre", Principio, ""

    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    Agregando = True

    Fg1.TextMatrix(Row, 2) = xRs.Fields("codrec") & "" ' CODIGO RECETA
    Fg1.TextMatrix(Row, 3) = xRs.Fields("abrev") & ""  ' UNIDAD
    Fg1.TextMatrix(Row, 4) = xRs.Fields("idrec") & ""  ' ID RECETA
       
    Agregando = False
    Set xRs = Nothing
    Exit Sub
    
SALIR:
    Set xRs = Nothing
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    
    If Col <= 5 Or Col = Fg1.Cols - 1 Then Exit Sub
        Dim N_SQL As String
        Dim M_IDREC As String
        Dim D_FECHA As String
        Dim M_ID As String
        Dim xRs As New ADODB.Recordset
        
        If NulosN(Fg1.TextMatrix(Row, 4)) = 0 Then
            MsgBox "Selecione la Receta" & vbCr & "Producto: " & Fg1.TextMatrix(Row, 1), vbExclamation, xTitulo
            Fg1.TextMatrix(Row, Col) = ""
            GoTo Reconfigurar_datos
        End If
        
        M_IDREC = Fg1.TextMatrix(Row, 4)
        If IsDate(Fg1.TextMatrix(0, Col)) = True Then
            D_FECHA = Format(Fg1.TextMatrix(0, Col), "dd/mm/yy")
        Else
            Exit Sub
        End If
        
        If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
            M_ID = -1
        Else
            M_ID = IIf(QueHace = 1, "-1", CStr(RstFrm.Fields("id")))
        End If
        
        ' BUSCANDO DATOS CON PRODUCCION
        ' SI SE GENERO LA PRODUCCION ANTES DE PROGRAMAR =>> NO PROGRAMAR PARA ESE DIA
        N_SQL = " SELECT pro_programa.id AS idprograma, pro_programa.fecha, pro_programa.fchini, pro_programa.fchfin, [pla_empleados].[nom] & ' ' & [pla_empleados].[nom] AS prog, pro_programadet.canpro, pro_producciondet.idpro, pro_produccion.num AS numprod,  [pla_empleados_1].[apepat] & ' ' & [pla_empleados_1].[apemat] & ' ' & [pla_empleados_1].[nom] AS res  " _
            + vbCr + " FROM pro_produccion RIGHT JOIN ((pla_empleados AS pla_empleados_1 RIGHT JOIN pro_emp AS pro_emp_1 ON pla_empleados_1.id = pro_emp_1.idemp) RIGHT JOIN (((pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) RIGHT JOIN pro_programa ON pro_emp.id = pro_programa.idprog) RIGHT JOIN (pro_programadet RIGHT JOIN pro_producciondet ON (pro_programadet.idpro = pro_producciondet.idpro) AND (pro_programadet.idrec = pro_producciondet.idrec)) ON pro_programa.id = pro_programadet.idprod) ON pro_emp_1.id = pro_producciondet.idres) ON pro_produccion.id = pro_producciondet.idpro " _
            + vbCr + " WHERE (((pro_produccion.dia)=CDATE('" + D_FECHA + "') ) AND ((pro_producciondet.idrec)=" + M_IDREC + "));"
        
        On Error GoTo error
        Me.MousePointer = vbHourglass
        RST_Busq xRs, N_SQL, xCon
        
        If xRs.BOF = False And xRs.EOF = False Or xRs.RecordCount <> 0 Then
            If IsNull(xRs.Fields("idpro")) = False Then
                MsgBox "El producto: " + Fg1.TextMatrix(Row, 1) + " ya se generó la producción para este día:" + vbCr + _
                "Id. Producción:       " + NulosC(xRs.Fields("idpro")) + vbCr + _
                "Num. Producción:   " + Format(NulosC(xRs.Fields("numprod")), FORMAT_NUM_PRODUCCION) + vbCr + _
                "Responsable de Producc.  " + NulosC(xRs.Fields("res")) + vbCr + _
                "No puede Programar......", vbExclamation, xTitulo
                Fg1.TextMatrix(Row, Col) = ""
                GoTo Reconfigurar_datos
            End If
        End If
        
        ' BUSCANDO DATOS CON PROGRAMA DE PRODUCCION
        N_SQL = Replace(N_SQL, "(pro_produccion.dia)", "(pro_programadet.dia)")
        N_SQL = Replace(N_SQL, "(pro_producciondet.idrec)", "(pro_programadet.idrec)")
        
        Set xRs = Nothing
        RST_Busq xRs, N_SQL, xCon
        If xRs.BOF = True Or xRs.EOF = True Or xRs.RecordCount = 0 Then GoTo poner_datos:
        
        If IsNull(xRs.Fields("idpro")) = False Then
            ' SI FUE REGISTRADO LA PRODUCCION =>> SALIR(NO HACER NADA)
            If xRs.Fields("idpro") <> "0" Then
                MsgBox "El producto: " + Fg1.TextMatrix(Row, 1) + " ya se segistró en producción para este día:" + vbCr + _
                "Id. Producción:       " + NulosC(xRs.Fields("idpro")) + vbCr + _
                "Num. Producción:   " + Format(NulosC(xRs.Fields("numprod")), FORMAT_NUM_PRODUCCION) + vbCr + _
                "Responsable de Producc.  " + NulosC(xRs.Fields("res")) + vbCr + _
                "No puede Modificar la cantidad......", vbExclamation, xTitulo
                ' COLOCANDO EL VALOR ORIGINAL
                Fg1.TextMatrix(Row, Col) = NulosN(xRs.Fields("canpro"))
                GoTo Reconfigurar_datos
            End If
        Else
            If RstFrm.Fields("idprograma") & "" <> M_ID Then    ' ESTA EN OTRA PROGRAMACION
                ' VALIDAR QUE LA CANTIDAD SEA UNICA EN UN DIA PARA CADA RECETA =>> SALIR(NO HACER NADA)
                MsgBox "El producto: " + Fg1.TextMatrix(Row, 1) + " esta programado para este día en otro registro:" + vbCr + _
                "Id. Registro " + NulosC(xRs.Fields("id")) + vbCr + _
                "Fch Programación: " + NulosC(xRs.Fields("fecha")) + vbCr + _
                "Cant Programada: " + Format(NulosN(xRs.Fields("canpro")), FORMAT_CANTIDAD) + vbCr + _
                "Programado Por:  " + NulosC(xRs.Fields("prog")) + vbCr + _
                "Periodo: Del " + NulosC(xRs.Fields("fchini")) + " Al " + NulosC(xRs.Fields("fchfin")) + vbCr + _
                "Modifique el registro en la programación indicada, no en esta......", vbExclamation, xTitulo
                Fg1.TextMatrix(Row, Col) = ""
                GoTo Reconfigurar_datos
            End If
        End If
        
poner_datos:
        '------------------------------------------------------------------
        
Reconfigurar_datos:
    SumarColumnas
    SumarFilas
    SumarColumnas True

SALIR:
    Set xRs = Nothing
    Me.MousePointer = vbDefault
    Exit Sub

error:
    Me.MousePointer = vbDefault
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Fg1_CellChanged"
End Sub

'*****************************************************************************************************
'* Nombre           : SumarColumnas
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : SUMA LAS COLUMNAS DEL CONTROL Fg1
'* Paranetros       : NOMBRE    |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    SUM_TOTAL |  Boolean    |
'* Devuelve         :
'*****************************************************************************************************
Sub SumarColumnas(Optional SUM_TOTAL As Boolean = False)
    Dim A As Integer
    Dim xTotal As Double
    Dim Q_COL As Integer
    If SUM_TOTAL = False Then Q_COL = Fg1.Col
    If SUM_TOTAL = True Then Q_COL = Fg1.Cols - 1
    
    For A = 1 To Fg1.Rows - 2
        If Fg1.ColHidden(Q_COL) = False Then xTotal = xTotal + NulosN(Fg1.TextMatrix(A, Q_COL))
    Next A
    
    Fg1.TextMatrix(Fg1.Rows - 1, Q_COL) = Format(NulosN(xTotal), "0.00")
End Sub

'*****************************************************************************************************
'* Nombre           : SumarFilas
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : SUMAS LAS FILAS DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub SumarFilas()
    Dim A As Integer
    Dim xTotal As Double
    
    For A = 9 To Fg1.Cols - 2
        xTotal = xTotal + NulosN(Fg1.TextMatrix(Fg1.Row, A))
    Next A
    
    Fg1.TextMatrix(Fg1.Row, Fg1.Cols - 1) = Format(xTotal, "0.00")
End Sub

Private Sub Fg1_EnterCell()
    If Agregando = True Then Exit Sub
    
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
        Exit Sub
    End If
    
    If Fg1.Col = 4 Then
        Fg1.Editable = flexEDNone
    Else
        If Fg1.Col = Fg1.Cols - 1 Then
            Fg1.Editable = flexEDNone
        Else
            Fg1.Editable = flexEDKbdMouse
        End If
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Row = Fg1.Rows - 1 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If Col <= 5 And KeyAscii <> 13 Then KeyAscii = 0
    
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Or KeyCode = 45 Then ' F3 = Agregar Item
        cmd_Click 0
    ElseIf KeyCode = 115 Or KeyCode = 46 Then
        cmd_Click 1                       ' F4 = Eliminar Item
    End If
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace = 3 Then
            PopupMenu menu2
        Else
            PopupMenu menu1
        End If
    End If
End Sub

Private Sub Fg1_RowColChange()
    If Agregando = True Then Exit Sub
    
    If Fg1.Rows = 1 Then
        Exit Sub
    End If
    
    If Fg1.TextMatrix(Fg1.Row, 1) = "" Then
        Exit Sub
    End If
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = True Then Exit Sub
   
    SeEjecuto = False
    mMesActivo = xMes
    pCargarGrid
    SeEjecuto = True
    
    If RstFrm.State = 0 Then Exit Sub
    
    If RstFrm.RecordCount = 0 Then
        If MsgBox("No se ha registrado ningún programa de producción, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
            Nuevo
        End If
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    SeEjecuto = False
    Agregando = False
    QueHace = 3
    Fg1.SelectionMode = flexSelectionByRow

    OCULTAR_COL Fg1, 4, 8
    
    Fg1.ColAlignment(3) = flexAlignLeftBottom
    Fg1.ColAlignment(4) = flexAlignCenterBottom
    Fg1.ColAlignment(5) = flexAlignLeftBottom
    Dg3.Columns("fecha").NumberFormat = FORMAT_DATE
    Dg3.Columns("fchini").NumberFormat = FORMAT_DATE
    Dg3.Columns("fchfin").NumberFormat = FORMAT_DATE
    
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Fg1.FrozenCols = 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 1
        Exit Sub
    End If
End Sub

Private Sub Menu1_1_Click()
    pRegistroAdd False
End Sub

Private Sub Menu1_3_Click()
    pRegistroDel
End Sub

Private Sub Menu1_4_Click()
    pRegistroAdd
End Sub

Private Sub Menu3_1_Click()
    pRegistroDel
End Sub

Private Sub Menu3_2_Click()
    Dim Q_ROW As Long
    
    If Fg1.Rows <= 2 Then Exit Sub
    
    If MsgBox("Seguro desea eliminar todos los Productos", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
    
    Agregando = True
    Fg1.Rows = 1
    GRID_ADD_TOTAL_ROW True
    Agregando = False
End Sub

Private Sub Menu4_1_Click()
    pRegistroAdd False
End Sub

Private Sub Menu4_2_Click()
    pRegistroAdd True
End Sub

Private Sub mn_exportar_Click()
    pExportar
End Sub

Private Sub mn_imprimir_Click()
    pImprimir False
End Sub

Private Sub mn_insumo_tprod_3_Click()
    ' CONSULTA DE INSUMO/TODO PRODUCTO/TODA PROGRAMACION
    If fCargarDatosArray() = False Then Exit Sub
    pCargarFrmLista E_INSUMO, 4
End Sub

Private Sub mn_tarea_prod_1_Click()
    ' CONSULTA DE TAREA/ X PRODUCTO/TODA PROGRAMACION
    If fValidarSeleccionProducto(Fg1.Row) = False Then Exit Sub
    If fValidarSeleccionCantidad(Fg1.Row) = False Then Exit Sub
    If fCargarDatosArray() = False Then Exit Sub
    pCargarFrmLista e_TAREA, 5, Fg1.TextMatrix(Fg1.Row, 4)
End Sub

Private Sub mn_tarea_prod_2_Click()
    ' CONSULTA DE TAREA/ X PRODUCTO/DIA ACTUAL
    If fValidarSeleccionProducto(Fg1.Row) = False Then Exit Sub
    If fValidarSeleccionCantidad(Fg1.Row) = False Then Exit Sub
    If fValidarSeleccionDia() = False Then Exit Sub
    If fCargarDatosArray(True) = False Then Exit Sub
    pCargarFrmLista e_TAREA, 6, Fg1.TextMatrix(Fg1.Row, 4)
End Sub

Private Sub mn_tarea_tprod_1_Click()
    ' CONSULTA DE TAREA/TODO PRODUCTO/TODA PROGRAMACION
    If fCargarDatosArray() = False Then Exit Sub
    pCargarFrmLista e_TAREA, 7
End Sub

Private Sub mn_tarea_tprod_2_Click()
    ' CONSULTA DE TAREA/TODO PRODUCTO/DIA ACTUAL
    If fValidarSeleccionDia() = False Then Exit Sub
    If fCargarDatosArray(True) = False Then Exit Sub
    pCargarFrmLista e_TAREA, 8
End Sub

Private Sub mn_tarea_tprod_3_Click()
    ' CONSULTA DE TAREA/TODO PRODUCTO/TODA PROGRAMACION
    If fCargarDatosArray() = False Then Exit Sub
    pCargarFrmLista e_TAREA, 9
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 3 Then MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then Modificar
    
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstFrm.Requery
            Dg3.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 8 Then Filtrar
    
    If Button.Index = 9 Then RstFrm.Filter = ""
    
    If Button.Index = 10 Then CambiarMes
    
    If Button.Index = 11 Then Buscar

    If Button.Index = 15 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : pImprimir
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : IMPRIME LOS DATOS DEL CONTROL Fg1
'* Paranetros       : NOMBRE      |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    IMP_LISTADO |  Boolean    |
'* Devuelve         :
'*****************************************************************************************************
Private Sub pImprimir(Optional IMP_LISTADO As Boolean = False)
    On Error GoTo error

    Me.MousePointer = vbHourglass
    
    If IMP_LISTADO = False Then
        Dim X_PRINT As New SGI2_funciones.formularios
        If Me.TabOne1.CurrTab = 1 Then
            X_PRINT.Imprimir_x_VSFlexGrid Fg1, "PROGRAMA DE PRODUCCION - " & txt(2).Text, "Programado Por: " + StrConv(lbl_cb(0).Caption, 3), "Del: " + TxtFecha(0).valor + " Al: " + TxtFecha(2).valor, False, True
        Else
            MsgBox "Primero muestre el detalle del Registro" + vbCr + _
                   "Luego inténtelo otra vez", vbExclamation, xTitulo
        End If
        Set X_PRINT = Nothing
    Else
        TDB_IMPRIMIR Dg3, "IMPRESIÓN DE PROGRAMA DE PRODUCCIÓN", "PROGRAMA DE PRODUCCIÓN  -  Periodo: " + MonthName(mMesActivo, False)
    End If

    Me.MousePointer = vbDefault
    Exit Sub

error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimir"
End Sub

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO DE LA TABLA pro_programa
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
        MsgBox "No hay registros", vbExclamation, xTitulo
        Exit Sub
    End If
    TabOne1.CurrTab = 0
    
    If MsgBox("¿Esta seguro de eliminar la programación seleccionada?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo) = vbYes Then
        xCon.Execute "DELETe * FROM pro_programa WHERE id = " & RstFrm("id") & ""
        MsgBox "La programación del dia " + Format(RstFrm("fecha"), "dd/mm/yy") + " fue eliminada con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstFrm.Requery
        Dg3.Refresh
        
        If RstFrm.RecordCount = 0 Then
            If MsgBox("No hay registrado ningún programa, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo) = vbYes Then
                Nuevo
            End If
        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL PROCESO INGRESAR O MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    TabOne1.TabEnabled(0) = True
    ActivaTool
    Habilitar_Obj False
    Label1.Caption = "Detalle Programa de Producción"
    Fg1.SelectionMode = flexSelectionByRow
    TabOne1.CurrTab = 0
    Dg3.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Modificar()
    QueHace = 2
    TabOne1.TabEnabled(0) = False
    ActivaTool
    Habilitar_Obj True
    habilitar TxtFecha, False
    
    If TabOne1.CurrTab = 0 Then
        TabOne1.CurrTab = 1
        MuestraSegundoTab
    End If
    
    GRID_COMBOLIST Fg1, 2
    Label1.Caption = "Modificando Programa de Producción"
    TxtFecha(0).SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DETALLADA DEL REGISTRO EN LA PESTALA DETALLE DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub MuestraSegundoTab()
On Error GoTo error
    With RstFrm
        Blanquea
        If .State = 0 Then Exit Sub
        
        If .EOF = True Or .BOF = True Or .RecordCount = 0 Then Exit Sub
        
        Me.MousePointer = vbHourglass
        txt(0).Text = .Fields("id") & ""
        txt(2).Text = .Fields("num") & ""
        TxtFecha(0).valor = .Fields("fecha") & ""
        lbl_cb_cod(0).Caption = .Fields("idprog") & ""
        txt_cb(0).Text = .Fields("prognum") & ""
        lbl_cb(0).Caption = .Fields("prog") & ""
        
        TxtFecha(1).valor = .Fields("fchini") & ""
        TxtFecha(2).valor = .Fields("fchfin") & ""
        txt(1).Text = .Fields("obs") & ""
        MuestraDetalle
    End With
    Me.MousePointer = vbDefault
    Exit Sub
    
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "MuestraSegundoTab"
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraDetalle
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DETALLADA DEL REGISTRO EN LA PESTALA DETALLE DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub MuestraDetalle()
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    Dim xCol, xFil As Integer
    Dim xSQL As String
    Dim xFch As Date
    Dim xFila  As Integer
    Agregando = True
    
    ' BUSCAR TODOS LOS PRODUCTOS PROGRAMADOS
    xSQL = "SELECT pro_programadet.idrec, pro_programadet.iditem, pro_receta.descripcion, pro_receta.codrec, mae_unidades.abrev " _
    + vbCr + " FROM mae_unidades INNER JOIN (pro_receta INNER JOIN pro_programadet ON pro_receta.id = pro_programadet.idrec) ON mae_unidades.id = pro_receta.idunimed " _
    + vbCr + " GROUP BY pro_programadet.idprod, pro_programadet.idrec, pro_programadet.iditem, pro_receta.descripcion, pro_receta.codrec, mae_unidades.abrev " _
    + vbCr + " Having (((pro_programadet.idprod) = " + CStr(RstFrm("id")) + "))  " _
    + vbCr + " ORDER BY pro_receta.descripcion;"

    RST_Busq Rst, xSQL, xCon
    Fg1.Rows = 1
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            xFila = Fg1.Rows
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(xFila, 1) = Rst.Fields("descripcion") & ""   ' ESCRIPPCION RECETA
            Fg1.TextMatrix(xFila, 2) = Rst.Fields("codrec") & ""        ' CODIGO RECETA
            Fg1.TextMatrix(xFila, 3) = Rst.Fields("abrev") & ""         ' IDUNIDAD
            Fg1.TextMatrix(xFila, 4) = Rst.Fields("idrec") & ""         ' ID RECETA
            Fg1.TextMatrix(xFila, 5) = Rst.Fields("iditem") & ""        ' ID ITEM
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
    Else
        GoTo PONER_TOTALES:
    End If
    
    ' CARGAMOS LAS CANTIDADES INGRESADAS POR FECHA
    Fg1.Cols = 9
    For xFch = RstFrm("fchini") To RstFrm("fchfin")
        Fg1.Cols = Fg1.Cols + 1
        Fg1.ColFormat(Fg1.Cols - 1) = FORMAT_CANTIDAD
        Fg1.ColAlignment(Fg1.Cols - 1) = flexAlignRightCenter
        Fg1.ColWidth(Fg1.Cols - 1) = 700
        Fg1.Col = Fg1.Cols - 1:     Fg1.Row = 0:    Fg1.CellAlignment = flexAlignRightCenter
        Fg1.TextMatrix(0, Fg1.Cols - 1) = Format(xFch, "dd/mm/yy")
        Fg1.ColWidth(Fg1.Cols - 1) = Q_ANCHO_COL_FECHA
        
        ' mostramos las cantidades a producir por cada producto
        For A = 1 To Fg1.Rows - 1
            xSQL = " SELECT pro_programadet.canpro " _
                + vbCr + " FROM pro_programadet " _
                + vbCr + " WHERE pro_programadet.idrec= " + CStr(Fg1.TextMatrix(A, 4)) + " AND pro_programadet.idprod=" + CStr(RstFrm("id")) + " AND pro_programadet.dia=CDate('" & Fg1.TextMatrix(0, Fg1.Cols - 1) & "');"
            
            RST_Busq Rst, xSQL, xCon
            
            If Rst.RecordCount <> 0 Then
                Fg1.TextMatrix(A, Fg1.Cols - 1) = Format(Rst("canpro"), "0.00")
            End If
            Set Rst = Nothing
        Next A
    Next xFch
    
PONER_TOTALES:
    Fg1.Cols = Fg1.Cols + 1
    Fg1.ColFormat(Fg1.Cols - 1) = FORMAT_CANTIDAD
    Fg1.Col = Fg1.Cols - 1:     Fg1.Row = 0:    Fg1.CellAlignment = flexAlignRightCenter
    Fg1.ColAlignment(Fg1.Cols - 1) = flexAlignRightCenter
    Fg1.ColWidth(Fg1.Cols - 1) = Q_ANCHO_COL_FECHA + 100
    Fg1.TextMatrix(0, Fg1.Cols - 1) = "Total"
    GRID_ADD_TOTAL_ROW False
    SumarTodasFilas
    Agregando = True
    
    With Fg1
        If QueHace = 3 Then
            GRID_COLOR_FONDO Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1, &HC0C0FF
        Else
            GRID_COLOR_FONDO Fg1, 1, Fg1.Cols - 1, Fg1.Rows - 1, Fg1.Cols - 1, &HFFC0C0
        End If
    End With
    
    Agregando = False
End Sub

'*****************************************************************************************************
'* Nombre           : Habilitar_Obj
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES ESPECIFICADOS DEL FORMULARIO
'* Paranetros       : NOMBRE    |  TIPO        |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    band      |  Boolean     |
'* Devuelve         :
'*****************************************************************************************************
Sub Habilitar_Obj(band As Boolean)
    habilitar_Locked TxtFecha, Not band
    habilitar_Locked txt, Not band
    habilitar_Locked txt_cb, Not band
    habilitar Me.cb, band
    habilitar Cmd, band
    txt(2).Enabled = False
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA LOS CONTROLES TEXTBOX PARA EL INGRESO DE NUEVOS REGISTROS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    LimpiaText TxtFecha
    LimpiaText txt
    LimpiaText txt_cb
    Fg1.Rows = 1
    Fg1.Cols = 9
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA Y DESACTIVA LOS BOTONES DE LA BARRA DE HERRAMIENTAS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

'*****************************************************************************************************
'* Nombre           : Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA EL INGRESO DE UN NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Nuevo()
    QueHace = 1
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    TxtFecha(0).valor = Date:    TxtFecha(1).valor = Date:    TxtFecha(2).valor = Date
    habilitar TxtFecha, True
    Habilitar_Obj True
    Blanquea
    Label1.Caption = "Programando La Producción"
    txt(2).Text = Format(HallaCodigoTabla("pro_programa", xCon, "id"), FORMAT_NUM_PROGRAMA)
    Fg1.Rows = 1
    Fg1.Cols = 9
    GRID_COMBOLIST Fg1, 2
   TxtFecha(0).SetFocus
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index = 1 Then pImprimir True

    If ButtonMenu.Index = 2 Then pImprimir
End Sub

Sub GeneraPeriodo()
    If NulosC(TxtFecha(1).valor) <> "" And NulosC(TxtFecha(2).valor) <> "" Then
        If CDate(TxtFecha(1).valor) >= CDate(TxtFecha(2).valor) Then
            MsgBox "La fecha de inicio no puede ser mayor a la fecha de final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFecha(1).SetFocus
            Exit Sub
        Else
            Dim A As Date
            Dim C, xCol, NumCols As Integer
            Agregando = True
            
            If Fg1.Cols = 9 Then
                For A = CDate(TxtFecha(1).valor) To CDate(TxtFecha(2).valor)
                    Fg1.Cols = Fg1.Cols + 1
                    Fg1.ColAlignment(Fg1.Cols - 1) = flexAlignRightCenter
                    Fg1.Col = Fg1.Cols - 1:     Fg1.Row = 0:    Fg1.CellAlignment = flexAlignRightCenter
                    Fg1.ColFormat(Fg1.Cols - 1) = FORMAT_CANTIDAD
                    Fg1.TextMatrix(0, Fg1.Cols - 1) = Format(A, "dd/mm/yy")
                    Fg1.ColWidth(Fg1.Cols - 1) = Q_ANCHO_COL_FECHA
                Next A
            Else
                xCol = 9
                Fg1.Cols = Fg1.Cols - 1       ' QUITANDO COLUMNA TOTAL
                NumCols = Fg1.Cols - 1
                
                If Fg1.Rows > 1 Then GRID_COLOR_FONDO Fg1, 1, Fg1.Cols - 1, Fg1.Rows - 1, Fg1.Cols - 1, &H80000005
                For A = CDate(TxtFecha(1).valor) To CDate(TxtFecha(2).valor)
                    If xCol > NumCols Then Fg1.Cols = Fg1.Cols + 1
                    Fg1.ColAlignment(Fg1.Cols - 1) = flexAlignRightCenter
                    Fg1.Col = Fg1.Cols - 1:     Fg1.Row = 0:    Fg1.CellAlignment = flexAlignRightCenter
                    Fg1.ColFormat(Fg1.Cols - 1) = FORMAT_CANTIDAD
                    Fg1.TextMatrix(0, Fg1.Cols - 1) = Format(A, "dd/mm/yy")
                    Fg1.ColWidth(Fg1.Cols - 1) = Q_ANCHO_COL_FECHA
                    
                    xCol = xCol + 1
                Next A
                
                ' ELIMINADO COLUMNAS SI LA FECHA FINAL ES MENOR A LA ULTIMA FECHA FINAL
                If xCol < NumCols Then
                    For A = xCol To NumCols
                        Fg1.Cols = Fg1.Cols - 1
                        If (xCol - 1) = (Fg1.Cols - 1) Then Exit For
                    Next A
                End If
            End If
            
            Fg1.Cols = Fg1.Cols + 1
            
            Fg1.Col = Fg1.Cols - 1:     Fg1.Row = 0:    Fg1.CellAlignment = flexAlignRightCenter
            Fg1.TextMatrix(0, Fg1.Cols - 1) = "Total"
            Fg1.ColWidth(Fg1.Cols - 1) = Q_ANCHO_COL_FECHA + 100
            
            If Fg1.Rows <> 1 Then GRID_COLOR_FONDO Fg1, 1, Fg1.Cols - 1, Fg1.Rows - 1, Fg1.Cols - 1, &HFFC0C0
            
            SumarTodasFilas
            
            Agregando = False
        End If
    End If
    
    If Fg1.Rows < 2 Then GRID_ADD_TOTAL_ROW
End Sub

'*****************************************************************************************************
'* Nombre           : SumarTodasFilas
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : SUMA LAS FILAS DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub SumarTodasFilas()
    Dim A, B As Integer
    Dim xTotal As Double
    
    ' TOTALES POR FILA
    Agregando = True
    For B = 1 To Fg1.Rows - 2
        xTotal = 0
        For A = 9 To Fg1.Cols - 2
            If Fg1.ColHidden(A) = False Then xTotal = xTotal + NulosN(Fg1.TextMatrix(B, A))
        Next A
        Fg1.TextMatrix(B, Fg1.Cols - 1) = Format(xTotal, "0.00")
    Next B
    
    ' TOTALES POR COLUMNA
    If Fg1.Rows = 2 Then
        For B = 6 To Fg1.Cols - 1
            Fg1.TextMatrix(Fg1.Rows - 1, B) = "0"
        Next B
        Exit Sub
    End If
    
    For B = 6 To Fg1.Cols - 1
        xTotal = 0
        For A = 1 To Fg1.Rows - 2
            If Fg1.ColHidden(B) = False Then xTotal = xTotal + NulosN(Fg1.TextMatrix(A, B))
        Next A
        If Fg1.Rows > 1 Then Fg1.TextMatrix(Fg1.Rows - 1, B) = Format(xTotal, "0.00")
    Next B
    Agregando = False
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA pro_programa, ESTA FUNCION DEVUELVE VERDADERO
'*                    CUANDO TIENE EXITO
'* Paranetros       :
'* Devuelve         : Boolean
'*****************************************************************************************************
Function Grabar() As Boolean
    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modificar") + " la Programación", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo SALIR
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xCod As Integer
    Dim xCol, xFil As Integer
    Dim RstTmp As New ADODB.Recordset           ' ALMACENARA LOS REGISTROS DEL DETALLE DEL PROGRAMA CUANDO SE MODIFIQUE
        
    On Error GoTo LaCague

    xCon.BeginTrans
    
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT top 1 * FROM pro_programa ", xCon
        RST_Busq RstDet, "SELECT top 1 * FROM pro_programadet", xCon
        xCod = HallaCodigoTabla("pro_programa", xCon, "id")
        RstCab.AddNew
        RstCab("id") = xCod
    Else
        ' CARGAR LOS DATOS ANTES DE ELIMINAR
        RST_Busq RstTmp, _
                "SELECT pro_programadet.idrec, pro_programadet.iditem, pro_programadet.canpro, Format([pro_programadet].[dia],'dd/mm/yy') AS dia, pro_programadet.idpro " _
                    + vbCr + " From pro_programadet " _
                    + vbCr + " WHERE (((pro_programadet.idprod)=" & RstFrm("id") & ")) ", xCon
        
        RST_Busq RstCab, "SELECT * FROM pro_programa WHERE id =" & RstFrm("id") & "", xCon
        xCon.Execute "DELETE * FROM pro_programadet WHERE idprod = " & RstFrm("id") & ""
        
        RST_Busq RstDet, "SELECT top 1 * FROM pro_programadet", xCon
        xCod = RstFrm("id")
    End If
    
    RstCab("fecha") = CDate(TxtFecha(0).valor)
    RstCab("num") = Format(xCod, FORMAT_NUM_PROGRAMA)
    RstCab("idprog") = Val(lbl_cb_cod(0).Caption)
    RstCab("fchini") = CDate(TxtFecha(1).valor)
    RstCab("fchfin") = CDate(TxtFecha(2).valor)
    RstCab("obs") = Trim(txt(1).Text) & ""
    RstCab.Update
    
    For xFil = 1 To Fg1.Rows - 2
        For xCol = 9 To Fg1.Cols - 2
            If NulosN(Fg1.TextMatrix(xFil, xCol)) > 0 And Fg1.ColHidden(xCol) = False Then
                RstDet.AddNew
                RstDet("idprod") = xCod
                RstDet("idrec") = Fg1.TextMatrix(xFil, 4)
                RstDet("iditem") = Fg1.TextMatrix(xFil, 5)
                RstDet("dia") = CDate(Fg1.TextMatrix(0, xCol))
                RstDet("canpro") = NulosN(Fg1.TextMatrix(xFil, xCol))
                
                If QueHace <> 1 Then
                    RstTmp.Filter = "idrec= " + Fg1.TextMatrix(xFil, 4) + " and dia='" + Format(Fg1.TextMatrix(0, xCol), "dd/mm/yy") + "'"
                    If RstTmp.RecordCount > 0 Then
                        ' ACTUALIZANDO SI YA PASO POR PRODUCCION
                        RstDet("idpro") = RstTmp.Fields("idpro") & ""
                    End If
                End If
                RstDet.Update
            End If
        Next xCol
    Next xFil
    MsgBox "La programación se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo
    xCon.CommitTrans
    Grabar = True

SALIR:
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstTmp = Nothing
    Exit Function
    
LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstTmp = Nothing
    
    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo:"
    Grabar = False
End Function

'*****************************************************************************************************
'* Nombre           : fValidarDatos
'* Tipo             : FUNCION
'* Descripcion      : VERIFICA QUE LOS DATOS INGRESADOS SON LOS CORRECTOS, DEVUELVE VERDADERO SI LOS
'*                    DATOS SON CORRECTOS
'* Paranetros       :
'* Devuelve         : Boolean
'*****************************************************************************************************
Private Function fValidarDatos() As Boolean
    If NulosC(TxtFecha(0).valor) = "" Or IsDate(TxtFecha(0).valor) = False Then
        MsgBox "No ha especificado la fecha de creación del programa", vbInformation, xTitulo
        TxtFecha(0).SetFocus
        Exit Function
    End If
    
    If lbl_cb_cod(0).Caption = "" Then
        MsgBox "No ha especificado el Programador", vbInformation, xTitulo
        txt_cb(0).SetFocus
        Exit Function
    End If
    
    If TxtFecha(1).valor = "" Or IsDate(TxtFecha(1).valor) = False Then
        MsgBox "No ha especificado la fecha de inicio del programa", vbInformation, xTitulo
        TxtFecha(1).SetFocus
        Exit Function
    End If
    
    If TxtFecha(2).valor = "" Then
        MsgBox "No ha especificado la fecha final del programa", vbInformation, xTitulo
        TxtFecha(2).SetFocus
        Exit Function
    End If
    
    If Fg1.Rows = 2 Then
        MsgBox "No ha especificado los productos para el programa de producción", vbInformation, xTitulo
        Fg1.SetFocus
        Exit Function
    End If
    
    fValidarDatos = True
End Function

Private Sub TxtFecha_Validate(Index As Integer, Cancel As Boolean)
    If Index = 0 Then Exit Sub
    
    If TxtFecha(Index).Enabled = False Then Exit Sub
    
    If Index = 2 Then
        If NulosC(TxtFecha(2).valor) = "" Then Exit Sub
        If TxtFecha(1).valor <> "" Then GeneraPeriodo
    Else
        If NulosC(TxtFecha(1).valor) = "" Then Exit Sub
        If TxtFecha(2).valor <> "" Then GeneraPeriodo
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : fCargarDatosArray
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS DATOS EN EL ARRAY ARR_TMP
'* Paranetros       : NOMBRE       |  TIPO        |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    UN_SOLO_DIA  |  Boolean     |
'* Devuelve         : Boolean
'*****************************************************************************************************
Private Function fCargarDatosArray(Optional UN_SOLO_DIA As Boolean = False) As Boolean
    Dim Q_TOTAL_DIAS As Integer
    Dim D_DIA As Date
    Dim POS_ARR As Integer
    
    Erase ARR_TMP()
    
    If UN_SOLO_DIA = False Then
        Q_TOTAL_DIAS = DateDiff("d", CDate(TxtFecha(1).valor), CDate(TxtFecha(2).valor))
        
        If Q_TOTAL_DIAS < 0 Then Exit Function
        ReDim ARR_TMP(Q_TOTAL_DIAS + 1, 4)
        POS_ARR = 0
        
        For D_DIA = CDate(TxtFecha(1).valor) To CDate(TxtFecha(2).valor)
            ARR_TMP(POS_ARR, 0) = Format(D_DIA, "dd/mm/yy")
            ARR_TMP(POS_ARR, 1) = "'" + Format(D_DIA, "dd/mm/yy") + "'"
            POS_ARR = POS_ARR + 1
        Next D_DIA
    Else
        Q_TOTAL_DIAS = 0
        ReDim ARR_TMP(Q_TOTAL_DIAS + 1, 4)
        ARR_TMP(Q_TOTAL_DIAS, 0) = Format(Fg1.TextMatrix(0, Fg1.Col), "dd/mm/yy")
        ARR_TMP(Q_TOTAL_DIAS, 1) = "'" + Format(Fg1.TextMatrix(0, Fg1.Col), "dd/mm/yy") + "'"
    End If
    
    ARR_TMP(Q_TOTAL_DIAS + 1, 0) = "Total"
    ARR_TMP(Q_TOTAL_DIAS + 1, 1) = "Total"
    
    fCargarDatosArray = True
End Function

Private Sub mn_insumo_prod_1_Click()
    ' CONSULTA DE INSUMO/ X PRODUCTO/TODA PROGRAMACION
    If fValidarSeleccionProducto(Fg1.Row) = False Then Exit Sub
    If fValidarSeleccionCantidad(Fg1.Row) = False Then Exit Sub
    If fCargarDatosArray() = False Then Exit Sub
    pCargarFrmLista E_INSUMO, 0, Fg1.TextMatrix(Fg1.Row, 4)
End Sub

'*****************************************************************************************************
'* Nombre           : fValidarSeleccionProducto
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : VERIFICA LA SELECCION DE FILAS EN EL CONTROL Fg1
'* Paranetros       : NOMBRE    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Q_ROW_INI |  Long      |
'* Devuelve         : Boolean
'*****************************************************************************************************
Private Function fValidarSeleccionProducto(Optional Q_ROW_INI As Long = -1) As Boolean
    Dim i_row, i_col, Q_INI As Long
    Dim F_HAYDATOS As Boolean
    
    If Q_ROW_INI = -1 Then Q_INI = Q_ROW_INI
    
    With Fg1
        For i_row = Q_INI To .Rows - 2
            If .TextMatrix(i_row, 1) <> "" Then
                F_HAYDATOS = True
                Exit For
            End If
            If Q_ROW_INI <> -1 Then Exit For
        Next i_row
        
        If F_HAYDATOS = False Then
            MsgBox "Seleccione de nuevo el Producto", vbExclamation, xTitulo
            .Row = 1
            .Col = 3
            .SetFocus
            Exit Function
        End If
    End With
       
    fValidarSeleccionProducto = True
End Function

'*****************************************************************************************************
'* Nombre           : fValidarSeleccionCantidad
'* Tipo             : PROCEDIMIENTO
'* Descripcion      :
'* Paranetros       : NOMBRE     |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Q_ROW_INI  |  Long       |
'* Devuelve         : Boolean
'*****************************************************************************************************
Private Function fValidarSeleccionCantidad(Optional Q_ROW_INI As Long = -1) As Boolean
    Dim i_row, i_col, Q_INI As Long
    Dim F_HAYDATOS As Boolean
    
    If Q_ROW_INI = -1 Then Q_INI = Q_ROW_INI
    
    With Fg1
        For i_row = Q_INI To .Rows - 2
            For i_col = 9 To .Cols - 2
                If IsNumeric(.TextMatrix(i_row, i_col)) = True Or Trim(.TextMatrix(i_row, i_col)) <> 0 Then
                    F_HAYDATOS = True
                    If Q_ROW_INI <> -1 Then GoTo REVIZA     ' SOLO VALIDA PARA UN PRODUCTO
                    If Q_ROW_INI = -1 Then Exit For         ' VALIDA PARA TODOS LOS PRODUCTO
                End If
            Next i_col
        Next i_row
REVIZA:
        If F_HAYDATOS = False Then
            MsgBox "No ha ingresado cantidades en el programa" + IIf(Q_ROW_INI <> -1, " para el producto seleccionado", "") + vbCr + "Si desea continuar Primero ingrese las cantidades", vbExclamation, xTitulo
            .Row = Q_ROW_INI
            .Col = 9
            .SetFocus
            Exit Function
        End If
    End With
    
    fValidarSeleccionCantidad = True
End Function

Private Sub mn_insumo_prod_2_Click()
    ' CONSULTA DE INSUMO/ X PRODUCTO/DIA ACTUAL
    If fValidarSeleccionProducto(Fg1.Row) = False Then Exit Sub
    If fValidarSeleccionCantidad(Fg1.Row) = False Then Exit Sub
    If fValidarSeleccionDia() = False Then Exit Sub
    If fCargarDatosArray(True) = False Then Exit Sub
    pCargarFrmLista E_INSUMO, 1, Fg1.TextMatrix(Fg1.Row, 4)
End Sub
 
Private Sub mn_insumo_tprod_1_Click()
    ' CONSULTA DE INSUMO/TODO PRODUCTO/TODA PROGRAMACION
    If fCargarDatosArray() = False Then Exit Sub
    pCargarFrmLista E_INSUMO, 2
End Sub

Private Sub mn_insumo_tprod_2_Click()
    ' CONSULTA DE INSUMO/TODO PRODUCTO/DIA ACTUAL
    If fValidarSeleccionDia() = False Then Exit Sub
    If fCargarDatosArray(True) = False Then Exit Sub
    pCargarFrmLista E_INSUMO, 3
End Sub

Private Function fValidarSeleccionDia() As Boolean
    If fValidarDatos() = False Then Exit Function
    
    If RstFrm.State = 0 Then Exit Function
    
    If RstFrm.RecordCount = 0 Then Exit Function
    
    If Fg1.Row >= 9 Or Fg1.Row = Fg1.Rows - 1 Or Fg1.Col = 0 Or Fg1.Col >= Fg1.Cols - 2 Then
        MsgBox "Selecione una Celda Correcta...", vbInformation, xTitulo
        Exit Function
    End If
    
    fValidarSeleccionDia = True
End Function

'*****************************************************************************************************
'* Nombre           : pCargarFrmLista
'* Tipo             : PROCEDIMIENTO
'* Descripcion      :
'* Paranetros       : NOMBRE        |  TIPO          |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    VENTANA       |  e_PROGRAMA    |
'*                    ESTILO_VISTA  |  Integer       |
'*                    ID_RECETA     |  String        |
'* Devuelve         :
'*****************************************************************************************************
Private Sub pCargarFrmLista(VENTANA As e_PROGRAMA, _
                                        ESTILO_VISTA As Integer, _
                                        Optional ID_RECETA As String = "-1")
    Dim N_PERIODO As String
                                        
    If RstFrm.State = 0 Then Exit Sub
    
    If RstFrm.RecordCount = 0 Then Exit Sub
    
    N_PERIODO = "Del: " + Format(TxtFecha(1).valor, FORMAT_DATE) + " Al: " + Format(TxtFecha(2).valor, FORMAT_DATE)
    
    If TxtFecha(1).valor = TxtFecha(2).valor Then N_PERIODO = "Al: " + Format(TxtFecha(2).valor, FORMAT_DATE)
    
    With FrmManProduccionPrograma_lista
        .RECIBE_LINK_FRM CStr(RstFrm.Fields("id")), ID_RECETA, VENTANA, ARR_TMP(), ESTILO_VISTA, Format(TxtFecha(0).valor, FORMAT_DATE), Trim(lbl_cb(0).Caption), N_PERIODO
        .Show
    End With
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarGrid
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS DATOS DE LA TABAL pro_programa EN EL CONTROL Dg3
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pCargarGrid()
    On Error GoTo error
    Dim nSQL  As String
    
    lblperiodo(0).Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    lblperiodo(1).Caption = lblperiodo(0).Caption
    
    nSQL = "SELECT pro_programa.id, pro_programa.num,pro_programa.fecha, pro_programa.idprog, pla_empleados.numdoc AS prognum, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] as prog, pro_programa.fchini, pro_programa.fchfin, pro_programa.obs " _
        + vbCr + " FROM (pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) RIGHT JOIN pro_programa ON pro_emp.id = pro_programa.idprog " _
        + vbCr + " WHERE YEAR(pro_programa.fecha)= " & AnoTra & " AND MONTH(pro_programa.fecha)= " & mMesActivo & " ;"
    
    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, nSQL, xCon
    
    Set Dg3.DataSource = RstFrm
    Me.MousePointer = vbDefault
    Exit Sub

error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pCargarGrid"
End Sub

'*****************************************************************************************************
'* Nombre           : CambiarMes
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CAMBIA EL MES DE TRABAJO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub CambiarMes()
    TabOne1.CurrTab = 0
    mMesActivo = SeleccionaMes(xCon)
    
    If mMesActivo = 0 Or mMesActivo = 13 Then
        MsgBox "Selecione un Periodo Correcto", vbExclamation, xTitulo
        CambiarMes
        Exit Sub
    End If
    
    pCargarGrid
End Sub

Private Sub cb_Click(Index As Integer)
    Dim xRs As New ADODB.Recordset
    Dim xCampos(1, 4) As String
    Dim N_SQL As String
    
    xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "apenom":   xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
        
    N_SQL = "SELECT  pla_empleados.numdoc, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS apenom , pro_emp.id " _
        + vbCr + " FROM (pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper " _
        + vbCr + " WHERE (((pro_empdet.idfun)=2)); "
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, N_SQL, xCampos(), "Buscando Personal", "apenom", "apenom", Principio

    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR

    txt_cb(Index) = xRs.Fields(0) & ""               ' TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & ""       ' NOMBRE
    lbl_cb_cod(Index).Caption = xRs.Fields(2) & ""   ' CODIGO

SALIR:
    Set xRs = Nothing
    Exit Sub

error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub

Private Sub txt_cb_Change(Index As Integer)
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cb_cod(Index).Caption = ""
        lbl_cb(Index).Tag = ""
    End If
End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo error
    
    If txt_cb(Index).Locked = True Then Exit Sub
    
    If KeyCode = vbKeyF5 Then
        cb_Click Index
        Exit Sub
    End If

    If txt_cb(Index).Text = "" Then Exit Sub
    
    If KeyCode <> 13 Then Exit Sub
    
    Dim RST_TMP As New ADODB.Recordset
    Dim N_SQL As String

    N_SQL = "SELECT pla_empleados.numdoc, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS apenom, pro_emp.id " _
        + vbCr + " FROM (pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper " _
        + vbCr + " WHERE (((pro_empdet.idfun)=2)) and pla_empleados.numdoc ='" + Trim(txt_cb(Index).Text) + "'" _
        + vbCr + " ORDER BY pla_empleados.apepat; "

    If xCon.State = 0 Then Exit Sub
    RST_Busq RST_TMP, N_SQL, xCon
    
    If RST_TMP.State = 0 Then Exit Sub
    
    If RST_TMP.RecordCount > 0 Then
        txt_cb(Index) = RST_TMP.Fields(0) & ""             ' TEXTO A MOSTRAR
        lbl_cb(Index).Caption = RST_TMP.Fields(1) & ""     ' NOMBRE
        lbl_cb_cod(Index).Caption = RST_TMP.Fields(2) & "" ' CODIGO
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cb_cod(Index).Caption = ""
    End If
    
    Set RST_TMP = Nothing
    Exit Sub

error:
    Set RST_TMP = Nothing
    SHOW_ERROR Me.Name, "txt_cb_KeyDown(" + CStr(Index) + ")"
End Sub

Private Sub txt_cb_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
    
    Select Case Index
        Case 3: If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case Else: If validar_numero(KeyAscii) = True And KeyAscii = 46 Then KeyAscii = 0
    End Select
End Sub

'*****************************************************************************************************
'* Nombre           : Buscar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA LA BUSQUEFA DE UN PROGRAMA DE PRODUCCION
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Buscar()
    On Error GoTo error
    TabOne1.CurrTab = 0
     
    Dim xRs As New ADODB.Recordset
    Dim N_SQL As String
   
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Fecha":            xCampos(0, 1) = "fecha":       xCampos(0, 2) = "850":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Num":              xCampos(1, 1) = "num":         xCampos(1, 2) = "900":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Programado Por":   xCampos(2, 1) = "prog":        xCampos(2, 2) = "2500":   xCampos(2, 3) = "C"
    xCampos(3, 0) = "Fch.Inicio":       xCampos(3, 1) = "fchini":      xCampos(3, 2) = "900":   xCampos(3, 3) = "C"
    xCampos(4, 0) = "Fch.Fin":          xCampos(4, 1) = "fchfin":      xCampos(4, 2) = "900":   xCampos(4, 3) = "C"
        
    N_SQL = "SELECT pro_programa.id, pro_programa.num, pro_programa.fecha, pro_programa.idprog, pla_empleados.numdoc AS prognum, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS prog, pro_programa.fchini, pro_programa.fchfin, pro_programa.obs " _
    + vbCr + " FROM (pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) RIGHT JOIN pro_programa ON pro_emp.id = pro_programa.idprog " _
    + vbCr + " WHERE YEAR(pro_programa.fecha)= " + AnoTra + " AND MONTH(pro_programa.fecha)= " + CStr(mMesActivo) + " ;"
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, N_SQL, xCampos(), "Buscando Producción", "num", "fecha", Principio
    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    
    RstFrm.MoveFirst
    RstFrm.Find "id = " + CStr(xRs("id"))

SALIR:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub

'*****************************************************************************************************
'* Nombre           : Filtrar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA UN FILTRO SOBRE EL RECORDSET RstFrm
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub Filtrar()
    TabOne1.CurrTab = 0
    
    Dim xCampos(5, 4) As String
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Programa N°":      xCampos(0, 1) = "num":      xCampos(0, 2) = "C":         xCampos(0, 3) = "800"
    xCampos(1, 0) = "Fecha":            xCampos(1, 1) = "fecha":    xCampos(1, 2) = "F":         xCampos(1, 3) = "1000"
    xCampos(2, 0) = "Programado Por.":  xCampos(2, 1) = "prog":     xCampos(2, 2) = "C":         xCampos(2, 3) = "1500"
    xCampos(3, 0) = "Fch. Inicio":      xCampos(3, 1) = "fchini":   xCampos(3, 2) = "F":         xCampos(3, 3) = "1000"
    xCampos(4, 0) = "Fch. Fin":         xCampos(4, 1) = "fchfin":   xCampos(4, 2) = "F":         xCampos(4, 3) = "1000"
    
    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg3
End Sub

'*****************************************************************************************************
'* Nombre           : pExportar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTA A EXCEL LOS DATOS DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pExportar()
    If IsDate(TxtFecha(0).valor) = False Then
        MsgBox "Falta ingresar la Fecha Inicial", vbExclamation, xTitulo
        TxtFecha(0).SetFocus
        Exit Sub
    ElseIf IsDate(TxtFecha(1).valor) = False Then
        MsgBox "Falta ingresar la Fecha Final", vbExclamation, xTitulo
        TxtFecha(1).SetFocus
        Exit Sub
    End If
    
    On Error GoTo error
    
    Dim oExport As New SGI2_funciones.formularios
    Dim nTitulo As String
    Dim nPeriodo As String
    Dim nTitulo1 As String
    
    nTitulo = "PROGRAMA DE PRODUCCION - " & txt(2).Text
    nPeriodo = "Del " & TxtFecha(0).valor & " Al " & TxtFecha(1).valor
    If NulosN(lbl_cb_cod(0).Caption) <> 0 Then nTitulo1 = "Programado Por : " & StrConv(lbl_cb(0).Caption, 3)

    Me.MousePointer = vbHourglass
    oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg1, nTitulo, nPeriodo, nTitulo1, "Programa de Producción"
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
    
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Exportar"
End Sub

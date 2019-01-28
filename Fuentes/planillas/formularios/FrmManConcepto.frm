VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmManConcepto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Concepto"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraEditor 
      BorderStyle     =   0  'None
      Height          =   5265
      Left            =   11985
      TabIndex        =   64
      Top             =   855
      Visible         =   0   'False
      Width           =   9210
      Begin VB.CommandButton cb 
         Height          =   225
         Index           =   4
         Left            =   1875
         Picture         =   "FrmManConcepto.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Seleccione la Categoría de Concepto"
         Top             =   465
         Width           =   210
      End
      Begin VB.CommandButton cb 
         Height          =   225
         Index           =   5
         Left            =   1875
         Picture         =   "FrmManConcepto.frx":0132
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Seleccione el Tipo de Concepto (Primero seleccione la categoría)"
         Top             =   795
         Width           =   210
      End
      Begin VB.CommandButton CmdEditor 
         Caption         =   "&Cancelar"
         Height          =   420
         Index           =   1
         Left            =   5205
         TabIndex        =   67
         Top             =   4755
         Width           =   1275
      End
      Begin VB.CommandButton CmdEditor 
         Caption         =   "&Grabar"
         Height          =   405
         Index           =   0
         Left            =   3735
         TabIndex        =   66
         Top             =   4755
         Width           =   1275
      End
      Begin VB.PictureBox pic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   8940
         Picture         =   "FrmManConcepto.frx":0264
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   65
         ToolTipText     =   "Cerrar"
         Top             =   90
         Width           =   195
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   5
         Left            =   1350
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   72
         Tag             =   "null"
         Text            =   "txt_cb(5)"
         Top             =   765
         Width           =   765
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   4
         Left            =   1350
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   71
         Tag             =   "null"
         Text            =   "txt_cb(4)"
         Top             =   435
         Width           =   765
      End
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   3435
         Index           =   2
         Left            =   105
         TabIndex        =   79
         Top             =   1155
         Width           =   9015
         _cx             =   15901
         _cy             =   6059
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
         SelectionMode   =   1
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
         FormatString    =   $"FrmManConcepto.frx":0550
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
      Begin VB.Label lbl_capt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoría"
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   76
         Top             =   540
         Width           =   705
      End
      Begin VB.Label lbl_cod 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cod(4)"
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
         Index           =   4
         Left            =   3870
         TabIndex        =   75
         Top             =   435
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lbl_cod 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cod(5)"
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
         Index           =   5
         Left            =   3870
         TabIndex        =   74
         Top             =   765
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lbl_capt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         Height          =   195
         Index           =   5
         Left            =   165
         TabIndex        =   73
         Top             =   855
         Width           =   315
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         Index           =   5
         X1              =   120
         X2              =   9030
         Y1              =   4665
         Y2              =   4665
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   4
         X1              =   0
         X2              =   15
         Y1              =   15
         Y2              =   6800
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   3
         X1              =   30
         X2              =   12000
         Y1              =   5250
         Y2              =   5250
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -330
         X2              =   12000
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   9195
         X2              =   9195
         Y1              =   30
         Y2              =   6770
      End
      Begin VB.Label LblTituloFrame 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "Activar Conceptos"
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
         Left            =   90
         TabIndex        =   68
         Top             =   90
         Width           =   1575
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   300
         Index           =   1
         Left            =   30
         Top             =   45
         Width           =   10275
      End
      Begin VB.Label lbl_cb 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cb(5)"
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
         Index           =   5
         Left            =   2115
         TabIndex        =   77
         Top             =   765
         Width           =   5295
      End
      Begin VB.Label lbl_cb 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cb(4)"
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
         Index           =   4
         Left            =   2115
         TabIndex        =   78
         Top             =   435
         Width           =   3495
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8235
      Top             =   45
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
            Picture         =   "FrmManConcepto.frx":05EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConcepto.frx":0B2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConcepto.frx":0CB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConcepto.frx":1106
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConcepto.frx":121E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConcepto.frx":1762
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConcepto.frx":1CA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConcepto.frx":1DBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConcepto.frx":1ECE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConcepto.frx":2322
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConcepto.frx":248E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConcepto.frx":29D6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Activar/Desactivar Concepto"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agregar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7215
      Left            =   30
      TabIndex        =   11
      Top             =   360
      Width           =   11835
      _cx             =   20876
      _cy             =   12726
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
      Appearance      =   2
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   12632256
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   12632256
      TabOutlineColor =   0
      FrontTabForeColor=   -2147483630
      Caption         =   "   Consulta   |   Detalles   "
      Align           =   0
      CurrTab         =   1
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
         Height          =   6795
         Left            =   -12390
         TabIndex        =   14
         Top             =   375
         Width           =   11745
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6405
            Left            =   45
            TabIndex        =   15
            Top             =   390
            Width           =   11550
            _ExtentX        =   20373
            _ExtentY        =   11298
            _LayoutType     =   4
            _RowHeight      =   14
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Id"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "CodSunat"
            Columns(1).DataField=   "codsun"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Concepto"
            Columns(2).DataField=   "descripcion"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Variable"
            Columns(3).DataField=   "variable"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Tipo"
            Columns(4).DataField=   "tiponombre"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Categoría"
            Columns(5).DataField=   "catnombre"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Fórmula"
            Columns(6).DataField=   "formula"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=926"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=2"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1667"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1588"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=1"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=6773"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=6694"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=3201"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=3122"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=4789"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=4710"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(30)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(33)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(34)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(35)=   "Column(6).Width=10028"
            Splits(0)._ColumnProps(36)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(37)=   "Column(6)._WidthInPix=9948"
            Splits(0)._ColumnProps(38)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(39)=   "Column(6).Order=7"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=38,.bgcolor=&HDBFDFD&,.fgcolor=&H0&,.bold=0"
            _StyleDefs(7)   =   ":id=1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.fgcolor=&HFF0000&,.bold=0"
            _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=38,.bgcolor=&H80&"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.bold=0,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(25)  =   ":id=13,.strikethrough=0,.charset=0"
            _StyleDefs(26)  =   ":id=13,.fontname=MS Sans Serif"
            _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.namedParent=33,.fgcolor=&H800000&"
            _StyleDefs(29)  =   ":id=14,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(30)  =   ":id=14,.fontname=MS Sans Serif"
            _StyleDefs(31)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(32)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(33)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H80&,.bold=0"
            _StyleDefs(34)  =   ":id=18,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(35)  =   ":id=18,.fontname=MS Sans Serif"
            _StyleDefs(36)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(37)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.bgcolor=&H80&"
            _StyleDefs(38)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(39)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(40)  =   ":id=21,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(41)  =   ":id=21,.fontname=MS Sans Serif"
            _StyleDefs(42)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(43)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(44)  =   "Splits(0).Columns(0).Style:id=58,.parent=13,.alignment=1"
            _StyleDefs(45)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(53)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14,.alignment=2"
            _StyleDefs(54)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(3).Style:id=62,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(6).Style:id=50,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
            _StyleDefs(72)  =   "Named:id=33:Normal"
            _StyleDefs(73)  =   ":id=33,.parent=0"
            _StyleDefs(74)  =   "Named:id=34:Heading"
            _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(76)  =   ":id=34,.wraptext=-1"
            _StyleDefs(77)  =   "Named:id=35:Footing"
            _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(79)  =   "Named:id=36:Selected"
            _StyleDefs(80)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(81)  =   "Named:id=37:Caption"
            _StyleDefs(82)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(83)  =   "Named:id=38:HighlightRow"
            _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(85)  =   "Named:id=39:EvenRow"
            _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(87)  =   "Named:id=40:OddRow"
            _StyleDefs(88)  =   ":id=40,.parent=33"
            _StyleDefs(89)  =   "Named:id=41:RecordSelector"
            _StyleDefs(90)  =   ":id=41,.parent=34"
            _StyleDefs(91)  =   "Named:id=42:FilterBar"
            _StyleDefs(92)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Conceptos"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   105
            TabIndex        =   16
            Top             =   60
            Width           =   11400
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6795
         Left            =   45
         TabIndex        =   12
         Top             =   375
         Width           =   11745
         Begin VB.TextBox txt 
            BackColor       =   &H0000FFFF&
            Height          =   285
            Index           =   0
            Left            =   10665
            TabIndex        =   19
            Tag             =   "null"
            Text            =   "txt(0)"
            Top             =   -15
            Visible         =   0   'False
            Width           =   945
         End
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   6315
            Left            =   -15
            TabIndex        =   17
            Top             =   420
            Width           =   11700
            _cx             =   20637
            _cy             =   11139
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
            Caption         =   "  Datos Principales  | Datos Contables | Conceptos Afecto a Impuesto  |   Lista de Conceptos - Aportes     "
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
               Height          =   5895
               Left            =   12945
               TabIndex        =   45
               Top             =   45
               Width           =   11610
               Begin VB.Frame fra 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame12"
                  Height          =   585
                  Index           =   3
                  Left            =   120
                  TabIndex        =   49
                  Top             =   5220
                  Width           =   11340
                  Begin VB.CommandButton CmdDet 
                     Caption         =   "&Eliminar"
                     Enabled         =   0   'False
                     Height          =   435
                     Index           =   1
                     Left            =   3090
                     TabIndex        =   52
                     Top             =   60
                     Width           =   1395
                  End
                  Begin VB.CommandButton CmdDet 
                     Caption         =   "&Agregar"
                     Enabled         =   0   'False
                     Height          =   435
                     Index           =   0
                     Left            =   120
                     TabIndex        =   51
                     Top             =   60
                     Width           =   1395
                  End
                  Begin VB.CommandButton CmdDet 
                     Caption         =   "&Seleccionar"
                     Enabled         =   0   'False
                     Height          =   435
                     Index           =   2
                     Left            =   1545
                     TabIndex        =   50
                     Top             =   60
                     Width           =   1395
                  End
                  Begin VB.Line lin 
                     BorderColor     =   &H80000009&
                     BorderWidth     =   2
                     Index           =   7
                     X1              =   -15
                     X2              =   13000
                     Y1              =   15
                     Y2              =   15
                  End
                  Begin VB.Line lin 
                     BorderColor     =   &H00808080&
                     BorderWidth     =   2
                     Index           =   6
                     X1              =   -30
                     X2              =   13000
                     Y1              =   570
                     Y2              =   570
                  End
                  Begin VB.Line Line1 
                     BorderColor     =   &H00808080&
                     BorderWidth     =   2
                     Index           =   3
                     X1              =   11325
                     X2              =   11325
                     Y1              =   0
                     Y2              =   985
                  End
                  Begin VB.Line Line5 
                     BorderColor     =   &H00FFFFFF&
                     BorderWidth     =   2
                     X1              =   15
                     X2              =   15
                     Y1              =   0
                     Y2              =   1000
                  End
               End
               Begin VB.Frame fra 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame12"
                  Height          =   405
                  Index           =   2
                  Left            =   120
                  TabIndex        =   46
                  Top             =   45
                  Width           =   11340
                  Begin VB.Line lin 
                     BorderColor     =   &H80000009&
                     BorderWidth     =   2
                     Index           =   5
                     X1              =   -15
                     X2              =   12000
                     Y1              =   15
                     Y2              =   15
                  End
                  Begin VB.Line lin 
                     BorderColor     =   &H00808080&
                     BorderWidth     =   2
                     Index           =   4
                     X1              =   -30
                     X2              =   12000
                     Y1              =   390
                     Y2              =   390
                  End
                  Begin VB.Line Line1 
                     BorderColor     =   &H00808080&
                     BorderWidth     =   2
                     Index           =   2
                     X1              =   11325
                     X2              =   11325
                     Y1              =   15
                     Y2              =   395
                  End
                  Begin VB.Line Line2 
                     BorderColor     =   &H00FFFFFF&
                     BorderWidth     =   2
                     Index           =   2
                     X1              =   15
                     X2              =   15
                     Y1              =   0
                     Y2              =   380
                  End
                  Begin VB.Label lbl_cabecera 
                     AutoSize        =   -1  'True
                     Caption         =   "lbl_cabecera(2)"
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
                     Index           =   2
                     Left            =   75
                     TabIndex        =   47
                     Top             =   75
                     Width           =   1335
                  End
               End
               Begin VSFlex7Ctl.VSFlexGrid fg 
                  Height          =   4680
                  Index           =   1
                  Left            =   120
                  TabIndex        =   48
                  Top             =   495
                  Width           =   11325
                  _cx             =   19976
                  _cy             =   8255
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   2
                  Cols            =   2
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmManConcepto.frx":2CF0
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
            Begin VB.Frame Frame5 
               BorderStyle     =   0  'None
               Height          =   5895
               Left            =   12645
               TabIndex        =   39
               Top             =   45
               Width           =   11610
               Begin VB.Frame fra 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame12"
                  Height          =   405
                  Index           =   0
                  Left            =   120
                  TabIndex        =   41
                  Top             =   45
                  Width           =   11340
                  Begin VB.Label lbl_cabecera 
                     AutoSize        =   -1  'True
                     Caption         =   "lbl_cabecera(0)"
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
                     Index           =   0
                     Left            =   75
                     TabIndex        =   42
                     Top             =   75
                     Width           =   1335
                  End
                  Begin VB.Line Line2 
                     BorderColor     =   &H00FFFFFF&
                     BorderWidth     =   2
                     Index           =   0
                     X1              =   15
                     X2              =   15
                     Y1              =   0
                     Y2              =   380
                  End
                  Begin VB.Line Line1 
                     BorderColor     =   &H00808080&
                     BorderWidth     =   2
                     Index           =   0
                     X1              =   11325
                     X2              =   11325
                     Y1              =   15
                     Y2              =   395
                  End
                  Begin VB.Line lin 
                     BorderColor     =   &H00808080&
                     BorderWidth     =   2
                     Index           =   1
                     X1              =   -30
                     X2              =   12000
                     Y1              =   390
                     Y2              =   390
                  End
                  Begin VB.Line lin 
                     BorderColor     =   &H80000009&
                     BorderWidth     =   2
                     Index           =   0
                     X1              =   -15
                     X2              =   12000
                     Y1              =   15
                     Y2              =   15
                  End
               End
               Begin VSFlex7Ctl.VSFlexGrid fg 
                  Height          =   5205
                  Index           =   0
                  Left            =   120
                  TabIndex        =   40
                  Top             =   495
                  Width           =   11325
                  _cx             =   19976
                  _cy             =   9181
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   2
                  Cols            =   2
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmManConcepto.frx":2D2C
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
               Caption         =   "Frame4"
               Height          =   5895
               Left            =   12345
               TabIndex        =   38
               Top             =   45
               Width           =   11610
               Begin VB.Frame Frame8 
                  Height          =   4920
                  Left            =   135
                  TabIndex        =   53
                  Top             =   675
                  Width           =   11340
                  Begin VB.CommandButton cb 
                     Height          =   225
                     Index           =   2
                     Left            =   3045
                     Picture         =   "FrmManConcepto.frx":2D68
                     Style           =   1  'Graphical
                     TabIndex        =   55
                     ToolTipText     =   "Seleccione la Cuenta Contable Debe"
                     Top             =   375
                     Width           =   210
                  End
                  Begin VB.CommandButton cb 
                     Height          =   225
                     Index           =   3
                     Left            =   3045
                     Picture         =   "FrmManConcepto.frx":2E9A
                     Style           =   1  'Graphical
                     TabIndex        =   54
                     ToolTipText     =   "Seleccione la Cuenta Contable Haber"
                     Top             =   780
                     Width           =   210
                  End
                  Begin VB.TextBox txt_cb 
                     Height          =   300
                     Index           =   2
                     Left            =   1830
                     Locked          =   -1  'True
                     MaxLength       =   20
                     TabIndex        =   56
                     Tag             =   "null"
                     Text            =   "txt_cb(2)"
                     Top             =   345
                     Width           =   1455
                  End
                  Begin VB.TextBox txt_cb 
                     Height          =   300
                     Index           =   3
                     Left            =   1830
                     Locked          =   -1  'True
                     MaxLength       =   20
                     TabIndex        =   57
                     Tag             =   "null"
                     Text            =   "txt_cb(3)"
                     Top             =   750
                     Width           =   1455
                  End
                  Begin VB.Label lbl_cb 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "lbl_cb(2)"
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
                     Index           =   2
                     Left            =   3285
                     TabIndex        =   63
                     Top             =   345
                     Width           =   5295
                  End
                  Begin VB.Label lbl_cb 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "lbl_cb(3)"
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
                     Index           =   3
                     Left            =   3285
                     TabIndex        =   62
                     Top             =   750
                     Width           =   5295
                  End
                  Begin VB.Label lbl_capt 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cta.Contable Debe"
                     Height          =   195
                     Index           =   2
                     Left            =   285
                     TabIndex        =   61
                     Top             =   435
                     Width           =   1350
                  End
                  Begin VB.Label lbl_cod 
                     BackColor       =   &H000000FF&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "lbl_cod(2)"
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
                     Index           =   2
                     Left            =   4500
                     TabIndex        =   60
                     Top             =   345
                     Visible         =   0   'False
                     Width           =   975
                  End
                  Begin VB.Label lbl_cod 
                     BackColor       =   &H000000FF&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "lbl_cod(3)"
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
                     Index           =   3
                     Left            =   4500
                     TabIndex        =   59
                     Top             =   750
                     Visible         =   0   'False
                     Width           =   975
                  End
                  Begin VB.Label lbl_capt 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cta.Contable Haber"
                     Height          =   195
                     Index           =   3
                     Left            =   270
                     TabIndex        =   58
                     Top             =   840
                     Width           =   1395
                  End
               End
               Begin VB.Frame fra 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame12"
                  Height          =   405
                  Index           =   1
                  Left            =   135
                  TabIndex        =   43
                  Top             =   60
                  Width           =   11340
                  Begin VB.Line lin 
                     BorderColor     =   &H80000009&
                     BorderWidth     =   2
                     Index           =   3
                     X1              =   -15
                     X2              =   12000
                     Y1              =   15
                     Y2              =   15
                  End
                  Begin VB.Line lin 
                     BorderColor     =   &H00808080&
                     BorderWidth     =   2
                     Index           =   2
                     X1              =   -30
                     X2              =   12000
                     Y1              =   390
                     Y2              =   390
                  End
                  Begin VB.Line Line1 
                     BorderColor     =   &H00808080&
                     BorderWidth     =   2
                     Index           =   1
                     X1              =   11325
                     X2              =   11325
                     Y1              =   15
                     Y2              =   395
                  End
                  Begin VB.Line Line2 
                     BorderColor     =   &H00FFFFFF&
                     BorderWidth     =   2
                     Index           =   1
                     X1              =   15
                     X2              =   15
                     Y1              =   0
                     Y2              =   380
                  End
                  Begin VB.Label lbl_cabecera 
                     AutoSize        =   -1  'True
                     Caption         =   "lbl_cabecera(0)"
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
                     Index           =   1
                     Left            =   75
                     TabIndex        =   44
                     Top             =   75
                     Width           =   1335
                  End
               End
            End
            Begin VB.Frame Frame3 
               BorderStyle     =   0  'None
               Caption         =   "Frame3"
               Height          =   5895
               Left            =   45
               TabIndex        =   18
               Top             =   45
               Width           =   11610
               Begin VB.TextBox txt 
                  Height          =   315
                  Index           =   4
                  Left            =   6135
                  Locked          =   -1  'True
                  MaxLength       =   4
                  TabIndex        =   5
                  Tag             =   "null"
                  Text            =   "txt(4)"
                  Top             =   1830
                  Width           =   1350
               End
               Begin VB.TextBox txt 
                  Height          =   315
                  Index           =   3
                  Left            =   1425
                  Locked          =   -1  'True
                  MaxLength       =   30
                  TabIndex        =   4
                  Tag             =   "null"
                  Text            =   "txt(3)"
                  Top             =   1815
                  Width           =   3585
               End
               Begin VB.TextBox txt 
                  BackColor       =   &H8000000F&
                  Height          =   315
                  Index           =   2
                  Left            =   1425
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   3
                  Text            =   "txt(2)"
                  Top             =   1440
                  Width           =   4245
               End
               Begin VB.TextBox txt 
                  Height          =   315
                  Index           =   1
                  Left            =   1425
                  Locked          =   -1  'True
                  MaxLength       =   200
                  TabIndex        =   2
                  Text            =   "txt(1)"
                  Top             =   1065
                  Width           =   9510
               End
               Begin VB.Frame fra_APlanilla 
                  Caption         =   "Seleccionar"
                  Enabled         =   0   'False
                  Height          =   645
                  Left            =   240
                  TabIndex        =   34
                  Top             =   2235
                  Width           =   5760
                  Begin VB.OptionButton opt_planilla 
                     Caption         =   "No Considerar en Planilla"
                     Height          =   225
                     Index           =   1
                     Left            =   3180
                     TabIndex        =   35
                     Top             =   300
                     Width           =   2100
                  End
                  Begin VB.OptionButton opt_planilla 
                     Caption         =   "Considerar en Planilla"
                     Height          =   225
                     Index           =   0
                     Left            =   420
                     TabIndex        =   6
                     Top             =   300
                     Value           =   -1  'True
                     Width           =   1905
                  End
               End
               Begin VB.CommandButton cb 
                  Height          =   225
                  Index           =   1
                  Left            =   1950
                  Picture         =   "FrmManConcepto.frx":2FCC
                  Style           =   1  'Graphical
                  TabIndex        =   27
                  ToolTipText     =   "Seleccione el Tipo de Concepto (Primero seleccione la categoría)"
                  Top             =   750
                  Width           =   210
               End
               Begin VB.CommandButton cb 
                  Height          =   225
                  Index           =   0
                  Left            =   1950
                  Picture         =   "FrmManConcepto.frx":30FE
                  Style           =   1  'Graphical
                  TabIndex        =   26
                  ToolTipText     =   "Seleccione la Categoría de Concepto"
                  Top             =   420
                  Width           =   210
               End
               Begin VB.TextBox txt 
                  Height          =   870
                  Index           =   5
                  Left            =   240
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   9
                  Tag             =   "null"
                  Text            =   "FrmManConcepto.frx":3230
                  Top             =   4920
                  Width           =   10800
               End
               Begin VB.Frame Frame6 
                  BeginProperty Font 
                     Name            =   "Trebuchet MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1575
                  Left            =   240
                  TabIndex        =   21
                  Top             =   2910
                  Width           =   10815
                  Begin VB.CheckBox chk_formula 
                     Caption         =   "Fórmula"
                     Enabled         =   0   'False
                     Height          =   225
                     Left            =   150
                     TabIndex        =   7
                     Top             =   45
                     Width           =   885
                  End
                  Begin VB.CommandButton cmd_formula 
                     Caption         =   "Editar Formula"
                     Enabled         =   0   'False
                     Height          =   645
                     Left            =   90
                     Picture         =   "FrmManConcepto.frx":3239
                     Style           =   1  'Graphical
                     TabIndex        =   8
                     Top             =   735
                     Width           =   1275
                  End
                  Begin VB.TextBox txt_formula 
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   1095
                     Left            =   1455
                     Locked          =   -1  'True
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   22
                     Text            =   "FrmManConcepto.frx":333B
                     Top             =   270
                     Width           =   9045
                  End
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   0
                  Left            =   1425
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   0
                  Text            =   "txt_cb(0)"
                  Top             =   390
                  Width           =   765
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   1
                  Left            =   1425
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   1
                  Text            =   "txt_cb(1)"
                  Top             =   720
                  Width           =   765
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cod.Sunat"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   4
                  Left            =   5235
                  TabIndex        =   37
                  ToolTipText     =   "Es nombre abreviado del concepto, con posibilidad de que aparesca en el reporte"
                  Top             =   1950
                  Width           =   750
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nombre Corto"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   3
                  Left            =   240
                  TabIndex        =   36
                  ToolTipText     =   "Es nombre abreviado del concepto, con posibilidad de que aparesca en el reporte"
                  Top             =   1935
                  Width           =   975
               End
               Begin VB.Label lbl_capt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tipo"
                  Height          =   195
                  Index           =   1
                  Left            =   240
                  TabIndex        =   33
                  Top             =   810
                  Width           =   315
               End
               Begin VB.Label lbl_cod 
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cod(1)"
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
                  Index           =   1
                  Left            =   3945
                  TabIndex        =   31
                  Top             =   720
                  Visible         =   0   'False
                  Width           =   975
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
                  Left            =   3945
                  TabIndex        =   29
                  Top             =   390
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.Label lbl_capt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Categoría"
                  Height          =   195
                  Index           =   0
                  Left            =   240
                  TabIndex        =   28
                  Top             =   495
                  Width           =   705
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Comentario :"
                  Height          =   210
                  Index           =   5
                  Left            =   240
                  TabIndex        =   25
                  Top             =   4665
                  Width           =   900
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Variable"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   2
                  Left            =   240
                  TabIndex        =   24
                  Top             =   1575
                  Width           =   570
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nombre"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   1
                  Left            =   240
                  TabIndex        =   23
                  Top             =   1200
                  Width           =   555
               End
               Begin VB.Label lbl_cb 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb(1)"
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
                  Index           =   1
                  Left            =   2190
                  TabIndex        =   32
                  Top             =   720
                  Width           =   5295
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
                  Left            =   2190
                  TabIndex        =   30
                  Top             =   390
                  Width           =   3495
               End
            End
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            Height          =   195
            Index           =   0
            Left            =   10125
            TabIndex        =   20
            Top             =   75
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Conceptos"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   90
            TabIndex        =   13
            Top             =   60
            Width           =   11400
         End
      End
   End
End
Attribute VB_Name = "FrmManConcepto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Dim RstFrm As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim QueHace As Integer
Dim Agregando As Boolean

Dim fOrdenLista As Boolean ''--especifica el orden de la lista de la consulta

Sub Cancelar()
    
    pHabilitarBotonEditor False
    
    Toolbar1.Buttons(7).Enabled = False
    Toolbar1.Buttons(8).Enabled = False
    
    Label5.Caption = "Detalle del Concepto"
    QueHace = 3
    Bloquea False
    TabOne1.TabEnabled(0) = True
    Dim K&
    For K = 1 To TabOne2.NumTabs - 1
        TabOne2.TabEnabled(K) = True
    Next K
    TabOne1.CurrTab = 0
End Sub

Private Sub chk_formula_Click()
    If chk_formula.Value = 1 Then
        cmd_formula.Enabled = True
        txt_formula.Text = txt_formula.Tag
    Else
        cmd_formula.Enabled = False
        txt_formula.Text = ""
    End If
End Sub

Private Sub chk_formula_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab

End Sub

Private Sub cmd_formula_Click()
    FrmManConceptoFormula.Show 1
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If ColIndex = 5 Then Exit Sub
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstFrm.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
    TabOne2.TabVisible(2) = False
    TabOne2.TabVisible(3) = True
    SeEjecuto = False
    pCargarGrid
    pConfigurarGrilla
    SeEjecuto = True
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.RecordCount = 0 Then
        If MsgBox("No se ha registrado Concepto, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
            nuevo
        End If
    End If
    
    
End Sub


Private Sub Form_Load()
    CentrarFrm Me
    SeEjecuto = False
    QueHace = 3
    TabOne1.CurrTab = 0
End Sub


Sub Blanquea()
    txt_formula.Text = ""
    txt_formula.Tag = ""
    LimpiaText txt
    LimpiaText txt_cb
    LimpiaText lbl_cb
    LimpiaText lbl_cod
End Sub

Sub Bloquea(band As Boolean)

    habilitar_Locked txt, Not band
    habilitar_Locked txt_cb, Not band
    
    chk_formula.Enabled = band
    fra_APlanilla.Enabled = band
    habilitar CmdDet, band
    If (QueHace = 1) Or (QueHace = 2 And NulosC(txt(2).Text) = "") Then
        txt(2).Enabled = True
        txt(2).BackColor = vbWhite
    Else
        Dim RstTmp As New ADODB.Recordset
        Dim nSQL As String
        
        nSQL = "SELECT pla_concepto.id FROM pla_concepto WHERE ucase(pla_concepto.formula) Like '%" & UCase(NulosC(RstFrm.Fields("variable"))) & "%' ;"
        RST_Busq RstTmp, nSQL, xCon
        If RstTmp.RecordCount <> 0 Then
            txt(2).Enabled = False
            txt(2).BackColor = &H8000000F
        Else
            txt(2).Enabled = True
            txt(2).BackColor = vbWhite
        End If
        Set RstTmp = Nothing
        
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 1
        Exit Sub
    End If
    Set RstFrm = Nothing
End Sub

Private Sub opt_planilla_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab

End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then MuestraSegundoTab
    End If
End Sub

Function Grabar() As Boolean

    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "Grabar", "Modificar") + " al Concepto ", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function

    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset '--relacionado a los aportes que se consideraran
    Dim RstTmp As New ADODB.Recordset '--temporal
    Dim nSQL As String
    Dim xId&, A&
    On Error GoTo LaCague

    xCon.BeginTrans

    If QueHace = 1 Then

        xId = HallaCodigoTabla("pla_concepto", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM pla_concepto", xCon

        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstFrm("id")
        RST_Busq RstCab, "SELECT * FROM pla_concepto WHERE id = " & xId & ";", xCon
        '--eliminar los conceptos relacionado con la formula
        xCon.Execute "DELETE FROM pla_conceptoformula WHERE idcpto = " & xId & ";"
        
        If NulosN(lbl_cod(0).Caption) <> 2 Then
            '--eliminar los conceptos de aportes relacionado con el concepto
            xCon.Execute "DELETE FROM pla_conceptoapo WHERE idcpto = " & xId & ";"
        Else
            '--eliminar los conceptos relacionado con el concepto de aporte
            xCon.Execute "DELETE FROM pla_conceptoapo WHERE idcptoref = " & xId & ";"
        End If
    End If
    RST_Busq RstDet, "SELECT TOP 1 * FROM pla_conceptoapo", xCon
    
    
    RstCab("idtipo") = NulosN(lbl_cod(1).Caption)
    RstCab("codsun") = NulosC(txt(4).Text)
    RstCab("descripcion") = NulosC(txt(1).Text)
    RstCab("variable") = NulosC(txt(2).Text)
    RstCab("formula") = NulosC(txt_formula.Text)
    RstCab("aplanilla") = IIf(opt_planilla(0).Value = True, -1, 0)
    RstCab("nomcorto") = NulosC(txt(3).Text)
    RstCab("observacion") = NulosC(txt(5).Text)
    
    RstCab("idctadeb") = NulosN(lbl_cod(2).Caption)
    RstCab("idctahab") = NulosN(lbl_cod(3).Caption)
    
    RstCab.Update
    
    '--grabar los conceptos usados en la formula
    '--buscar en concepto, tipos de horas, conceptos varios
    If NulosC(txt_formula.Text) <> "" Then
        '--insertando los conceptos que estan relacionado a la formula
        xCon.Execute "INSERT INTO pla_conceptoformula (idcpto,idcptoref,origen) " _
            + vbCr + " SELECT " & xId & " as IdCpto, pla_concepto.id as idref,0 as idtipo " _
            + vbCr + " FROM pla_concepto " _
            + vbCr + " GROUP BY '" & NulosC(txt_formula.Text) & "',pla_concepto.variable, pla_concepto.id " _
            + vbCr + " HAVING ((('" & NulosC(txt_formula.Text) & "') Like '%' & [pla_concepto].[variable] & '%') AND ((pla_concepto.variable) Is Not Null)); "
        
        xCon.Execute "INSERT INTO pla_conceptoformula (idcpto,idcptoref,origen) " _
            + vbCr + " SELECT " & xId & " as IdCpto, mae_tipohora.id as idref, 1 as idtipo " _
            + vbCr + " FROM  mae_tipohora " _
            + vbCr + " GROUP BY '" & NulosC(txt_formula.Text) & "',mae_tipohora.variable, mae_tipohora.id " _
            + vbCr + " HAVING ((('" & NulosC(txt_formula.Text) & "') Like '%' & mae_tipohora.variable & '%') AND ((mae_tipohora.variable) Is Not Null)); "
        
        xCon.Execute "INSERT INTO pla_conceptoformula (idcpto,idcptoref,origen) " _
            + vbCr + " SELECT " & xId & " as IdCpto, pla_conceptovarios.id as idref,2 as idtipo " _
            + vbCr + " FROM pla_conceptovarios " _
            + vbCr + " GROUP BY '" & NulosC(txt_formula.Text) & "',pla_conceptovarios.variable, pla_conceptovarios.id " _
            + vbCr + " HAVING ((('" & NulosC(txt_formula.Text) & "') Like '%' & [pla_conceptovarios].[variable] & '%') AND ((pla_conceptovarios.variable) Is Not Null)); "
            
    End If
    '--
    If NulosN(lbl_cod(0).Caption) <> 2 Then
        '--relacionado a los aportes que se consideraran
        For A = Fg(0).FixedRows To Fg(0).Rows - 1
            DoEvents
            RstDet.AddNew
            RstDet("idcpto") = xId
            RstDet("idcptoref") = NulosN(Fg(0).TextMatrix(A, 1))
            If NulosN(Fg(0).TextMatrix(A, 4)) = -1 Then
                RstDet("activo") = -1
            Else
                RstDet("activo") = 0
            End If
            RstDet.Update
        Next A
    Else
        '--relacionado a los aportes que se consideraran
        For A = Fg(1).FixedRows To Fg(1).Rows - 1
            DoEvents
            RstDet.AddNew
            RstDet("idcpto") = NulosN(Fg(1).TextMatrix(A, 1))
            RstDet("idcptoref") = xId
            RstDet("activo") = -1
            RstDet.Update
        Next
    End If
    xCon.CommitTrans
    Set RstCab = Nothing
    MsgBox "Los datos del Concepto " + IIf(QueHace = 1, "grabaron", "modificaron") + " con éxito", vbInformation, xTitulo

    Grabar = True
    Exit Function
LaCague:
    Set RstCab = Nothing
    xCon.RollbackTrans
    MsgBox "No se pudo guardar al Personal por el siguiente motivo: " + vbCr + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Function

Sub nuevo()
    QueHace = 1
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Bloquea True
    Blanquea
    
    TabOne2.TabVisible(2) = False
    TabOne2.TabVisible(3) = False
    
    Label5.Caption = "Agregando Conceptos"
    TabOne2.CurrTab = 0
    '--cargando la lista de conceptos de aportes
    pCargarDatosDet True
    '-------------------------------------------
    txt_cb(0).SetFocus
    
End Sub

Sub Modificar()
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay registro para modificar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    
    Label5.Caption = "Modificando Concepto"

    ActivaTool
    
    QueHace = 2
    If TabOne1.CurrTab = 0 Then
        TabOne1.CurrTab = 1
    End If
    
    Bloquea True
    
    TabOne1.TabEnabled(0) = False
    
    Agregando = False
    If TabOne2.CurrTab <> 0 Then TabOne2.CurrTab = 0
    txt_cb(0).SetFocus

End Sub

Sub ActivaTool()
    Dim A As Integer
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Sub Eliminar()
    TabOne1.CurrTab = 0
    Dim Rpta As Integer
    Dim xId&
    Dim nSQL As String
    
    Dim RstBus As New ADODB.Recordset
    xId = RstFrm.Fields("id")
    nSQL = "SELECT pla_concepto.id, pla_concepto.descripcion, pla_concepto.variable, pla_concepto.formula " _
        + vbCr + " FROM pla_concepto " _
        + vbCr + " WHERE (((pla_concepto.id)<>" & xId & ") AND ((pla_concepto.formula) Like '%" & RstFrm.Fields("variable") & "%'));"
        
    RST_Busq RstBus, nSQL, xCon
    If RstBus.RecordCount <> 0 Then
        MsgBox "No se puede eliminar el Concepto, figura en fórmulas de otros conceptos" + vbCr + "Ej. " & RstBus.Fields("descripcion") & vbCr & "Eliminar primero el Concepto Mensionado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Set RstBus = Nothing
        Exit Sub
    End If
    Set RstBus = Nothing
    '--falta validar que el concepto no este ya en planilla
    
    
    Rpta = MsgBox("Esta seguro de eliminar al Concepto seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
    
        xCon.Execute "DELETE FROM pla_conceptoproc WHERE idproc = " & xId & ";" '--relacionados a las categorias
        xCon.Execute "DELETE FROM pla_conceptoformula WHERE idcpto = " & xId & ";" '--relacionado a las formulas
        
        xCon.Execute "DELETE FROM pla_conceptoregpen WHERE idcpto = " & xId & ";" '--replaciona al regimen pensionario
        xCon.Execute "DELETE FROM pla_conceptoemp WHERE idcpto = " & xId & ";" '--relacionado a la asignacion de sueldos
        
        xCon.Execute "DELETE FROM pla_concepto WHERE id = " & xId & ";"
        
        MsgBox "El Concepto se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstFrm.Requery
        Dg1.Refresh
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pActivarConcepto
    If Button.Index = 3 Then nuevo
    If Button.Index = 4 Then Modificar
    If Button.Index = 5 Then Eliminar
    If Button.Index = 7 Then Cancelar
    If Button.Index = 8 Then
        If Grabar = True Then
            RstFrm.Requery
            Dg1.Refresh
            Cancelar
        End If
    End If

    If Button.Index = 10 Then Filtrar
    If Button.Index = 11 Then
        RstFrm.Filter = adFilterNone
    End If
    If Button.Index = 12 Then Buscar
    If Button.Index = 16 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub

Sub Buscar()
    TabOne1.CurrTab = 0
    Dim xRs As New ADODB.Recordset
    Dim nSQL  As String
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "CodSunat":    xCampos(0, 1) = "codsun":      xCampos(0, 2) = "800":     xCampos(0, 3) = "C"
    xCampos(1, 0) = "Descripción": xCampos(1, 1) = "descripcion": xCampos(1, 2) = "3200":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Tipo":        xCampos(2, 1) = "tiponombre":  xCampos(2, 2) = "2200":    xCampos(2, 3) = "C"
    xCampos(3, 0) = "Categoría":   xCampos(3, 1) = "catnombre":   xCampos(3, 2) = "2000":    xCampos(3, 3) = "C"
    
    nSQL = "SELECT pla_conceptotipo.idcat, pla_conceptocat.descripcion AS catnombre, pla_conceptotipo.descripcion AS tiponombre, pla_concepto.*, pla_conceptotipo.codsun AS tiposun " _
        + vbCr + " FROM (pla_conceptocat RIGHT JOIN pla_conceptotipo ON pla_conceptocat.id = pla_conceptotipo.idcat) RIGHT JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo " _
        + vbCr + " ORDER BY pla_conceptocat.descripcion desc,pla_concepto.codsun asc, pla_conceptotipo.descripcion, pla_concepto.descripcion; "

    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Conceptos", "descripcion", "descripcion", Principio
    
    If xRs.State = 1 Then
        RstFrm.MoveFirst
        RstFrm.Find "id = " & xRs("id") & ""
    End If
    Set xRs = Nothing
    
End Sub

Sub Filtrar()
    TabOne1.CurrTab = 0
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(4, 3) As String

    xCampos(0, 0) = "CodSunat":    xCampos(0, 1) = "codsun":      xCampos(0, 2) = "C":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripción": xCampos(1, 1) = "descripcion": xCampos(1, 2) = "c":    xCampos(1, 3) = "3200"
    xCampos(2, 0) = "Variable":    xCampos(2, 1) = "variable":    xCampos(2, 2) = "c":    xCampos(2, 3) = "3200"
    xCampos(3, 0) = "Tipo":        xCampos(3, 1) = "tiponombre":  xCampos(3, 2) = "c":    xCampos(3, 3) = "2200"
    xCampos(4, 0) = "Categoría":   xCampos(4, 1) = "catnombre":   xCampos(4, 2) = "c":    xCampos(4, 3) = "2000"

    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg1

End Sub



'*******************************************************************************************

Private Sub cb_Click(Index As Integer)
    If QueHace = 3 And FraEditor.Visible = False Then Exit Sub
    Dim xCampos() As String
    Dim nTitulo As String
    Dim nSQL As String
    On Error GoTo error
    Select Case Index
        Case 0, 4 '--categoria de concepto
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "600":    xCampos(1, 3) = "N"
            nTitulo = "Categoría de Concepto"
            nSQL = "SELECT pla_conceptocat.id, pla_conceptocat.descripcion AS nombre, pla_conceptocat.id AS cod " _
                + vbCr + " FROM pla_conceptocat;"
        
        Case 1, 5 '--tipo de concepto
            If NulosN(lbl_cod(Index - 1).Caption) = 0 Then
                MsgBox "Seleccione la Categoría de Concepto", vbExclamation, xTitulo
                txt_cb(Index - 1).SetFocus
                Exit Sub
            End If
        
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "600":    xCampos(1, 3) = "N"
            
            nTitulo = "Buscando Tipo Concepto"

            nSQL = "SELECT pla_conceptotipo.id, pla_conceptotipo.descripcion AS nombre, pla_conceptotipo.id AS cod " _
                + vbCr + " FROM pla_conceptotipo " _
                + vbCr + " WHERE (((pla_conceptotipo.idcat)=" & NulosN(txt_cb(Index - 1).Text) & ")); "
        Case 2, 3 '--cta contable
        
            ReDim xCampos(2, 4) As String
    
            xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "cuenta":             xCampos(0, 2) = "2000":    xCampos(0, 3) = "C"
            xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":        xCampos(1, 2) = "5000":    xCampos(1, 3) = "C"
            
            nSQL = "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id " _
            + vbCr + " From con_planctas " _
            + vbCr + " ORDER BY con_planctas.cuenta "

    End Select

    Dim xRs As New ADODB.Recordset
    Select Case Index
        Case 2, 3
            CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "cuenta", "cuenta", Principio
        Case Else
            CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
    End Select
    
    If xRs.State = 0 Then GoTo salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo salir

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    txt_cb(Index) = NulosC(xRs.Fields(0)) '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = NulosC(xRs.Fields(1))  '--NOMBRE
    lbl_cod(Index).Caption = NulosN(xRs.Fields(2)) '--CODIGO
    txt_cb(Index).Tag = NulosC(xRs.Fields(1))
    '************************************************************************
    If Index = 0 Then
        If NulosN(lbl_cod(0).Caption) = 2 Then
            TabOne2.TabVisible(2) = False
            TabOne2.TabVisible(3) = True
        ElseIf NulosN(lbl_cod(0).Caption) = 1 Then
            TabOne2.TabVisible(2) = True
            TabOne2.TabVisible(3) = False
        Else
            TabOne2.TabVisible(2) = False
            TabOne2.TabVisible(3) = False
        End If
    End If
    '************************************************************************
    Select Case Index
        Case 0, 4 '--categoria
            txt_cb(Index + 1).SetFocus
        Case 1 '--tipo
            txt(1).SetFocus
        Case 2 '--cta debe
            txt_cb(3).SetFocus
    End Select
    
    '-- si desea activar los conceptos
    If Index = 5 And NulosN(lbl_cod(5).Caption) <> 0 Then pCargarConceptoActivar
        
salir:
    Set xRs = Nothing
Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub

Private Sub txt_cb_Change(Index As Integer)
    If QueHace = 3 And FraEditor.Visible = False Then Exit Sub
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cod(Index).Caption = ""
        lbl_cb(Index).Tag = ""
        If Index = 0 Or Index = 4 Then
            txt_cb(Index + 1).Text = ""
            Fg(2).Rows = 1
        End If
        If Index = 5 Then Fg(2).Rows = 1
    End If

End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
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
    Select Case Index
        Case 0, 1
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case Else
    
    End Select
   
End Sub

Private Sub txt_cb_LostFocus(Index As Integer)
    If Index <> 0 Then Exit Sub
    If NulosN(lbl_cod(0).Caption) = 2 Then
        TabOne2.TabVisible(2) = False
        TabOne2.TabVisible(3) = True
    ElseIf NulosN(lbl_cod(0).Caption) = 1 Then
        TabOne2.TabVisible(2) = True
        TabOne2.TabVisible(3) = False
    Else
        TabOne2.TabVisible(2) = False
        TabOne2.TabVisible(3) = False
    End If

End Sub

Private Sub txt_cb_Validate(Index As Integer, Cancel As Boolean)
    If QueHace = 3 And FraEditor.Visible = False Then Exit Sub
    If txt_cb(Index).Text = "" Then Exit Sub
    
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
'    On Error GoTo error
    Select Case Index

        Case 0, 4 '--categoria de concepto
            nSQL = "SELECT pla_conceptocat.id, pla_conceptocat.descripcion AS nombre, pla_conceptocat.id AS cod " _
                + vbCr + " FROM pla_conceptocat " _
                + vbCr + " WHERE pla_conceptocat.id = " & NulosN(txt_cb(Index).Text) & ";"
        
        Case 1, 5 '--tipo de concepto
            If NulosN(lbl_cod(Index - 1).Caption) = 0 Then
                MsgBox "Seleccione la Categoría de Concepto", vbExclamation, xTitulo
                Exit Sub
            End If
            nSQL = "SELECT pla_conceptotipo.id, pla_conceptotipo.descripcion AS nombre, pla_conceptotipo.id AS cod " _
                + vbCr + " FROM pla_conceptotipo " _
                + vbCr + " WHERE (((pla_conceptotipo.idcat)=" & NulosN(lbl_cod(Index - 1).Caption) & ")) and pla_conceptotipo.id = " & NulosN(txt_cb(Index).Text) & ";"
                
        Case 2, 3 '--cuenta contable
            nSQL = "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id " _
                + vbCr + " FROM con_planctas " _
                + vbCr + " WHERE con_planctas.cuenta= '" & NulosC(txt_cb(Index).Text) & "' ;"

    End Select

    If xCon.State = 0 Then GoTo salir
    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then GoTo salir

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    If RstTmp.RecordCount > 0 Then
        txt_cb(Index).Text = RstTmp.Fields(0) & ""  '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = RstTmp.Fields(1) & "" '--NOMBRE
        lbl_cod(Index).Caption = RstTmp.Fields(2) & "" '--CODIGO
        txt_cb(Index).Tag = RstTmp.Fields(1)
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cod(Index).Caption = ""
    End If
    '--------------
    If Trim(lbl_cod(Index).Tag) <> Trim(lbl_cod(Index).Caption) Then
        Select Case Index
            Case 0 '--categoria de concepto
                txt_cb(1).Text = ""
        End Select
    End If
        
    '-- si desea activar los conceptos
    If Index = 5 And NulosN(lbl_cod(5).Caption) <> 0 Then pCargarConceptoActivar
    
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
'****************************************************************************************
'****************************************************************************************

Sub MuestraSegundoTab()
    Dim QueHaceTmp As Integer
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay Registros para Mostrar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    Blanquea
    TabOne2.CurrTab = 0
    QueHaceTmp = QueHace
    QueHace = -1 '--comodin para entrar a [txt_cb_Validate]
    txt(0).Text = RstFrm("id")
    '--datos de concepto
    If NulosN(RstFrm("idcat")) <> 0 Then
        txt_cb(0).Text = NulosN(RstFrm("idcat"))
        txt_cb_Validate 0, False
    End If
    If NulosN(RstFrm("idtipo")) <> 0 Then
        txt_cb(1).Text = NulosN(RstFrm("idtipo"))
        txt_cb_Validate 1, False
    End If
    '--de las cuentas contables
    '--obtener el numero de cuenta
    Dim mCtaDescripcion As String
    If NulosN(RstFrm("idctadeb")) <> 0 Then
        mCtaDescripcion = Busca_Codigo(NulosN(RstFrm("idctadeb")), "id", "cuenta", "con_planctas", "N", xCon)
        txt_cb(2).Text = mCtaDescripcion
        txt_cb_Validate 2, False
    End If
    If NulosN(RstFrm("idctahab")) <> 0 Then
        mCtaDescripcion = Busca_Codigo(NulosN(RstFrm("idctahab")), "id", "cuenta", "con_planctas", "N", xCon)
        txt_cb(3).Text = mCtaDescripcion
        txt_cb_Validate 3, False
    End If
    '------------------
    txt(1).Text = NulosC(RstFrm("descripcion"))
    
    txt(2).Text = NulosC(RstFrm("variable"))
    txt(3).Text = NulosC(RstFrm("nomcorto"))
    txt(4).Text = NulosC(RstFrm("codsun"))
    txt_formula.Tag = ""
    If NulosC(RstFrm("formula")) <> "" Then
        chk_formula.Value = 1
        txt_formula.Text = RstFrm("formula")
        txt_formula.Tag = RstFrm("formula")
    Else
        chk_formula.Value = 0
    End If
    txt(5).Text = NulosC(RstFrm("observacion"))
    txt_formula.Text = NulosC(RstFrm("formula"))
    chk_formula.Tag = txt_formula.Text
    
    If NulosN(RstFrm("aplanilla")) = -1 Then
        opt_planilla(0).Value = True
    Else
        opt_planilla(1).Value = True
    End If
    '--cabeceras de cada tab
    lbl_cabecera(0).Caption = txt(1).Text
    lbl_cabecera(1).Caption = txt(1).Text
    lbl_cabecera(2).Caption = txt(1).Text
    
    If NulosN(RstFrm("idcat")) <> 2 Then
        TabOne2.TabVisible(2) = True
        TabOne2.TabVisible(3) = False
    Else
        TabOne2.TabVisible(2) = False
        TabOne2.TabVisible(3) = True
    End If
    
    '--datos de contables
    
    QueHace = QueHaceTmp
    
    pCargarDatosDet False
    
    '--centro de costo
'    pCargarDatosDet
End Sub


Private Sub pCargarGrid()
    On Error GoTo error
    Dim nSQL As String
    
    nSQL = "SELECT pla_conceptotipo.idcat, pla_conceptocat.descripcion AS catnombre, pla_conceptotipo.descripcion AS tiponombre, pla_concepto.*, pla_conceptotipo.codsun AS tiposun " _
        + vbCr + " FROM (pla_conceptocat RIGHT JOIN pla_conceptotipo ON pla_conceptocat.id = pla_conceptotipo.idcat) RIGHT JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo " _
        + vbCr + " ORDER BY pla_conceptocat.descripcion desc,pla_concepto.codsun asc, pla_conceptotipo.descripcion, pla_concepto.descripcion; "

    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, nSQL, xCon
    
    Set Dg1.DataSource = RstFrm
    TabOne1.CurrTab = 0
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pCargarGrid"
End Sub

Private Function fValidarDatos() As Boolean
    Dim band As Integer
    TabOne2.CurrTab = 0
    
    band = Validar(txt_cb)
    If band <> -1 Then
       MsgBox "Llene el Campo de " & lbl_capt(band).Caption, vbInformation, xTitulo
       TabOne2.CurrTab = 0
       txt_cb(band).SetFocus
       Exit Function
    End If

    
    band = Validar(txt)
    If band <> -1 Then
       MsgBox "Llene el Campo de " & lbl(band).Caption, vbInformation, xTitulo
       txt(band).SetFocus
       Exit Function
    End If

    If chk_formula.Value = 1 And NulosC(txt_formula.Text) = "" Then
        MsgBox "Ingrese la fórmula", vbExclamation, xTitulo
        cmd_formula.SetFocus
        Exit Function
    End If
    
    '--verificar que no tengan cietos caracteres
    If txt(2).Enabled = True Then
        Dim mCantCarateres&
        For mCantCarateres = 1 To Len(txt(2).Text)
            If InStr("()=+*-/[],: .'?¿!¡%&$#@<>áéíóúñ|°", Mid(txt(2).Text, mCantCarateres, 1)) <> 0 Then
                MsgBox "Caracter no Permitido: [ " & Mid(txt(2).Text, mCantCarateres, 1) & " ]" + vbCr + "Modifique la variable", vbInformation, xTitulo
                Exit Function
            End If
        Next
    End If
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    If QueHace = 1 Then
        nSQL = "SELECT pla_concepto.descripcion, UCase([pla_concepto].[variable])  FROM pla_concepto WHERE (((UCase([pla_concepto].[variable]))='" & UCase(txt(2).Text) & "'));"
    Else
        nSQL = "SELECT pla_concepto.descripcion, UCase([pla_concepto].[variable]) AS Expr1  FROM pla_concepto WHERE (((UCase([pla_concepto].[variable]))='" & UCase(txt(2).Text) & "') AND ((pla_concepto.id)<>" & NulosN(RstFrm.Fields("id")) & "));"
    End If
    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.RecordCount <> 0 Then
        MsgBox "Existe un Concepto que tiene asignado la misma variable" + vbCr + "Concepto: " + NulosC(RstFrm("descripcion")) + vbCr + "Cambie el nombre de la Variable", vbExclamation, xTitulo
        Exit Function
    End If
    '--
    fValidarDatos = True
    
End Function

Private Sub txt_Change(Index As Integer)
    If QueHace = 3 Then Exit Sub
    lbl_cabecera(0).Caption = txt(1).Text
    lbl_cabecera(1).Caption = txt(1).Text
    lbl_cabecera(2).Caption = txt(1).Text
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If Index <> 4 Then Exit Sub
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
    
End Sub

'***************************************************************************************************
Private Sub fg_EnterCell(Index As Integer)
    If FraEditor.Visible = True Then
        Fg(Index).Editable = flexEDKbdMouse
        Exit Sub
    End If
    If QueHace = 3 Then
        Fg(Index).Editable = flexEDNone
        Exit Sub
    End If
    If Fg(Index).Col = 4 Or Fg(Index).Col = 5 Then
        Fg(Index).Editable = flexEDKbdMouse
    Else
        Fg(Index).Editable = flexEDNone
    End If
End Sub

Private Sub fg_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    KeyAscii = 0
End Sub

Private Sub pConfigurarGrilla()
    Agregando = True
    With Fg(0) '--del detalle de la planilla
        .Rows = 2
        .Cols = 5
        .FixedRows = 2
        .RowHeight(0) = 300
        .RowHeight(1) = 250
        .FrozenCols = 1
        GRID_COMBINAR Fg(0), 0, 1, 1, 1, "IdCptoRef", flexAlignLeftCenter, False, , vbBlack, &HD8E9EC, True
        GRID_COMBINAR Fg(0), 0, 2, 1, 2, "Tipo", flexAlignLeftCenter, False, , vbBlack, &HD8E9EC, True
        GRID_COMBINAR Fg(0), 0, 3, 1, 3, "Descripción", flexAlignLeftCenter, False, , vbBlack, &HD8E9EC, True
        GRID_COMBINAR Fg(0), 0, 4, 0, 4, "Seleccionar", flexAlignCenterCenter, True, , vbBlack, &HD8E9EC, True
        
        .ColWidth(1) = 0:
        .ColWidth(2) = 3000:    .ColAlignment(2) = flexAlignLeftCenter:  .Row = 1: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .ColWidth(3) = 6100:    .ColAlignment(3) = flexAlignLeftCenter:   .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 4) = "Si":   .ColWidth(4) = 1200:    .ColAlignment(4) = flexAlignCenterCenter:   .Row = 1: .Col = 4: .CellAlignment = flexAlignCenterCenter

        .SelectionMode = flexSelectionByRow
        .ColDataType(4) = flexDTBoolean
    End With

    With Fg(1) '--
        .Rows = 1
        .Cols = 7
        .FixedRows = 1
        .RowHeight(0) = 250
        .TextMatrix(0, 1) = "IdCpto":       .ColWidth(1) = 0:
        .TextMatrix(0, 2) = "CodSunat":     .ColWidth(2) = 800:  .ColAlignment(2) = flexAlignCenterCenter: .Row = 0: .Col = 2: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 3) = "Tipo":         .ColWidth(3) = 2500: .ColAlignment(3) = flexAlignLeftCenter:   .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 4) = "Descripción":  .ColWidth(4) = 6500: .ColAlignment(4) = flexAlignLeftCenter:   .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 5) = "Variable":     .ColWidth(5) = 0: .ColAlignment(5) = flexAlignLeftCenter:   .Row = 0: .Col = 5: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 6) = "Fórmula":      .ColWidth(6) = 0: .ColAlignment(6) = flexAlignLeftCenter:   .Row = 0: .Col = 6: .CellAlignment = flexAlignLeftCenter

        .SelectionMode = flexSelectionByRow
    End With


    Agregando = False
End Sub

Private Sub pCargarDatosDet(Optional fEsNuevo As Boolean = False)
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    If NulosN(RstFrm("idcat")) <> 2 Then '--is es diferente a aportaciones
        If fEsNuevo = True Then
            nSQL = "SELECT pla_concepto.id AS idcpto, pla_conceptotipo.descripcion AS tipo, pla_concepto.descripcion AS concepto, -2 AS activo " _
                + vbCr + " FROM (pla_conceptocat INNER JOIN pla_conceptotipo ON pla_conceptocat.id = pla_conceptotipo.idcat) INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo " _
                + vbCr + " WHERE (((pla_conceptotipo.idcat) = 2)) AND pla_concepto.afecto =-1 " _
                + vbCr + " ORDER BY pla_conceptotipo.descripcion,pla_concepto.descripcion; "
        Else
            nSQL = "SELECT * FROM " _
                + vbCr + " (SELECT pla_concepto.id AS idcpto, pla_conceptotipo.descripcion AS tipo, pla_concepto.descripcion AS concepto, pla_conceptoapo.activo " _
                + vbCr + " FROM pla_conceptoapo LEFT JOIN (pla_conceptotipo RIGHT JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo) ON pla_conceptoapo.idcptoref = pla_concepto.id " _
                + vbCr + " WHERE (((pla_conceptoapo.idcpto) = " & RstFrm("id") & ")) " _
                + vbCr + " UNION " _
                + vbCr + " SELECT pla_concepto.id AS idcpto, pla_conceptotipo.descripcion AS tipo, pla_concepto.descripcion AS concepto, -2 AS activo " _
                + vbCr + " FROM (pla_conceptocat INNER JOIN pla_conceptotipo ON pla_conceptocat.id = pla_conceptotipo.idcat) INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo " _
                + vbCr + " WHERE (((pla_concepto.id) Not In (SELECT idcptoref FROM pla_conceptoapo WHERE idcpto=" & RstFrm("id") & ")) AND ((pla_conceptotipo.idcat)=2)) AND pla_concepto.afecto =-1 " _
                + vbCr + " ) AS vw " _
                + vbCr + " ORDER BY vw.tipo,vw.concepto ; "
        End If
        
        RST_Busq RstTmp, nSQL, xCon
        Fg(0).Rows = Fg(0).FixedRows
        Agregando = True
        If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            Fg(0).Rows = Fg(0).Rows + 1
            Fg(0).TextMatrix(Fg(0).Rows - 1, 1) = NulosN(RstTmp("idcpto"))
            Fg(0).TextMatrix(Fg(0).Rows - 1, 2) = NulosC(RstTmp("tipo"))
            Fg(0).TextMatrix(Fg(0).Rows - 1, 3) = NulosC(RstTmp("concepto"))
            If NulosN(RstTmp("activo")) = -1 Then
                Fg(0).TextMatrix(Fg(0).Rows - 1, 4) = -1
            Else
                Fg(0).TextMatrix(Fg(0).Rows - 1, 4) = 0
            End If
            
            RstTmp.MoveNext
        Loop
    
    Else '--cargar datos para concepto de aportes
        nSQL = "SELECT pla_concepto.id AS idcpto, pla_concepto.codsun, pla_conceptotipo.descripcion AS tipo, pla_concepto.descripcion AS concepto, pla_concepto.variable, pla_concepto.formula " _
            + vbCr + " FROM pla_conceptoapo INNER JOIN (pla_conceptotipo INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo) ON pla_conceptoapo.idcpto = pla_concepto.id " _
            + vbCr + " Where (((pla_conceptoapo.idcptoref) = " & RstFrm("id") & ")) " _
            + vbCr + " ORDER BY pla_conceptotipo.descripcion, pla_concepto.descripcion;"

        RST_Busq RstTmp, nSQL, xCon
        Fg(1).Rows = Fg(1).FixedRows
        Agregando = True
        If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            Fg(1).Rows = Fg(1).Rows + 1
            Fg(1).TextMatrix(Fg(1).Rows - 1, 1) = NulosN(RstTmp("idcpto"))
            Fg(1).TextMatrix(Fg(1).Rows - 1, 2) = NulosC(RstTmp("codsun"))
            Fg(1).TextMatrix(Fg(1).Rows - 1, 3) = NulosC(RstTmp("tipo"))
            Fg(1).TextMatrix(Fg(1).Rows - 1, 4) = NulosC(RstTmp("concepto"))
            RstTmp.MoveNext
        Loop
    
    End If
    
    Agregando = False
    Set RstTmp = Nothing
End Sub

'**********************************************************************************
Private Sub CmdDet_Click(Index As Integer)
    Select Case Index
        Case 0 '--agregar
            pRegistroAdd False
        Case 1 '--eliminar
            pRegistroDel
        Case 2 '--seleccionar
            pRegistroAdd True
    End Select
End Sub

Private Sub pRegistroAdd(Optional fSeleccionVarios As Boolean = True)
    On Error GoTo error
    Dim xCampos(3, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim nSQLId As String '--almacenara los codigos de documentos ya seleccionados
    Dim nSQLNotInDocumentos As String
    Dim nSQL As String
    Dim nTitulo As String
    xCampos(0, 0) = "CodSun":       xCampos(0, 1) = "codsun":       xCampos(0, 2) = "800":   xCampos(0, 3) = "C":   xCampos(0, 4) = "S"
    If fSeleccionVarios = True Then
        xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "concepto":  xCampos(1, 2) = "6500":  xCampos(1, 3) = "C":  xCampos(1, 4) = "N"
        xCampos(2, 0) = "Tipo":         xCampos(2, 1) = "tipo":    xCampos(2, 2) = "2800":  xCampos(2, 3) = "C":    xCampos(2, 4) = "N"
    Else
        xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "concepto":  xCampos(1, 2) = "4800":  xCampos(1, 3) = "C":  xCampos(1, 4) = "N"
        xCampos(2, 0) = "Tipo":         xCampos(2, 1) = "tipo":    xCampos(2, 2) = "2500":  xCampos(2, 3) = "C":    xCampos(2, 4) = "N"
    End If
    '*************************************************************
    nSQLId = GRID_GENERAR_SQL_ID(Fg(1), 1, "pla_concepto.id", " NOT IN ")
    If nSQLId <> "" Then nSQLId = " AND " & nSQLId
    '*************************************************************
    nSQL = "SELECT pla_concepto.id AS idcpto, pla_concepto.codsun, pla_conceptotipo.descripcion AS tipo, pla_concepto.descripcion AS concepto, pla_concepto.variable, pla_concepto.formula " _
        + vbCr + " FROM pla_conceptotipo INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo " _
        + vbCr + " WHERE (((pla_concepto.codsun)<>'' And (pla_concepto.codsun) Is Not Null) AND ((pla_conceptotipo.idcat)=1)) " & nSQLId _
        + vbCr + " ORDER BY pla_conceptotipo.descripcion, pla_concepto.descripcion;"


    nTitulo = "Buscando Conceptos - Remuneraciones"
    '*************************************************************
    If fSeleccionVarios = True Then
        CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), nTitulo
    Else
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "concepto", "concepto", Principio
    End If
    If xRs.State = 0 Then GoTo salir
    If xRs.RecordCount = 0 Then GoTo salir
    If fSeleccionVarios = True Then xRs.MoveFirst
    Agregando = True
    Do While Not xRs.EOF
        Fg(1).Rows = Fg(1).Rows + 1
        Fg(1).TextMatrix(Fg(1).Rows - 1, 1) = NulosC(xRs("idcpto"))
        Fg(1).TextMatrix(Fg(1).Rows - 1, 2) = NulosC(xRs("codsun"))
        Fg(1).TextMatrix(Fg(1).Rows - 1, 3) = NulosC(xRs("tipo"))
        Fg(1).TextMatrix(Fg(1).Rows - 1, 4) = NulosC(xRs("concepto"))
        If fSeleccionVarios = False Then Exit Do
        xRs.MoveNext
    Loop
    Agregando = False
    Fg(1).Row = Fg(1).Rows - 1: Fg(1).Col = 4:  Fg(1).SetFocus
salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    Agregando = False
    SHOW_ERROR Me.Name, "fg(1)_CellButtonClick"
End Sub

Private Sub pRegistroDel()
    If Fg(1).Rows = 1 Then Exit Sub
    If Fg(1).Row < 1 Then Exit Sub
    If MsgBox("Seguro desea eliminar el Concepto", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    Fg(1).RemoveItem Fg(1).Row
End Sub

'**********************************************************************************
Private Sub pActivarConcepto()
    TabOne1.CurrTab = 0
    Blanquea
    Bloquea True
    pHabilitarBotonEditor True
End Sub

Private Sub pHabilitarBotonEditor(band As Boolean)
    '--TRUE= MUESTRA LA OPCION PARA SELECCIONAR LA RUTA
    Dim K&
    Fg(2).ColWidth(1) = 0 '--idcpto
    Fg(2).Rows = 1
    If band = True Then
        TabOne1.Enabled = False
        
        FraEditor.Top = 1365
        FraEditor.Left = 1050
    Else
        TabOne1.Enabled = True
    End If
    FraEditor.Visible = band
    
    For K = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(K).Enabled = Not band
    Next K
    
End Sub

Private Sub pic_Click()
    CmdEditor_Click 1
End Sub

Private Sub CmdEditor_Click(Index As Integer)
    Select Case Index
        Case 0 'grabar
            If Fg(2).Rows = 1 Then
                MsgBox "No hay registros para grabar", vbExclamation, xTitulo
                Exit Sub
            End If
            If MsgBox("Seguro desea grabar", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
            Dim mRow&
            For mRow = 1 To Fg(2).Rows - 1
                DoEvents
                Fg(2).Row = mRow
                xCon.Execute "UPDATE pla_concepto SET activo=" & NulosN(Fg(2).TextMatrix(mRow, 4)) & " where id = " & NulosN(Fg(2).TextMatrix(mRow, 1)) & ";"
            Next mRow
            MsgBox "Los registros de grabaron Correctamente", vbInformation, xTitulo
            
        Case 1 'cancelar
            Cancelar
    End Select
End Sub

Private Sub pCargarConceptoActivar()
    Dim nSQL  As String
    Dim RstTmp As New ADODB.Recordset
    On Error GoTo error
    
    If NulosN(lbl_cod(4).Caption) = 0 Then
        MsgBox "Seleccione la Categoría", vbExclamation, xTitulo
        Exit Sub
    End If
    
    Fg(2).Rows = 1
    CmdEditor(0).Enabled = False
    
    nSQL = "SELECT pla_concepto.id, pla_concepto.codsun, pla_concepto.descripcion, pla_concepto.activo " _
        + vbCr + " From pla_concepto " _
        + vbCr + " WHERE (((pla_concepto.idtipo) = " & NulosN(lbl_cod(5).Caption) & ")) " _
        + vbCr + " ORDER BY pla_concepto.codsun;"

    Me.MousePointer = vbHourglass
    RST_Busq RstTmp, nSQL, xCon
    '---------------
    If RstTmp.RecordCount <> 0 Then
        Agregando = True
        RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            With Fg(2)
                DoEvents
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = NulosN(RstTmp.Fields("id"))
                .TextMatrix(.Rows - 1, 2) = NulosC(RstTmp.Fields("codsun"))
                .TextMatrix(.Rows - 1, 3) = NulosC(RstTmp.Fields("descripcion"))
                .TextMatrix(.Rows - 1, 4) = NulosN(RstTmp.Fields("activo"))
                RstTmp.MoveNext
            End With
        Loop
    End If
    If Fg(2).Rows > 1 Then
        Fg(2).Row = 1
        CmdEditor(0).Enabled = True
    End If
    '---------------
    Me.MousePointer = vbDefault
    Exit Sub
error:
    SHOW_ERROR Me.Name, "pCargarConceptoActivar"
End Sub



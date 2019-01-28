VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmGenerarOrdProd 
   Caption         =   "Produccion - Generar Orden de Produccion"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   9120
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   3120
      Left            =   6195
      TabIndex        =   14
      Top             =   2415
      Visible         =   0   'False
      Width           =   8325
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   315
         Left            =   3465
         TabIndex        =   17
         Top             =   2700
         Width           =   1410
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg2 
         Height          =   1950
         Left            =   120
         TabIndex        =   16
         Top             =   675
         Width           =   8130
         _cx             =   14340
         _cy             =   3440
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
         Rows            =   50
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmGenerarOrdProd.frx":0000
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
      Begin VB.Label LblProducto 
         AutoSize        =   -1  'True
         Caption         =   "LblProducto"
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
         Left            =   135
         TabIndex        =   18
         Top             =   435
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receta del Producto"
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
         Left            =   180
         TabIndex        =   15
         Top             =   90
         Width           =   1770
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   315
         Left            =   45
         Top             =   30
         Width           =   8250
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   -15
         Y2              =   3090
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   8310
         X2              =   8310
         Y1              =   15
         Y2              =   3075
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   8295
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   8295
         Y1              =   3105
         Y2              =   3105
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
            Picture         =   "FrmGenerarOrdProd.frx":007F
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd.frx":05C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd.frx":0955
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd.frx":0AD9
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd.frx":0F2D
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd.frx":1045
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd.frx":1589
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd.frx":1ACD
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd.frx":1BE1
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd.frx":1CF5
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd.frx":2149
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd.frx":22B5
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd.frx":27FD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Listado"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   5955
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   9120
      _cx             =   16087
      _cy             =   10504
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
         Height          =   5535
         Left            =   45
         TabIndex        =   4
         Top             =   375
         Width           =   9030
         Begin TrueOleDBGrid70.TDBGrid Dg3 
            Height          =   5775
            Left            =   0
            TabIndex        =   5
            Top             =   345
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   10186
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "IdPer"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "IdEmp"
            Columns(1).DataField=   "idemp"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Empleado"
            Columns(2).DataField=   "nomemp"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "T.D."
            Columns(3).DataField=   "abrev"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Nº Documento"
            Columns(4).DataField=   "numdoc"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Nº Funciones"
            Columns(5).DataField=   "totalfunc"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=794"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=714"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1111"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1032"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Visible=0"
            Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(15)=   "Column(2).Width=7250"
            Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=7170"
            Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(21)=   "Column(3).Width=979"
            Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=900"
            Splits(0)._ColumnProps(24)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(25)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(27)=   "Column(4).Width=2328"
            Splits(0)._ColumnProps(28)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._WidthInPix=2249"
            Splits(0)._ColumnProps(30)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(31)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(32)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(33)=   "Column(5).Width=2223"
            Splits(0)._ColumnProps(34)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._WidthInPix=2143"
            Splits(0)._ColumnProps(36)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(37)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(38)=   "Column(5).Order=6"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
            _StyleDefs(60)  =   "Named:id=33:Normal"
            _StyleDefs(61)  =   ":id=33,.parent=0"
            _StyleDefs(62)  =   "Named:id=34:Heading"
            _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
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
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Control de Personal"
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
            Left            =   105
            TabIndex        =   6
            Top             =   45
            Width           =   7905
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5535
         Left            =   9765
         TabIndex        =   2
         Top             =   375
         Width           =   9030
         Begin VB.CommandButton CmdLoadCrono 
            Caption         =   "Cargar Cronograma"
            Height          =   420
            Left            =   6720
            TabIndex        =   11
            Top             =   525
            Width           =   2205
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   3870
            Left            =   45
            TabIndex        =   10
            Top             =   1020
            Width           =   8910
            _cx             =   15716
            _cy             =   6826
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
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmGenerarOrdProd.frx":2B8F
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
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchPro 
            Height          =   300
            Left            =   1050
            TabIndex        =   8
            Top             =   585
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
            Locked          =   -1  'True
         End
         Begin VB.CommandButton cb 
            Enabled         =   0   'False
            Height          =   240
            Index           =   0
            Left            =   5955
            Picture         =   "FrmGenerarOrdProd.frx":2C80
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   480
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Frame Frame3 
            Height          =   690
            Left            =   45
            TabIndex        =   12
            Top             =   4845
            Width           =   8940
            Begin VB.CommandButton Command4 
               Caption         =   "Generar Nº Prod."
               Height          =   345
               Left            =   4020
               TabIndex        =   22
               Top             =   225
               Visible         =   0   'False
               Width           =   1545
            End
            Begin VB.CommandButton Command3 
               Caption         =   "Imprimir"
               Height          =   345
               Left            =   5610
               TabIndex        =   21
               Top             =   225
               Width           =   1545
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Eliminar Producto"
               Height          =   345
               Left            =   1845
               TabIndex        =   20
               Top             =   225
               Width           =   1545
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Agregar Producto"
               Height          =   345
               Left            =   255
               TabIndex        =   19
               Top             =   225
               Width           =   1545
            End
            Begin VB.CommandButton CmdVerReceta 
               Caption         =   "Ver Receta"
               Height          =   345
               Left            =   7185
               TabIndex        =   13
               Top             =   225
               Width           =   1545
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Prod."
            Height          =   195
            Left            =   180
            TabIndex        =   9
            Top             =   630
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Orden de Produccion"
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
            Left            =   75
            TabIndex        =   3
            Top             =   75
            Width           =   7605
         End
      End
   End
End
Attribute VB_Name = "FrmGenerarOrdProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rst As New ADODB.Recordset
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO

Private Sub CmdLoadCrono_Click()
    If TxtFchPro.Valor = "" Then
        MsgBox "No ha especificado la fecha de produccion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchPro.SetFocus
        Exit Sub
    End If
    
    Dim Rst As New ADODB.Recordset
    Dim xSQL As String
    xSQL = "SELECT 0 as xsel, pro_cronogramadetprod.idpro, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev AS unimed, pro_cronogramadet.cantidad, " _
        & " pro_cronogramadetprod.fchpro, alm_inventario_1.descripcion AS matpri, pro_receta.codrec, pro_receta.id AS idrec, 'M' AS tipo " _
        & " FROM (((((pro_cronogramadetprod LEFT JOIN pro_cronogramadet ON (pro_cronogramadetprod.id = pro_cronogramadet.id) AND " _
        & " (pro_cronogramadetprod.iditem = pro_cronogramadet.iditem) AND (pro_cronogramadetprod.fchpro = pro_cronogramadet.fchpro) " _
        & " AND (pro_cronogramadetprod.horpro = pro_cronogramadet.Horpro)) LEFT JOIN alm_inventario ON pro_cronogramadetprod.idpro = alm_inventario.id) " _
        & " LEFT JOIN alm_inventario AS alm_inventario_1 ON pro_cronogramadetprod.iditem = alm_inventario_1.id) LEFT JOIN pro_cronograma " _
        & " ON pro_cronogramadet.id = pro_cronograma.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) LEFT JOIN pro_receta " _
        & " ON pro_cronogramadetprod.idpro = pro_receta.iditem WHERE (((pro_cronogramadetprod.fchpro)=CDate('" & TxtFchPro.Valor & "')) AND ((pro_cronograma.idtippro)=1) AND ((pro_receta.prirec)=1))" _
        & " Union " _
        & " SELECT 0 as xsel, pro_cronogramadet.iditem, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev AS unimed, pro_cronogramadet.cantidad, " _
        & " pro_cronogramadet.fchpro, '' AS matpri, pro_receta.codrec, pro_receta.id AS idrec, 'P' AS tipo FROM (((pro_cronogramadet LEFT JOIN pro_cronograma " _
        & " ON pro_cronogramadet.id = pro_cronograma.id) LEFT JOIN alm_inventario ON pro_cronogramadet.iditem = alm_inventario.id) LEFT JOIN mae_unidades " _
        & " ON alm_inventario.idunimed = mae_unidades.id) LEFT JOIN pro_receta ON pro_cronogramadet.iditem = pro_receta.iditem " _
        & " WHERE (((pro_cronogramadet.fchpro)=CDate('" & TxtFchPro.Valor & "')) AND ((pro_cronograma.idtippro)=3) AND ((pro_receta.prirec)=1))"

    Dim xFrm As New eps_librerias.FormSeleccion
    Dim xCampos(4, 5) As String
    Dim xRs As New ADODB.Recordset
    
    xCampos(0, 0) = "Producto":    xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "2500":   xCampos(0, 3) = "C":     xCampos(0, 4) = "S"
    xCampos(1, 0) = "Uni. Med.":   xCampos(1, 1) = "unimed":        xCampos(1, 2) = "800":    xCampos(1, 3) = "C":     xCampos(1, 4) = "N"
    xCampos(2, 0) = "Cantidad":    xCampos(2, 1) = "cantidad":      xCampos(2, 2) = "1000":   xCampos(2, 3) = "N":     xCampos(2, 4) = "N"
    xCampos(3, 0) = "Receta":      xCampos(3, 1) = "codrec":        xCampos(3, 2) = "1200":   xCampos(3, 3) = "C":     xCampos(3, 4) = "N"
    
    
    xFrm.SQLCad = xSQL
    xFrm.Titulo = "Buscando Entradas a Almacen"
    
    Set xFrm.Coneccion = xCon
    Set xRs = xFrm.Seleccionar(xCampos)
    Dim A As Integer
    
    If xRs.State = 1 Then
        xRs.MoveFirst
        For A = 1 To xRs.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = xRs("descripcion")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = xRs("unimed")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(xRs("cantidad"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = xRs("codrec")
            
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = xRs("idpro")
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = xRs("idrec")
            xRs.MoveNext
            
            If xRs.EOF = True Then Exit For
        Next
    End If
    Set xFrm = Nothing
End Sub

Private Sub Command1_Click()
    If Fg1.TextMatrix(Fg1.Rows - 1, 1) = "" Then Exit Sub
    
    Fg1.Rows = Fg1.Rows + 1
    Fg1_CellButtonClick Fg1.Rows - 1, 1
End Sub

Private Sub CmdSalir_Click()
    ActivaEntorno True
    Frame4.Visible = False
End Sub

Sub ActivaEntorno(xQueFue As Boolean)
    TabOne1.Enabled = xQueFue
    Toolbar1.Enabled = xQueFue
End Sub

Private Sub CmdVerReceta_Click()
    If Fg1.Rows = 1 Then
        MsgBox "No hay productos que mostrar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.SetFocus
        Exit Sub
    End If
    
    ActivaEntorno False
    Frame4.Left = 465
    Frame4.Top = 1845
    LblProducto.Caption = Fg1.TextMatrix(Fg1.Row, 1)
    Frame4.Visible = True
    
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    RST_Busq Rst, "SELECT pro_recetains.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro, [pro_recetains]![canpro]* " & NulosN(Fg1.TextMatrix(Fg1.Row, 3)) & " AS canreq " _
        & " FROM (pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id " _
        & " WHERE (((pro_recetains.idrec)=" & NulosN(Fg1.TextMatrix(Fg1.Row, 7)) & "))", xCon

    If Rst.RecordCount <> 0 Then
        Fg2.Rows = 1
        
        For A = 1 To Rst.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = Rst("descripcion")
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = Rst("abrev")
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(Rst("canreq"), "0.000000")
            
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
    End If
End Sub

Private Sub Command2_Click()
    If Fg1.Rows = 1 Then Exit Sub
    
    Fg1.RemoveItem Fg1.Row
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Mandar a imprimir
Private Sub Command3_Click()
    If NulosC(TxtFchPro.Valor) = "" Then
        MsgBox "No ha especificado la fecha de produccion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchPro.SetFocus
        Exit Sub
    End If
    
    Dim A As Integer
    
    With FrmVsPrinter.Vs
        'FrmVsPrinter.Vs.PaperSize = 512
        FrmVsPrinter.Vs.BrushColor = &H80000005
        FrmVsPrinter.Vs.FontSize = 11
        FrmVsPrinter.Vs.TextAlign = taCenterMiddle
        FrmVsPrinter.Vs.StartDoc
        Dim xLinea As Integer
        
        xLinea = 500
        For A = 1 To Fg1.Rows - 1
            If xLinea >= 13000 Then
                xLinea = 500
                FrmVsPrinter.Vs.NewPage
            End If
            'LADO A
            FrmVsPrinter.Vs.FontSize = 13
            FrmVsPrinter.Vs.TextAlign = taCenterMiddle
            FrmVsPrinter.Vs.TextBox "SOLICITUD DE MATERIALES", 500, xLinea, 6500, 500, True, False, True
            FrmVsPrinter.Vs.FontSize = 10
            FrmVsPrinter.Vs.TextAlign = taCenterTop
            FrmVsPrinter.Vs.TextBox "Nº ", 7300, xLinea, 1700, 250, True, False, True
            xLinea = xLinea + 240
            FrmVsPrinter.Vs.TextBox "0001-0000000___", 7300, xLinea, 1700, 250, True, False, True
            
            FrmVsPrinter.Vs.TextAlign = taLeftMiddle
            FrmVsPrinter.Vs.FontSize = 9
            xLinea = xLinea + 500
            FrmVsPrinter.Vs.TextBox "Nº ORDEN     ", 500, xLinea, 1500, 250, True, False, False
            FrmVsPrinter.Vs.TextBox "_____________", 2000, xLinea, 1500, 250, True, False, False
            FrmVsPrinter.Vs.TextBox "Fch. Prod.   ", 4000, xLinea, 900, 250, True, False, False
            FrmVsPrinter.Vs.TextBox TxtFchPro.Valor, 5200, xLinea, 1500, 250, True, False, False
            xLinea = xLinea + 250
            FrmVsPrinter.Vs.TextBox "PRODUCTO    ", 500, xLinea, 1500, 250, True, False, False
            FrmVsPrinter.Vs.TextBox Fg1.TextMatrix(A, 1), 2000, xLinea, 6000, 250, True, False, False
            xLinea = xLinea + 250
            FrmVsPrinter.Vs.TextBox "Receta ", 500, xLinea, 1500, 250, True, False, False
            FrmVsPrinter.Vs.TextBox Fg1.TextMatrix(A, 4), 2000, xLinea, 6000, 250, True, False, False
            
            
            FrmVsPrinter.Vs.TextBox "CANTIDAD   ", 4000, xLinea, 1500, 250, True, False, False
            FrmVsPrinter.Vs.TextBox Fg1.TextMatrix(A, 3), 5200, xLinea, 6000, 250, True, False, False
            
            
            
            
            xLinea = xLinea + 360
            FrmVsPrinter.Vs.TextAlign = taCenterMiddle
            FrmVsPrinter.Vs.TextBox "Item", 500, xLinea, 400, 500, True, False, True
            FrmVsPrinter.Vs.TextBox "INSUMO / PRODUCTO / MP", 900, xLinea, 3700, 500, True, False, True
            FrmVsPrinter.Vs.TextBox "Uni. Med.", 4600, xLinea, 400, 500, True, False, True
            FrmVsPrinter.Vs.TextBox "Cantidad Teorica", 5000, xLinea, 1000, 500, True, False, True
            FrmVsPrinter.Vs.TextBox "Cantidad Real", 6000, xLinea, 1000, 500, True, False, True
            FrmVsPrinter.Vs.TextBox "Adicional", 7000, xLinea, 1000, 500, True, False, True
            FrmVsPrinter.Vs.TextBox "Devolucion", 8000, xLinea, 1000, 500, True, False, True
            
            Dim Rst As New ADODB.Recordset
            'Dim A As Integer
            Dim B As Integer
            RST_Busq Rst, "SELECT pro_recetains.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro, [pro_recetains]![canpro]*" & NulosN(Fg1.TextMatrix(A, 3)) & " AS canreq " _
                & " FROM (pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id " _
                & " WHERE (((pro_recetains.idrec)=" & NulosN(Fg1.TextMatrix(A, 7)) & "))", xCon
        
            If Rst.RecordCount <> 0 Then
                Fg2.Rows = 1
                Dim xFila As Integer
                xLinea = xLinea + 500
                xFila = xLinea
                For B = 1 To Rst.RecordCount
                    FrmVsPrinter.Vs.FontSize = 8
                    Fg2.Rows = Fg2.Rows + 1
                    FrmVsPrinter.Vs.TextAlign = taLeftMiddle
                    FrmVsPrinter.Vs.TextBox Format(B, "00"), 500, xLinea, 400, 250, True, False, True
                    FrmVsPrinter.Vs.TextBox Rst("descripcion"), 900, xLinea, 3700, 250, True, False, True
                    FrmVsPrinter.Vs.TextAlign = taCenterMiddle
                    FrmVsPrinter.Vs.TextBox Rst("abrev"), 4600, xLinea, 400, 250, True, False, True
                    FrmVsPrinter.Vs.TextAlign = taRightMiddle
                    FrmVsPrinter.Vs.TextBox Format(Rst("canreq"), "0.000000"), 5000, xLinea, 1000, 250, True, False, True
                    FrmVsPrinter.Vs.TextBox "", 6000, xLinea, 1000, 250, True, False, True
                    FrmVsPrinter.Vs.TextBox "", 7000, xLinea, 1000, 250, True, False, True
                    FrmVsPrinter.Vs.TextBox "", 8000, xLinea, 1000, 250, True, False, True
                    
                    Rst.MoveNext
                    If Rst.EOF = True Then Exit For
                    
                    xLinea = xLinea + 250
                    
                    If xLinea >= 16200 Then
                        xLinea = 500
                        FrmVsPrinter.Vs.NewPage
                    End If
                Next B
            End If
            
            ' POSICION ANTES DEL DETALLE + ALTO DE DE 10 ITEMS + 500 DE ESPACIO
            'xLinea = (xFila + 2500 + 100)
            xLinea = xLinea + 750
            
            If xLinea >= 16400 Then
                xLinea = 2000
                FrmVsPrinter.Vs.NewPage
            End If
            
            'LADO A
            FrmVsPrinter.Vs.TextBox "-------------------------", 1950, xLinea, 1400, 200, True, False, False
            FrmVsPrinter.Vs.TextBox "-------------------------", 5300, xLinea, 1400, 200, True, False, False
            
            xLinea = xLinea + 200
            FrmVsPrinter.Vs.TextBox "VºBº Ger. Prod. ", 2000, xLinea, 1200, 250, True, False, False
            FrmVsPrinter.Vs.TextBox "Entregado Por ", 5300, xLinea, 1200, 250, True, False, False
            
            xLinea = xLinea + 500
            
            'If xLinea >= 16400 Then
            '    xLinea = 500
            '    FrmVsPrinter.Vs.NewPage
            'End If
        Next A
        
        FrmVsPrinter.Vs.EndDoc
    End With
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Muestra la preimagen de la impresion
    FrmVsPrinter.Show vbModal
    
End Sub

Private Sub Command4_Click()
    Dim Rst As New ADODB.Recordset
    Dim xNumOP As Integer
    Dim A As Integer
    
    RST_Busq Rst, "SELECT Max(pro_producciondet.numparte) AS MaxDenumparte From pro_producciondet ORDER BY Max(pro_producciondet.numparte)", xCon
    
    If Rst.RecordCount <> 0 Then
        xNumOP = NulosN(Rst("MaxDenumparte")) + 1
        For A = 1 To Fg1.Rows - 1
            If NulosN(Fg1.TextMatrix(A, 0)) = -1 Then
                Fg1.TextMatrix(A, 5) = Format(xNumOP, "00000000")
                xNumOP = xNumOP + 1
            End If
        Next A
    End If
    
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    
    If Col = 1 Then
        xCampos(0, 0) = "Producto":   xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":     xCampos(1, 1) = "codpro":        xCampos(1, 2) = "1200":         xCampos(1, 3) = "C"
        
        xform.SQLCad = "SELECT alm_inventario.codpro, alm_inventario.descripcion, alm_inventario.tippro, alm_inventario.id, pro_receta.codrec, " _
            & " pro_receta.id AS idrec, mae_unidades.abrev AS unimed FROM (alm_inventario LEFT JOIN pro_receta ON alm_inventario.id = pro_receta.iditem) " _
            & " LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id Where (((alm_inventario.tippro) = 3) And ((pro_receta.prirec) = 1))" _
            & " ORDER BY alm_inventario.descripcion"

        xform.Titulo = "Buscando Producto"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "descripcion"
        xform.CampoBusca = "descripcion"
        Set xform.Coneccion = xCon
        
        'Inicia tabla de busqueda
        Set xRs = xform.BuscarReg(xCampos)
        
        If xRs.State = 1 Then
            Fg1.TextMatrix(Row, 1) = xRs("descripcion")
            Fg1.TextMatrix(Row, 2) = xRs("unimed")
            Fg1.TextMatrix(Row, 4) = xRs("codrec")
            Fg1.TextMatrix(Row, 6) = xRs("id")
            Fg1.TextMatrix(Row, 7) = xRs("idrec")
        End If
    End If
    
    If Col = 4 Then
        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Receta":     xCampos(1, 1) = "codrec":        xCampos(1, 2) = "1200":         xCampos(1, 3) = "C"
        
        xform.SQLCad = "SELECT pro_receta.codrec, pro_receta.descripcion, pro_receta.prirec, pro_receta.id " _
            & " From pro_receta Where (((pro_receta.iditem) = " & NulosN(Fg1.TextMatrix(Fg1.Row, 6)) & ")) ORDER BY pro_receta.prirec"
        
        xform.Titulo = "Buscando Recetas del Producto"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "descripcion"
        xform.CampoBusca = "descripcion"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        
        If xRs.State = 1 Then
            Fg1.TextMatrix(Fg1.Row, 4) = xRs("codrec")
            Fg1.TextMatrix(Fg1.Row, 7) = xRs("id")
        End If
    End If
End Sub


'Funcion que se encarga de verificar que se introdusca solo numeros
Function solo_numero(texto As String) As Boolean
    
    Dim i As Integer
    Dim tArray() As String
    
    ReDim tArray(1 To Len(texto)) As String
    For i = 1 To Len(texto)
        tArray(i) = Mid(texto, i, 1)
    Next
    
    For i = LBound(tArray) To UBound(tArray)
        If InStr(1, ".0123456789", tArray(i)) = 0 Then
            solo_numero = False
        Else
            solo_numero = True
        End If
    Next
End Function

'Cambiar cantidad manualmente
Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 3 Then
        'Verifico que sea un dato numerico
        If solo_numero(Fg1.TextMatrix(Row, Col)) Then
            Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), "0.00")
        Else
            MsgBox "Ingrese solo datos numericos"
            Fg1.TextMatrix(Row, Col) = ""
        End If
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
    
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
    End If
End Sub

Private Sub Form_Load()
    QueHace = 3
    SeEjecuto = False
    TabOne1.CurrTab = 1
    
    Fg1.ColWidth(6) = 0
    Fg1.ColWidth(7) = 0
    
    Fg2.SelectionMode = flexSelectionByRow
    Fg2.Editable = flexEDNone
    TabOne1.TabEnabled(0) = False
    Toolbar1.Enabled = Not Toolbar1.Enabled
    Nuevo
End Sub

Sub Nuevo()
    QueHace = 1
    'TabOne1.CurrTab = 1
    'TabOne1.TabEnabled(0) = False
    Bloquea
    Blanquea
    Fg1.Rows = 1
    
    Fg1.ColComboList(1) = "|..."
    Fg1.ColComboList(4) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    Fg1.Editable = flexEDKbdMouse
    
    'TxtFchPro.SetFocus
End Sub

Sub Bloquea()
    TxtFchPro.Locked = Not TxtFchPro.Locked
End Sub

Sub Blanquea()
    TxtFchPro.Valor = Date
    TxtFchPro.Valor = ""
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then MsgBox "1"
    If Button.Index = 3 Then MsgBox "2"
        
End Sub


VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmIngresoAlmacen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Almacen - Ingresos / Salida de Almacen"
   ClientHeight    =   7410
   ClientLeft      =   165
   ClientTop       =   1590
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8265
      Top             =   30
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
            Picture         =   "FrmIngresoAlmacen.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen.frx":277E
            Key             =   "IMG13"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7020
      Left            =   0
      TabIndex        =   15
      Top             =   360
      Width           =   11880
      _cx             =   20955
      _cy             =   12382
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
      Appearance      =   1
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
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6600
         Left            =   45
         TabIndex        =   28
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6210
            Left            =   30
            TabIndex        =   29
            Top             =   360
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   10954
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
            Columns(1).Caption=   "Nº Registro"
            Columns(1).DataField=   "id"
            Columns(1).NumberFormat=   "0000"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Ing"
            Columns(2).DataField=   "fching1"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "T. Mov."
            Columns(3).DataField=   "movi"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "T.D."
            Columns(4).DataField=   "abrev"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Nº Documento"
            Columns(5).DataField=   "numdoc2"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Fch. Emi"
            Columns(6).DataField=   "fchdoc1"
            Columns(6).NumberFormat=   "Short Date"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Cliente / Proveedor"
            Columns(7).DataField=   "nombre"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Nº Documento"
            Columns(8).DataField=   "numdoccom"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   9
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=9"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1931"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1852"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=1799"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1720"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=1323"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1244"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=979"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=900"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=2619"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=2540"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=1773"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1693"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=513"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=6244"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=6165"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=512"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(50)=   "Column(8).Width=2937"
            Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=2858"
            Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=516"
            Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
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
            _StyleDefs(11)  =   ":id=2,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
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
         Begin VB.Label LblMes 
            Caption         =   "LblMes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   9930
            TabIndex        =   45
            Top             =   30
            Width           =   1815
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Movimientos ( Ingresos/Salidas )"
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
            TabIndex        =   31
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label LblPeriodo 
            Alignment       =   2  'Center
            Caption         =   "LblMes"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   9810
            TabIndex        =   30
            Top             =   0
            Width           =   1860
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6600
         Left            =   12525
         TabIndex        =   16
         Top             =   375
         Width           =   11790
         Begin VB.Frame Frame5 
            Height          =   495
            Left            =   9570
            TabIndex        =   46
            Top             =   300
            Width           =   2115
            Begin VB.Label LblPeriodo2 
               Alignment       =   2  'Center
               Caption         =   "LblPeriodo2"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   300
               Left            =   120
               TabIndex        =   47
               Top             =   150
               Width           =   1860
            End
         End
         Begin VB.CommandButton CmdBusAlm 
            Height          =   240
            Left            =   9195
            Picture         =   "FrmIngresoAlmacen.frx":2B10
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   1800
            Width           =   240
         End
         Begin VB.TextBox TxtIdAlm 
            Height          =   300
            Left            =   8700
            Locked          =   -1  'True
            TabIndex        =   4
            Text            =   "TxtIdAlm"
            Top             =   1770
            Width           =   765
         End
         Begin VB.CommandButton CmdBusArea 
            Height          =   240
            Left            =   2445
            Picture         =   "FrmIngresoAlmacen.frx":2C42
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   3060
            Width           =   240
         End
         Begin VB.CommandButton CmdBusRes 
            Height          =   240
            Left            =   2445
            Picture         =   "FrmIngresoAlmacen.frx":2D74
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   2745
            Width           =   240
         End
         Begin VB.Frame Frame3 
            Height          =   660
            Left            =   240
            TabIndex        =   33
            Top             =   375
            Width           =   3570
            Begin VB.OptionButton OptSal 
               Caption         =   "Salida"
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
               Height          =   240
               Left            =   2010
               TabIndex        =   14
               Top             =   240
               Width           =   1000
            End
            Begin VB.OptionButton OptIng 
               Caption         =   "Ingreso"
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
               Height          =   240
               Left            =   585
               TabIndex        =   13
               Top             =   240
               Width           =   1080
            End
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   5
            Text            =   "TxtNumSer"
            Top             =   2085
            Width           =   915
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2895
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   6
            Text            =   "TxtNumDoc"
            Top             =   2085
            Width           =   1440
         End
         Begin VB.CommandButton CmdBusProv 
            Height          =   240
            Left            =   7005
            Picture         =   "FrmIngresoAlmacen.frx":2EA6
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   2430
            Width           =   240
         End
         Begin VB.CommandButton CmdBusTipoCompra 
            Height          =   240
            Left            =   2445
            Picture         =   "FrmIngresoAlmacen.frx":2FD8
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   1485
            Width           =   240
         End
         Begin VB.Frame Frame4 
            Height          =   3225
            Left            =   9345
            TabIndex        =   20
            Top             =   3345
            Width           =   2190
            Begin VB.CommandButton CmdDelItem 
               Caption         =   "Eliminar Item"
               Enabled         =   0   'False
               Height          =   525
               Left            =   420
               TabIndex        =   12
               Top             =   1545
               Width           =   1380
            End
            Begin VB.CommandButton CmdAddItem 
               Caption         =   "Agregar Item"
               Enabled         =   0   'False
               Height          =   525
               Left            =   420
               TabIndex        =   11
               Top             =   1005
               Width           =   1380
            End
         End
         Begin VB.TextBox TxtTipCom 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   2
            Text            =   "TxtTipCom"
            Top             =   1455
            Width           =   915
         End
         Begin VB.TextBox TxtProv 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   7
            Text            =   "TxtProv"
            Top             =   2400
            Width           =   5475
         End
         Begin VB.CommandButton CmdBusTipDoc 
            Height          =   240
            Left            =   2445
            Picture         =   "FrmIngresoAlmacen.frx":310A
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   1800
            Width           =   240
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   3195
            Left            =   240
            TabIndex        =   10
            Top             =   3375
            Width           =   9075
            _cx             =   16007
            _cy             =   5636
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
            Rows            =   1
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmIngresoAlmacen.frx":323C
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
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchIng 
            Height          =   300
            Left            =   1800
            TabIndex        =   0
            Top             =   1140
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
            Valor           =   "18/09/2007"
         End
         Begin VB.TextBox TxtIdRes 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   8
            Text            =   "TxtIdRes"
            Top             =   2715
            Width           =   915
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchDoc 
            Height          =   300
            Left            =   6015
            TabIndex        =   1
            Top             =   1140
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
            Valor           =   "18/09/2007"
         End
         Begin VB.TextBox TxtIdArea 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   9
            Text            =   "TxtIdArea"
            Top             =   3030
            Width           =   915
         End
         Begin VB.TextBox TxtTipDoc 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   3
            Text            =   "TxtTipDoc"
            Top             =   1770
            Width           =   915
         End
         Begin VB.Label LblDescDoc 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDescDoc"
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
            Left            =   9495
            TabIndex        =   44
            Top             =   1770
            Width           =   2085
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Area"
            Height          =   195
            Index           =   8
            Left            =   255
            TabIndex        =   43
            Top             =   3030
            Width           =   330
         End
         Begin VB.Label LblArea 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblArea"
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
            Left            =   2760
            TabIndex        =   42
            Top             =   3030
            Width           =   4515
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Almacén"
            Height          =   195
            Index           =   5
            Left            =   7830
            TabIndex        =   40
            Top             =   1815
            Width           =   615
         End
         Begin VB.Label LblTipDoc 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipDoc"
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
            Left            =   2760
            TabIndex        =   38
            Top             =   1770
            Width           =   4515
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Documento"
            Height          =   195
            Index           =   4
            Left            =   4470
            TabIndex        =   37
            Top             =   1185
            Width           =   1185
         End
         Begin VB.Label LblResp 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblResp"
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
            Left            =   2760
            TabIndex        =   36
            Top             =   2715
            Width           =   4515
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Responsable"
            Height          =   195
            Index           =   2
            Left            =   255
            TabIndex        =   35
            Top             =   2745
            Width           =   930
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000001&
            BackStyle       =   1  'Opaque
            Height          =   90
            Left            =   2745
            Top             =   2190
            Width           =   105
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Documento"
            Height          =   195
            Index           =   0
            Left            =   255
            TabIndex        =   32
            Top             =   2115
            Width           =   1275
         End
         Begin VB.Label LblIdProveedor 
            AutoSize        =   -1  'True
            Caption         =   "LblIdProveedor"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   7320
            TabIndex        =   27
            Top             =   2445
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Documento"
            Height          =   195
            Index           =   1
            Left            =   255
            TabIndex        =   26
            Top             =   1815
            Width           =   1410
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   255
            TabIndex        =   25
            Top             =   2430
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Movimiento"
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
            TabIndex        =   24
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Ingreso"
            Height          =   195
            Index           =   3
            Left            =   255
            TabIndex        =   23
            Top             =   1185
            Width           =   885
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Producto"
            Height          =   195
            Index           =   6
            Left            =   255
            TabIndex        =   22
            Top             =   1500
            Width           =   1005
         End
         Begin VB.Label LblTipoCompra 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipoCompra"
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
            Left            =   2760
            TabIndex        =   21
            Top             =   1455
            Width           =   4515
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
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
            Object.ToolTipText     =   "Agregar "
            ImageKey        =   "IMG1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageKey        =   "IMG4"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageKey        =   "IMG6"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageKey        =   "IMG5"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageKey        =   "IMG7"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageKey        =   "IMG8"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageKey        =   "IMG9"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageKey        =   "IMG13"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageKey        =   "IMG11"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageKey        =   "IMG10"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu menu1 
      Caption         =   "menu1"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Agregar                "
      End
      Begin VB.Menu menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Eliminar"
      End
   End
End
Attribute VB_Name = "FrmIngresoAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMINGRESOALMACEN.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : FORMULARIO PARA EL INGRESO DE DOCUMENTOS NO CONTABLES DE INGRESO O SALIDA,
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 17/09/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim RstIng As New ADODB.Recordset                    ' RECORDSET PRINCIPAL QUE CARGARA TODAS LAS OPERACIONES REGISTRADAS
Dim QueHace As Integer                               ' VARIABLE QUE INDICA EL ESTADO DEL FORMULARIO 1 = NUEVO, 2 = MODIFICA; 3 = SOLO LECTURA
Dim SeEjecuto As Boolean                             ' VARIABLE UTILIZADA PARA EJECUTAR UNA SOLA VEZ EL EVENTO ACTIVATE
Dim Agregando As Boolean                             ' VARIABLE QUE INFORMA A LOS CONTROLES FlexGrid QUE SE ESTA AGREGADO UNA FILA
Dim Mostrando As Boolean
Dim CaracteresNumericos As String                    ' ESPECIFICA LOS CARACTERES NUMERICOS QUE PODRA SOPORTAR LOS CONTROLES TextBox
Dim CaracteresNumericos2 As String, vStr As String   ' ESPECIFICA LOS CARACTERES NUMERICOS QUE PODRA SOPORTAR LOS CONTROLES TextBox
Dim mIdRegistro&                                     ' identificador del registro
Dim fOrdenLista As Boolean                           ' especfica el orden de la lista de la consulta
Dim xHorIni As Date                                  ' ESPECIFICA LA HORA DE INICIO

Dim mMesActivo As Integer                  ' --indica el mes activo
Dim fCierrePeriodo As Boolean              ' --indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO


'*****************************************************************************************************
'* Nombre           : fVerifSiExistDocum
'* Tipo             : FUNCION
'* Descripcion      : BUSCA UN NUMERO DE DOCUMENTO EN LA TABLA alm_ingreso, SI EL NUMERO DOCUMENTO
'*                    EXISTE DEVUELVE VERDADERO
'* Paranetros       : NOMBRE    |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                              |                   |
'* Devuelve         : Boolean
'*****************************************************************************************************
Function fVerifSiExistDocum() As Boolean
    Dim rsVerifDoc As New ADODB.Recordset
    vStr = "SELECT alm_ingreso.numser, alm_ingreso.numdoc"
    vStr = vStr & " FROM alm_ingreso"
    vStr = vStr & " WHERE tipdoc = " & NulosN(TxtTipDoc.Text) & " AND numser = '" + Trim(TxtNumSer.Text) + "' AND numdoc = '" + Trim(TxtNumDoc.Text) + "'"
    If OptIng.Value = True Then
        vStr = vStr & " AND tipmov = true"
    Else
        vStr = vStr & " AND tipmov = false"
    End If
    If OptIng.Value = True And NulosN(TxtTipDoc.Text) = 9 Then
        vStr = vStr & " AND idpro = " & NulosN(LblIdProveedor.Caption) & ""
    End If
    
    RST_Busq rsVerifDoc, vStr, xCon
    If rsVerifDoc.RecordCount >= 1 Then
        fVerifSiExistDocum = True
    Else
        fVerifSiExistDocum = False
    End If
    Set rsVerifDoc = Nothing
End Function

'*****************************************************************************************************
'* Nombre           : Buscar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : BUSCA UN REGISTRO EN EL RECORDSET RstIng
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Buscar()
    TabOne1.CurrTab = 0
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(6, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Num Serie":        xCampos(0, 1) = "numser":   xCampos(0, 2) = "1000":         xCampos(0, 3) = "N"
    xCampos(1, 0) = "Nº Doc.":          xCampos(1, 1) = "numdoc":   xCampos(1, 2) = "1200":         xCampos(1, 3) = "N"
    xCampos(2, 0) = "Nom. Prov/Clie":   xCampos(2, 1) = "nombre":   xCampos(2, 2) = "2100":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Tipo Mov.":        xCampos(3, 1) = "tipomov":  xCampos(3, 2) = "2000":         xCampos(3, 3) = "C"
    xCampos(4, 0) = "Tipo Doc.":        xCampos(4, 1) = "tipodoc":  xCampos(4, 2) = "2000":         xCampos(4, 3) = "C"
    xCampos(5, 0) = "IdDoc.":           xCampos(5, 1) = "Id":       xCampos(5, 2) = "1000":         xCampos(5, 3) = "N"
    
    xform.SQLCad = "SELECT alm_ingreso.numser, alm_ingreso.numdoc, alm_ingreso.nombre, IIF(alm_ingreso.tipmov = -1, 'Ingresos', 'Salidas') as tipomov, mae_documento.descripcion AS tipodoc, alm_ingreso.id " _
        & " FROM alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id "
    
    xform.Titulo = "Buscando Movimiento de items"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "numdoc"
    xform.CampoBusca = "numdoc"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        RstIng.MoveFirst
        RstIng.Find "id = " & xRs("id") & ""
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : fFuncEspecNomSolic
'* Tipo             : FUNCION
'* Descripcion      : DEVUELEVE EL NOMBRE Y APELLIDO COMPLETO DE UN EMPLEADO, DEVOLVERA VACIO EN CASO
'*                    DE NO TENER EXITO
'* Paranetros       : NOMBRE    |  TIPO         |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pId       |  LONG         |  ESPECIFICA EL ID DEL PERSONAL QUE SE ESTS BUSCANDO
'* Devuelve         : STRING
'*****************************************************************************************************
Private Function fFuncEspecNomSolic(pId As Long) As String
    Dim RsFunEsp As New ADODB.Recordset
    RsFunEsp.CursorLocation = adUseClient
    RST_Busq RsFunEsp, "SELECT id, LTRIM(UCASE(apepat)) +' '+LTRIM(UCASE(apemat)) + ', ' + LTRIM(nom) AS DatAJalar FROM pla_empleados WHERE id = " & pId & "", xCon
    
    If RsFunEsp.RecordCount > 0 Then
        If NulosC(RsFunEsp("DatAJalar")) <> "" Then
            fFuncEspecNomSolic = Trim(RsFunEsp("DatAJalar"))
        Else
            fFuncEspecNomSolic = ""
        End If
    Else
        fFuncEspecNomSolic = ""
    End If
    Set RsFunEsp = Nothing
End Function

Private Sub CmdAddItem_Click()
    ' PERMITE AGREGAR UN ITEM
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtTipCom.Text) = "" Then
        MsgBox "Seleccione el tipo de Producto que se está Ingresando", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipCom.SetFocus
        Exit Sub
    End If
    
    Dim fInsertar As Boolean
    If QueHace = 3 Then Exit Sub
    Agregando = True
    If Fg1.Rows > Fg1.FixedRows Then
        If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 4)) = 0 Then
            MsgBox "Seleccione un Producto", vbExclamation, xTitulo
        Else
            fInsertar = True
        End If
    Else
        fInsertar = True
    End If
    
    If fInsertar = True Then Fg1.AddItem ""
        
    Fg1.Row = Fg1.Rows - 1
    Fg1.Col = 1
    
    Fg1.SetFocus
    Agregando = False
End Sub

Private Sub CmdBusAlm_Click()
    ' PERMITE BUSCAR UN ALMACEN
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT alm_almacenes.* FROM alm_almacenes"
    
    xform.Titulo = "Buscando Almacenes"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdAlm.Text = xRs("id")
        LblDescDoc.Caption = xRs("descripcion")
        TxtNumSer.SetFocus
        
        If TxtTipDoc.Text <> "" Then
            Dim Rst As New ADODB.Recordset
            Set Rst = BuscaConCriterio("SELECT * FROM alm_numseries WHERE idtipdoc = " & NulosN(TxtTipDoc.Text) & " AND idalm = " & NulosN(TxtIdAlm.Text) & "", xCon)
            If Rst.RecordCount <> 0 Then
                TxtNumSer.Text = NulosC(Rst("numser"))
                TxtNumSer_Validate True
            End If
            Set Rst = Nothing
        Else
            TxtNumSer.Text = ""
            TxtNumDoc.Text = ""
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusArea_Click()
    ' PERMITE BUSCAR UN AREA
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":  xCampos(0, 2) = "4000":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":       xCampos(1, 1) = "id":           xCampos(1, 2) = "2000":   xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT id, descripcion FROM mae_area ORDER BY descripcion" _
    
    xform.Titulo = "Buscando el Area"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdArea.Text = xRs("id")
        LblArea.Caption = xRs("descripcion")
        Fg1.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusProv_Click()
    ' PERMITE BUSCAR UN PROVEEDOR O UN CLIENTE SEGUN SEA EL CASO
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(3, 4) As String
    
    xCampos(0, 0) = "Nombre":      xCampos(0, 1) = "nombre":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":   xCampos(1, 1) = "numruc":      xCampos(1, 2) = "1200":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Codigo":      xCampos(2, 1) = "id":          xCampos(2, 2) = "1000":         xCampos(2, 3) = "N"
    
    If OptIng.Value = True Then
        ' MOSTRARA LA LISTA DE PROVEEDORES
        xform.SQLCad = "SELECT mae_prov.id, mae_prov.numruc, mae_prov.nombre, mae_prov.activo From mae_prov " _
            & " Where (((mae_prov.activo) = -1)) ORDER BY mae_prov.nombre"
        xform.Titulo = "Buscando Proveedores"
    Else
        ' MOSTRARA LA LISTA DE CLIENTES
        xform.SQLCad = "SELECT mae_cliente.id, mae_cliente.numruc, mae_cliente.nombre, mae_cliente.activo From mae_cliente " _
            & " Where (((mae_cliente.activo) = -1)) ORDER BY mae_cliente.nombre"
        xform.Titulo = "Buscando Clientes"
    End If
    
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtProv.Text = xRs("nombre")
        LblIdProveedor.Caption = xRs("id")
        TxtIdRes.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusRes_Click()
    ' BUSCA EL RESPONSABLE QUE GENERA LA OPERACION
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Apellidos y Nombres":  xCampos(0, 1) = "apenom":     xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":               xCampos(1, 1) = "id":         xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT UCase([pla_empleados]![apepat]& ' ' & [pla_empleados]![apemat])& ', ' & [pla_empleados]![nom] AS apenom, pla_empleados.id " _
        & " From pla_empleados ORDER BY UCase([pla_empleados]![apepat])& ' ' & UCase([pla_empleados]![apemat])& ', ' & [pla_empleados]![nom]"
    
    xform.Titulo = "Buscando Responsables"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "apenom"
    xform.CampoBusca = "apenom"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdRes.Text = xRs("id")
        LblResp.Caption = xRs("apenom")
    Fg1.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipDoc_Click()
    ' BUSCA EL TIPO DE DOCUMENTO
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_documento.* FROM mae_documento WHERE (tipo = 1 OR tipo = 3)"
    
    xform.Titulo = "Buscando Tipo de Documento"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtTipDoc.Text = xRs("id")
        LblTipDoc.Caption = NulosC(xRs("descripcion"))
        TxtIdAlm.SetFocus
        If NulosN(TxtIdAlm.Text) <> 0 Then
            Dim Rst As New ADODB.Recordset
            Set Rst = BuscaConCriterio("SELECT * FROM alm_numseries WHERE idtipdoc = " & NulosN(TxtTipDoc.Text) & " AND idalm = " & NulosN(TxtIdAlm.Text) & "", xCon)
            If Rst.RecordCount <> 0 Then
                TxtNumSer.Text = NulosC(Rst("numser"))
                TxtNumSer_Validate True
            Else
                TxtNumSer.Text = ""
                TxtNumDoc.Text = ""
            End If
            Set Rst = Nothing
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipoCompra_Click()
    ' BUSCA EL TIPO DE PRODUCTO
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_tipoproducto.* FROM mae_tipoproducto"
    
    xform.Titulo = "Buscando Tipo de Producto"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtTipCom.Text = xRs("id")
        LblTipoCompra.Caption = xRs("descripcion")
        TxtTipDoc.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdDelItem_Click()
    ' ELIMINA UNA FILA DEL CONTROL FlexGrid Fg1
    If QueHace = 3 Then Exit Sub
    
    If Fg1.Row < 1 Then
        MsgBox "Elija el registro correcto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    If Fg1.Rows = 1 Then
        MsgBox "No hay Registro para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.SetFocus
        Exit Sub
    End If
    If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    
    Fg1.RemoveItem Fg1.Row
    
    If Fg1.Rows > 1 Then
        Fg1.Row = Fg1.Rows - 1
        Fg1.Col = 1
        Fg1.SetFocus
    Else
        CmdAddItem.SetFocus
    End If
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstIng
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA LA COLUMNA SELECCIONADA DEL CONTROL DataGrid Dg1
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstIng.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        If fCierrePeriodo = False Then Exit Sub
        Nuevo
    End If
    If KeyCode = 46 Then
        If fCierrePeriodo = False Then Exit Sub
        Eliminar
    End If
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then VerMovimientos1 IdMenuActivo, NulosN(RstIng("id")), xCon
End Sub

Private Sub fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtTipCom.Text) = "" Then
        MsgBox "Seleccione el tipo de producto que se esta iungresando", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipCom.SetFocus
        Exit Sub
    End If
    
    If Col = 1 Then
        ' BUSCA UN ITEM
        Dim xform As New eps_librerias.FormBuscar
        Dim xRs As New ADODB.Recordset
        Dim nSQLTmp As String
        Dim xCampos(3, 4) As String
        
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Producto":    xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":      xCampos(1, 1) = "codpro":        xCampos(1, 2) = "1500":         xCampos(1, 3) = "C"
        xCampos(2, 0) = "Uni. Med":    xCampos(2, 1) = "abrev":         xCampos(2, 2) = "1200":         xCampos(2, 3) = "C"
        
        If NulosC(Fg1.TextMatrix(Row, Col)) <> "" Then
            nSQLTmp = " AND UCASE(alm_inventario.descripcion) LIKE '%" & UCase(NulosC(Fg1.TextMatrix(Row, Col))) & "%'"
        End If
        
        xform.SQLCad = "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, " _
            & " alm_inventario.stckact, alm_inventario.activo FROM mae_unidades RIGHT JOIN alm_inventario ON " _
            & " mae_unidades.id = alm_inventario.idunimed " _
            & " WHERE (((alm_inventario.activo)=-1) AND ((alm_inventario.tippro)=" & NulosN(TxtTipCom.Text) & ")) " & nSQLTmp & " ORDER BY alm_inventario.codpro"
        
        xform.Titulo = "Buscando Productos"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "descripcion"
        xform.CampoBusca = "descripcion"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        If xRs.State = 1 Then
            Fg1.TextMatrix(Row, 1) = NulosC(xRs("descripcion"))
            Fg1.TextMatrix(Row, 2) = NulosC(xRs("abrev"))
            Fg1.TextMatrix(Row, 4) = NulosN(xRs("id"))
            Fg1.Col = 3
            Fg1.SetFocus
        End If
        Set xform = Nothing
        Set xRs = Nothing
    End If
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If Col = 3 Then
        Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), "0.0000")
    End If
End Sub

Private Sub Fg1_EnterCell()
    If Fg1.Col = 1 Or Fg1.Col = 3 Then
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = 3 Then
        If KeyAscii <> 13 Then If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    
    If KeyCode = 45 Then
        CmdAddItem_Click
    End If
    If KeyCode = 46 Then
        CmdDelItem_Click
    End If
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If QueHace = 3 Then Exit Sub
    If Button = 2 Then PopupMenu menu1
End Sub

Private Sub Form_Activate()
'Modificado 13/01/11 Johan Castro
'           Eliminar

    ' SEGUNDO EVENTO AL CARGARSE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        
        mMesActivo = xMes
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        'CONSULTA ANTERIOR
        
'        RST_Busq RstIng, "SELECT alm_ingreso.*, mae_documento.abrev, mae_documento.descripcion, alm_ingreso!numser+'-'+alm_ingreso!numdoc AS numdoc2, " _
'            & " UCase(pla_empleados!apepat) + ' ' + UCase(pla_empleados!apemat) + ', ' + pla_empleados!nom AS nomres, IIf(alm_ingreso!tipmov=-1,'Ingreso','Salida') AS movi, alm_almacenes.descripcion AS descalm, " _
'            & " alm_ingreso.idsol AS idsolicitante, alm_ingreso.idare AS idarea, mae_area.descripcion AS nomarea, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoccom, " _
'            & " com_compras.numreg,alm_ingreso.fching & '' as fching1,alm_ingreso.fchdoc & '' as fchdoc1 FROM ((((alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id) LEFT JOIN pla_empleados " _
'            & " ON alm_ingreso.idres = pla_empleados.id) LEFT JOIN alm_almacenes ON alm_ingreso.idalm = alm_almacenes.id) LEFT JOIN mae_area ON alm_ingreso.idare = mae_area.id) " _
'            & " LEFT JOIN com_compras ON alm_ingreso.idfac = com_compras.id ORDER BY alm_ingreso.fching DESC", xCon
'
'        Set Dg1.DataSource = RstIng
        
        
        OpcionesPeriodo
        
    End If
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
    ActivaTool
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Agregando Movimiento"
    Bloquea
    Blanquea
    Fg1.Editable = flexEDKbdMouse
    Fg1.Rows = 1
    Fg1.Rows = Fg1.Rows + 1
    Fg1.ColComboList(1) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    
    OptIng.Value = True
    TxtTipDoc.Text = "71"
    TxtTipDoc_Validate True
    xHorIni = Time
    TxtFchIng.SetFocus
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A AJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    QueHace = 3
    'Dg1.Columns("fching1").NumberFormat = FORMAT_DATE
    Dg1.Columns("fchdoc1").NumberFormat = FORMAT_DATE
    TxtFchIng.Valor = Date
    TxtFchDoc.Valor = Date
    
    TxtFchIng.Valor = ""
    TxtFchDoc.Valor = ""
    Fg1.ColWidth(4) = 0
    SeEjecuto = False
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    CaracteresNumericos = "0123456789." & Chr(8)
    CaracteresNumericos2 = "0123456789." & Chr(8) & Chr(13)
    Mostrando = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario miestras este agregando o editando un registro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub Menu1_1_Click()
    CmdAddItem_Click
End Sub

Private Sub Menu1_3_Click()
    Fg1.RemoveItem Fg1.Row
End Sub

Private Sub OptIng_Click()
    ' PREPARAMOS EL FORMULARIO PARA UN INGRESO
    TxtTipDoc.Text = "71"
    TxtTipDoc_Validate True
    Label3(3).Caption = "Fch. Ingreso"
    Label33.Caption = "Proveedor"
    
    If Mostrando = False Then
        TxtProv.Text = ""
        LblIdProveedor.Caption = ""
    End If
    TxtIdArea.Locked = True
    CmdBusArea.Enabled = False
    TxtIdArea.Text = "": LblArea.Caption = ""
End Sub

Private Sub OptIng_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then TxtFchIng.SetFocus
End Sub

Private Sub OptSal_Click()
    ' PREPARAMOS EL FORMULARIO PARA UNA SALIDA
    TxtTipDoc.Text = "70"
    TxtTipDoc_Validate True
    Label3(3).Caption = "Fch. Salida"
    Label33.Caption = "Cliente"
    If Mostrando = False Then
        TxtProv.Text = ""
        LblIdProveedor.Caption = ""
    End If
    TxtIdArea.Locked = False
    CmdBusArea.Enabled = True
End Sub

Private Sub OptSal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then TxtFchIng.SetFocus
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If RstIng.State = 0 Then Exit Sub
        If RstIng.RecordCount = 0 And QueHace <> 1 Then
            Cancel = 1
            Exit Sub
        End If
        If QueHace = 3 Then MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            RstIng.Requery
            Dg1.Refresh
            Cancelar
            
            If RstIng.RecordCount <> 0 Then
                RstIng.MoveFirst
                RstIng.Find "id=" & mIdRegistro
                If RstIng.EOF = True Then RstIng.MoveFirst
            End If
            
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then
'        TDB_FiltroLimpiar Dg1
'        RstIng.Requery
'        Set Dg1.DataSource = RstIng
        TDB_Actualizar Me, TabOne1, Dg1, RstIng
    End If
    
    If Button.Index = 10 Then Buscar
    
    If Button.Index = 11 Then
        mMesActivo = SeleccionaMes(xCon)
        OpcionesPeriodo
    End If
        
    If Button.Index = 12 Then FrmConsIngAlmacen.Show
    
    If Button.Index = 16 Then
        Unload Me
        Set RstIng = Nothing
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL INGRESO O MODIFICACION DE UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    Bloquea
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    ActivaTool
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : Function
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA Alm_ingreso, DEVUELVE VERDADERO SI TIENE EXITO
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Function Grabar() As Boolean
    ' VALIDAMOS QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
    If TxtFchIng.Valor = "" Then
        MsgBox "No ha especificado la fecha de ingreso", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIng.SetFocus
        Exit Function
    End If
    
    If IsDate(TxtFchIng.Valor) = False Then
        MsgBox "La fecha de ingreso esta mal ingresado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIng.SetFocus
        Exit Function
    End If
    
    If TxtFchDoc.Valor = "" Then
        MsgBox "No ha especificado la fecha del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchDoc.SetFocus
        Exit Function
    End If
    
    If IsDate(TxtFchDoc.Valor) = False Then
        MsgBox "La fecha del documento está mal ingresado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchDoc.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtTipCom.Text) = "" Then
        MsgBox "No ha especificado el tipo de item", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipCom.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtTipDoc.Text) = "" Then
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtNumSer.Text) = "" Then
        MsgBox "No ha especificado el número de serie del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumSer.SetFocus
        Exit Function
    End If
    If NulosC(TxtNumDoc.Text) = "" Then
        MsgBox "No ha especificado el número de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDoc.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtProv.Text) = "" Then
        MsgBox "No ha especificado el nombre del proveedor", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtProv.SetFocus
        Exit Function
    End If
    
    If OptIng.Value = True Then
        If Trim(TxtTipDoc.Text) <> "" Then
            If NulosN(TxtTipDoc.Text) = 9 Then
                If Trim(LblIdProveedor.Caption) = "" Then
                    MsgBox "Falta especificar el proveedor", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    TxtProv.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    
    If NulosC(TxtIdRes.Text) = "" Then
        MsgBox "No ha especificado el responsable el movimiento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdRes.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtIdAlm.Text) = "" Then
        MsgBox "No ha especificado el almacen", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdAlm.SetFocus
        Exit Function
    End If
    
    If OptSal.Value = True Then
        If NulosC(TxtIdArea.Text) = "" Then
            MsgBox "No ha especificado el area", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtIdArea.SetFocus
            Exit Function
        End If
    End If
    
    If Fg1.Rows = 2 Then
        If Fg1.TextMatrix(1, 1) = "" Then
            MsgBox "Falta ingresar el detalle...!", vbExclamation, "Mensaje...!"
            Exit Function
        End If
    ElseIf Fg1.Rows = 1 Then
        MsgBox "Falta ingresar el detalle...!", vbExclamation, "Mensaje...!"
        Exit Function
    End If
    
    Dim xId As Double
    Dim A As Integer
    
    For A = 1 To Fg1.Rows - 1
        If NulosN(Format(Fg1.TextMatrix(A, 3), "####0.00")) = 0 Then
            MsgBox "Falta ingresar la cantidad en el detalle", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Fg1.SetFocus
            Exit Function
        End If
    Next
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    
    
    On Error GoTo LaCague
    xCon.BeginTrans
    
    If QueHace = 1 Then 'NUEVO
        If fVerifSiExistDocum = True Then
            MsgBox "El numero de documento ingresado ya existe.", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtNumDoc.SetFocus
            Exit Function
        End If
        xId = HallaCodigoTabla("alm_ingreso", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM alm_ingreso", xCon
        RST_Busq RstDet, "SELECT TOP 1 * FROM alm_ingresodet", xCon
        RstCab.AddNew
        RstCab("id") = xId
        
    Else 'MODIFICAR
        xId = RstIng("id")
        RST_Busq RstCab, "SELECT * FROM alm_ingreso WHERE id = " & RstIng("id") & "", xCon
        xCon.Execute "DELETE * FROM alm_ingresodet WHERE id = " & RstIng("id") & " "
        RST_Busq RstDet, "SELECT * FROM alm_ingresodet", xCon
    End If
    
    mIdRegistro = xId
    
    RstCab("tipcom") = NulosN(TxtTipCom.Text)
    RstCab("tipdoc") = NulosN(TxtTipDoc.Text)
    RstCab("fching") = CDate(TxtFchIng.Valor)
    RstCab("fchdoc") = CDate(TxtFchDoc.Valor)
    RstCab("numser") = NulosC(TxtNumSer.Text)
    RstCab("numdoc") = NulosC(TxtNumDoc.Text)
    RstCab("idres") = NulosN(TxtIdRes.Text)
    RstCab("idalm") = NulosN(TxtIdAlm.Text)
    If NulosN(LblIdProveedor.Caption) <> 0 Then RstCab("idpro") = NulosN(LblIdProveedor.Caption)
    RstCab("nombre") = Trim(TxtProv.Text)
    
    If OptIng.Value = True Then
        RstCab("tipmov") = -1
    Else
        RstCab("tipmov") = 0
    End If
    RstCab("idare") = NulosN(TxtIdArea.Text)
    
    RstCab("ano") = AnoTra
    RstCab("idmes") = mMesActivo
    
    RstCab.Update
    
    ' GRABAMOS EL DETALLE DEL MOVIMIENTO
    For A = 1 To Fg1.Rows - 1
        RstDet.AddNew
        RstDet("id") = xId
        RstDet("iditem") = NulosN(Fg1.TextMatrix(A, 4))
        RstDet("cantidad") = NulosN(Fg1.TextMatrix(A, 3))
        RstDet.Update
    Next A
    
    ' grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    xCon.CommitTrans
    
    Grabar = True
    MsgBox "El registro se guardo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Exit Function

LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO DE LA TABLA alm_ingreso
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    Dim Rpta As Integer
    TabOne1.CurrTab = 0
    If RstIng.State = 0 Then Exit Sub
    If RstIng.RecordCount = 0 Then
        MsgBox "No hay Registros de Ingreso/Salida de Almacén para eliminar", vbInformation, xTitulo
        Exit Sub
    End If
    
    Rpta = MsgBox("¿Esta seguro de eliminar el ingreso Nº " + Trim(RstIng("numdoc")) + "?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM alm_ingreso WHERE id = " & NulosN(RstIng("id")) & ""
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstIng("id") & " AND idform = " & IdMenuActivo

        RstIng.Requery
        Dg1.Refresh
        MsgBox "El ingreso se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Modificar()
    TabOne1.CurrTab = 1
    ActivaTool
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Modificando Movimiento"
    QueHace = 2
    Bloquea
    Blanquea
 
    OptIng.Value = True
    Fg1.Editable = flexEDKbdMouse
    Fg1.Rows = 1
    Fg1.Rows = Fg1.Rows + 1
    Fg1.ColComboList(1) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    MuestraSegundoTab
    xHorIni = Time
    TxtFchIng.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA Y DESACTIVA LA BARRA DE HERRAMIENTAS
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
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : BLANQUEA LOS CONTROLES TextBox, PARA EL INGRESO DE UN NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    TxtFchIng.Valor = ""
    TxtFchDoc.Valor = ""
    TxtTipCom.Text = ""
    TxtTipDoc.Text = ""
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    TxtProv.Text = ""
    TxtIdRes.Text = ""
    LblTipoCompra.Caption = ""
    LblResp.Caption = ""
    LblTipDoc.Caption = ""
    LblIdProveedor.Caption = ""
    TxtIdAlm.Text = ""
    LblDescDoc.Caption = ""
    TxtIdArea.Text = ""
    LblArea.Caption = ""
End Sub

'*****************************************************************************************************
'* Nombre           : Bloquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA Y DESACTIVA LOS CONTROLES TEXTBOX, PREPARA PARA AGREGAR O MODIFICAR UN
'*                    REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Bloquea()
    TxtFchIng.Locked = Not TxtFchIng.Locked
    TxtFchDoc.Locked = Not TxtFchDoc.Locked
    TxtTipCom.Locked = Not TxtTipCom.Locked
    TxtTipDoc.Locked = Not TxtTipDoc.Locked
    TxtNumSer.Locked = Not TxtNumSer.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    TxtProv.Locked = Not TxtProv.Locked
    TxtIdRes.Locked = Not TxtIdRes.Locked
    TxtIdAlm.Locked = Not TxtIdAlm.Locked
    CmdAddItem.Enabled = Not CmdAddItem.Enabled
    CmdDelItem.Enabled = Not CmdDelItem.Enabled
End Sub

Private Sub TxtAlmacen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtAlmacen_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusAlm_Click
    End If
End Sub

Private Sub TxtIdAlm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        TxtIdAlm_Validate True
    End If
End Sub

Private Sub TxtIdAlm_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusAlm_Click
    End If
End Sub

Private Sub TxtIdAlm_Validate(Cancel As Boolean)
    If NulosN(TxtIdAlm.Text) = 0 Then Exit Sub
    
    LblDescDoc.Caption = Busca_Codigo(TxtIdAlm.Text, "id", "descripcion", "alm_almacenes", "N", xCon)
    If LblDescDoc.Caption = "" Then
        TxtIdAlm.Text = ""
    End If
End Sub

Private Sub TxtIdArea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CmdBusArea.Enabled = False Then Exit Sub
        SendKeys vbTab
        
        If NulosC(TxtIdArea.Text) = "" Then Exit Sub
        Dim xRs As New ADODB.Recordset
        xRs.CursorLocation = adUseClient
        RST_Busq xRs, "SELECT id, descripcion FROM mae_area WHERE id = " & NulosN(TxtIdArea.Text) & "", xCon
        
        If xRs.RecordCount = 0 Then
            TxtIdArea.Text = ""
            LblArea.Caption = ""
        Else
            LblArea.Caption = xRs("descripcion")
        End If
        Set xRs = Nothing
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdArea_KeyUp(KeyCode As Integer, Shift As Integer)
    If CmdBusArea.Enabled = False Then Exit Sub
    If KeyCode = 116 Then  'TECLA F5
        CmdBusArea.Value = True
    End If
End Sub

Private Sub TxtIdRes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        If NulosC(TxtIdRes.Text) = "" Then Exit Sub
        Dim xRs As New ADODB.Recordset
        xRs.CursorLocation = adUseClient
        RST_Busq xRs, "SELECT id, ltrim(apepat) + ' '+ltrim(apemat) + ', ' + ltrim(nom) AS nomsolic FROM pla_empleados WHERE id = " & NulosN(TxtIdRes.Text) & "", xCon
        
        If xRs.RecordCount = 0 Then
            TxtIdRes.Text = ""
            LblResp.Caption = ""
        Else
            LblResp.Caption = NulosC(xRs("nomsolic"))
        End If
        Set xRs = Nothing
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdRes_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusRes_Click
    End If
End Sub

Private Sub TxtIdRes_Validate(Cancel As Boolean)
    If TxtIdRes.Text = "" Then Exit Sub
    
    LblResp.Caption = Busca_Codigo(NulosN(TxtIdRes.Text), "id", "apepat", "pla_empleados", "N", xCon)
    LblResp.Caption = UCase(Trim(LblResp.Caption)) + ", " + Trim(Busca_Codigo(NulosN(TxtIdRes.Text), "id", "nom", "pla_empleados", "N", xCon))
    If NulosC(LblResp.Caption) = "" Then
        TxtIdRes.Text = ""
    End If
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumDoc_Validate(Cancel As Boolean)
    If NulosC(TxtNumDoc.Text) <> "" Then
        TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000000")
    End If
End Sub

Private Sub TxtNumSer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumSer_Validate(Cancel As Boolean)
    If NulosC(TxtNumSer.Text) <> "" Then
        If TxtNumDoc.Text = "" Then
            TxtNumSer.Text = Format(TxtNumSer.Text, "0000")
            TxtNumDoc.Text = HallarNumIngresoAlmacen(TxtNumSer)
        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : HallarNumIngresoAlmacen
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : DEVUELVE EL NUMERO DE DOCUMENTO ACTUAL EN FUNCION AL NUMERO DE SERIE, DEVUELVE
'*                    UNA CADENA
'* Paranetros       : NOMBRE    |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    NumSerie  |  STRING     |  ESPECIFICA EL NUMERO DE SERIE
'* Devuelve         : STRING
'*****************************************************************************************************
Function HallarNumIngresoAlmacen(NumSerie As String) As String
    Dim vFiltro As String
    Dim Rst As New ADODB.Recordset
    Dim xNum As Double
    If OptIng.Value = True Then
        If NulosN(TxtTipDoc.Text) <> 9 Then
            vFiltro = " AND tipdoc = " & NulosN(TxtTipDoc.Text) & ""
        End If
    Else
        vFiltro = " AND tipdoc = " & NulosN(TxtTipDoc.Text) & ""
    End If

    vStr = "SELECT * FROM alm_ingreso WHERE numser = '" & NulosC(NumSerie) & "'" & vFiltro
    If OptIng.Value = True Then
        vStr = vStr & " AND tipmov = true"
    Else
        vStr = vStr & " AND tipmov = false"
    End If
    vStr = vStr & " ORDER BY numdoc"
    RST_Busq Rst, vStr, xCon
    
    If Rst.RecordCount = 0 Then
        ' SI ESTA VACIO INICIALIZA LA NUMERACION
        If OptIng.Value = True And NulosN(TxtTipDoc.Text) <> 9 Then
            HallarNumIngresoAlmacen = "0000000001"
        ElseIf OptSal.Value = True Then
            HallarNumIngresoAlmacen = "0000000001"
        Else
            HallarNumIngresoAlmacen = ""
        End If
    Else
        ' SI ESTA LLENO SUMA UNO AL ULTIMO NUMERO
        Rst.MoveLast
        If OptIng.Value = True And NulosN(TxtTipDoc.Text) <> 9 Then
            xNum = NulosN(Rst("numdoc")) + 1
            HallarNumIngresoAlmacen = Format(xNum, "0000000000")
        ElseIf OptSal.Value = True Then
            xNum = NulosN(Rst("numdoc")) + 1
            HallarNumIngresoAlmacen = Format(xNum, "0000000000")
        Else
            HallarNumIngresoAlmacen = ""
        End If
    End If
    Set Rst = Nothing
End Function

Private Sub TxtProv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        LblIdProveedor.Caption = ""
    End If
End Sub

Private Sub TxtProv_KeyUp(KeyCode As Integer, Shift As Integer)
'    MsgBox KeyCode
    If KeyCode = 116 Then
        CmdBusProv_Click
    End If
End Sub

Private Sub TxtTipCom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipCom_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipoCompra_Click
    End If
End Sub

Private Sub TxtTipCom_Validate(Cancel As Boolean)
    If TxtTipCom.Text = "" Then Exit Sub
    
    LblTipoCompra.Caption = Busca_Codigo(NulosN(TxtTipCom.Text), "id", "descripcion", "mae_tipoproducto", "N", xCon)
    If NulosC(LblTipoCompra.Caption) = "" Then
        TxtTipCom.Text = ""
    End If
End Sub

Private Sub TxtTipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipDoc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipDoc_Click
    End If
End Sub

Private Sub TxtTipDoc_Validate(Cancel As Boolean)
    If TxtTipDoc.Text = "" Then Exit Sub
    
    LblTipDoc.Caption = Busca_Codigo(NulosN(TxtTipDoc.Text), "id", "descripcion", "mae_documento", "N", xCon)
    If NulosC(LblTipDoc.Caption) = "" Then
        TxtTipDoc.Text = ""
    End If
    
    If NulosN(TxtIdAlm.Text) <> 0 Then
        Dim Rst As New ADODB.Recordset
        Set Rst = BuscaConCriterio("SELECT * FROM alm_numseries WHERE idtipdoc = " & NulosN(TxtTipDoc.Text) & " AND idalm = " & NulosN(TxtIdAlm.Text) & "", xCon)
        If Rst.RecordCount <> 0 Then
            TxtNumSer.Text = Rst("numser")
            TxtNumSer_Validate True
        End If
        Set Rst = Nothing
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EL DETALLE DEL REGISTRO ACTUAL EN LA PESTAÑA DETALLE DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    If RstIng.RecordCount = 0 Then Exit Sub
    If RstIng.BOF = True Or RstIng.EOF = True Then Exit Sub
    TxtFchIng.Valor = RstIng("fching")
    TxtFchDoc.Valor = RstIng("fchdoc")
    TxtTipCom.Text = RstIng("tipcom")
    TxtTipDoc.Text = RstIng("tipdoc")
    TxtNumSer.Text = RstIng("numser")
    TxtNumDoc.Text = RstIng("numdoc")
    TxtProv.Text = RstIng("nombre")
    LblIdProveedor.Caption = NulosN(RstIng("idpro"))
    TxtIdRes.Text = RstIng("idres")
    LblResp.Caption = NulosC(RstIng("nomres"))
    LblTipoCompra.Caption = Busca_Codigo(NulosN(TxtTipCom.Text), "id", "descripcion", "mae_tipoproducto", "N", xCon)
    LblTipDoc.Caption = Busca_Codigo(NulosN(TxtTipDoc.Text), "id", "descripcion", "mae_documento", "N", xCon)
    TxtIdAlm.Text = RstIng("idalm")
    LblDescDoc.Caption = Busca_Codigo(NulosN(TxtIdAlm.Text), "id", "descripcion", "alm_almacenes", "N", xCon)
    Mostrando = True
    
    If RstIng("tipmov") = -1 Then
        OptIng.Value = True
    Else
        OptSal.Value = True
    End If
    
    If NulosN(RstIng("idare")) <= 0 Then
        TxtIdArea.Text = ""
        LblArea.Caption = ""
    Else
        TxtIdArea.Text = NulosN(RstIng("idare"))
        LblArea.Caption = NulosC(RstIng("nomarea"))
    End If
    
    Mostrando = False
    
    Dim RstDet As New ADODB.Recordset
    Dim A As Integer
    
    RST_Busq RstDet, "SELECT alm_ingresodet.*, alm_inventario.descripcion, mae_unidades.abrev, alm_ingresodet.id " _
        & " FROM mae_unidades RIGHT JOIN (alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) " _
        & " ON mae_unidades.id = alm_inventario.idunimed Where (((alm_ingresodet.id) = " & RstIng("id") & "))" _
        & " ", xCon
    Fg1.Rows = 1
    If RstDet.RecordCount <> 0 Then
        RstDet.MoveFirst
        For A = 1 To RstDet.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = NulosC(RstDet("descripcion"))
            Fg1.TextMatrix(A, 2) = NulosC(RstDet("abrev"))
            Fg1.TextMatrix(A, 3) = NulosN(RstDet("cantidad"))
            Fg1.TextMatrix(A, 4) = NulosN(RstDet("iditem"))
            RstDet.MoveNext
            If RstDet.EOF = True Then
                Exit For
            End If
        Next A
    End If
End Sub





Sub OpcionesPeriodo()
    '===================================================================================================
    'Creado : 13/01/11 Por: Johan Castro
    'Propósito: 1.- Cargar datos en lista.
    '           2.- Controlar el acceso segun el usuario
    '
    'Entradas:  Ninguna
    '
    'Resultados:formulario bloqueado o listo para editar
    '
    '===================================================================================================


     Dim NomMes As String
     Dim Cerrado As Boolean
     
     LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    '------------------------------------------------------------------------------------------
    ' bloqueamos los botones del toolbar
    CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
    '------------------------------------------------------------------------------------------
    TDB_FiltroLimpiar Dg1
    Set RstIng = Nothing
    '------------------------------------------
    
    LblPeriodo.Caption = LblMes.Caption
    LblPeriodo2.Caption = LblPeriodo.Caption
    
    
    '------------------------------------------
    
    RST_Busq RstIng, "SELECT alm_ingreso.*, mae_documento.abrev, mae_documento.descripcion, alm_ingreso!numser+'-'+alm_ingreso!numdoc AS numdoc2, " _
            & " UCase(pla_empleados!apepat) + ' ' + UCase(pla_empleados!apemat) + ', ' + pla_empleados!nom AS nomres, IIf(alm_ingreso!tipmov=-1,'Ingreso','Salida') AS movi, alm_almacenes.descripcion AS descalm, " _
            & " alm_ingreso.idsol AS idsolicitante, alm_ingreso.idare AS idarea, mae_area.descripcion AS nomarea, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoccom, " _
            & " com_compras.numreg,alm_ingreso.fching & '' as fching1,alm_ingreso.fchdoc & '' as fchdoc1 FROM ((((alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id) LEFT JOIN pla_empleados " _
            & " ON alm_ingreso.idres = pla_empleados.id) LEFT JOIN alm_almacenes ON alm_ingreso.idalm = alm_almacenes.id) LEFT JOIN mae_area ON alm_ingreso.idare = mae_area.id) " _
            & " LEFT JOIN com_compras ON alm_ingreso.idfac = com_compras.id " _
            & " WHERE (((alm_ingreso.ano)=" & AnoTra & ") AND ((alm_ingreso.idmes)=" & mMesActivo & ")) ORDER BY alm_ingreso.fching DESC ", xCon
        
        Set Dg1.DataSource = RstIng
    
    Set Dg1.DataSource = RstIng
    
End Sub


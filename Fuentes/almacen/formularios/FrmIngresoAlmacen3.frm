VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmIngresoAlmacen3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Almacén - Movimientos de Almacén"
   ClientHeight    =   7410
   ClientLeft      =   165
   ClientTop       =   1590
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   11835
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
            Picture         =   "FrmIngresoAlmacen3.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen3.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen3.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen3.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen3.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen3.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen3.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen3.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen3.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen3.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen3.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen3.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen3.frx":277E
            Key             =   "IMG13"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7020
      Left            =   0
      TabIndex        =   16
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
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6600
         Left            =   45
         TabIndex        =   21
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6090
            Left            =   30
            TabIndex        =   22
            Top             =   480
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   10742
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
            Columns(1).Caption=   "Fch. Mov."
            Columns(1).DataField=   "fchdoc"
            Columns(1).NumberFormat=   "Short Date"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Mov."
            Columns(2).DataField=   "movi"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Nº Documento"
            Columns(3).DataField=   "numdoc"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Cliente/Proveedor"
            Columns(4).DataField=   "nombre"
            Columns(4).NumberFormat=   "Short Date"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Almacén"
            Columns(5).DataField=   "desalm"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "T.D. Ref."
            Columns(6).DataField=   "destipdocref"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Nº Doc. Ref."
            Columns(7).DataField=   "numdocref"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Estado"
            Columns(8).DataField=   "desestado"
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
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1667"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1588"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=1058"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=979"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=131585"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=2646"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2566"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=4974"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=4895"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=131588"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=3334"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=3254"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=1296"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1217"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=513"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=2566"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=2487"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=516"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(50)=   "Column(8).Width=2090"
            Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=2011"
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
            HeadLines       =   1
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
            _StyleDefs(24)  =   "Splits(0).Style:id=75,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=84,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=76,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=77,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=78,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=80,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=79,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=81,.parent=8,.bgcolor=&H80&"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=82,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=83,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=85,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=86,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=90,.parent=75"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=76"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=77"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=79"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=98,.parent=75"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=95,.parent=76"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=96,.parent=77"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=97,.parent=79"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=102,.parent=75,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=99,.parent=76"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=100,.parent=77,.alignment=2"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=101,.parent=79"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=110,.parent=75"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=107,.parent=76"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=108,.parent=77"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=109,.parent=79"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=118,.parent=75,.alignment=3"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=115,.parent=76,.alignment=2"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=116,.parent=77,.alignment=2"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=117,.parent=79"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=16,.parent=75"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=13,.parent=76"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=14,.parent=77"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=15,.parent=79"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=28,.parent=75,.alignment=2"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=76"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=77"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=79"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=122,.parent=75"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=119,.parent=76"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=120,.parent=77"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=121,.parent=79"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=126,.parent=75"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=123,.parent=76"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=124,.parent=77"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=125,.parent=79"
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
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label LblPeriodo 
            Alignment       =   2  'Center
            Caption         =   "LblPeriodo"
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
            TabIndex        =   23
            Top             =   30
            Width           =   1860
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Consulta de Movimientos ( Ingresos/Salidas )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
            TabIndex        =   24
            Top             =   30
            Width           =   11610
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6600
         Left            =   12525
         TabIndex        =   17
         Top             =   375
         Width           =   11790
         Begin VB.Frame Frame7 
            Caption         =   "[ Movimientos Relacionados ]"
            ForeColor       =   &H00800000&
            Height          =   2390
            Left            =   120
            TabIndex        =   40
            Top             =   4080
            Width           =   11565
            Begin VSFlex7Ctl.VSFlexGrid fg 
               Height          =   2025
               Index           =   1
               Left            =   60
               TabIndex        =   41
               Top             =   270
               Width           =   11385
               _cx             =   20082
               _cy             =   3572
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
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   11
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmIngresoAlmacen3.frx":2B10
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
         Begin VB.OptionButton OptIng 
            Caption         =   "Ingreso"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   1065
            TabIndex        =   38
            Top             =   390
            Width           =   1080
         End
         Begin VB.OptionButton OptSal 
            Caption         =   "Salida"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   2280
            TabIndex        =   37
            Top             =   390
            Width           =   1000
         End
         Begin VB.CommandButton cmd 
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   11460
            Picture         =   "FrmIngresoAlmacen3.frx":2C83
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   420
            Width           =   240
         End
         Begin VB.TextBox GlosaText 
            Height          =   300
            Left            =   7065
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   12
            Text            =   "GlosaText"
            Top             =   1305
            Width           =   4640
         End
         Begin VB.CommandButton cmd 
            Enabled         =   0   'False
            Height          =   240
            Index           =   0
            Left            =   1730
            Picture         =   "FrmIngresoAlmacen3.frx":2DB5
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1335
            Width           =   240
         End
         Begin VB.CommandButton cmd 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   7710
            Picture         =   "FrmIngresoAlmacen3.frx":2EE7
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   720
            Width           =   240
         End
         Begin VB.CommandButton cmd 
            Enabled         =   0   'False
            Height          =   240
            Index           =   3
            Left            =   11430
            Picture         =   "FrmIngresoAlmacen3.frx":3019
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1020
            Width           =   240
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1070
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   2
            Text            =   "TxtNumSer"
            Top             =   990
            Width           =   915
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2265
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   3
            Text            =   "TxtNumDoc"
            Top             =   990
            Width           =   1440
         End
         Begin VSFlex7Ctl.VSFlexGrid fg 
            Height          =   2355
            Index           =   0
            Left            =   105
            TabIndex        =   15
            Top             =   1680
            Width           =   10080
            _cx             =   17780
            _cy             =   4154
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
            Rows            =   2
            Cols            =   9
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmIngresoAlmacen3.frx":314B
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
            Left            =   1065
            TabIndex        =   0
            Top             =   690
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
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchDoc 
            Height          =   300
            Left            =   4455
            TabIndex        =   1
            Top             =   690
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
         Begin VB.TextBox txtNumDocRef 
            Height          =   300
            Left            =   7065
            Locked          =   -1  'True
            TabIndex        =   10
            Text            =   "txtNumDocRef"
            Top             =   990
            Width           =   4635
         End
         Begin VB.TextBox TxtIdTipDocRef 
            Height          =   300
            Left            =   7065
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   8
            Text            =   "TxtTipDocR"
            Top             =   690
            Width           =   915
         End
         Begin VB.TextBox txtIdAlm 
            Height          =   300
            Left            =   1070
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   4
            Text            =   "txtIdAlm"
            Top             =   1305
            Width           =   915
         End
         Begin VB.TextBox TxtProv 
            Height          =   300
            Left            =   7065
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   6
            Text            =   "TxtProv"
            Top             =   390
            Width           =   4660
         End
         Begin VB.Frame Frame4 
            Height          =   2485
            Left            =   10250
            TabIndex        =   18
            Top             =   1560
            Width           =   1450
            Begin VB.CommandButton cmd 
               Caption         =   "Eliminar Item"
               Enabled         =   0   'False
               Height          =   330
               Index           =   5
               Left            =   50
               TabIndex        =   14
               Top             =   600
               Width           =   1305
            End
            Begin VB.CommandButton cmd 
               Caption         =   "Agregar Item"
               Enabled         =   0   'False
               Height          =   330
               Index           =   4
               Left            =   50
               TabIndex        =   13
               Top             =   180
               Width           =   1305
            End
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tip. Mov."
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   39
            Top             =   420
            Width           =   675
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   5955
            TabIndex        =   36
            Top             =   420
            Width           =   735
         End
         Begin VB.Label LblIdProveedor 
            AutoSize        =   -1  'True
            Caption         =   "LblIdProveedor"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   9600
            TabIndex        =   35
            Top             =   0
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Glosa"
            Height          =   195
            Index           =   9
            Left            =   5955
            TabIndex        =   34
            Top             =   1335
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Almacén"
            Height          =   195
            Index           =   1
            Left            =   105
            TabIndex        =   33
            Top             =   1335
            Width           =   615
         End
         Begin VB.Label lblAlmacen 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblAlmacen"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   2025
            TabIndex        =   32
            Top             =   1305
            Width           =   3690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tip. Doc. Ref."
            Height          =   195
            Index           =   5
            Left            =   5955
            TabIndex        =   31
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label LblTipDocRef 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipDocRef"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   8025
            TabIndex        =   30
            Top             =   690
            Width           =   3675
         End
         Begin VB.Label lbliddocref 
            AutoSize        =   -1  'True
            Caption         =   "lbliddocref"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   10980
            TabIndex        =   29
            Top             =   0
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Doc. Ref."
            Height          =   195
            Index           =   7
            Left            =   5955
            TabIndex        =   28
            Top             =   1020
            Width           =   915
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Documento"
            Height          =   195
            Index           =   4
            Left            =   3075
            TabIndex        =   26
            Top             =   735
            Width           =   1185
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000001&
            BackStyle       =   1  'Opaque
            Height          =   90
            Left            =   2115
            Top             =   1110
            Width           =   105
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Num. Doc."
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   25
            Top             =   1020
            Width           =   765
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Movimiento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   60
            TabIndex        =   20
            Top             =   30
            Width           =   11670
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Mov."
            Height          =   195
            Index           =   3
            Left            =   105
            TabIndex        =   19
            Top             =   720
            Width           =   720
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   1058
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
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "FrmIngresoAlmacen3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim mMesActivo As Integer                            ' --indica el mes activo
Dim fCierrePeriodo As Boolean                        ' --indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)
Dim IdMenuActivo As Integer                          ' INDICA EL CODIGO DEL MENU ACTIVO
Dim cSQL As String
Dim F As New SistemaLogica.Funciones

Const ESTADOPENDIENTE_ = 1
Const ESTADOPROCESADO_ = 2
Const ESTADOAPROBADO_ = 3
Const ESTADOANULADO_ = 4

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

Private Sub cmd_Click(Index As Integer)
    Dim xRs As New ADODB.Recordset
    Dim xRsAux As New ADODB.Recordset
    Dim nTitulo As String
    Dim xCampos() As String
    Dim nSQLId As String
    Dim IDTIPDOCREF_ As Integer
    Dim IDDOCREF_ As Integer
        
    If QueHace = 3 Then Exit Sub
    
    Select Case Index
        Case 0 ' Almacen
            ReDim xCampos(2, 4) As String
            
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
            
            nTitulo = "Buscando Almacenes"
            cSQL = "SELECT alm_almacenes.* FROM alm_almacenes"
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "descripcion", "descripcion", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            txtIdAlm.Text = NulosN(xRs("id"))
            lblAlmacen.Caption = UCase(NulosC(xRs("descripcion")))
            TxtProv.SetFocus
            Set xRs = Nothing
        
        Case 1 ' Proveedor
            ReDim xCampos(3, 4) As String
            
            xCampos(0, 0) = "Nombre":      xCampos(0, 1) = "nombre":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Nº R.U.C.":   xCampos(1, 1) = "numruc":      xCampos(1, 2) = "1200":         xCampos(1, 3) = "C"
            xCampos(2, 0) = "Codigo":      xCampos(2, 1) = "id":          xCampos(2, 2) = "1000":         xCampos(2, 3) = "N"
            
            If OptIng.Value = True Then
                ' MOSTRARA LA LISTA DE PROVEEDORES
                cSQL = "SELECT mae_prov.id, mae_prov.numruc, mae_prov.nombre, mae_prov.activo " _
                    + vbCr + "FROM mae_prov " _
                    + vbCr + "WHERE (((mae_prov.activo) = -1))"
                nTitulo = "Buscando Proveedores"
            Else
                ' MOSTRARA LA LISTA DE CLIENTES
                cSQL = "SELECT mae_cliente.id, mae_cliente.numruc, mae_cliente.nombre, mae_cliente.activo " _
                    + vbCr + "FROM mae_cliente " _
                    + vbCr + "WHERE (((mae_cliente.activo) = -1))"
                nTitulo = "Buscando Clientes"
            End If
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "nombre", "nombre", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            LblIdProveedor.Caption = NulosN(xRs("id"))
            TxtProv.Text = NulosC(xRs("nombre"))
            TxtIdTipDocRef.SetFocus
            Set xRs = Nothing
        
        Case 2 ' Tipo Documento de Referencia
            ReDim xCampos(2, 4) As String
            
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
                  
            nTitulo = "Buscando Tipo de Documento de Referencia"
            
            cSQL = "SELECT mae_documento.* FROM mae_documento WHERE (id In (70,71,92,110,112,113,114))"
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "descripcion", "descripcion", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            TxtIdTipDocRef.Text = NulosN(xRs("id"))
            LblTipDocRef.Caption = UCase(NulosC(xRs("descripcion")))
            txtNumDocRef.SetFocus
            Set xRs = Nothing
            
        Case 3 ' Documento de Referencia
            'ReDim xCampos(3, 4) As String
            ReDim xCampos(5, 4) As String
            
            'descripcion                        'campo                              'tamaño                     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Fecha":            xCampos(0, 1) = "fecha":            xCampos(0, 2) = "1000":       xCampos(0, 3) = "F"
            xCampos(1, 0) = "Descripcion":      xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "3000":      xCampos(1, 3) = "C"
            xCampos(2, 0) = "Num. Doc.":        xCampos(2, 1) = "numdoc":           xCampos(2, 2) = "1500":       xCampos(2, 3) = "C"
            
            
            xCampos(3, 0) = "Num. Ord.":        xCampos(3, 1) = "numdocref":        xCampos(3, 2) = "1500":       xCampos(3, 3) = "C"
            xCampos(4, 0) = "Lote":             xCampos(4, 1) = "loteref":          xCampos(4, 2) = "900":       xCampos(4, 3) = "C"
            
            nTitulo = "Buscando Tipos"
            
            IDTIPDOCREF_ = NulosN(TxtIdTipDocRef.Text)
            
            ' Se genera la consulta
            cSQL = "SELECT cDOCREF.id, cDOCREF.fecha, cDOCREF.numdoc, cDOCREF.descripcion, cDOCREF.numdocref, cDOCREF.loteref " _
                + vbCr + "FROM ( " _
                + vbCr + "SELECT alm_devolucion.id, alm_devolucion.fching AS fecha, [alm_devolucion].[numser] & '-' & [alm_devolucion].[numdoc] AS numdoc, pla_empleados.nombre AS descripcion, 114 AS iddoc, '' AS numdocref, '' AS loteref " _
                + vbCr + "FROM alm_devolucion LEFT JOIN pla_empleados ON alm_devolucion.idresp = pla_empleados.id " _
                + vbCr + "UNION " _
                + vbCr + "SELECT alm_recepcion.id, alm_recepcion.fching AS fecha, [alm_recepcion].[numser] & '-' & [alm_recepcion].[numdoc] AS numdoc, alm_inventario.descripcion, 71 AS iddoc, '' AS numdocref, '' AS loteref " _
                + vbCr + "FROM alm_recepcion LEFT JOIN alm_inventario ON alm_recepcion.iditem = alm_inventario.id " _
                + vbCr + "UNION " _
                + vbCr + "SELECT pro_solicitudmat.id, pro_solicitudmat.fchdoc AS fecha, [pro_solicitudmat].[numser] & '-' & [pro_solicitudmat].[numdoc] AS numdoc, pla_empleados.nombre AS descripcion, 110 AS iddoc, [pro_ordenprod].[numser] & '-' & [pro_ordenprod].[numdoc] AS numdocref, pro_ordenprod.lote AS loteref " _
                + vbCr + "FROM (pro_solicitudmat LEFT JOIN pla_empleados ON pro_solicitudmat.idresp = pla_empleados.id) LEFT JOIN pro_ordenprod ON pro_solicitudmat.iddocref = pro_ordenprod.id " _
                + vbCr + "WHERE (((pro_solicitudmat.estado)=" & ESTADOPROCESADO_ & ") AND ((pro_solicitudmat.idmes) In (" & Month(TxtFchIng.Valor) & "," & Month(TxtFchIng.Valor) - 1 & "))) " _
                + vbCr + "UNION " _
                + vbCr + "SELECT com_ordencompra.id, com_ordencompra.fchemi AS fecha, [com_ordencompra].[numser] & '-' & [com_ordencompra].[numdoc] AS numdoc, pla_empleados.nombre AS descripcion, 92 AS iddoc, '' AS numdocref, '' AS loteref " _
                + vbCr + "FROM com_ordencompra LEFT JOIN pla_empleados ON com_ordencompra.idaut = pla_empleados.id " _
                + vbCr + "WHERE ((com_ordencompra.idest)=2) " _
                + vbCr + ") AS cDOCREF " _
                + vbCr + "WHERE (((cDOCREF.iddoc)=" & IDTIPDOCREF_ & "));"
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "numdoc", "numdoc", CualquierParte, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            IDDOCREF_ = NulosN(xRs("id"))
            txtNumDocRef.Text = NulosC(xRs("numdoc"))
            lbliddocref.Caption = IDDOCREF_
            
            cSQL = ""
            Select Case IDTIPDOCREF_
                Case 71 ' Guia Interna de Recepcion
                    cSQL = "SELECT alm_recepcion.iditem, alm_inventario.descripcion AS desitem, Sum(alm_recepciondet.pesnettot) AS cantidad, '' AS idlotedet, alm_recepciondet.idunimed, alm_inventario.tippro AS idtippro, mae_unidades.descripcion AS desunimed, mae_tipoproducto.descripcion AS destippro, '' AS deslote " _
                        + vbCr + "FROM (((alm_recepcion LEFT JOIN alm_recepciondet ON alm_recepcion.id = alm_recepciondet.idrecep) LEFT JOIN alm_inventario ON alm_recepcion.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON alm_recepciondet.idunimed = mae_unidades.id) LEFT JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id " _
                        + vbCr + "GROUP BY alm_recepcion.iditem, alm_inventario.descripcion, '', alm_recepciondet.idunimed, alm_inventario.tippro, mae_unidades.descripcion, mae_tipoproducto.descripcion, '', alm_recepciondet.idestado, alm_recepcion.id " _
                        + vbCr + "HAVING (((alm_recepciondet.idestado)>1 And (alm_recepciondet.idestado)<>4) AND ((alm_recepcion.id)=" & IDDOCREF_ & "));"

                Case 92 ' Orden de Compra
                    cSQL = "SELECT com_ordencompradet.iditem, alm_inventario.descripcion AS desitem, com_ordencompradet.canpro As cantidad, '' As idlotedet, alm_inventario.idunimed, alm_inventario.tippro AS idtippro, mae_unidades.abrev AS desunimed, mae_tipoproducto.descripcion AS destippro, '' AS deslote " _
                        + vbCr + "FROM ((com_ordencompradet LEFT JOIN alm_inventario ON com_ordencompradet.iditem = alm_inventario.id) LEFT JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
                        + vbCr + "WHERE (((com_ordencompradet.idcom)=" & IDDOCREF_ & "));"
                    
                Case 110 ' Solicitud de Materiales
                    cSQL = "SELECT pro_solicitudmatdet.iditem, alm_inventario.descripcion AS desitem, pro_solicitudmatdet.cantidad, pro_solicitudmatdet.idlotedet, alm_inventario.idunimed, alm_inventario.tippro AS idtippro, mae_unidades.abrev AS desunimed, mae_tipoproducto.descripcion AS destippro, alm_inventariolote.descripcion AS deslote " _
                        + vbCr + "FROM ((((pro_solicitudmatdet LEFT JOIN alm_inventario ON pro_solicitudmatdet.iditem = alm_inventario.id) LEFT JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) LEFT JOIN alm_inventariolotedet ON pro_solicitudmatdet.idlotedet = alm_inventariolotedet.id) LEFT JOIN alm_inventariolote ON alm_inventariolotedet.idlote = alm_inventariolote.id " _
                        + vbCr + "WHERE (((pro_solicitudmatdet.idsol)=" & IDDOCREF_ & "));"
                
                Case 114 ' Nota de Devolucion
                    cSQL = "SELECT alm_devoluciondet.iditem, alm_inventario.descripcion AS desitem, alm_devoluciondet.cantidad, '' AS idlotedet, alm_devoluciondet.idunimed, alm_inventario.tippro AS idtippro, mae_unidades.abrev AS desunimed, mae_tipoproducto.descripcion AS destippro, '' AS deslote " _
                        + vbCr + "FROM ((alm_devoluciondet LEFT JOIN alm_inventario ON alm_devoluciondet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON alm_devoluciondet.idunimed = mae_unidades.id) LEFT JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id " _
                        + vbCr + "WHERE (((alm_devoluciondet.iddev)=" & IDDOCREF_ & "));"
            End Select
            
            Set xRs = Nothing
            RST_Busq xRs, cSQL, xCon
            
            fg(0).Rows = fg(0).FixedRows
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            xRs.MoveFirst
            While Not xRs.EOF
                fg(0).Rows = fg(0).Rows + 1
                fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("IDITEM")) = NulosN(xRs("iditem"))
                fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("CODIGO")) = Busca_Codigo(NulosN(xRs("iditem")), "id", "codpro", "alm_inventario", "N", xCon)
                fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("STOCK")) = Format(SaldoActual(NulosN(xRs("iditem")), "01/01/" & AnoTra, Date, xCon), FORMAT_CANTIDAD)
                If OptIng.Value Then
                    fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("CANTEO")) = Format(NulosN(xRs("cantidad")), FORMAT_CANTIDADDECIMAL)
                    fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("CANMOV")) = Format(NulosN(xRs("cantidad")), FORMAT_CANTIDADDECIMAL)
                Else
                    fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("CANTEO")) = Format(NulosN(xRs("cantidad")), FORMAT_CANTIDADDECIMAL)
                    fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("CANMOV")) = Format(NulosN(xRs("cantidad")), FORMAT_CANTIDADDECIMAL)
                End If
                fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("ITEM")) = NulosC(xRs("desitem"))
                fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("UM")) = NulosC(xRs("desunimed"))
                
                xRs.MoveNext
            Wend
            txtIdAlm.SetFocus
        
        Case 4 ' Agregar Item
            AddItem
        
        Case 5 ' Eliminar Item
            DelItem
        
    End Select
End Sub

Private Sub AddItem()
    Dim fInsertar As Boolean
    
    ' PERMITE AGREGAR UN ITEM
    If QueHace = 3 Then Exit Sub
    
    Agregando = True
    If fg(0).Rows > fg(0).FixedRows Then
        If NulosN(fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("IDITEM"))) = 0 Then
            MsgBox "Seleccione un Producto", vbExclamation, xTitulo
        Else
            fInsertar = True
        End If
    Else
        fInsertar = True
    End If
    
    If fInsertar = True Then
        fg(0).AddItem ""
    End If
        
    fg(0).Row = fg(0).Rows - 1
    fg(0).Col = fg(0).ColIndex("ITEM")
    
    fg(0).SetFocus
    Agregando = False
End Sub

Private Sub DelItem()
    ' ELIMINA UNA FILA DEL CONTROL FlexGrid fg(0)
    If QueHace = 3 Then Exit Sub
    
    If fg(0).Row < 1 Then
        MsgBox "Elija el registro correcto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    If fg(0).Rows = 1 Then
        MsgBox "No hay Registro para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fg(0).SetFocus
        Exit Sub
    End If
    If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    
    fg(0).RemoveItem fg(0).Row
    
    If fg(0).Rows > 1 Then
        fg(0).Row = fg(0).Rows - 1
        fg(0).Col = fg(0).ColIndex("ITEM")
        fg(0).SetFocus
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

Private Sub fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim xRs As New ADODB.Recordset
    Dim nTitulo As String
    Dim xCampos() As String
    Dim TIPOPRODUCTO_ As Double
    Dim iDITEM_ As Double
        
    If QueHace = 3 Then Exit Sub
    Select Case Index
        Case 0
            If Col = fg(0).ColIndex("ITEM") Then
                ' BUSCA UN ITEM
                ' Se verifica el Almacen
                If NulosN(txtIdAlm.Text) = 0 Then
                    MsgBox "Seleccione el Almacén para el registro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    txtIdAlm.SetFocus
                    Exit Sub
                End If
                
                ReDim xCampos(3, 4) As String
                
                xCampos(0, 0) = "Producto":    xCampos(0, 1) = "desitem":   xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
                xCampos(1, 0) = "Codigo":      xCampos(1, 1) = "codigo":        xCampos(1, 2) = "1500":         xCampos(1, 3) = "C"
                xCampos(2, 0) = "Uni. Med":    xCampos(2, 1) = "abrev":         xCampos(2, 2) = "1200":         xCampos(2, 3) = "C"
                        
                nTitulo = "Buscando Ítems"
                
'                cSQL = "SELECT alm_almacenesdet.iditem, alm_inventario.descripcion, alm_inventario.codpro, mae_unidades.abrev " _
'                    + vbCr + "FROM ((alm_almacenes INNER JOIN alm_almacenesdet ON alm_almacenes.id = alm_almacenesdet.idalm) INNER JOIN alm_inventario ON alm_almacenesdet.iditem = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
'                    + vbCr + "WHERE (((alm_almacenes.id) = " & NulosN(txtIdAlm.Text) & ") And ((alm_almacenes.idtippro) = 0)) " _
'                    + vbCr + "UNION " _
'                    + vbCr + "SELECT alm_inventario.id AS iditem, alm_inventario.descripcion, alm_inventario.codpro, mae_unidades.abrev " _
'                    + vbCr + "FROM (alm_almacenes INNER JOIN alm_inventario ON alm_almacenes.idtippro = alm_inventario.tippro) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
'                    + vbCr + "WHERE (((alm_almacenes.id) = " & NulosN(txtIdAlm.Text) & "))"
                    
                cSQL = "SELECT alm_inventario.id AS iditem, alm_inventario.descripcion AS desitem, alm_inventario.codpro AS codigo, alm_inventario.idunimed, mae_tipoproducto.descripcion AS tippro, mae_familia.descripcion AS familia, mae_clase.descripcion AS clase, mae_subclase.descripcion AS subclase " _
                    + vbCr + "FROM (((alm_inventario LEFT JOIN mae_familia ON alm_inventario.idfam = mae_familia.id) LEFT JOIN mae_clase ON alm_inventario.idclas = mae_clase.id) LEFT JOIN mae_subclase ON alm_inventario.idsubclas = mae_subclase.id) LEFT JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id " _
                    + vbCr + "WHERE (((alm_inventario.activo)=-1))"
                     
                Set xRs = Nothing
                CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                                "desitem", "codigo", Principio, ""
                
                If xRs.State = 0 Then Exit Sub
                If xRs.RecordCount = 0 Then Exit Sub
                
                fg(0).TextMatrix(Row, fg(0).ColIndex("CODIGO")) = NulosC(xRs("codigo"))
                fg(0).TextMatrix(Row, fg(0).ColIndex("ITEM")) = NulosC(xRs("desitem"))
                fg(0).TextMatrix(Row, fg(0).ColIndex("UM")) = Busca_Codigo(NulosN(xRs("idunimed")), "id", "abrev", "mae_unidades", "N", xCon)
                fg(0).TextMatrix(Row, fg(0).ColIndex("STOCK")) = Format(SaldoActual(NulosN(xRs("iditem")), "01/01/" & AnoTra, Date, xCon), FORMAT_CANTIDAD)
                fg(0).TextMatrix(Row, fg(0).ColIndex("IDITEM")) = NulosN(xRs("iditem"))
                fg(0).Col = fg(0).ColIndex("CANMOV")
                fg(0).SetFocus
            End If
    End Select
End Sub

Private Sub fg_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If Row <= 0 Then Exit Sub
    Select Case Index
        Case 0
            Select Case Col
                Case fg(0).ColIndex("CANMOV"), fg(0).ColIndex("CANTEO")
                    fg(0).TextMatrix(Row, Col) = Format(fg(0).TextMatrix(Row, Col), "0.0000")
                    
            End Select
    End Select
End Sub

Private Sub fg_EnterCell(Index As Integer)
    Select Case Index
        Case 0
            If QueHace = 3 Then fg(0).Editable = flexEDNone: Exit Sub
            Select Case fg(0).Col
                Case fg(0).ColIndex("ITEM"), fg(0).ColIndex("CANMOV"), fg(0).ColIndex("CANTEO")
                    fg(0).Editable = flexEDKbdMouse
                    
                Case Else
                    fg(0).Editable = flexEDNone
            End Select
    End Select
End Sub

Private Sub fg_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case Index
        Case 0
            Select Case Col
                Case fg(0).ColIndex("CANMOV"), fg(0).ColIndex("CANTEO")
                    If validar_numero(KeyAscii) = False Then KeyAscii = 0
            End Select
    End Select
End Sub

Private Sub fg_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    Select Case Index
        Case 0
            If KeyCode = 45 Then
                AddItem
            End If
            If KeyCode = 46 Then
                DelItem
            End If
    End Select
End Sub

Private Sub fg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If QueHace = 3 Then Exit Sub
    Select Case Index
        Case 0
            If Button = 2 Then PopupMenu menu1
    End Select
End Sub

Private Sub Form_Activate()
    'Modificado 13/01/11 Johan Castro
    '           Eliminar

    If SeEjecuto = False Then
        SeEjecuto = True
        
        mMesActivo = xMes
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        pCargarDatos
    End If
End Sub

Private Sub Form_Load()
    QueHace = 3
    iniciarCampos
End Sub

Private Sub iniciarCampos()
    Dim xRs As New ADODB.Recordset
    
    SeEjecuto = False
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Mostrando = False
        
    GRID_COMBOLIST fg(0), fg(0).ColIndex("ITEM")
    fg(0).ColWidth(fg(0).ColIndex("IDITEM")) = 0
    With fg(0)
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExSortShow
        .SelectionMode = flexSelectionByRow
        .ForeColorSel = &H80000005
        .BackColorSel = &H80&
    End With
    With fg(1)
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExSortShow
        .SelectionMode = flexSelectionByRow
        .ForeColorSel = &H80000005
        .BackColorSel = &H80&
    End With
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
    
    fg(0).Editable = flexEDKbdMouse
    fg(0).Rows = 1
    fg(0).Rows = fg(0).Rows + 1
    fg(0).SelectionMode = flexSelectionFree
    
    OptIng.Value = True
    OptIng_Click
    xHorIni = Time
    TxtFchIng.Valor = Date
    TxtFchDoc.Valor = Date
    TxtFchIng.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario miestras este agregando o editando un registro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub Menu1_1_Click()
    AddItem
End Sub

Private Sub Menu1_3_Click()
    DelItem
End Sub

Private Sub OptIng_Click()
    ' PREPARAMOS EL FORMULARIO PARA UN INGRESO
    Label3(3).Caption = "Fch. Ingreso"
    fg(0).TextMatrix(0, fg(0).ColIndex("CANMOV")) = "Cant. Ingreso"
    Label33.Caption = "Proveedor"
End Sub

Private Sub OptIng_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then TxtFchDoc.SetFocus
End Sub

Private Sub OptSal_Click()
    ' PREPARAMOS EL FORMULARIO PARA UNA SALIDA
    Label3(3).Caption = "Fch. Salida"
    fg(0).TextMatrix(0, fg(0).ColIndex("CANMOV")) = "Cant. Salida"
    Label33.Caption = "Cliente"
End Sub

Private Sub OptSal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then TxtFchDoc.SetFocus
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If RstIng.State = 0 Then Exit Sub
        If RstIng.RecordCount = 0 And QueHace <> 1 Then
            Cancel = 1
            Exit Sub
        End If
        If QueHace = 3 Then Blanquea: MuestraSegundoTab
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
        RstIng.Filter = ""
        TDB_FiltroLimpiar Dg1
    End If
    
    If Button.Index = 10 Then Buscar
    
    If Button.Index = 11 Then
        mMesActivo = SeleccionaMes(xCon)
        pCargarDatos
    End If
        
    If Button.Index = 13 Then pExportar
    
    If Button.Index = 16 Then
        Unload Me
        Set RstIng = Nothing
    End If
End Sub

Private Sub pExportar()
    TabOne1.CurrTab = 0
    Dim nSQL As String
    Dim oExport As New SGI2_funciones.formularios
    Dim RstTmp  As New ADODB.Recordset
    Dim xCampos(7, 3) As String
    
    ' 0 Nombre a Mostrar;
    ' 1 nombre de Campo del Rst;
    ' 2 alineacion(0::derecha, 1::centro, 2::izquierda);
    ' 3 ancho de columna
    ' obs: el rst puede tener mas columnas solo se consideran los campos del array
    xCampos(0, 0) = "Fch. Mov.":                xCampos(0, 1) = "fchdoc":           xCampos(0, 2) = 0:  xCampos(0, 3) = "1000"
    xCampos(1, 0) = "Mov.":                     xCampos(1, 1) = "movi":             xCampos(1, 2) = 0:  xCampos(1, 3) = "1000"
    xCampos(2, 0) = "Nombre":                   xCampos(2, 1) = "nombre":           xCampos(2, 2) = 0:  xCampos(2, 3) = "1000"
    xCampos(3, 0) = "Nº Documento":             xCampos(3, 1) = "numdoc":           xCampos(3, 2) = 0:  xCampos(3, 3) = "1050"
    xCampos(4, 0) = "Almacén":                  xCampos(4, 1) = "desalm":           xCampos(4, 2) = 0:  xCampos(4, 3) = "4500"
    xCampos(5, 0) = "Tip. Doc. Ref.":           xCampos(5, 1) = "destipdocref":     xCampos(5, 2) = 0:  xCampos(5, 3) = "1050"
    xCampos(6, 0) = "Num. Doc. ref.":           xCampos(6, 1) = "numdocref":        xCampos(6, 2) = 0:  xCampos(6, 3) = "1050"
    xCampos(7, 0) = "Estado":                   xCampos(7, 1) = "desestado":        xCampos(7, 2) = 0:  xCampos(7, 3) = "1050"
    '**********************************************************************************************************************************
        
    Set RstTmp = RstIng
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "Lista de Movimientos", "Periodo: " & LblPeriodo.Caption & "  -  " & AnoTra, "", "Listado de Producción", RstTmp, xCampos
    Set oExport = Nothing
    Set RstTmp = Nothing
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

Function validarDatos() As Boolean
    ' VALIDAMOS QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
    If Year(TxtFchIng.Valor) <> AnoTra Then
        MsgBox "El año ingresado en la " & Label3(3).Caption & " no coincide con el Ejercicio" & vbCr & "Corrija la fecha o registre en su año que corresponde", vbInformation, xTitulo
        TxtFchIng.Valor = ""
        TxtFchIng.SetFocus
        validarDatos = False
        Exit Function
    End If
    
    If TxtFchIng.Valor = "" Then
        MsgBox "No ha especificado la fecha de ingreso", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIng.SetFocus
        validarDatos = False
        Exit Function
    End If
    
    If TxtFchDoc.Valor = "" Then
        MsgBox "No ha especificado la fecha del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchDoc.SetFocus
        validarDatos = False
        Exit Function
    End If
        
    If NulosC(txtIdAlm.Text) = "" Then
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        txtIdAlm.SetFocus
        validarDatos = False
        Exit Function
    End If
    
    If NulosC(TxtNumSer.Text) = "" Then
        MsgBox "No ha especificado el numero de serie del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumSer.SetFocus
        validarDatos = False
        Exit Function
    End If
    If NulosC(TxtNumDoc.Text) = "" Then
        MsgBox "No ha especificado el numero de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDoc.SetFocus
        validarDatos = False
        Exit Function
    End If
    
    If NulosN(txtIdAlm.Text) = 0 Then
        MsgBox "No ha especificado el nombre del almacén", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        txtIdAlm.SetFocus
        validarDatos = False
        Exit Function
    End If
    
    If fg(0).Rows = 2 Then
        If fg(0).TextMatrix(1, fg(0).ColIndex("ITEM")) = "" Then
            MsgBox "Falta ingresar el detalle...!", vbExclamation, "Mensaje...!"
            validarDatos = False
            Exit Function
        End If
    ElseIf fg(0).Rows = 1 Then
        MsgBox "Falta ingresar el detalle...!", vbExclamation, "Mensaje...!"
        validarDatos = False
        Exit Function
    End If
    
    validarDatos = True
End Function

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : Function
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA Alm_ingreso, DEVUELVE VERDADERO SI TIENE EXITO
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Function Grabar() As Boolean
    Dim Movimiento As New AlmacenEntidad.EMovimiento
    Dim F As New SistemaLogica.Funciones
    Dim A As Integer

On Error GoTo ERROR_
        
    If Not validarDatos Then Grabar = False: Exit Function
    ' Se llenan los detalles
    If QueHace = 1 Then Movimiento.IdMovimiento = 0 Else Movimiento.IdMovimiento = NulosN(RstIng("id"))
    Movimiento.FechaMovimiento = TxtFchIng.Valor
    Movimiento.NumeroSerie = NulosC(TxtNumSer.Text)
    Movimiento.NumeroDocumento = NulosC(TxtNumDoc.Text)
    Movimiento.IdEstado = F.NuloNumeric(F.KeyValue("EstadoAprobadoMovimiento", xCon))
    Movimiento.IdAlmacen = NulosN(txtIdAlm.Text)
    Movimiento.IdProveedor = NulosN(LblIdProveedor.Caption)
    Movimiento.Proveedor = NulosC(TxtProv.Text)
    Movimiento.Glosa = NulosC(GlosaText.Text)
    If OptIng.Value Then Movimiento.IdTipoMovimiento = -1 Else Movimiento.IdTipoMovimiento = 0
    Movimiento.IdTipoDocumentoReferencia = NulosN(TxtIdTipDocRef.Text)
    Movimiento.IdDocumentoReferencia = NulosN(lbliddocref.Caption)
    Movimiento.DocumentoReferencia = NulosC(txtNumDocRef.Text)
    Movimiento.MesTrabajo = mMesActivo
    Movimiento.AnhoTrabajo = AnoTra
    
    ' Se llenan los detalles
    For A = 1 To fg(0).Rows - 1
        Dim MovimientoDet As New AlmacenEntidad.EMovimientoDet
        MovimientoDet.IdItem = NulosN(fg(0).TextMatrix(A, fg(0).ColIndex("IDITEM")))
        MovimientoDet.Cantidad = NulosN(fg(0).TextMatrix(A, fg(0).ColIndex("CANMOV")))
        MovimientoDet.CantidadTeorica = NulosN(fg(0).TextMatrix(A, fg(0).ColIndex("CANTEO")))
        ' Se agrega al padre
        Movimiento.LMovimientoDet.Add MovimientoDet
        Set MovimientoDet = Nothing
    Next A
    
    ' Se graba el movimiento
    Set Movimiento.Conexion = xCon
    If Not Movimiento.Save(0, "") Then Err.Raise &HFFFFFF01, , "No se puedo registrar el movimiento"
    MsgBox "El movimiento se grabó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    mIdRegistro = Movimiento.IdMovimiento
    Set Movimiento = Nothing
    Grabar = True
    Exit Function

ERROR_:
    Set Movimiento = Nothing
    Grabar = False
    MsgBox "No se pudo registrar el movimiento por el siguiente motivo :" + Trim(Err.Description)
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
        xCon.Execute "DELETE * FROM alm_ingresodet WHERE id = " & NulosN(RstIng("id"))
        xCon.Execute "DELETE * FROM alm_ingreso WHERE id = " & NulosN(RstIng("id"))
        
        ' Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & NulosN(RstIng("id")) & " AND idform = " & IdMenuActivo

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
    Dim F As New SistemaLogica.Funciones
    
    If NulosN(RstIng("idtipdocref")) = F.NuloNumeric(F.KeyValue("Transferencia", xCon)) Then
        If F.NuloNumeric(F.KeyValue("ModificaMovTransferencia", xCon)) = 0 Then
            MsgBox "Los Movimientos de Transferencias no son editables desde este formulario, " _
            + vbCr + "ingrese a: Mantenimiento de Transferencias de Almacèn", vbInformation, xTitulo
            Exit Sub
        End If
    End If
    
    If NulosN(RstIng("idtipdocref")) = F.NuloNumeric(F.KeyValue("TomaInventario", xCon)) Then
        If F.NuloNumeric(F.KeyValue("ModificaMovTomaInventario", xCon)) = 0 Then
            MsgBox "Los Movimientos de Ajuste de Inventario no son editables desde este formulario, " _
            + vbCr + "ingrese a: Ajuste de Inventario", vbInformation, xTitulo
            Exit Sub
        End If
    End If
    
    If NulosN(RstIng("idtipdocref")) = F.NuloNumeric(F.KeyValue("ParteProduccion", xCon)) Then
        If F.NuloNumeric(F.KeyValue("ModificaMovParteProduccion", xCon)) = 0 Then
            MsgBox "Los Movimientos de Parte de Produccion no son editables desde este formulario, " _
            + vbCr + "ingrese a: Parte de Produccion", vbInformation, xTitulo
            Exit Sub
        End If
    End If
    
    If NulosN(RstIng("idtipdocref")) = F.NuloNumeric(F.KeyValue("IdDocumentoInventarioInicial", xCon)) Then
        If F.NuloNumeric(F.KeyValue("ModificaMovInventarioInicial", xCon)) = 0 Then
            MsgBox "Los Movimientos de Inventario Inicial no son editables desde este formulario. " _
            + vbCr + "Ingrese a: Ajustes de Inventario", vbInformation, xTitulo
            Exit Sub
        End If
    End If
    
    If NulosN(RstIng("idtipdocref")) = F.NuloNumeric(F.KeyValue("IdDocumentoGuiaRemision", xCon)) Then
        If F.NuloNumeric(F.KeyValue("ModificaMovGuiaRemision", xCon)) = 0 Then
            MsgBox "Los Movimientos de Guia de Remision no son editables desde este formulario. " _
            + vbCr + "Ingrese a: Guia de Remision", vbInformation, xTitulo
            Exit Sub
        End If
    End If
    
    If NulosN(RstIng("idtipdocref")) = F.NuloNumeric(F.KeyValue("IdDocumentoFactura", xCon)) Then
        If F.NuloNumeric(F.KeyValue("ModificaMovFactura", xCon)) = 0 Then
            MsgBox "Los Movimientos de Factura no son editables desde este formulario. " _
            + vbCr + "Ingrese a: Ventas", vbInformation, xTitulo
            Exit Sub
        End If
    End If
    
    TabOne1.CurrTab = 1
    ActivaTool
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Modificando Movimiento"
    QueHace = 2
    Bloquea
    Blanquea
 
    OptIng.Value = True
    fg(0).Editable = flexEDKbdMouse
    fg(0).Rows = 1
    fg(0).Rows = fg(0).Rows + 1
    fg(0).SelectionMode = flexSelectionFree
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
    txtIdAlm.Text = ""
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    lblAlmacen.Caption = ""
    txtNumDocRef.Text = ""
    TxtIdTipDocRef.Text = ""
    TxtProv.Text = ""
    LblTipDocRef.Caption = ""
    GlosaText.Text = ""
    fg(0).Rows = fg(0).FixedRows
    fg(1).Rows = fg(1).FixedRows
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
    txtIdAlm.Locked = Not txtIdAlm.Locked
    TxtNumSer.Locked = Not TxtNumSer.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    TxtProv.Locked = Not TxtProv.Locked
    GlosaText.Locked = Not GlosaText.Locked
    OptIng.Enabled = Not OptIng.Enabled
    OptSal.Enabled = Not OptSal.Enabled
    habilitar cmd, Not cmd(0).Enabled
End Sub

Private Sub TxtIdTipDocRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdTipDocRef_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        cmd(2).Value = True
    End If
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumDoc_Validate(Cancel As Boolean)
    Dim idDocumento As Long
    
    If NulosC(TxtNumDoc.Text) = "" Then Exit Sub
    TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000000")
    
    If QueHace = 1 Then idDocumento = 0 Else idDocumento = F.NuloNumeric(RstIng("id"))
    If F.ExisteDocumento("alm_ingreso", "'" & F.NuloString(TxtNumDoc.Text) & "'", xCon, , "'" & F.NuloString(TxtNumSer.Text) & "'", , , , idDocumento, "id") Then
        MsgBox "El documento ingresado ya existe" & vbCr & "Corrija el numero de documento", vbInformation, xTitulo
        TxtNumDoc.Text = ""
        TxtNumDoc.SetFocus
        Exit Sub
    End If
End Sub

Private Sub TxtNumSer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumSer_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    
    If NulosC(TxtNumSer.Text) <> "" Then
        TxtNumSer.Text = Format(TxtNumSer.Text, "0000")
        TxtNumDoc.Text = hallarNumDoc("alm_ingreso", "'" & NulosC(TxtNumSer.Text) & "'", "numser")
        If NulosC(TxtNumDoc.Text) = "" Then TxtNumSer.Text = ""
    End If
End Sub

Private Sub txtNumDocRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtNumDocRef_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then ' F5
        cmd(3).Value = True
    End If
End Sub

Private Sub txtIdAlm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub txtIdAlm_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        cmd_Click 0
    End If
End Sub

Private Sub txtIdAlm_Validate(Cancel As Boolean)
    If txtIdAlm.Text = "" Then
        lblAlmacen.Caption = ""
        Exit Sub
    End If
    
    lblAlmacen.Caption = Busca_Codigo(NulosN(txtIdAlm.Text), "id", "descripcion", "alm_almacenes", "N", xCon)
    If NulosC(lblAlmacen.Caption) = "" Then
        txtIdAlm.Text = ""
    End If
End Sub

Private Sub TxtProv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        LblIdProveedor.Caption = ""
    End If
End Sub

Private Sub TxtProv_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        cmd(1).Value = True
    End If
End Sub

Sub MuestraSegundoTab()
    Dim xRs As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim A As Integer
    Dim database As New SistemaData.EDataBase
    Dim F As New SistemaLogica.Funciones
    
    If RstIng.RecordCount = 0 Then Exit Sub
    If RstIng.BOF = True Or RstIng.EOF = True Then Exit Sub
    
    cSQL = "SELECT alm_ingreso.* " _
        + vbCr + "FROM alm_ingreso " _
        + vbCr + "WHERE alm_ingreso.id=" & NulosN(RstIng("id"))
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    
    TxtFchIng.Valor = xRs("fching")
    TxtFchDoc.Valor = xRs("fchdoc")
    TxtNumSer.Text = NulosC(xRs("numser"))
    TxtNumDoc.Text = NulosC(xRs("numdoc"))
    txtIdAlm.Text = NulosN(xRs("idalm"))
    lblAlmacen.Caption = UCase(Busca_Codigo(NulosN(xRs("idalm")), "id", "descripcion", "alm_almacenes", "N", xCon))
    LblIdProveedor.Caption = NulosN(xRs("idpro"))
    TxtProv.Text = NulosC(xRs("nombre"))
    GlosaText.Text = NulosC(xRs("glosa"))
    
    If NulosN(xRs("idtipdocref")) = 0 Then
        TxtIdTipDocRef.Text = ""
        LblTipDocRef.Caption = ""
        lbliddocref.Caption = ""
        txtNumDocRef.Text = ""
    Else
        TxtIdTipDocRef.Text = NulosN(xRs("idtipdocref"))
        LblTipDocRef.Caption = UCase(Busca_Codigo(NulosN(xRs("idtipdocref")), "id", "descripcion", "mae_documento", "N", xCon))
        lbliddocref.Caption = NulosN(xRs("iddocref"))
        txtNumDocRef.Text = NulosC(RstIng("numdocref"))
    End If
    
    Mostrando = True
    
    If xRs("tipmov") = -1 Then
        OptIng.Value = True
    Else
        OptSal.Value = True
    End If
    
    Mostrando = False
    cSQL = "SELECT alm_ingresodet.*, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventariolote.descripcion AS deslote, alm_inventariolotedet.idlote " _
        + vbCr + "FROM (mae_unidades RIGHT JOIN (alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) ON mae_unidades.id = alm_inventario.idunimed) LEFT JOIN (alm_inventariolote RIGHT JOIN alm_inventariolotedet ON alm_inventariolote.id = alm_inventariolotedet.idlote) ON alm_ingresodet.idlotedet = alm_inventariolotedet.id " _
        + vbCr + "WHERE (((alm_ingresodet.id) = " & NulosN(RstIng("id")) & "));"
    
    Set RstDet = Nothing
    RST_Busq RstDet, cSQL, xCon

    fg(0).Rows = 1
    If RstDet.RecordCount <> 0 Then
        RstDet.MoveFirst
        For A = 1 To RstDet.RecordCount
            fg(0).Rows = fg(0).Rows + 1
            fg(0).TextMatrix(A, fg(0).ColIndex("CODIGO")) = NulosC(RstDet("codpro"))
            fg(0).TextMatrix(A, fg(0).ColIndex("ITEM")) = NulosC(RstDet("descripcion"))
            fg(0).TextMatrix(A, fg(0).ColIndex("UM")) = NulosC(RstDet("abrev"))
            fg(0).TextMatrix(A, fg(0).ColIndex("STOCK")) = Format(F.SaldoActual(NulosN(RstDet("iditem")), NulosN(xRs("idalm")), "01/01/" & AnoTra, Date, xCon), FORMAT_CANTIDAD)
            fg(0).TextMatrix(A, fg(0).ColIndex("CANMOV")) = Format(NulosN(RstDet("cantidad")), FORMAT_CANTIDAD)
            fg(0).TextMatrix(A, fg(0).ColIndex("IDITEM")) = NulosN(RstDet("iditem"))
            fg(0).TextMatrix(A, fg(0).ColIndex("CANTEO")) = Format(NulosN(RstDet("cantteo")), FORMAT_CANTIDAD)
            
            RstDet.MoveNext
            
            If RstDet.EOF = True Then
                Exit For
            End If
        Next A
    End If
    
    ' Se cargan los movimientos relacionados
    If F.NuloNumeric(xRs("idtipdocref")) > 0 And F.NuloNumeric(xRs("iddocref")) > 0 Then
        Set database.Connection = xCon
        database.CommandText = "SELECT alm_ingreso.id AS idmov, alm_ingresodet.idmovdet, alm_ingreso.idalm, alm_almacenes.descripcion AS alm, [alm_ingreso].[numser] & '-' & [alm_ingreso].[numdoc] AS numdoc, alm_ingreso.fching AS fchmov, IIf([alm_ingreso].[tipmov]=-1,'I','S') AS tipmov, alm_ingresodet.iditem, alm_inventario.codpro, alm_inventario.descripcion AS item, alm_inventario.idunimed, mae_unidades.abrev AS unimed, alm_ingresodet.cantidad " _
                    + vbCr + "FROM (((alm_ingreso INNER JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id) INNER JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) INNER JOIN alm_almacenes ON alm_ingreso.idalm = alm_almacenes.id " _
                    + vbCr + "WHERE (((alm_ingreso.id)<>" & NulosN(RstIng("id")) & ") AND ((alm_ingreso.idtipdocref)=" & NulosN(xRs("idtipdocref")) & ") AND ((alm_ingreso.iddocref)=" & NulosN(xRs("iddocref")) & "))"
        Set xRs = Nothing
        Set xRs = database.GetRecordset
        
        If xRs.RecordCount > 0 Then
            xRs.MoveFirst
            While Not xRs.EOF
                With fg(1)
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, .ColIndex("ALM")) = NulosC(xRs("alm"))
                    .TextMatrix(.Rows - 1, .ColIndex("NUMMOV")) = NulosC(xRs("numdoc"))
                    .TextMatrix(.Rows - 1, .ColIndex("FCHMOV")) = NulosC(xRs("fchmov"))
                    .TextMatrix(.Rows - 1, .ColIndex("TIPMOV")) = NulosC(xRs("tipmov"))
                    .TextMatrix(.Rows - 1, .ColIndex("CODPRO")) = NulosC(xRs("codpro"))
                    .TextMatrix(.Rows - 1, .ColIndex("ITEM")) = NulosC(xRs("item"))
                    .TextMatrix(.Rows - 1, .ColIndex("UNIMED")) = NulosC(xRs("unimed"))
                    .TextMatrix(.Rows - 1, .ColIndex("CANTIDAD")) = NulosC(xRs("cantidad"))
                End With
                xRs.MoveNext
            Wend
        End If
    End If
End Sub

Sub pCargarDatos()
    Dim mFiltroSQL As String
    
    TDB_FiltroLimpiar Dg1
    Set RstIng = Nothing
    
    mFiltroSQL = " AND ((alm_ingreso.idalm) IN (SELECT alm_almacenes.id FROM alm_almacenes WHERE alm_almacenes.vismov = -1)) "
    
    cSQL = "SELECT [alm_ingreso].[id] & '' AS id, Format([alm_ingreso].[fching],'Short Date') AS fchdoc, IIf(alm_ingreso!tipmov=-1,'ING.','SAL.') AS movi, [alm_ingreso]![numser]+'-'+[alm_ingreso]![numdoc] AS numdoc, alm_ingreso.nombre, alm_almacenes.descripcion AS desalm, mae_documento.abrev AS destipdocref, " _
            & "IIf([alm_ingreso].[idtipdocref]=110,[pro_solicitudmat].[numser] & '-' & [pro_solicitudmat].[numdoc],IIf([alm_ingreso].[idtipdocref]=71,[alm_recepcion].[numser] & '-' & [alm_recepcion].[numdoc],IIf([alm_ingreso].[idtipdocref]=114,[alm_devolucion].[numser] & '-' & [alm_devolucion].[numdoc],IIf([alm_ingreso].[idtipdocref]=92,[com_ordencompra].[numser] & '-' & [com_ordencompra].[numdoc],IIf([alm_ingreso].[idtipdocref]=119,[alm_transferencia].[numser] & '-' & [alm_transferencia].[numdoc],IIf([alm_ingreso].[idtipdocref]=120,[pro_produccion].[numser] & '-' & [pro_produccion].[numdoc],IIf([alm_ingreso].[idtipdocref]=111,[alm_tomainventario].[numser] & '-' & [alm_tomainventario].[numdoc],IIf([alm_ingreso].[idtipdocref]=9,[vta_guia].[numser] & '-' & [vta_guia].[numdoc],IIf([alm_ingreso].[idtipdocref]=1,[vta_ventas].[numser] & '-' & [vta_ventas].[numdoc],''))))))))) AS numdocref, " _
            & "UCase([mae_estados].[descripcion]) AS desestado, alm_ingreso.idtipdocref " _
        + vbCr + "FROM ((((((((((((alm_ingreso LEFT JOIN alm_almacenes ON alm_ingreso.idalm = alm_almacenes.id) LEFT JOIN mae_area ON alm_ingreso.idare = mae_area.id) LEFT JOIN mae_estados ON alm_ingreso.estado = mae_estados.id) LEFT JOIN pro_solicitudmat ON alm_ingreso.iddocref = pro_solicitudmat.id) LEFT JOIN alm_recepcion ON alm_ingreso.iddocref = alm_recepcion.id) LEFT JOIN alm_devolucion ON alm_ingreso.iddocref = alm_devolucion.id) LEFT JOIN mae_documento ON alm_ingreso.idtipdocref = mae_documento.id) LEFT JOIN com_ordencompra ON alm_ingreso.iddocref = com_ordencompra.id) LEFT JOIN alm_transferencia ON alm_ingreso.iddocref = alm_transferencia.idtransferencia) LEFT JOIN pro_produccion ON alm_ingreso.iddocref = pro_produccion.id) LEFT JOIN alm_tomainventario ON alm_ingreso.iddocref = alm_tomainventario.idtomainventario) LEFT JOIN vta_guia ON alm_ingreso.iddocref = vta_guia.id) LEFT JOIN vta_ventas ON alm_ingreso.iddocref = vta_ventas.id " _
        + vbCr + "WHERE (((alm_ingreso.ano) = " & AnoTra & ") And ((alm_ingreso.idmes) = " & mMesActivo & ")) " & mFiltroSQL _
        + vbCr + "ORDER BY Format([alm_ingreso].[fching],'Short Date') DESC;"

    RST_Busq RstIng, cSQL, xCon
    Set Dg1.DataSource = RstIng
    
    '********************************************************************************************
    LblPeriodo.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    '********************************************************************************************

    '------------------------------------------------------------------------------------------
    ' bloqueamos los botones del toolbar
    CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
    '------------------------------------------------------------------------------------------
End Sub


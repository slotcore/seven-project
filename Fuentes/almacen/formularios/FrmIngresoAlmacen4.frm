VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmIngresoAlmacen4 
   Caption         =   "Almacén - Movimientos de Almacén"
   ClientHeight    =   7215
   ClientLeft      =   180
   ClientTop       =   1605
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
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
            Picture         =   "FrmIngresoAlmacen4.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen4.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen4.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen4.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen4.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen4.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen4.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen4.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen4.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen4.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen4.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen4.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoAlmacen4.frx":277E
            Key             =   "IMG13"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6900
      Left            =   0
      TabIndex        =   11
      Top             =   360
      Width           =   11880
      _cx             =   20955
      _cy             =   12171
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
         Height          =   6480
         Left            =   45
         TabIndex        =   16
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6090
            Left            =   30
            TabIndex        =   17
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
            Columns(4).Caption=   "Cliente/Proveedor/Responsable"
            Columns(4).DataField=   "nombre"
            Columns(4).NumberFormat=   "Short Date"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Operación"
            Columns(5).DataField=   "desope"
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
            TabIndex        =   18
            Top             =   0
            Width           =   1860
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            TabIndex        =   19
            Top             =   30
            Width           =   11610
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6480
         Left            =   12525
         TabIndex        =   12
         Top             =   375
         Width           =   11790
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   2175
            Left            =   90
            TabIndex        =   8
            Top             =   1950
            Width           =   10125
            _cx             =   17859
            _cy             =   3836
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
            Cols            =   13
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmIngresoAlmacen4.frx":2B10
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
         Begin VB.Frame Frame5 
            Caption         =   "[ Detalle de Movimiento ]"
            Height          =   2205
            Left            =   90
            TabIndex        =   43
            Top             =   4260
            Width           =   11685
            Begin VSFlex7Ctl.VSFlexGrid Fg2 
               Height          =   1300
               Left            =   120
               TabIndex        =   44
               Top             =   300
               Width           =   11440
               _cx             =   20179
               _cy             =   2293
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
               Cols            =   14
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmIngresoAlmacen4.frx":2C7F
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
            Begin VB.Frame Frame6 
               Height          =   615
               Left            =   120
               TabIndex        =   45
               Top             =   1530
               Width           =   11445
               Begin VB.CommandButton cmd 
                  Caption         =   "Agregar Pesaje"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   3
                  Left            =   90
                  TabIndex        =   47
                  Top             =   180
                  Width           =   1305
               End
               Begin VB.CommandButton cmd 
                  Caption         =   "Eliminar Pesaje"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   2
                  Left            =   1410
                  TabIndex        =   46
                  Top             =   180
                  Width           =   1305
               End
               Begin VB.Label lblIdIngreso 
                  AutoSize        =   -1  'True
                  Caption         =   "lblIdIngreso"
                  ForeColor       =   &H00000080&
                  Height          =   195
                  Left            =   10200
                  TabIndex        =   48
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   810
               End
            End
         End
         Begin VB.ComboBox cbOperacion 
            Height          =   315
            Left            =   1170
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   840
            Width           =   4635
         End
         Begin VB.CommandButton cmd 
            Height          =   240
            Index           =   1
            Left            =   1820
            Picture         =   "FrmIngresoAlmacen4.frx":2E11
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   1530
            Width           =   240
         End
         Begin VB.CommandButton cmd 
            Height          =   240
            Index           =   0
            Left            =   7710
            Picture         =   "FrmIngresoAlmacen4.frx":2F43
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   1245
            Width           =   240
         End
         Begin VB.CommandButton cmd 
            Height          =   240
            Index           =   6
            Left            =   11430
            Picture         =   "FrmIngresoAlmacen4.frx":3075
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   1575
            Width           =   240
         End
         Begin VB.CommandButton cmd 
            Height          =   240
            Index           =   7
            Left            =   7710
            Picture         =   "FrmIngresoAlmacen4.frx":31A7
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   870
            Width           =   240
         End
         Begin VB.Frame Frame3 
            Enabled         =   0   'False
            Height          =   525
            Left            =   8310
            TabIndex        =   21
            Top             =   240
            Width           =   3405
            Begin VB.ComboBox cbEstado 
               Height          =   315
               Left            =   900
               Style           =   2  'Dropdown List
               TabIndex        =   34
               Top             =   120
               Width           =   2445
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Estado"
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
               Index           =   6
               Left            =   90
               TabIndex        =   35
               Top             =   180
               Width           =   600
            End
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1170
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   2
            Text            =   "TxtNumSer"
            Top             =   1170
            Width           =   915
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2265
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   3
            Text            =   "TxtNumDoc"
            Top             =   1170
            Width           =   1440
         End
         Begin VB.Frame Frame4 
            Height          =   2265
            Left            =   10260
            TabIndex        =   13
            Top             =   1830
            Width           =   1500
            Begin VB.CommandButton cmd 
               Caption         =   "Eliminar Item"
               Enabled         =   0   'False
               Height          =   330
               Index           =   5
               Left            =   90
               TabIndex        =   10
               Top             =   510
               Width           =   1305
            End
            Begin VB.CommandButton cmd 
               Caption         =   "Agregar Item"
               Enabled         =   0   'False
               Height          =   330
               Index           =   4
               Left            =   90
               TabIndex        =   9
               Top             =   180
               Width           =   1305
            End
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchIng 
            Height          =   300
            Left            =   1170
            TabIndex        =   0
            Top             =   540
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
            Left            =   4560
            TabIndex        =   1
            Top             =   540
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
            Left            =   7065
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   5
            Text            =   "TxtIdRes"
            Top             =   840
            Width           =   915
         End
         Begin VB.TextBox txtNumDocRef 
            Height          =   300
            Left            =   7065
            TabIndex        =   7
            Text            =   "txtNumDocRef"
            Top             =   1545
            Width           =   4635
         End
         Begin VB.TextBox TxtIdTipDocRef 
            Height          =   300
            Left            =   7065
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   6
            Text            =   "TxtTipDocR"
            Top             =   1200
            Width           =   915
         End
         Begin VB.TextBox txtIdAlm 
            Height          =   300
            Left            =   1170
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   4
            Text            =   "txtIdAlm"
            Top             =   1500
            Width           =   915
         End
         Begin VB.Label lbltipmov 
            AutoSize        =   -1  'True
            Caption         =   "lbltipmov"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   9240
            TabIndex        =   42
            Top             =   60
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblidtipdoc 
            AutoSize        =   -1  'True
            Caption         =   "lblidtipdoc"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   10020
            TabIndex        =   41
            Top             =   60
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Operación"
            Height          =   195
            Index           =   9
            Left            =   105
            TabIndex        =   40
            Top             =   885
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Almacén"
            Height          =   195
            Index           =   1
            Left            =   105
            TabIndex        =   38
            Top             =   1545
            Width           =   615
         End
         Begin VB.Label lblAlmacen 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblAlmacen"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   2115
            TabIndex        =   37
            Top             =   1515
            Width           =   3660
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tip. Doc. Ref."
            Height          =   195
            Index           =   5
            Left            =   6030
            TabIndex        =   33
            Top             =   1245
            Width           =   1005
         End
         Begin VB.Label LblTipDocRef 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipDocRef"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   8025
            TabIndex        =   32
            Top             =   1230
            Width           =   3675
         End
         Begin VB.Label lbliddocref 
            AutoSize        =   -1  'True
            Caption         =   "lbliddocref"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   10830
            TabIndex        =   30
            Top             =   60
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label LblIdProveedor 
            AutoSize        =   -1  'True
            Caption         =   "LblIdProveedor"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   4710
            TabIndex        =   29
            Top             =   1230
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Doc. Ref."
            Height          =   195
            Index           =   7
            Left            =   6030
            TabIndex        =   28
            Top             =   1590
            Width           =   915
         End
         Begin VB.Label lblRespon 
            AutoSize        =   -1  'True
            Caption         =   "lblRespon"
            Height          =   195
            Left            =   6030
            TabIndex        =   25
            Top             =   885
            Width           =   705
         End
         Begin VB.Label lblResp 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblResp"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   8025
            TabIndex        =   24
            Top             =   855
            Width           =   3675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Documento"
            Height          =   195
            Index           =   4
            Left            =   3180
            TabIndex        =   22
            Top             =   585
            Width           =   1185
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000001&
            BackStyle       =   1  'Opaque
            Height          =   90
            Index           =   0
            Left            =   2115
            Top             =   1290
            Width           =   105
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Num. Doc."
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   20
            Top             =   1185
            Width           =   765
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
            TabIndex        =   15
            Top             =   30
            Width           =   11670
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Mov."
            Height          =   195
            Index           =   3
            Left            =   105
            TabIndex        =   14
            Top             =   585
            Width           =   720
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   26
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
Attribute VB_Name = "FrmIngresoAlmacen4"
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
Dim mMesActivo As Integer                            ' --indica el mes activo
Dim fCierrePeriodo As Boolean                        ' --indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)
Dim IdMenuActivo As Integer                          ' INDICA EL CODIGO DEL MENU ACTIVO
Dim cSQL As String
Dim RSTDETALLE_ As New ADODB.Recordset
Dim CORRELATIVOCAB_ As Integer
Dim CORRELATIVODET_ As Integer

Private Enum COLUMNACABECERA_
    COLUMNAHORA_ = 1
    COLUMNAITEM_
    COLUMNAUNIDAD_
    COLUMNACANENT_
    COLUMNALOTE_
    COLUMNAIDINGDET_
    COLUMNAIDITEM_
    COLUMNAIDLOTE_
    COLUMNAIDLOTEDET_
    COLUMNACANANT_
    COLUMNAIDLOTEANT_
    COLUMNAIDLOTEDETANT_
End Enum

Private Enum COLUMNADETALLE_
    COLUMNAHORA_ = 1
    COLUMNAENVASE_
    COLUMNAUNIMED_
    COLUMNAPESOENV_
    COLUMNAPESOPARIHUELA_
    COLUMNANUMEROENV_
    COLUMNAPBRUTOENV_
    COLUMNAPBRUTOTOTAL_
    COLUMNAPNETOTOTAL_
    COLUMNAOBS_
    COLUMNAID_
    COLUMNAIDENV_
    COLUMNAIDUNIMED_
End Enum

Const ESTADOPENDIENTE_ = 1
Const ESTADOPROCESADO_ = 2
Const ESTADOAPROBADO_ = 3
Const ESTADOANULADO_ = 4

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
    
    If QueHace = 1 Then
        vStr = "SELECT alm_ingreso.numser, alm_ingreso.numdoc " _
            + vbCr + "FROM alm_ingreso " _
            + vbCr + "WHERE ((tipdoc = " & NulosN(txtIdAlm.Text) & ") AND (numser = '" + NulosC(TxtNumSer.Text) + "') AND (numdoc = '" + NulosC(TxtNumDoc.Text) + "'))"
    Else
        vStr = "SELECT alm_ingreso.numser, alm_ingreso.numdoc " _
            + vbCr + "FROM alm_ingreso " _
            + vbCr + "WHERE ((id<>" & NulosN(RstIng("id")) & ") AND (tipdoc = " & NulosN(txtIdAlm.Text) & ") AND (numser = '" + NulosC(TxtNumSer.Text) + "') AND (numdoc = '" + NulosC(TxtNumDoc.Text) + "'))"
    End If
    
    If NulosN(lbltipmov.Caption) = -1 Then
        vStr = vStr & " AND tipmov = true"
    Else
        vStr = vStr & " AND tipmov = false"
    End If
    
    RST_Busq rsVerifDoc, vStr, xCon
    
    If rsVerifDoc.RecordCount > 0 Then
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

Private Sub cbOperacion_Click()
    Select Case NulosN(cbOperacion.ItemData(cbOperacion.ListIndex))
        Case 1 ' RECEPCION
            lblRespon.Caption = "Proveedor"
            TxtIdRes.Text = ""
            lblResp.Caption = ""
            
            LblTipDocRef.Caption = ""
            TxtIdTipDocRef.Text = ""
            txtNumDocRef.Text = ""
            lbliddocref.Caption = ""
            
            lbltipmov.Caption = -1
            lblidtipdoc.Caption = 71
            
        Case 2 ' DESPACHO
            lblRespon.Caption = "Cliente"
            TxtIdRes.Text = ""
            lblResp.Caption = ""
            lbltipmov.Caption = 0
            lblidtipdoc.Caption = 70
            
            LblTipDocRef.Caption = ""
            TxtIdTipDocRef.Text = ""
            txtNumDocRef.Text = ""
            lbliddocref.Caption = ""
            
        Case 3 ' ENTRADA PRODUCCION
            lblRespon.Caption = "Responsable"
            TxtIdRes.Text = ""
            lblResp.Caption = ""
            lbltipmov.Caption = -1
            lblidtipdoc.Caption = 0
            
            LblTipDocRef.Caption = ""
            TxtIdTipDocRef.Text = ""
            txtNumDocRef.Text = ""
            lbliddocref.Caption = ""
            
        Case 4 ' SALIDA PRODUCCION
            lblRespon.Caption = "Responsable"
            TxtIdRes.Text = ""
            lblResp.Caption = ""
            lbltipmov.Caption = 0
            lblidtipdoc.Caption = 0
            
            LblTipDocRef.Caption = ""
            TxtIdTipDocRef.Text = ""
            txtNumDocRef.Text = ""
            lbliddocref.Caption = ""
        
    End Select
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
        Case 0 ' -----------------------------------TIPO DE DOCUMENTO DE REFERENCIA
            ReDim xCampos(2, 4) As String
            
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
                  
            nTitulo = "Buscando Tipo de Documento de Referencia"
            
            Select Case NulosN(cbOperacion.ItemData(cbOperacion.ListIndex))
                Case 1 ' RECEPCION
                    nSQLId = "105,9,1" ' BOLETA DE PAGO/GUIS DE REMISION/FACTURA
                    
                Case 2 ' DESPACHO
                    nSQLId = "101" ' ORDEN DE SALIDA
                    
                Case 3 ' ENTRADA DE PRODUCCION
                    nSQLId = "115" ' ORDEN DE PRODUCCION
                    
                Case 4 ' SALIDA DE PRODUCCION
                    nSQLId = "110" ' SOLICITUD DE MATERIALES
                    
            End Select
            
            cSQL = "SELECT mae_documento.* FROM mae_documento WHERE (id In (" & nSQLId & "))"
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "descripcion", "descripcion", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            TxtIdTipDocRef.Text = NulosN(xRs("id"))
            LblTipDocRef.Caption = UCase(NulosC(xRs("descripcion")))
            txtNumDocRef.SetFocus
            Set xRs = Nothing
        
        Case 1 ' -----------------------------------ALMACEN
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
            TxtIdRes.SetFocus
            Set xRs = Nothing
            
        Case 2 ' ELIMINAR PESAJE
            If Fg2.Rows = 3 Then Fg2.Rows = 1: Exit Sub
            If Fg2.Rows <= 1 Then Exit Sub
            If Fg2.Row = Fg2.Rows - 1 Then Exit Sub
            
            Fg2.RemoveItem Fg2.Row
            ' SE ELIMINA EL RECORDSET
            RSTDETALLE_.Filter = "id=" & NulosN(Fg2.TextMatrix(Fg2.Row, COLUMNAID_))
            limpiarRST RSTDETALLE_, False
            ' SE SELECCIONA LA FILA ANTERIOR
            Fg2.Select Fg2.Row - 1, 1
            Fg2.SetFocus
        
        Case 3 ' AGREGAR PESAJE
            If Fg2.Rows > 2 Then Fg2.Rows = Fg2.Rows - 1
            Fg2.Rows = Fg2.Rows + 1
            ' SE AGREGA EL RECORDSET
            RSTDETALLE_.AddNew
            RSTDETALLE_("id") = CORRELATIVODET_
            RSTDETALLE_("idingdet") = NulosN(Fg1.TextMatrix(Fg1.Row, COLUMNAIDINGDET_))
            RSTDETALLE_.Update
            ' SE AGREGA EL CORRELATIVO
            Fg2.TextMatrix(Fg2.Rows - 1, COLUMNAID_) = CORRELATIVODET_
            CORRELATIVODET_ = CORRELATIVODET_ + 1
            hallarTotales Fg2.Rows - 1, 1
            Fg2.Select Fg2.Rows - 2, 1
            Fg2_EnterCell
            Fg2.SetFocus
            
        Case 4 ' AGREGAR ITEM
            Dim fInsertar As Boolean
            
            ' PERMITE AGREGAR UN ITEM
            If QueHace = 3 Then Exit Sub
            
            Agregando = True
            If Fg1.Rows > Fg1.FixedRows Then
                If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAIDITEM_)) = 0 Then
                    MsgBox "Seleccione un Producto", vbExclamation, xTitulo
                Else
                    fInsertar = True
                End If
            Else
                fInsertar = True
            End If
            
            If fInsertar = True Then
                Fg1.Rows = Fg1.Rows + 1
                Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAIDINGDET_) = CORRELATIVOCAB_
                CORRELATIVOCAB_ = CORRELATIVOCAB_ + 1
            End If
                
            Fg1.Row = Fg1.Rows - 1
            Fg1.Col = COLUMNACABECERA_.COLUMNAHORA_
            
            Fg1.SetFocus
            Agregando = False
        
        Case 5 ' ELIMINAR ITEM
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
            
            ' SE ELIMINAN EL DETALLE DEL DETALLE
            RSTDETALLE_.Filter = "idingdet=" & NulosN(Fg1.TextMatrix(Fg1.Row, COLUMNAIDINGDET_))
            limpiarRST RSTDETALLE_, False
            ' SE ELIMINA EL DETALLE
            Fg1.RemoveItem Fg1.Row
            
            If Fg1.Rows > Fg1.FixedRows Then
                Fg1.Row = Fg1.Rows - 1
                Fg1.Col = COLUMNACABECERA_.COLUMNAHORA_
                Fg1.SetFocus
            Else
                cmd(4).SetFocus
            End If
        
            
        Case 6 ' ------------------------------------DOCUMENTO DE REFERENCIA
            ReDim xCampos(3, 4) As String
            
            'descripcion                        'campo                              'tamaño                     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Fecha":            xCampos(0, 1) = "fecha":            xCampos(0, 2) = "1000":       xCampos(0, 3) = "F"
            xCampos(1, 0) = "Descripcion":      xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "4500":      xCampos(1, 3) = "C"
            xCampos(2, 0) = "Num. Doc.":        xCampos(2, 1) = "numdoc":           xCampos(2, 2) = "1500":       xCampos(2, 3) = "C"
            
            nTitulo = "Buscando Tipos"
            
            IDTIPDOCREF_ = NulosN(TxtIdTipDocRef.Text)
            
            Select Case IDTIPDOCREF_
                Case 115 ' ORDEN DE PRODUCCION
                    cSQL = "SELECT pro_ordenprod.id, pro_ordenprod.fchpro AS fecha, alm_inventario.descripcion, [pro_ordenprod].[numser] & '-' & [pro_ordenprod].[numdoc] AS numdoc " _
                        + vbCr + "FROM (pro_ordenprod INNER JOIN pro_receta ON pro_ordenprod.idrec = pro_receta.id) INNER JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id " _
                        + vbCr + "WHERE (((pro_ordenprod.idmes) In(" & Month(TxtFchIng.Valor) & "," & Month(TxtFchIng.Valor) - 1 & ")));"

                Case 110 ' SOLICITUD DE MATERIALES
                    cSQL = "SELECT pro_solicitudmat.id, pro_solicitudmat.fchdoc AS fecha, [pro_solicitudmat].[numser] & '-' & [pro_solicitudmat].[numdoc] AS numdoc, pla_empleados.nombre AS descripcion, 110 AS iddoc " _
                        + vbCr + "FROM pro_solicitudmat LEFT JOIN pla_empleados ON pro_solicitudmat.idresp = pla_empleados.id " _
                        + vbCr + "WHERE (((pro_solicitudmat.estado)=" & ESTADOPROCESADO_ & ") AND ((pro_solicitudmat.idmes) In (" & Month(TxtFchIng.Valor) & "," & Month(TxtFchIng.Valor) - 1 & ")))"
            End Select
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "numdoc", "numdoc", CualquierParte, ""
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            IDDOCREF_ = NulosN(xRs("id"))
            txtNumDocRef.Text = NulosC(xRs("numdoc"))
            lbliddocref.Caption = IDDOCREF_
            
            cSQL = ""
            Select Case IDTIPDOCREF_
                Case 115 ' ORDEN DE PRODUCCION
                    cSQL = "SELECT pro_receta.iditem, alm_inventario.descripcion AS desitem, pro_ordenprod.cantidad, '' AS idlotedet, pro_ordenprod.idunimed, alm_inventario.tippro AS idtippro, mae_unidades.abrev AS desunimed, mae_tipoproducto.descripcion AS destippro, '' AS deslote " _
                        + vbCr + "FROM (((pro_ordenprod INNER JOIN pro_receta ON pro_ordenprod.idrec = pro_receta.id) INNER JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id) INNER JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id) INNER JOIN mae_unidades ON pro_ordenprod.idunimed = mae_unidades.id " _
                        + vbCr + "WHERE (((pro_ordenprod.id)=" & IDDOCREF_ & "));"
                
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
            
            Fg1.Rows = Fg1.FixedRows
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            xRs.MoveFirst
            While Not xRs.EOF
                Fg1.Rows = Fg1.Rows + 1
                Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAIDITEM_) = NulosN(xRs("iditem"))
                If NulosN(lbltipmov.Caption) = -1 Then
                    Fg1.TextMatrix(Fg1.Rows - 1, COLUMNACANENT_) = Format(NulosN(xRs("cantidad")), FORMAT_CANTIDADDECIMAL)
                Else
                    Fg1.TextMatrix(Fg1.Rows - 1, COLUMNACANENT_) = 0
                End If
                Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAITEM_) = NulosC(xRs("desitem"))
                Fg1.TextMatrix(Fg1.Rows - 1, COLUMNALOTE_) = NulosC(xRs("deslote"))
                Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAIDLOTEDET_) = NulosC(xRs("idlotedet"))
                Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAUNIDAD_) = NulosC(xRs("desunimed"))
                
                xRs.MoveNext
            Wend
            Fg1.SetFocus
            
        Case 7 ' SELECCIONAR RESPONSABLE
            Select Case NulosN(cbOperacion.ItemData(cbOperacion.ListIndex))
                Case 1 ' RECEPCION
                    ' SELECCIONAR PROVEEDOR
                    If QueHace = 3 Then Exit Sub
                    ReDim xCampos(3, 4) As String
                    
                    xCampos(0, 0) = "Nombre":      xCampos(0, 1) = "nombre":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
                    xCampos(1, 0) = "Nº R.U.C.":   xCampos(1, 1) = "numruc":      xCampos(1, 2) = "1200":         xCampos(1, 3) = "C"
                    xCampos(2, 0) = "Codigo":      xCampos(2, 1) = "id":          xCampos(2, 2) = "1000":         xCampos(2, 3) = "N"
                    
                    nTitulo = "Buscando Proveedores"
                    
                    cSQL = "SELECT mae_prov.id, mae_prov.numruc, mae_prov.nombre, mae_prov.activo " _
                        + vbCr + "FROM mae_prov " _
                        + vbCr + "WHERE (((mae_prov.activo) = -1));"
                    
                    Set xRs = Nothing
                    CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                                    "nombre", "nombre", Principio, ""
                    If xRs.State = 0 Then Exit Sub
                    
                    TxtIdRes.Text = NulosN(xRs("id"))
                    lblResp.Caption = NulosC(xRs("nombre"))
                    TxtIdTipDocRef.SetFocus
                    
                Case 2 ' DESPACHO
                    ' SELECCIONAR CLIENTE
                    If QueHace = 3 Then Exit Sub
                    ReDim xCampos(3, 4) As String
                    
                    xCampos(0, 0) = "Nombre":      xCampos(0, 1) = "nombre":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
                    xCampos(1, 0) = "Nº R.U.C.":   xCampos(1, 1) = "numruc":      xCampos(1, 2) = "1200":         xCampos(1, 3) = "C"
                    xCampos(2, 0) = "Codigo":      xCampos(2, 1) = "id":          xCampos(2, 2) = "1000":         xCampos(2, 3) = "N"
                    
                    nTitulo = "Buscando Clientes"
                    
                    cSQL = "SELECT mae_cliente.id, mae_cliente.numruc, mae_cliente.nombre, mae_cliente.activo " _
                        + vbCr + "FROM mae_cliente " _
                        + vbCr + "WHERE (((mae_cliente.activo) = -1));"
                    
                    Set xRs = Nothing
                    CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                                    "nombre", "nombre", Principio, ""
                    If xRs.State = 0 Then Exit Sub
                    
                    TxtIdRes.Text = NulosN(xRs("id"))
                    lblResp.Caption = NulosC(xRs("nombre"))
                    TxtIdTipDocRef.SetFocus
                    
                Case 3, 4 ' ENTREDA PRODUCCION / SALIDA PRODUCCION
                    ' ---SELECCIONAR RESPONSABLE
                    If QueHace = 3 Then Exit Sub
                    ReDim xCampos(2, 4) As String
                    
                    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
                    xCampos(0, 0) = "Apellidos y Nombres":  xCampos(0, 1) = "apenom":     xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
                    xCampos(1, 0) = "Codigo":               xCampos(1, 1) = "id":         xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
                    
                    nTitulo = "Buscando Responsable"
                    
                    cSQL = "SELECT pla_empleados.nombre AS apenom, pla_empleados.id " _
                        + vbCr + "FROM pla_empleados " _
                        + vbCr + "ORDER BY pla_empleados.nombre;"
                    
                    Set xRs = Nothing
                    CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                                    "apenom", "apenom", Principio, ""
                    
                    If xRs.State = 0 Then Exit Sub
                    If xRs.RecordCount = 0 Then Exit Sub
                    
                    TxtIdRes.Text = NulosN(xRs("id"))
                    lblResp.Caption = NulosC(xRs("apenom"))
                    TxtIdTipDocRef.SetFocus
                    
            End Select
            
    End Select
End Sub

Private Sub Fg1_RowColChange()
    pMostrarDetalle NulosN(Fg1.TextMatrix(Fg1.Row, COLUMNAIDINGDET_))
End Sub

Private Sub Fg2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim xRs As New ADODB.Recordset
    Dim nTitulo As String
    Dim xCampos() As String
        
    If QueHace = 3 Then Exit Sub
        
    If Col = COLUMNAENVASE_ Then
        ' BUSCA EL TIPO DE PRODUCTO
        ReDim xCampos(3, 4) As String
        
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Unidad":        xCampos(1, 1) = "abrev":          xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
        xCampos(2, 0) = "Peso":          xCampos(2, 1) = "peso":           xCampos(2, 2) = "1000":         xCampos(2, 3) = "N"
        
        nTitulo = "Buscando Envases"
        
        cSQL = "SELECT alm_inventario.descripcion, mae_equivalencia.idunimed, mae_equivalencia.peso, mae_unidades.abrev, mae_equivalencia.iditem " _
            + vbCr + "FROM (mae_equivalencia LEFT JOIN alm_inventario ON mae_equivalencia.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON mae_equivalencia.idunimed = mae_unidades.id " _
            + vbCr + "WHERE (((alm_inventario.idfam)=122));"
        
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                        "descripcion", "descripcion", Principio, ""
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        Fg2.TextMatrix(Fg2.Row, COLUMNAENVASE_) = NulosC(xRs("descripcion"))
        Fg2.TextMatrix(Fg2.Row, COLUMNAIDENV_) = NulosN(xRs("iditem"))
        Fg2.TextMatrix(Fg2.Row, COLUMNAUNIMED_) = NulosC(xRs("abrev"))
        Fg2.TextMatrix(Fg2.Row, COLUMNAIDUNIMED_) = NulosN(xRs("idunimed"))
        Fg2.TextMatrix(Fg2.Row, COLUMNAPESOENV_) = NulosN(xRs("peso"))
        Fg2.Select Fg2.Row, COLUMNAPESOPARIHUELA_
        
        Set xRs = Nothing
    End If
    
    If Col = COLUMNAUNIMED_ Then
    End If
End Sub

Private Sub Fg2_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim NUEVO_ As Boolean
    
    If Agregando = True Then Exit Sub
    If Row <= 0 Then Exit Sub
    
    Select Case Col
        Case COLUMNAIDENV_
            RSTDETALLE_.Filter = "id=" & NulosN(Fg2.TextMatrix(Row, COLUMNAID_))
            RSTDETALLE_("idenvase") = NulosN(Fg2.TextMatrix(Row, Col))
            RSTDETALLE_.Update
        
        Case COLUMNAIDUNIMED_
            RSTDETALLE_.Filter = "id=" & NulosN(Fg2.TextMatrix(Row, COLUMNAID_))
            RSTDETALLE_("idunimed") = NulosN(Fg2.TextMatrix(Row, Col))
            RSTDETALLE_.Update
            
        Case COLUMNADETALLE_.COLUMNAHORA_
            If IsDate(Fg2.TextMatrix(Row, Col)) Then
                Fg2.TextMatrix(Row, Col) = Format(Fg2.TextMatrix(Row, Col), FORMAT_HORA_SIN_SEGUNDO)
                
                RSTDETALLE_.Filter = "id=" & NulosN(Fg2.TextMatrix(Row, COLUMNAID_))
                RSTDETALLE_("hora") = Fg2.TextMatrix(Row, Col)
                RSTDETALLE_.Update
            Else
                MsgBox "Ingrese una hora correcta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Fg2.TextMatrix(Row, Col) = ""
                
                RSTDETALLE_.Filter = "id=" & NulosN(Fg2.TextMatrix(Row, COLUMNAID_))
                RSTDETALLE_("hora") = Null
                RSTDETALLE_.Update
                Fg2.Col = Col
            End If
        
        Case COLUMNAPESOENV_
            Fg2.TextMatrix(Row, Col) = Format(NulosN(Fg2.TextMatrix(Row, Col)), FORMAT_CANTIDAD)
                
            RSTDETALLE_.Filter = "id=" & NulosN(Fg2.TextMatrix(Row, COLUMNAID_))
            RSTDETALLE_("pesoenv") = NulosN(Fg2.TextMatrix(Row, Col))
            
            Fg2.TextMatrix(Row, COLUMNAPBRUTOENV_) = NulosN(Fg2.TextMatrix(Row, COLUMNANUMEROENV_)) * NulosN(Fg2.TextMatrix(Row, COLUMNAPESOENV_))
            Fg2.TextMatrix(Row, COLUMNAPBRUTOENV_) = Format(Fg2.TextMatrix(Row, COLUMNAPBRUTOENV_), FORMAT_CANTIDAD)
                
            RSTDETALLE_("pesbruenv") = NulosN(Fg2.TextMatrix(Row, COLUMNAPBRUTOENV_))
            
            Fg2.TextMatrix(Row, COLUMNAPNETOTOTAL_) = NulosN(Fg2.TextMatrix(Row, COLUMNAPBRUTOTOTAL_)) - NulosN(Fg2.TextMatrix(Row, COLUMNAPBRUTOENV_)) - NulosN(Fg2.TextMatrix(Row, COLUMNAPESOPARIHUELA_))
            Fg2.TextMatrix(Row, COLUMNAPNETOTOTAL_) = Format(Fg2.TextMatrix(Row, COLUMNAPNETOTOTAL_), FORMAT_CANTIDAD)
                
            RSTDETALLE_("pesnettot") = NulosN(Fg2.TextMatrix(Row, COLUMNAPNETOTOTAL_))
            RSTDETALLE_.Update
            
            hallarTotales Row, Col
            
        Case COLUMNAPESOPARIHUELA_
            Fg2.TextMatrix(Row, Col) = Format(NulosN(Fg2.TextMatrix(Row, Col)), FORMAT_CANTIDAD)
                
            RSTDETALLE_.Filter = "id=" & NulosN(Fg2.TextMatrix(Row, COLUMNAID_))
            RSTDETALLE_("pesopar") = NulosN(Fg2.TextMatrix(Row, Col))
            
            Fg2.TextMatrix(Row, COLUMNAPNETOTOTAL_) = NulosN(Fg2.TextMatrix(Row, COLUMNAPBRUTOTOTAL_)) - NulosN(Fg2.TextMatrix(Row, COLUMNAPBRUTOENV_)) - NulosN(Fg2.TextMatrix(Row, COLUMNAPESOPARIHUELA_))
            Fg2.TextMatrix(Row, COLUMNAPNETOTOTAL_) = Format(Fg2.TextMatrix(Row, COLUMNAPNETOTOTAL_), FORMAT_CANTIDAD)
            
            RSTDETALLE_("pesnettot") = NulosN(Fg2.TextMatrix(Row, COLUMNAPNETOTOTAL_))
            RSTDETALLE_.Update
            
            hallarTotales Row, Col
            
        Case COLUMNANUMEROENV_
            Fg2.TextMatrix(Row, Col) = Format(NulosN(Fg2.TextMatrix(Row, Col)), "000")
                
            RSTDETALLE_.Filter = "id=" & NulosN(Fg2.TextMatrix(Row, COLUMNAID_))
            RSTDETALLE_("numenv") = NulosN(Fg2.TextMatrix(Row, Col))
            
            Fg2.TextMatrix(Row, COLUMNAPBRUTOENV_) = NulosN(Fg2.TextMatrix(Row, COLUMNANUMEROENV_)) * NulosN(Fg2.TextMatrix(Row, COLUMNAPESOENV_))
            Fg2.TextMatrix(Row, COLUMNAPBRUTOENV_) = Format(Fg2.TextMatrix(Row, COLUMNAPBRUTOENV_), FORMAT_CANTIDAD)
            
            RSTDETALLE_("pesbruenv") = NulosN(Fg2.TextMatrix(Row, COLUMNAPBRUTOENV_))
            
            Fg2.TextMatrix(Row, COLUMNAPNETOTOTAL_) = NulosN(Fg2.TextMatrix(Row, COLUMNAPBRUTOTOTAL_)) - NulosN(Fg2.TextMatrix(Row, COLUMNAPBRUTOENV_)) - NulosN(Fg2.TextMatrix(Row, COLUMNAPESOPARIHUELA_))
            Fg2.TextMatrix(Row, COLUMNAPNETOTOTAL_) = Format(Fg2.TextMatrix(Row, COLUMNAPNETOTOTAL_), FORMAT_CANTIDAD)
            
            RSTDETALLE_("pesnettot") = NulosN(Fg2.TextMatrix(Row, COLUMNAPNETOTOTAL_))
            RSTDETALLE_.Update
            
            hallarTotales Row, Col
        
        Case COLUMNAPBRUTOTOTAL_
            Fg2.TextMatrix(Row, Col) = Format(NulosN(Fg2.TextMatrix(Row, Col)), FORMAT_CANTIDAD)
                
            RSTDETALLE_.Filter = "id=" & NulosN(Fg2.TextMatrix(Row, COLUMNAID_))
            RSTDETALLE_("pesbrutot") = NulosN(Fg2.TextMatrix(Row, Col))
            
            Fg2.TextMatrix(Row, COLUMNAPNETOTOTAL_) = NulosN(Fg2.TextMatrix(Row, COLUMNAPBRUTOTOTAL_)) - NulosN(Fg2.TextMatrix(Row, COLUMNAPBRUTOENV_)) - NulosN(Fg2.TextMatrix(Row, COLUMNAPESOPARIHUELA_))
            Fg2.TextMatrix(Row, COLUMNAPNETOTOTAL_) = Format(Fg2.TextMatrix(Row, COLUMNAPNETOTOTAL_), FORMAT_CANTIDAD)
            
            RSTDETALLE_("pesnettot") = NulosN(Fg2.TextMatrix(Row, COLUMNAPNETOTOTAL_))
            RSTDETALLE_.Update
            
            hallarTotales Row, Col
            
    End Select
End Sub

Private Sub hallarTotales(Optional FILA_ As Long, Optional COLUMNA_ As Long)
    Dim NUEVO_ As Boolean
    Dim TOTALBRUTO_ As Double
    Dim TOTALNETO_ As Double
    Dim TOTALPPARIH_ As Double
    Dim TOTALNUMENV_ As Double
    Dim TOTALPBRUTOENV_ As Double
    Dim A As Integer
    
    If Agregando Then Exit Sub
    
    If FILA_ < 0 Or COLUMNA_ < 0 Then Exit Sub
    
    If Fg2.TextMatrix(Fg2.Rows - 1, COLUMNAENVASE_) = "TOTAL" Then
        NUEVO_ = False
    Else
        NUEVO_ = True
    End If
    
    If Not NUEVO_ Then
        Fg2.Rows = Fg2.Rows - 1
    End If
       
    For A = 1 To Fg2.Rows - 1
        TOTALPPARIH_ = TOTALPPARIH_ + NulosN(Fg2.TextMatrix(A, COLUMNAPESOPARIHUELA_))
        TOTALNUMENV_ = TOTALNUMENV_ + NulosN(Fg2.TextMatrix(A, COLUMNANUMEROENV_))
        TOTALPBRUTOENV_ = TOTALPBRUTOENV_ + NulosN(Fg2.TextMatrix(A, COLUMNAPBRUTOENV_))
        TOTALBRUTO_ = TOTALBRUTO_ + NulosN(Fg2.TextMatrix(A, COLUMNAPBRUTOTOTAL_))
        TOTALNETO_ = TOTALNETO_ + NulosN(Fg2.TextMatrix(A, COLUMNAPNETOTOTAL_))
    Next A
    
    Agregando = True
    Fg2.Rows = Fg2.Rows + 1
    Fg2.TextMatrix(Fg2.Rows - 1, COLUMNAENVASE_) = "TOTAL"
    Fg2.TextMatrix(Fg2.Rows - 1, COLUMNAPESOPARIHUELA_) = Format(TOTALPPARIH_, FORMAT_CANTIDAD)
    Fg2.TextMatrix(Fg2.Rows - 1, COLUMNANUMEROENV_) = Format(TOTALNUMENV_, FORMAT_CANTIDAD)
    Fg2.TextMatrix(Fg2.Rows - 1, COLUMNAPBRUTOENV_) = Format(TOTALPBRUTOENV_, FORMAT_CANTIDAD)
    Fg2.TextMatrix(Fg2.Rows - 1, COLUMNAPBRUTOTOTAL_) = Format(TOTALBRUTO_, FORMAT_CANTIDAD)
    Fg2.TextMatrix(Fg2.Rows - 1, COLUMNAPNETOTOTAL_) = Format(TOTALNETO_, FORMAT_CANTIDAD)
    
    ' SE ACTUALIZA LA CANTIDAD GLOBAL
    Fg1.TextMatrix(Fg1.Row, COLUMNACANENT_) = Format(TOTALNETO_, FORMAT_CANTIDADDECIMAL)
    
    Fg2.Select Fg2.Rows - 1, COLUMNAENVASE_, Fg2.Rows - 1, COLUMNAPNETOTOTAL_
    Fg2.FillStyle = flexFillRepeat
    Fg2.CellBackColor = &H8000000F
    Fg2.CellFontBold = True
    Fg2.Select FILA_, COLUMNA_
    
    Agregando = False
    
    'Fg2.SetFocus
End Sub

Private Sub Fg2_EnterCell()
    If QueHace = 3 Then
        Fg2.Editable = flexEDNone
        Fg2.SelectionMode = flexSelectionByRow
        Exit Sub
    Else
        Fg2.SelectionMode = flexSelectionFree
    End If
    
    If Agregando Then Exit Sub
    If Fg2.Rows - 1 < Fg2.FixedRows Then Exit Sub
    If Fg2.Row = Fg2.Rows - 1 Then Exit Sub
    
    Select Case Fg2.Col
        Case COLUMNAPBRUTOTOTAL_, COLUMNANUMEROENV_, COLUMNAPESOPARIHUELA_, _
                                COLUMNAENVASE_, COLUMNAPESOENV_, COLUMNAOBS_, COLUMNADETALLE_.COLUMNAHORA_
            Fg2.Editable = flexEDKbdMouse

        Case Else
            Fg2.Editable = flexEDNone
    End Select
End Sub

Private Sub Fg2_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    
    If KeyCode = 45 Then
        cmd_Click 3
    End If
    If KeyCode = 46 Then
        cmd_Click 2
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
    Dim xRs As New ADODB.Recordset
    Dim nTitulo As String
    Dim xCampos() As String
    Dim TIPOPRODUCTO_ As Double
    Dim iDITEM_ As Double
        
    If QueHace = 3 Then Exit Sub
    
    If Col = COLUMNAITEM_ Then
        ' BUSCA UN ITEM
        ' Se verifica el Almacen
        If NulosN(txtIdAlm.Text) = 0 Then
            MsgBox "Seleccione el Almacén para el registro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            txtIdAlm.SetFocus
            Exit Sub
        End If
        
        ReDim xCampos(3, 4) As String
        
        xCampos(0, 0) = "Producto":    xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":      xCampos(1, 1) = "codpro":        xCampos(1, 2) = "1500":         xCampos(1, 3) = "C"
        xCampos(2, 0) = "Uni. Med":    xCampos(2, 1) = "abrev":         xCampos(2, 2) = "1200":         xCampos(2, 3) = "C"
                
        nTitulo = "Buscando Ítems"
        
        cSQL = "SELECT alm_almacenesdet.iditem, alm_inventario.descripcion, alm_inventario.codpro, mae_unidades.abrev " _
            + vbCr + "FROM ((alm_almacenes INNER JOIN alm_almacenesdet ON alm_almacenes.id = alm_almacenesdet.idalm) INNER JOIN alm_inventario ON alm_almacenesdet.iditem = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
            + vbCr + "WHERE (((alm_almacenes.id) = " & NulosN(txtIdAlm.Text) & ") And ((alm_almacenes.idtippro) = 0)) " _
            + vbCr + "UNION " _
            + vbCr + "SELECT alm_inventario.id AS iditem, alm_inventario.descripcion, alm_inventario.codpro, mae_unidades.abrev " _
            + vbCr + "FROM (alm_almacenes INNER JOIN alm_inventario ON alm_almacenes.idtippro = alm_inventario.tippro) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
            + vbCr + "WHERE (((alm_almacenes.id) = " & NulosN(txtIdAlm.Text) & "))"
        
        Set xRs = Nothing
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                        "descripcion", "descripcion", Principio, ""
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        Fg1.TextMatrix(Row, COLUMNAITEM_) = NulosC(xRs("descripcion"))
        Fg1.TextMatrix(Row, COLUMNAUNIDAD_) = NulosC(xRs("abrev"))
        Fg1.TextMatrix(Row, COLUMNAIDITEM_) = NulosN(xRs("iditem"))
        Fg1.Col = COLUMNACANENT_
        Fg1.SetFocus
    End If
    
    If Col = COLUMNALOTE_ Then
        ReDim xCampos(4, 4) As String
        
        ' Se verifica el Almacen
        If NulosN(txtIdAlm.Text) = 0 Then
            MsgBox "Seleccione el Almacén para el registro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            txtIdAlm.SetFocus
            Exit Sub
        End If
        ' Se verifica si se escogio el producto
        If NulosN(Fg1.TextMatrix(Fg1.Row, COLUMNAIDITEM_)) = 0 Then
            MsgBox "Seleccione el Ítem para el registro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Fg1.Col = COLUMNAIDITEM_
            Exit Sub
        End If
        
        'descripcion                    'campo                          'tamaño                         'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Lote":         xCampos(0, 1) = "deslote":      xCampos(0, 2) = "2000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Fch. Ing.":    xCampos(1, 1) = "fching":       xCampos(1, 2) = "1000":         xCampos(1, 3) = "D"
        xCampos(2, 0) = "Almacen":      xCampos(2, 1) = "desalm":       xCampos(2, 2) = "3000":         xCampos(2, 3) = "C"
        xCampos(3, 0) = "Cantidad":     xCampos(3, 1) = "cantidad":     xCampos(3, 2) = "1000":         xCampos(3, 3) = "N"
         
        nTitulo = "Buscando Lotes de " & NulosC(Fg1.TextMatrix(Fg1.Row, COLUMNAITEM_))
        
        cSQL = "SELECT alm_inventariolotedet.idlote, alm_inventariolotedet.id AS idlotedet, alm_inventariolote.iditem, alm_inventariolotedet.idalm, alm_inventariolote.fching, alm_almacenes.descripcion AS desalm, alm_inventariolotedet.cantidad, alm_inventariolote.descripcion AS deslote " _
            + vbCr + "FROM (alm_inventariolote LEFT JOIN alm_inventariolotedet ON alm_inventariolote.id = alm_inventariolotedet.idlote) LEFT JOIN alm_almacenes ON alm_inventariolotedet.idalm = alm_almacenes.id " _
            + vbCr + "WHERE (((alm_inventariolotedet.idalm)=" & NulosN(txtIdAlm.Text) & ") AND ((alm_inventariolote.iditem)=" & NulosN(Fg1.TextMatrix(Fg1.Row, COLUMNAIDITEM_)) & "))"
        
        Set xRs = Nothing
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                        "deslote", "deslote", Principio, ""
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
                
        ' Lote
        If OptSal.Value Then
            If xRs("cantidad") < NulosN(Fg1.TextMatrix(Row, COLUMNACANENT_)) Then
                MsgBox "El lote seleccionado no contiene stock suficiente", vbCritical + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
        End If
        
        Agregando = True
        Fg1.TextMatrix(Row, COLUMNAIDLOTEDET_) = NulosN(xRs("idlotedet"))
        Fg1.TextMatrix(Row, COLUMNAIDLOTE_) = NulosN(xRs("idlote"))
        Fg1.TextMatrix(Row, COLUMNALOTE_) = NulosC(xRs("deslote"))
        Agregando = False
    End If
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If Row <= 0 Then Exit Sub
    
    Select Case Col
        Case COLUMNACANENT_
            Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), "0.0000")
        
        Case COLUMNACABECERA_.COLUMNAHORA_
            If IsDate(Fg1.TextMatrix(Row, Col)) Then
                Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), FORMAT_HORA_SIN_SEGUNDO)
            Else
                MsgBox "Ingrese una hora correcta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Fg1.TextMatrix(Row, Col) = ""
                Fg1.Col = Col
            End If
            
    End Select
End Sub

'Private Sub Fg1_DblClick()
'    CentrarFrm Frm4
'    Frm4.Visible = True
'End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
        Fg1.SelectionMode = flexSelectionByRow
        Exit Sub
    Else
        Fg1.SelectionMode = flexSelectionFree
    End If
    
    Select Case Fg1.Col
        Case COLUMNAITEM_, COLUMNACANENT_, COLUMNALOTE_, COLUMNACABECERA_.COLUMNAHORA_
            Fg1.Editable = flexEDKbdMouse
            
        Case Else
            Fg1.Editable = flexEDNone
    End Select
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case Col
        Case COLUMNACANENT_
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
            
        Case COLUMNALOTE_
            KeyAscii = 0
    End Select
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    
    If KeyCode = 45 Then
        cmd_Click 4
    End If
    If KeyCode = 46 Then
        cmd_Click 5
    End If
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If QueHace = 3 Then Exit Sub
    If Button = 2 Then PopupMenu menu1
End Sub

Private Sub Form_Activate()
    'Modificado 13/01/11 Johan Castro
    '           Eliminar

    If SeEjecuto = False Then
        SeEjecuto = True
        
        mMesActivo = xMes
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
    
         Dim NomMes As String
         Dim Cerrado As Boolean
        '------------------------------------------------------------------------------------------
        ' bloqueamos los botones del toolbar
        CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
        '------------------------------------------------------------------------------------------
        pCargarDatos
    End If
End Sub

Private Sub iniciarCampos()
    Dim xRs As New ADODB.Recordset
    
    SeEjecuto = False
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Mostrando = False
        
    GRID_COMBOLIST Fg1, COLUMNAITEM_
    GRID_COMBOLIST Fg1, COLUMNALOTE_
    Fg1.ColEditMask(COLUMNACABECERA_.COLUMNAHORA_) = "##:##"
    
    Fg1.ColWidth(COLUMNAIDITEM_) = 0
    Fg1.ColWidth(COLUMNAIDLOTE_) = 0
    Fg1.ColWidth(COLUMNAIDLOTEDET_) = 0
    Fg1.ColWidth(COLUMNACANANT_) = 0
    Fg1.ColWidth(COLUMNAIDLOTEANT_) = 0
    Fg1.ColWidth(COLUMNAIDLOTEDETANT_) = 0
    Fg1.ColWidth(COLUMNAIDINGDET_) = 0
    
    Fg1.AllowUserResizing = flexResizeColumns
    Fg1.ExplorerBar = flexExSortShow
    Fg1.SelectionMode = flexSelectionFree
    Fg1.ForeColorSel = &H80000005
    Fg1.BackColorSel = &H80&
    
    GRID_COMBOLIST Fg2, COLUMNAENVASE_
    GRID_COMBOLIST Fg2, COLUMNAUNIMED_
    Fg2.ColEditMask(COLUMNADETALLE_.COLUMNAHORA_) = "##:##"
    
    Fg2.ColWidth(COLUMNAIDENV_) = 0
    Fg2.ColWidth(COLUMNAIDUNIMED_) = 0
    
    Fg2.AllowUserResizing = flexResizeColumns
    Fg2.ExplorerBar = flexExMove
    Fg2.SelectionMode = flexSelectionByRow
    Fg2.ForeColorSel = &H80000005
    Fg2.BackColorSel = &H80&
    Fg2.WordWrap = True
    
    Fg2.ColWidth(COLUMNADETALLE_.COLUMNAIDENV_) = 0
    Fg2.ColWidth(COLUMNADETALLE_.COLUMNAID_) = 0
    Fg2.ColWidth(COLUMNADETALLE_.COLUMNAIDUNIMED_) = 0
        
    ' SE LLENA LOS TIPO DE OPERACION
    cbOperacion.Clear
    cbOperacion.AddItem "RECEPCION"
    cbOperacion.ItemData(cbOperacion.NewIndex) = 1
    
    cbOperacion.AddItem "DESPACHO"
    cbOperacion.ItemData(cbOperacion.NewIndex) = 2
    
    cbOperacion.AddItem "ENTRADA PRODUCCION"
    cbOperacion.ItemData(cbOperacion.NewIndex) = 3
    
    cbOperacion.AddItem "SALIDA PRODUCCION"
    cbOperacion.ItemData(cbOperacion.NewIndex) = 4
    cbOperacion.ListIndex = 0
    
    
    ' SE AGREGA LOS ESTADOS PARA EL COMBO
    cSQL = "SELECT * FROM mae_estados"
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then GoTo SIGUIENTE_
    If xRs.RecordCount = 0 Then GoTo SIGUIENTE_
    
    cbEstado.Clear
    xRs.MoveFirst
    While Not xRs.EOF
        cbEstado.AddItem UCase(NulosC(xRs("descripcion")))
        cbEstado.ItemData(cbEstado.NewIndex) = NulosN(xRs("id"))
        xRs.MoveNext
    Wend
    
    cbEstado.ListIndex = 0
    CORRELATIVOCAB_ = -999
    CORRELATIVODET_ = -999
    
SIGUIENTE_:
End Sub
'*****************************************************************************************************
'* Nombre           : Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA EL INGRESO DE UN NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Nuevo()
    Dim A As Integer
    Dim xRs As New ADODB.Recordset
    
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
    Fg1.SelectionMode = flexSelectionFree
    
    lbltipmov.Caption = -1
    
    ' SE LLENA TIPO DE OPERACION
    For A = 0 To cbOperacion.ListCount - 1
        If cbOperacion.ItemData(A) = 1 Then
            cbOperacion.ListIndex = A
            cbOperacion_Click
            Exit For
        End If
    Next A
    
    If RSTDETALLE_.State = 0 Then
        cSQL = "SELECT TOP 1 * FROM alm_ingresodetdet"
        Set xRs = Nothing
        RST_Busq xRs, cSQL, xCon
        
        DEFINIR_RST_TMP RSTDETALLE_, xRs
    End If
    
    xHorIni = Time
    TxtFchIng.Valor = Date
    TxtFchDoc.Valor = Date
    TxtFchIng.SetFocus
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A AJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    QueHace = 3
    iniciarCampos
End Sub

Private Sub Form_Resize()
    TabOne1.Width = Me.Width - 90
    TabOne1.Height = Me.Height - 705
    
    Dg1.Width = TabOne1.Width - 135
    Dg1.Height = TabOne1.Height - 930
    
    Fg1.Width = TabOne1.Width - 1755
    Frame4.Left = TabOne1.Width - 1620
    
    Frame5.Width = TabOne1.Width - 195
    Frame5.Height = TabOne1.Height - 4815
    
    Fg2.Width = Frame5.Width - 245
    Fg2.Height = Frame5.Height - 905
    Frame6.Top = Frame5.Height - 675
    Frame6.Width = Frame5.Width - 240
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
    xCampos(0, 0) = "Fch. Mov.":                xCampos(0, 1) = "fching":           xCampos(0, 2) = 0:  xCampos(0, 3) = "1000"
    xCampos(1, 0) = "Mov.":                     xCampos(1, 1) = "movi":             xCampos(1, 2) = 0:  xCampos(1, 3) = "1000"
    xCampos(2, 0) = "Tip. Doc.":                xCampos(2, 1) = "abrev":            xCampos(2, 2) = 0:  xCampos(2, 3) = "1000"
    xCampos(3, 0) = "Nº Documento":             xCampos(3, 1) = "numdoc2":          xCampos(3, 2) = 0:  xCampos(3, 3) = "1050"
    xCampos(4, 0) = "Ítem":                     xCampos(4, 1) = "desitem":          xCampos(4, 2) = 0:  xCampos(4, 3) = "4500"
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

Sub preparaRST(ByRef RST_ As ADODB.Recordset, ByRef RSTDET_ As ADODB.Recordset)
    Dim xRs As New ADODB.Recordset
    
    ' SE DEIFINE EL DETALLE
    cSQL = "SELECT TOP 1 *, 0 AS idlote, 0 AS idloteant, 0 AS idlotedetant, 0 AS canant FROM alm_ingresodet;"
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then Exit Sub
    DEFINIR_RST_TMP RST_, xRs
    
    ' SE DEFINE EL DETALLE DEL DETALLE
    cSQL = "SELECT TOP 1 * FROM alm_ingresodetdet"
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then Exit Sub
    DEFINIR_RST_TMP RSTDET_, xRs
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : Function
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA Alm_ingreso, DEVUELVE VERDADERO SI TIENE EXITO
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Function Grabar() As Boolean
    Dim FCHMOV_ As String
    Dim TIPDOC_ As Integer
    Dim NUMSER_ As String
    Dim IDOPE_ As Integer
    Dim IDPROV_ As Integer
    Dim DESPROV_ As String
    Dim IDESTADO_ As Integer
    Dim IDTIPMOV_ As Integer
    Dim IDTIPDOCREF_ As Integer
    Dim IDDOCREF_ As Integer
    Dim DESDOCREF_ As String
    Dim IDING_ As Integer
    Dim NUMDOC_ As String
    Dim iDITEM_ As Integer
    Dim IDALM_ As Integer
    Dim xRs As New ADODB.Recordset
    Dim xRsDet As New ADODB.Recordset
    Dim xRsAux As New ADODB.Recordset
    Dim A As Integer
    
On Error GoTo ERROR_
    ' VALIDAMOS QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
    If Year(TxtFchIng.Valor) <> AnoTra Then
        MsgBox "El año ingresado en la " & Label3(3).Caption & " no coincide con el Ejercicio" & vbCr & "Corrija la fecha o registre en su año que corresponde", vbInformation, xTitulo
        TxtFchIng.Valor = ""
        TxtFchIng.SetFocus
        Exit Function
    End If
    
    If TxtFchIng.Valor = "" Then
        MsgBox "No ha especificado la fecha de ingreso", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIng.SetFocus
        Exit Function
    End If
    
    If TxtFchDoc.Valor = "" Then
        MsgBox "No ha especificado la fecha del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchDoc.SetFocus
        Exit Function
    End If
        
    If NulosC(txtIdAlm.Text) = "" Then
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        txtIdAlm.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtNumSer.Text) = "" Then
        MsgBox "No ha especificado el numero de serie del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumSer.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtNumDoc.Text) = "" Then
        MsgBox "No ha especificado el numero de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDoc.SetFocus
        Exit Function
    End If
    
    If NulosN(txtIdAlm.Text) = 0 Then
        MsgBox "No ha especificado el nombre del almacén", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        txtIdAlm.SetFocus
        Exit Function
    End If
    
    If NulosN(TxtIdRes.Text) = 0 Then
        MsgBox "Falta especificar " & lblRespon, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdRes.SetFocus
        Exit Function
    End If
    
    If Fg1.Rows = 2 Then
        If Fg1.TextMatrix(1, COLUMNAITEM_) = "" Then
            MsgBox "Falta ingresar el detalle...!", vbExclamation, "Mensaje...!"
            Exit Function
        End If
    ElseIf Fg1.Rows = 1 Then
        MsgBox "Falta ingresar el detalle...!", vbExclamation, "Mensaje...!"
        Exit Function
    End If
    
    ' Se llenan los detalles
    If QueHace = 1 Then IDING_ = 0 Else IDING_ = NulosN(RstIng("id"))
    FCHMOV_ = Format(TxtFchIng.Valor, "dd/mm/yyyy")
    TIPDOC_ = NulosN(lblidtipdoc.Caption)
    NUMSER_ = NulosC(TxtNumSer.Text)
    NUMDOC_ = NulosC(TxtNumDoc.Text)
    
    IDOPE_ = NulosN(cbOperacion.ItemData(cbOperacion.ListIndex))
    IDPROV_ = NulosN(TxtIdRes.Text)
    DESPROV_ = NulosC(lblResp.Caption)
    
    IDESTADO_ = NulosN(cbEstado.ItemData(cbEstado.ListIndex))
    IDALM_ = NulosN(txtIdAlm.Text)
    IDTIPMOV_ = NulosN(lbltipmov.Caption)
    IDTIPDOCREF_ = NulosN(TxtIdTipDocRef.Text)
    IDDOCREF_ = NulosN(lbliddocref.Caption)
    DESDOCREF_ = NulosC(txtNumDocRef.Text)
    ' Se prepara el Recordset
    If xRs.State = 0 Then preparaRST xRs, xRsDet
    limpiarRST xRs
    ' Se llena el recordset
    For A = 1 To Fg1.Rows - 1
        iDITEM_ = NulosN(Fg1.TextMatrix(A, COLUMNAIDITEM_))
        xRsAux.Filter = adFilterNone
        xRsAux.Filter = "iditem=" & iDITEM_
        xRs.AddNew
        xRs("id") = NulosN(Fg1.TextMatrix(A, COLUMNAIDINGDET_))
        xRs("iditem") = iDITEM_
        xRs("cantidad") = NulosN(Fg1.TextMatrix(A, COLUMNACANENT_))
        'xRs("cantteo") = NulosN(Fg1.TextMatrix(A, COLUMNACANTEO_))
        xRs("idlote") = NulosN(Fg1.TextMatrix(A, COLUMNAIDLOTE_))
        xRs("idlotedet") = NulosN(Fg1.TextMatrix(A, COLUMNAIDLOTEDET_))
        xRs("idloteant") = NulosN(Fg1.TextMatrix(A, COLUMNAIDLOTEANT_))
        xRs("idlotedetant") = NulosN(Fg1.TextMatrix(A, COLUMNAIDLOTEDETANT_))
        xRs("canant") = NulosN(Fg1.TextMatrix(A, COLUMNACANANT_))
        xRs("hora") = NulosC(Fg1.TextMatrix(A, COLUMNACABECERA_.COLUMNAHORA_))
        xRs.Update
    Next A
    
    ' SE LLENA DEL DETALLE DEL DETALLE
    RSTDETALLE_.Filter = adFilterNone
    CARGAR_RST_TMP xRsDet, RSTDETALLE_
    
    ' Se graba el movimiento
    Grabar = grabarMovimiento(FCHMOV_, TIPDOC_, NUMSER_, IDOPE_, IDPROV_, DESPROV_, IDESTADO_, _
                                IDTIPMOV_, IDTIPDOCREF_, IDDOCREF_, DESDOCREF_, IDALM_, xRs, xRsDet, IDING_, _
                                NUMDOC_, QueHace, mMesActivo, CInt(AnoTra))
    
    mIdRegistro = IDING_
    Exit Function
    
ERROR_:
    'Resume
    Set xRs = Nothing
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
    TabOne1.CurrTab = 1
    ActivaTool
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Modificando Movimiento"
    QueHace = 2
    Bloquea
    Blanquea
 
    Fg1.Editable = flexEDKbdMouse
    Fg1.Rows = 1
    Fg1.Rows = Fg1.Rows + 1
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
    Dim A As Integer
    
    TxtFchIng.Valor = ""
    TxtFchDoc.Valor = ""
    txtIdAlm.Text = ""
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    TxtIdRes.Text = ""
    lblResp.Caption = ""
    lblAlmacen.Caption = ""
    txtNumDocRef.Text = ""
    TxtIdTipDocRef.Text = ""
    LblTipDocRef.Caption = ""
    
    For A = 1 To cbEstado.ListCount - 1
        cbEstado.ListIndex = A
        If cbEstado.ItemData(A) = 2 Then Exit For
    Next A
    
    lbltipmov.Caption = 0
    lblidtipdoc.Caption = 0
    lbliddocref.Caption = 0
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
    TxtIdRes.Locked = Not TxtIdRes.Locked
    
    habilitar cmd, Not TxtNumDoc.Locked
        
    Frame3.Enabled = Not Frame3.Enabled
    cbOperacion.Locked = Not cbOperacion.Locked
End Sub

Private Sub TxtIdRes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdRes_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        cmd_Click 7
    End If
End Sub

Private Sub TxtIdRes_Validate(Cancel As Boolean)
    If NulosC(TxtIdRes.Text) = "" Then Exit Sub
    Dim xRs As New ADODB.Recordset
    xRs.CursorLocation = adUseClient
    
    cSQL = "SELECT pla_empleados.id, pla_empleados.nombre AS nomsolic " _
        + vbCr + "FROM pla_empleados " _
        + vbCr + "WHERE id = " & NulosN(TxtIdRes.Text) & ""
    
    RST_Busq xRs, cSQL, xCon
    
    If xRs.RecordCount = 0 Then
        TxtIdRes.Text = ""
        lblResp.Caption = ""
    Else
        lblResp.Caption = NulosC(xRs("nomsolic"))
    End If
    
    Set xRs = Nothing
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
        cmd_Click 0
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
    If NulosC(TxtNumDoc.Text) = "" Then Exit Sub
    
    TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000000")
    
    If fVerifSiExistDocum() Then
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
        'TxtNumDoc.Text = HallarNumIngresoAlmacen(TxtNumSer)
        If NulosC(TxtNumDoc.Text) = "" Then TxtNumSer.Text = ""
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
        If NulosN(txtIdAlm.Text) <> 9 Then
            vFiltro = " AND tipdoc = " & NulosN(txtIdAlm.Text) & ""
        End If
    Else
        vFiltro = " AND tipdoc = " & NulosN(txtIdAlm.Text) & ""
    End If

    vStr = "SELECT * FROM alm_ingreso WHERE numser = '" & NulosC(NumSerie) & "'" & vFiltro
    
    vStr = vStr & " ORDER BY numdoc"
    RST_Busq Rst, vStr, xCon
    
    If Rst.RecordCount = 0 Then
        ' SI ESTA VACIO INICIALIZA LA NUMERACION
        If OptIng.Value = True And NulosN(txtIdAlm.Text) <> 9 Then
            HallarNumIngresoAlmacen = "0000000001"
        ElseIf OptSal.Value = True Then
            HallarNumIngresoAlmacen = "0000000001"
        Else
            HallarNumIngresoAlmacen = ""
        End If
    Else
        ' SI ESTA LLENO SUMA UNO AL ULTIMO NUMERO
        Rst.MoveLast
        If OptIng.Value = True And NulosN(txtIdAlm.Text) <> 9 Then
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

Private Sub txtNumDocRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        lbliddocref.Caption = 0
    End If
End Sub

Private Sub txtNumDocRef_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then ' F5
        cmd(6).Value = True
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
        cmd_Click 1
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

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EL DETALLE DEL REGISTRO ACTUAL EN LA PESTAÑA DETALLE DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    '********************************************************************
    ' Modificado: 02/04/2012 - Jose Chacon - Modificar referencias a lote
    '********************************************************************
    Dim xRs As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim A As Integer
    
    If RstIng.RecordCount = 0 Then Exit Sub
    If RstIng.BOF = True Or RstIng.EOF = True Then Exit Sub
    
    cSQL = "SELECT alm_ingreso.* " _
        + vbCr + "FROM alm_ingreso " _
        + vbCr + "WHERE alm_ingreso.id=" & NulosN(RstIng("id"))
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    
    ' SE LLENA ESTADO
    For A = 0 To cbEstado.ListCount - 1
        If cbEstado.ItemData(A) = NulosN(xRs("estado")) Then
            cbEstado.ListIndex = A
            Exit For
        End If
    Next A
    
    ' SE LLENA TIPO DE OPERACION
    For A = 0 To cbOperacion.ListCount - 1
        If cbOperacion.ItemData(A) = NulosN(xRs("idope")) Then
            cbOperacion.ListIndex = A
            cbOperacion_Click
            Exit For
        End If
    Next A
    
    TxtFchIng.Valor = xRs("fching")
    TxtFchDoc.Valor = xRs("fchdoc")
    TxtNumSer.Text = NulosC(xRs("numser"))
    TxtNumDoc.Text = NulosC(xRs("numdoc"))
    
    TxtIdRes.Text = NulosN(xRs("idpro"))
    lblResp.Caption = NulosC(xRs("nombre"))
    
    txtIdAlm.Text = NulosN(xRs("idalm"))
    lblAlmacen.Caption = UCase(Busca_Codigo(NulosN(xRs("idalm")), "id", "descripcion", "alm_almacenes", "N", xCon))
    
    If NulosN(xRs("idtipdocref")) = 0 Then
        TxtIdTipDocRef.Text = ""
        LblTipDocRef.Caption = ""
        lbliddocref.Caption = ""
        txtNumDocRef.Text = ""
    Else
        TxtIdTipDocRef.Text = NulosN(xRs("idtipdocref"))
        LblTipDocRef.Caption = UCase(Busca_Codigo(NulosN(xRs("idtipdocref")), "id", "descripcion", "mae_documento", "N", xCon))
        lbliddocref.Caption = NulosN(xRs("iddocref"))
        txtNumDocRef.Text = NulosC(xRs("desdocref"))
    End If
    
    Mostrando = True
    ' -1: INGRESO, 0 : SALIDA
    lbltipmov.Caption = NulosN(xRs("tipmov"))
        
    Mostrando = False
    cSQL = "SELECT alm_ingresodet.*, alm_inventario.descripcion, mae_unidades.abrev, alm_inventariolote.descripcion AS deslote, alm_inventariolotedet.idlote " _
        + vbCr + "FROM (mae_unidades RIGHT JOIN (alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) ON mae_unidades.id = alm_inventario.idunimed) LEFT JOIN (alm_inventariolote RIGHT JOIN alm_inventariolotedet ON alm_inventariolote.id = alm_inventariolotedet.idlote) ON alm_ingresodet.idlotedet = alm_inventariolotedet.id " _
        + vbCr + "WHERE (((alm_ingresodet.iding) = " & NulosN(RstIng("id")) & "));"
    
    Set RstDet = Nothing
    RST_Busq RstDet, cSQL, xCon

    Fg1.Rows = Fg1.FixedRows
    Fg2.Rows = Fg2.FixedRows
    If RstDet.State = 0 Then Exit Sub
    If RstDet.RecordCount = 0 Then Exit Sub
    
    RstDet.MoveFirst
    While Not RstDet.EOF
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAITEM_) = NulosC(RstDet("descripcion"))
        Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAUNIDAD_) = NulosC(RstDet("abrev"))
        Fg1.TextMatrix(Fg1.Rows - 1, COLUMNACANENT_) = NulosN(RstDet("cantidad"))
        Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAIDITEM_) = NulosN(RstDet("iditem"))
        'Fg1.TextMatrix(Fg1.Rows-1, COLUMNACANTEO_) = NulosN(RstDet("cantteo"))
        Fg1.TextMatrix(Fg1.Rows - 1, COLUMNALOTE_) = NulosC(RstDet("deslote"))
        Fg1.TextMatrix(Fg1.Rows - 1, COLUMNACABECERA_.COLUMNAHORA_) = Format(RstDet("hora"), FORMAT_HORA_SIN_SEGUNDO)
        Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAIDINGDET_) = NulosN(RstDet("id"))
        Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAIDLOTE_) = NulosN(RstDet("idlote"))
        Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAIDLOTEDET_) = NulosN(RstDet("idlotedet"))
        Fg1.TextMatrix(Fg1.Rows - 1, COLUMNACANANT_) = NulosN(RstDet("cantidad"))
        Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAIDLOTEANT_) = NulosN(RstDet("idlote"))
        Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAIDLOTEDETANT_) = NulosN(RstDet("idlotedet"))
        
        RstDet.MoveNext
    Wend
    
    ' SE LLENA EL RECORDSET DE DETALLE
    cSQL = "SELECT alm_ingresodetdet.* " _
        + vbCr + "FROM alm_ingresodetdet " _
        + vbCr + "WHERE (((alm_ingresodetdet.iding) = " & NulosN(RstIng("id")) & ")) " _
        + vbCr + "ORDER BY alm_ingresodetdet.hora;"
        
    Set xRs = Nothing
    Set RSTDETALLE_ = Nothing
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then Exit Sub
    DEFINIR_RST_TMP RSTDETALLE_, xRs
    If xRs.RecordCount = 0 Then Exit Sub
    CARGAR_RST_TMP RSTDETALLE_, xRs
    
    ' SE SELECCIONAD EL PRIMER REGISTRO
    Fg1.Row = 1
    pMostrarDetalle Fg1.TextMatrix(Fg1.Row, COLUMNAIDINGDET_)
End Sub

Private Sub pMostrarDetalle(IDINGDET_ As Integer)
    Agregando = True
    With Fg2
        .Rows = .FixedRows
        
        If RSTDETALLE_.State = 0 Then Agregando = False: Exit Sub
        
        RSTDETALLE_.Filter = "idingdet=" & IDINGDET_
        If RSTDETALLE_.RecordCount = 0 Then Agregando = False: Exit Sub
        
        While Not RSTDETALLE_.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = Format(RSTDETALLE_("hora"), FORMAT_HORA_SIN_SEGUNDO)
            .TextMatrix(.Rows - 1, 2) = NulosC(Busca_Codigo(NulosN(RSTDETALLE_("idenvase")), "id", "descripcion", "alm_inventario", "N", xCon))
            .TextMatrix(.Rows - 1, 3) = NulosC(Busca_Codigo(NulosN(RSTDETALLE_("idunimed")), "id", "abrev", "mae_unidades", "N", xCon))
            .TextMatrix(.Rows - 1, 4) = Format(NulosN(RSTDETALLE_("pesoenv")), FORMAT_CANTIDAD)
            .TextMatrix(.Rows - 1, 5) = Format(NulosN(RSTDETALLE_("pesopar")), FORMAT_CANTIDAD)
            .TextMatrix(.Rows - 1, 6) = Format(NulosN(RSTDETALLE_("numenv")), "000")
            .TextMatrix(.Rows - 1, 7) = Format(NulosN(RSTDETALLE_("pesbruenv")), FORMAT_CANTIDAD)
            .TextMatrix(.Rows - 1, 8) = Format(NulosN(RSTDETALLE_("pesbrutot")), FORMAT_CANTIDAD)
            .TextMatrix(.Rows - 1, 9) = Format(NulosN(RSTDETALLE_("pesnettot")), FORMAT_CANTIDAD)
            .TextMatrix(.Rows - 1, 10) = NulosC(RSTDETALLE_("obs"))
            .TextMatrix(.Rows - 1, 11) = NulosN(RSTDETALLE_("id"))
            .TextMatrix(.Rows - 1, 12) = NulosN(RSTDETALLE_("idenvase"))
            .TextMatrix(.Rows - 1, 13) = NulosN(RSTDETALLE_("idunimed"))
            
            RSTDETALLE_.MoveNext
        Wend
        
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 2) = "TOTAL"
        .TextMatrix(.Rows - 1, 4) = Format(GRID_SUMAR_COL(Fg2, 4), FORMAT_CANTIDAD)
        .TextMatrix(.Rows - 1, 5) = Format(GRID_SUMAR_COL(Fg2, 5), FORMAT_CANTIDAD)
        .TextMatrix(.Rows - 1, 6) = Format(GRID_SUMAR_COL(Fg2, 6), FORMAT_CANTIDAD)
        .TextMatrix(.Rows - 1, 7) = Format(GRID_SUMAR_COL(Fg2, 7), FORMAT_CANTIDAD)
        .TextMatrix(.Rows - 1, 8) = Format(GRID_SUMAR_COL(Fg2, 8), FORMAT_CANTIDAD)
        .TextMatrix(.Rows - 1, 9) = Format(GRID_SUMAR_COL(Fg2, 9), FORMAT_CANTIDAD)
        
        .Select .Rows - 1, 2, .Rows - 1, 9
        .FillStyle = flexFillRepeat
        .CellBackColor = &H8000000F
        .CellFontBold = True
        .Row = .FixedRows
    
    End With
    Agregando = False
End Sub

Sub pCargarDatos()
    TDB_FiltroLimpiar Dg1
    Set RstIng = Nothing
    
    cSQL = "SELECT [alm_ingreso].[id] & '' AS id, Format([alm_ingreso].[fching],'Short Date') AS fchdoc, IIf(alm_ingreso!tipmov=-1,'ING.','SAL.') AS movi, [alm_ingreso]![numser]+'-'+[alm_ingreso]![numdoc] AS numdoc, alm_ingreso.nombre, mae_documento.abrev AS destipdocref, alm_ingreso.desdocref AS numdocref, IIf([alm_ingreso].[idope]=1,'RECEPCION',IIf([alm_ingreso].[idope]=2,'DESPACHO',IIf([alm_ingreso].[idope]=3,'ENTRADA PRODUCCION',IIf([alm_ingreso].[idope]=4,'SALIDA PRODUCCION','')))) AS desope, UCase([mae_estados].[descripcion]) AS desestado " _
        + vbCr + "FROM (alm_ingreso LEFT JOIN mae_estados ON alm_ingreso.estado = mae_estados.id) LEFT JOIN mae_documento ON alm_ingreso.idtipdocref = mae_documento.id " _
        + vbCr + "WHERE (((alm_ingreso.ano) = " & AnoTra & ") And ((alm_ingreso.idmes) = " & mMesActivo & ")) " _
        + vbCr + "ORDER BY Format([alm_ingreso].[fching],'Short Date') DESC;"
        
    RST_Busq RstIng, cSQL, xCon
    Set Dg1.DataSource = RstIng
    
    LblPeriodo.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
End Sub


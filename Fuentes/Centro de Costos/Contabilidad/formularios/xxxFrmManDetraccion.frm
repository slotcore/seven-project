VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmManDetraccion 
   Caption         =   "Contabilidad - Mantenimiento de las Detracciones"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8250
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
            Picture         =   "FrmManDetraccion.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDetraccion.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDetraccion.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDetraccion.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDetraccion.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDetraccion.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDetraccion.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDetraccion.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDetraccion.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDetraccion.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDetraccion.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDetraccion.frx":2236
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7215
      Left            =   0
      TabIndex        =   10
      Top             =   360
      Width           =   11880
      _cx             =   20955
      _cy             =   12726
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
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6795
         Left            =   -12435
         TabIndex        =   13
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6420
            Left            =   30
            TabIndex        =   14
            Top             =   345
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11324
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
            Columns(1).Caption=   "Nº Comp. de Detrac."
            Columns(1).DataField=   "numdet"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Pago"
            Columns(2).DataField=   "fchpag1"
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
            Columns(5).Caption=   "M"
            Columns(5).DataField=   "abremon"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Fch. Doc."
            Columns(6).DataField=   "fchdoc1"
            Columns(6).NumberFormat=   "Short Date"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Proveedor"
            Columns(7).DataField=   "nombre"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Imp. Doc."
            Columns(8).DataField=   "imptot1"
            Columns(8).NumberFormat=   "0.00"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Imp. Det."
            Columns(9).DataField=   "imp1"
            Columns(9).NumberFormat=   "0.00"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   10
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=10"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=953"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=873"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=3440"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=3360"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=1852"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1773"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=873"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=794"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=2566"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2487"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=900"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=820"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=1773"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1693"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=513"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=4763"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=4683"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=512"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(50)=   "Column(8).Width=1720"
            Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=1640"
            Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=514"
            Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(56)=   "Column(9).Width=1799"
            Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=1720"
            Splits(0)._ColumnProps(59)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(60)=   "Column(9)._ColStyle=514"
            Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=74,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=71,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=72,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=73,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=62,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=28,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=51,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=52,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=53,.parent=17"
            _StyleDefs(76)  =   "Named:id=33:Normal"
            _StyleDefs(77)  =   ":id=33,.parent=0"
            _StyleDefs(78)  =   "Named:id=34:Heading"
            _StyleDefs(79)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(80)  =   ":id=34,.wraptext=-1"
            _StyleDefs(81)  =   "Named:id=35:Footing"
            _StyleDefs(82)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(83)  =   "Named:id=36:Selected"
            _StyleDefs(84)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(85)  =   "Named:id=37:Caption"
            _StyleDefs(86)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(87)  =   "Named:id=38:HighlightRow"
            _StyleDefs(88)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(89)  =   "Named:id=39:EvenRow"
            _StyleDefs(90)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(91)  =   "Named:id=40:OddRow"
            _StyleDefs(92)  =   ":id=40,.parent=33"
            _StyleDefs(93)  =   "Named:id=41:RecordSelector"
            _StyleDefs(94)  =   ":id=41,.parent=34"
            _StyleDefs(95)  =   "Named:id=42:FilterBar"
            _StyleDefs(96)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ventas"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   8835
            TabIndex        =   43
            Top             =   0
            Width           =   735
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta Detracciones"
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
            Index           =   0
            Left            =   105
            TabIndex        =   16
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
            TabIndex        =   15
            Top             =   0
            Width           =   1860
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6795
         Left            =   45
         TabIndex        =   11
         Top             =   375
         Width           =   11790
         Begin VB.Frame Frame3 
            Height          =   4365
            Left            =   810
            TabIndex        =   18
            Top             =   1350
            Width           =   9960
            Begin VB.TextBox TxtGlosa 
               Height          =   300
               Left            =   2460
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   45
               Text            =   "TxtGlosa"
               Top             =   3900
               Width           =   6690
            End
            Begin VB.TextBox TxtNumDet 
               Height          =   300
               Left            =   2460
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   8
               Text            =   "TxtNumDet"
               Top             =   3270
               Width           =   1260
            End
            Begin VB.CommandButton CmdBusDetra 
               Height          =   240
               Left            =   3105
               Picture         =   "FrmManDetraccion.frx":277E
               Style           =   1  'Graphical
               TabIndex        =   37
               Top             =   2355
               Width           =   240
            End
            Begin VB.CommandButton CmdBusDoc 
               Height          =   240
               Left            =   3930
               Picture         =   "FrmManDetraccion.frx":28B0
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   1320
               Width           =   240
            End
            Begin VB.CommandButton CmdBusPro 
               Height          =   240
               Left            =   3930
               Picture         =   "FrmManDetraccion.frx":29E2
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   1005
               Width           =   240
            End
            Begin VB.CommandButton CmdBusTipDoc 
               Height          =   240
               Left            =   3105
               Picture         =   "FrmManDetraccion.frx":2B14
               Style           =   1  'Graphical
               TabIndex        =   34
               Top             =   690
               Width           =   240
            End
            Begin VB.TextBox TxtTasa 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2460
               Locked          =   -1  'True
               MaxLength       =   8
               TabIndex        =   6
               Text            =   "TxtTasa"
               Top             =   2640
               Width           =   915
            End
            Begin VB.TextBox TxtIdDet 
               Height          =   300
               Left            =   2460
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   5
               Text            =   "TxtIdDet"
               Top             =   2325
               Width           =   915
            End
            Begin VB.TextBox TxtImpDoc 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2460
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   4
               Text            =   "TxtImpDoc"
               Top             =   1920
               Width           =   1260
            End
            Begin VB.TextBox TxtImpDet 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2460
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   7
               Text            =   "TxtImpDet"
               Top             =   2955
               Width           =   1260
            End
            Begin VB.TextBox TxtNumDoc 
               Height          =   300
               Left            =   2460
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   2
               Text            =   "TxtNumDoc"
               Top             =   1290
               Width           =   1740
            End
            Begin VB.TextBox TxtTipDoc 
               Height          =   300
               Left            =   2460
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   0
               Text            =   "TxtTipDoc"
               Top             =   660
               Width           =   915
            End
            Begin VB.TextBox TxtNumRuc 
               Height          =   300
               Left            =   2460
               Locked          =   -1  'True
               MaxLength       =   11
               TabIndex        =   1
               Text            =   "TxtNumRuc"
               Top             =   975
               Width           =   1740
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchPag 
               Height          =   300
               Left            =   2460
               TabIndex        =   9
               Top             =   3585
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
               Valor           =   "25/03/2008"
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchDoc 
               Height          =   300
               Left            =   2460
               TabIndex        =   41
               Top             =   1590
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
               Valor           =   "25/03/2008"
            End
            Begin VB.TextBox TxtIdMon 
               Height          =   300
               Left            =   6195
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   3
               Text            =   "TxtIdMon"
               Top             =   1290
               Width           =   915
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               Caption         =   "Glosa"
               Height          =   195
               Index           =   4
               Left            =   915
               TabIndex        =   46
               Top             =   3975
               Width           =   405
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Fecha de Emisión"
               Height          =   195
               Index           =   2
               Left            =   915
               TabIndex        =   42
               Top             =   1695
               Width           =   1260
            End
            Begin VB.Label LblIdDocumento 
               AutoSize        =   -1  'True
               Caption         =   "LblIdDocumento"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   6825
               TabIndex        =   40
               Top             =   195
               Visible         =   0   'False
               Width           =   1170
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Fecha de Pago"
               Height          =   195
               Index           =   10
               Left            =   915
               TabIndex        =   39
               Top             =   3645
               Width           =   1095
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Nº Detracción"
               Height          =   195
               Index           =   9
               Left            =   915
               TabIndex        =   38
               Top             =   3320
               Width           =   1005
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Tasa"
               Height          =   195
               Index           =   8
               Left            =   915
               TabIndex        =   33
               Top             =   2670
               Width           =   360
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Detracción"
               Height          =   195
               Index           =   6
               Left            =   915
               TabIndex        =   32
               Top             =   2345
               Width           =   780
            End
            Begin VB.Label LblDetraccion 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblDetraccion"
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
               Left            =   3420
               TabIndex        =   31
               Top             =   2325
               Width           =   5760
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               Caption         =   "Moneda"
               Height          =   195
               Index           =   0
               Left            =   5400
               TabIndex        =   30
               Top             =   1370
               Visible         =   0   'False
               Width           =   585
            End
            Begin VB.Label LblTipoCambio 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblTipoCambio"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   7260
               TabIndex        =   29
               Top             =   1605
               Visible         =   0   'False
               Width           =   1920
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
               Left            =   7260
               TabIndex        =   28
               Top             =   1290
               Width           =   1920
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Proveedor"
               Height          =   195
               Index           =   7
               Left            =   915
               TabIndex        =   27
               Top             =   1045
               Width           =   735
            End
            Begin VB.Label LblNomPro 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblNomPro"
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
               Left            =   4245
               TabIndex        =   26
               Top             =   975
               Width           =   4935
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Documento"
               Height          =   195
               Index           =   1
               Left            =   915
               TabIndex        =   25
               Top             =   720
               Width           =   1410
            End
            Begin VB.Label LblNomDoc 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblNomDoc"
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
               Left            =   3420
               TabIndex        =   24
               Top             =   660
               Width           =   5760
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Nº de Documento"
               Height          =   195
               Index           =   0
               Left            =   915
               TabIndex        =   23
               Top             =   1370
               Width           =   1275
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Moneda"
               Height          =   195
               Index           =   5
               Left            =   7620
               TabIndex        =   22
               Top             =   2310
               Width           =   585
            End
            Begin VB.Label LblIdProveedor 
               AutoSize        =   -1  'True
               Caption         =   "LblIdProveedor"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   8070
               TabIndex        =   21
               Top             =   195
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Importe Documento"
               Height          =   195
               Index           =   3
               Left            =   915
               TabIndex        =   20
               Top             =   2020
               Width           =   1395
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Imp. Detracción"
               Height          =   195
               Index           =   4
               Left            =   915
               TabIndex        =   19
               Top             =   2995
               Width           =   1125
            End
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Ventas"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   8835
            TabIndex        =   44
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Detracciones"
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
            Left            =   120
            TabIndex        =   12
            Top             =   30
            Width           =   11565
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   17
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
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Guia"
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
End
Attribute VB_Name = "FrmManDetraccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim RstDetra As New ADODB.Recordset
Public xTIPO_MOVIMIETO As Integer
Dim xHorIni As Date

Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta
Dim mIdRegistro& '--identificador del registro



Private Sub CmdBusDetra_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Detraccion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Tasa":          xCampos(1, 1) = "tasa":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SqlCad = "SELECT mae_detraccion.id, mae_detraccion.descripcion, mae_detraccion.tasa From mae_detraccion ORDER BY mae_detraccion.descripcion"
    
    xform.Titulo = "Buscando Detraccion"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdDet.Text = xRs("id")
            LblDetraccion.Caption = xRs("descripcion")
            TxtTasa.Text = Format(xRs("tasa"), "0.00")
            TxtImpDet.Text = NulosN(TxtImpDoc.Text) * ((NulosN(TxtTasa.Text) / 100))
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusDoc_Click()
    If QueHace = 3 Then Exit Sub

    If NulosN(TxtTipDoc.Text) = 0 Then
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc.SetFocus
        Exit Sub
    End If
    
    If NulosC(TxtNumRuc.Text) = "" Then
        MsgBox "No ha especificado el proveedor", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumRuc.SetFocus
        Exit Sub
    End If
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(5, 4) As String
    
    xCampos(0, 0) = "Nº Documento":     xCampos(0, 1) = "numdoc":      xCampos(0, 2) = "2000":        xCampos(0, 3) = "C"
    xCampos(1, 0) = "Fch. Emisión":     xCampos(1, 1) = "fchdoc":      xCampos(1, 2) = "1200":        xCampos(1, 3) = "C"
    xCampos(2, 0) = "Fch. Venc":        xCampos(2, 1) = "fchven":      xCampos(2, 2) = "1200":        xCampos(2, 3) = "C"
    xCampos(3, 0) = "M":                xCampos(3, 1) = "simbolo":     xCampos(3, 2) = "500":         xCampos(3, 3) = "C"
    xCampos(4, 0) = "Importe":          xCampos(4, 1) = "imptot":      xCampos(4, 2) = "1200":        xCampos(4, 3) = "N"
    
    If xTIPO_MOVIMIETO = 1 Then
        xform.SqlCad = "SELECT [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, com_compras.fchdoc, com_compras.fchven, mae_moneda.simbolo, idmon, " _
            & " com_compras.imptot, com_compras.id, mae_moneda.descripcion AS nommon FROM mae_moneda RIGHT JOIN com_compras ON mae_moneda.id = com_compras.idmon " _
            & " Where (((com_compras.tipdoc) = " & NulosN(TxtTipDoc.Text) & ") And ((com_compras.idpro) = " & NulosN(LblIdProveedor.Caption) & ")) ORDER BY [com_compras]![numser]+'-'+[com_compras]![numdoc]"
        xform.Titulo = "Buscando Documentos del Proveedor"
    Else
        xform.SqlCad = "SELECT [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, vta_ventas.fchdoc, vta_ventas.fchven, mae_moneda.simbolo, idmon, " _
            & " vta_ventas.imptotdoc as imptot, vta_ventas.id, mae_moneda.descripcion AS nommon FROM mae_moneda RIGHT JOIN vta_ventas ON mae_moneda.id = vta_ventas.idmon " _
            & " Where (((vta_ventas.tipdoc) = " & NulosN(TxtTipDoc.Text) & ") And ((vta_ventas.idcli) = " & NulosN(LblIdProveedor.Caption) & ")) ORDER BY [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc]"
        xform.Titulo = "Buscando Documentos del Cliente"
        
    End If
   
    'SELECT [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, com_compras.fchdoc, com_compras.fchven, mae_moneda.simbolo, " _
        & " com_compras.imptot FROM mae_moneda RIGHT JOIN com_compras ON mae_moneda.id = com_compras.idmon Where " _
        & " (((com_compras.tipdoc) = " & NulosN(TxtTipDoc.Text) & ") And ((com_compras.idpro) = " & NulosN(LblIdProveedor.Caption) & ")) ORDER BY [com_compras]![numser]+'-'+[com_compras]![numdoc]"
    
    
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "numdoc"
    xform.CampoBusca = "numdoc"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
        
            TxtNumDoc.Text = NulosC(xRs("numdoc"))
            LblIdDocumento.Caption = NulosN(xRs("id"))
            
            TxtIdMon.Text = NulosN(xRs("idmon"))
            LblMoneda.Caption = NulosC(xRs("nommon"))
            
            TxtFchDoc.Valor = xRs("fchdoc")
            TxtImpDoc.Text = Format(NulosN(xRs("imptot")), FORMAT_MONTO)
            TxtNumDoc.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusPro_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Proveedor":    xCampos(0, 1) = "nombre":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":    xCampos(1, 1) = "numruc":      xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    If xTIPO_MOVIMIETO = 1 Then
        xform.SqlCad = "SELECT mae_prov.nombre, mae_prov.numruc, mae_prov.id From mae_prov"
        xform.Titulo = "Buscando Proveedor"
    Else
        xform.SqlCad = "SELECT mae_cliente.nombre, mae_cliente.numruc, mae_prov.id From mae_cliente"
        xform.Titulo = "Buscando Cliente"
    End If
    
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtNumRuc.Text = xRs("numruc")
            LblNomPro.Caption = xRs("nombre")
            LblIdProveedor.Caption = xRs("id")
            TxtNumDoc.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipDoc_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
        xform.SqlCad = "SELECT mae_documento.*, mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, con_planctas.cuenta, " _
            & " mae_impuestos.Abrev AS abreimp, mae_impuestos.idcuen  as cuentaimp" _
            & " FROM mae_documento LEFT JOIN (mae_impuestos LEFT JOIN con_planctas ON mae_impuestos.idcuen = con_planctas.id) " _
            & " ON mae_documento.idimp = mae_impuestos.id "
    
    xform.Titulo = "Buscando Tipo de Documento"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtTipDoc.Text = xRs("id")
            LblNomDoc.Caption = xRs("descripcion")
            TxtNumRuc.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstDetra
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstDetra.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 7, NulosN(RstDetra("id")), xCon
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        If xTIPO_MOVIMIETO = 1 Then
            RST_Busq RstDetra, "SELECT con_detraccion.*, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, mae_prov.nombre, " _
                & " com_compras.fchdoc, mae_documento.abrev, mae_moneda.descripcion AS descmon, mae_moneda.simbolo AS abremon, " _
                & " mae_detraccion.descripcion AS descdetra, com_compras.imptot, com_compras.tipdoc, mae_documento.descripcion AS descdoc, " _
                & " mae_prov.numruc, com_compras.idpro, con_detraccion.tipo, " _
                & " con_detraccion.fchpag & '' as fchpag1, com_compras.fchdoc & '' as fchdoc1,com_compras.imptot & '' as imptot1, con_detraccion.imp & '' as imp1 " _
                & " FROM mae_documento RIGHT JOIN (mae_detraccion RIGHT JOIN " _
                & " ((mae_prov RIGHT JOIN com_compras ON mae_prov.id = com_compras.idpro) INNER JOIN (con_detraccion LEFT JOIN mae_moneda " _
                & " ON con_detraccion.idmon = mae_moneda.id) ON com_compras.id = con_detraccion.iddoc) ON mae_detraccion.id = con_detraccion.iddet) " _
                & " ON mae_documento.id = com_compras.tipdoc Where (((con_detraccion.Tipo) = 1)) ORDER BY com_compras.fchdoc DESC", xCon
        Else
            RST_Busq RstDetra, "SELECT con_detraccion.*, vta_ventas!numser+'-'+vta_ventas!numdoc AS numdoc, mae_cliente.nombre, vta_ventas.fchdoc, " _
                & " mae_documento.abrev, mae_moneda.descripcion AS descmon, mae_moneda.simbolo AS abremon, mae_detraccion.descripcion AS descdetra, " _
                & " vta_ventas.imptotdoc AS imptot, vta_ventas.tipdoc, mae_documento.descripcion AS descdoc, mae_cliente.numruc, vta_ventas.idcli AS idpro, " _
                & " con_detraccion.fchpag & '' as fchpag1, vta_ventas.fchdoc & '' as fchdoc1,vta_ventas.imptotdoc & '' as imptot1, con_detraccion.imp & '' as imp1 " _
                & " FROM mae_detraccion RIGHT JOIN ((con_detraccion LEFT JOIN mae_moneda ON con_detraccion.idmon = mae_moneda.id) LEFT JOIN ((vta_ventas " _
                & " LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id) " _
                & " ON con_detraccion.iddoc = vta_ventas.id) ON mae_detraccion.id = con_detraccion.iddet WHERE (((con_detraccion.tipo)=2)) ", xCon
                
        End If
        Set Dg1.DataSource = RstDetra
        If RstDetra.State = 1 Then
            If RstDetra.RecordCount = 0 Then
                MsgBox "No se ha registrado ninguna detraccion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo

            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    TabOne1.CurrTab = 0
    
    Dg1.Columns("fchpag1").NumberFormat = FORMAT_DATE
    Dg1.Columns("fchdoc1").NumberFormat = FORMAT_DATE
    
    Dg1.Columns("imptot1").NumberFormat = FORMAT_MONTO
    Dg1.Columns("imp1").NumberFormat = FORMAT_MONTO
    
    If xTIPO_MOVIMIETO = 1 Then
        'si tipo de movimiento es = 1 es compras
        Label1.Caption = "Compras"
        Label2.Caption = "Compras"
        Dg1.Columns.Item(6).Caption = "Proveedor"
        Label3(7).Caption = "Proveedor"
    Else
        Label1.Caption = "Ventas"
        Label2.Caption = "Ventas"
        Dg1.Columns.Item(6).Caption = "Cliente"
        Label3(7).Caption = "Cliente"
    End If
End Sub

Sub MuestraSegundoTab()
    Blanquea
    If RstDetra.EOF = True Or RstDetra.BOF = True Or RstDetra.RecordCount = 0 Then Exit Sub
    TxtTipDoc.Text = RstDetra("tipdoc")
    LblNomDoc.Caption = RstDetra("descdoc")
    TxtNumRuc.Text = RstDetra("numruc")
    LblNomPro.Caption = RstDetra("nombre")
    LblIdProveedor.Caption = RstDetra("idpro")
    TxtNumDoc.Text = RstDetra("numdoc")
    TxtIdMon.Text = RstDetra("idmon")
    LblMoneda.Caption = RstDetra("descmon")
    TxtFchDoc.Valor = RstDetra("fchdoc")
    TxtImpDoc.Text = Format(RstDetra("imptot"), FORMAT_MONTO)
    TxtIdDet.Text = RstDetra("iddet")
    LblDetraccion.Caption = RstDetra("descdetra")
    TxtTasa.Text = Format(RstDetra("por"), "0.00")
    TxtImpDet.Text = Format(RstDetra("imp"), FORMAT_MONTO)
    
    TxtNumDet.Text = NulosC(RstDetra("numdet"))
    TxtFchPag.Valor = NulosC(RstDetra("fchpag"))
    
    TxtGlosa.Text = NulosC(RstDetra("glosa"))
    
    LblIdDocumento.Caption = RstDetra("iddoc")
End Sub

Sub Bloquea()
    
    TxtTipDoc.Locked = Not TxtTipDoc.Locked
    TxtNumRuc.Locked = Not TxtNumRuc.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    TxtIdMon.Locked = Not TxtIdMon.Locked
    TxtImpDoc.Locked = Not TxtImpDoc.Locked
    TxtIdDet.Locked = Not TxtIdDet.Locked
    TxtTasa.Locked = Not TxtTasa.Locked
    TxtImpDet.Locked = Not TxtImpDet.Locked
    
    TxtGlosa.Locked = Not TxtGlosa.Locked
    
    TxtNumDet.Locked = Not TxtNumDet.Locked
    TxtFchPag.Locked = Not TxtFchPag.Locked
    TxtFchDoc.Locked = Not TxtFchDoc.Locked
End Sub

Sub Blanquea()
    TxtTipDoc.Text = ""
    TxtNumRuc.Text = ""
    TxtNumDoc.Text = ""
    TxtIdMon.Text = ""
    TxtImpDoc.Text = ""
    TxtIdDet.Text = ""
    TxtTasa.Text = ""
    TxtImpDet.Text = ""
    
    TxtFchDoc.Valor = ""
    TxtNumDet.Text = ""
    TxtFchPag.Valor = ""
    
    TxtGlosa.Text = ""
    
    LblNomDoc.Caption = ""
    LblNomPro.Caption = ""
    LblMoneda.Caption = ""
    LblIdDocumento.Caption = ""
    LblDetraccion.Caption = ""
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        MuestraSegundoTab
    End If
End Sub

Sub Nuevo()
    QueHace = 1
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    Bloquea
    Blanquea
    
    Label5.Caption = "Agregando Detraccion"
    xHorIni = Time
    TxtTipDoc.SetFocus
End Sub

Sub Modificar()
    QueHace = 2
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    Bloquea
    Label5.Caption = "Modificando Detraccion"
    xHorIni = Time
    TxtTipDoc.SetFocus
    
End Sub

Sub ActivaTool()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    Toolbar1.Buttons(2).Enabled = Not Toolbar1.Buttons(2).Enabled
    Toolbar1.Buttons(3).Enabled = Not Toolbar1.Buttons(3).Enabled
    
    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
    Toolbar1.Buttons(6).Enabled = Not Toolbar1.Buttons(6).Enabled
    
    Toolbar1.Buttons(8).Enabled = Not Toolbar1.Buttons(8).Enabled
    Toolbar1.Buttons(9).Enabled = Not Toolbar1.Buttons(9).Enabled
    Toolbar1.Buttons(10).Enabled = Not Toolbar1.Buttons(10).Enabled
    Toolbar1.Buttons(11).Enabled = Not Toolbar1.Buttons(11).Enabled
    
    Toolbar1.Buttons(13).Enabled = Not Toolbar1.Buttons(13).Enabled
    Toolbar1.Buttons(15).Enabled = Not Toolbar1.Buttons(15).Enabled
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    
    Rpta = MsgBox("Esta seguro de eliminar la detraccion registrada", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM con_detraccion WHERE id =" & RstDetra("id") & ""
        MsgBox "La detraccion se elimino cone exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstDetra.Requery
        Dg1.Refresh
    End If
End Sub

Function Grabar() As Boolean
    If NulosN(TxtTipDoc.Text) = 0 Then
        MsgBox "No ha especificado el tipo de documento de compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtNumRuc.Text) = "" Then
        MsgBox "No ha especificado en Nº R.U.C.", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumRuc.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtNumDoc.Text) = "" Then
        MsgBox "No ha especificado el numero de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDoc.SetFocus
        Exit Function
    End If
    
    If NulosN(TxtIdDet.Text) = 0 Then
        MsgBox "No ha especificado el tipo de detraccion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdDet.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtNumDet.Text) = "" Then
        MsgBox "No ha especificado el Nº del comprobante de detraccion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDet.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtFchPag.Valor) = "" Then
        MsgBox "No ha especificado la fecha de pago de la detraccion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchPag.SetFocus
        Exit Function
    End If
    
    Dim RstCab As New ADODB.Recordset
    Dim xId As Integer
    
    If QueHace = 1 Then
        xId = HallaCodigoTabla("con_detraccion", xCon, "id")
        RST_Busq RstCab, "SELECT * FROM con_detraccion", xCon
        
        RstCab.AddNew
        RstCab("id") = xId
    Else
        RST_Busq RstCab, "SELECT * FROM con_detraccion WHERE id = " & RstDetra("id") & "", xCon
    End If
    
    RstCab("iddet") = NulosN(TxtIdDet.Text)
    RstCab("por") = NulosN(TxtTasa.Text)
    RstCab("iddoc") = NulosN(LblIdDocumento.Caption)
    RstCab("idmon") = NulosN(TxtIdMon.Text)
    RstCab("tipo") = xTIPO_MOVIMIETO
    RstCab("fchmov") = NulosC(TxtFchDoc.Valor)
    RstCab("glosa") = ""
    RstCab("imp") = NulosN(TxtImpDet.Text)
    RstCab("numdet") = NulosC(TxtNumDet.Text)
    RstCab("fchpag") = CDate(TxtFchPag.Valor)
    RstCab("glosa") = NulosC(TxtGlosa.Text)
    
    RstCab.Update
    
     'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, 7, QueHace, xHorIni, Time, Date, xCon, CDbl(xId)

    Set RstCab = Nothing
    MsgBox "La detracción se grabo con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Exit Function
End Function

Sub Cancelar()
    QueHace = 3
    Label5.Caption = "Detalle de Detracciones"
    ActivaTool
    Bloquea
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo

    If Button.Index = 2 Then Modificar

    If Button.Index = 3 Then Eliminar

    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstDetra.Requery
            Dg1.Refresh
        End If
    End If

    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then
        RstDetra.Filter = ""
        TDB_FiltroLimpiar Dg1
        RstDetra.Requery
    End If
    
    If Button.Index = 15 Then
        Set RstDetra = Nothing
        Unload Me
    End If
    
End Sub

Private Sub TxtIdDet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdDet_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusDetra_Click
    End If
End Sub

Private Sub TxtIdDet_Validate(Cancel As Boolean)
    If NulosN(TxtIdDet.Text) <> 0 Then
        LblDetraccion.Caption = Busca_Codigo(NulosN(TxtIdDet.Text), "id", "descripcion", "mae_detraccion", "N", xCon)
            
        If LblDetraccion.Caption <> "" Then
            TxtTasa.Text = Format(NulosN(Busca_Codigo(NulosN(TxtIdDet.Text), "id", "tasa", "mae_detraccion", "N", xCon)), "0.00")
            TxtImpDet.Text = NulosN(TxtImpDoc.Text) * ((NulosN(TxtTasa.Text) / 100))
            TxtImpDet.Text = Format(TxtImpDet.Text, "0.00")
        Else
            LblDetraccion.Caption = ""
            TxtIdDet.Text = ""
            TxtTasa.Text = ""
            TxtImpDet.Text = ""
        End If
    End If
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtImpDet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtImpDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
    
End Sub

Private Sub TxtNumDet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumDoc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusDoc_Click
    End If
End Sub

Private Sub TxtNumRuc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumRuc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusPro_Click
    End If
End Sub

Private Sub TxtNumRuc_Validate(Cancel As Boolean)
    If TxtNumRuc.Text <> "" Then
        LblNomPro.Caption = Busca_Codigo(NulosC(TxtNumRuc.Text), "numruc", "nombre", "mae_prov", "C", xCon)
        If LblNomPro.Caption = "" Then
            TxtNumRuc.Text = ""
            LblNomPro.Caption = ""
            LblIdProveedor.Caption = ""
            TxtNumDoc.Text = ""
            TxtIdMon.Text = ""
            TxtFchDoc.Valor = ""
            TxtImpDoc.Text = ""
        Else
            LblIdProveedor.Caption = NulosN(Busca_Codigo(NulosC(TxtNumRuc.Text), "numruc", "id", "mae_prov", "C", xCon))
            'TxtIdMon.Text = ""
            'TxtFchDoc.Valor = ""
            'TxtImpDoc.Text = ""
        End If
    Else
            TxtNumRuc.Text = ""
            LblNomPro.Caption = ""
            LblIdProveedor.Caption = ""
            TxtNumDoc.Text = ""
            TxtIdMon.Text = ""
            TxtFchDoc.Valor = ""
            TxtImpDoc.Text = ""
    End If
End Sub

Private Sub TxtTasa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtTipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtTipDoc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipDoc_Click
    End If
End Sub

Private Sub TxtTipDoc_Validate(Cancel As Boolean)
    If NulosN(TxtTipDoc.Text) <> 0 Then
        LblNomDoc.Caption = Busca_Codigo(NulosN(TxtTipDoc.Text), "id", "descripcion", "mae_documento", "N", xCon)
        If LblNomDoc.Caption = "" Then
            TxtTipDoc.Text = ""
        End If
    Else
        LblNomDoc.Caption = ""
    End If
End Sub

VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmManDetraccion 
   Caption         =   "Contabilidad - Registro de las Detracciones"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11895
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
         NumListImages   =   13
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
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDetraccion.frx":277E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7215
      Left            =   0
      TabIndex        =   13
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
         Height          =   6795
         Left            =   45
         TabIndex        =   16
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6420
            Left            =   30
            TabIndex        =   17
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
            Columns(1).Caption=   "Nº Detracción"
            Columns(1).DataField=   "numdet"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Pago"
            Columns(2).DataField=   "fchpag1"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Num.Reg."
            Columns(3).DataField=   "registro"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Proveedor"
            Columns(4).DataField=   "nombre"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "T.D."
            Columns(5).DataField=   "abrev"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Nº Documento"
            Columns(6).DataField=   "numerodoc"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Fch. Doc."
            Columns(7).DataField=   "fchdoc1"
            Columns(7).NumberFormat=   "Short Date"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "M"
            Columns(8).DataField=   "abremon"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Imp. Doc."
            Columns(9).DataField=   "imptot1"
            Columns(9).NumberFormat=   "0.00"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Imp. Det."
            Columns(10).DataField=   "imp1"
            Columns(10).NumberFormat=   "0.00"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   11
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=11"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2355"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2275"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=1773"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1693"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=1640"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1561"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=4260"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=4180"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=847"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=767"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=2566"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=2487"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=1640"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=1561"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=513"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(50)=   "Column(8).Width=900"
            Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=820"
            Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=513"
            Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(56)=   "Column(9).Width=1720"
            Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=1640"
            Splits(0)._ColumnProps(59)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(60)=   "Column(9)._ColStyle=514"
            Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(62)=   "Column(10).Width=1799"
            Splits(0)._ColumnProps(63)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(64)=   "Column(10)._WidthInPix=1720"
            Splits(0)._ColumnProps(65)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(66)=   "Column(10)._ColStyle=514"
            Splits(0)._ColumnProps(67)=   "Column(10).Order=11"
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
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=78,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=75,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=76,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=77,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=43,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=44,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=45,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=58,.parent=13,.alignment=2"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=55,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=56,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=57,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=63,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=64,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=65,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=51,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=52,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=53,.parent=17"
            _StyleDefs(80)  =   "Named:id=33:Normal"
            _StyleDefs(81)  =   ":id=33,.parent=0"
            _StyleDefs(82)  =   "Named:id=34:Heading"
            _StyleDefs(83)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(84)  =   ":id=34,.wraptext=-1"
            _StyleDefs(85)  =   "Named:id=35:Footing"
            _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(87)  =   "Named:id=36:Selected"
            _StyleDefs(88)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(89)  =   "Named:id=37:Caption"
            _StyleDefs(90)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(91)  =   "Named:id=38:HighlightRow"
            _StyleDefs(92)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(93)  =   "Named:id=39:EvenRow"
            _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(95)  =   "Named:id=40:OddRow"
            _StyleDefs(96)  =   ":id=40,.parent=33"
            _StyleDefs(97)  =   "Named:id=41:RecordSelector"
            _StyleDefs(98)  =   ":id=41,.parent=34"
            _StyleDefs(99)  =   "Named:id=42:FilterBar"
            _StyleDefs(100) =   ":id=42,.parent=33"
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
            Index           =   1
            Left            =   8850
            TabIndex        =   45
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
            TabIndex        =   19
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
            TabIndex        =   18
            Top             =   0
            Width           =   1860
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6795
         Left            =   12525
         TabIndex        =   14
         Top             =   375
         Width           =   11790
         Begin VB.CheckBox ChkDoc 
            Alignment       =   1  'Right Justify
            Caption         =   "Mostrar Documentos con Detracción"
            Height          =   285
            Left            =   7740
            TabIndex        =   76
            Top             =   660
            Visible         =   0   'False
            Width           =   2985
         End
         Begin VB.Frame FrmMasivo 
            Height          =   6285
            Left            =   11340
            TabIndex        =   61
            Top             =   5550
            Visible         =   0   'False
            Width           =   11625
            Begin MSComDlg.CommonDialog CommonDialog1 
               Left            =   10980
               Top             =   360
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchPag1 
               Height          =   315
               Left            =   1575
               TabIndex        =   53
               Top             =   720
               Width           =   1485
               _ExtentX        =   2619
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
               Locked          =   -1  'True
            End
            Begin VB.CommandButton CmdBusArch 
               Height          =   240
               Left            =   9060
               Picture         =   "FrmManDetraccion.frx":2B10
               Style           =   1  'Graphical
               TabIndex        =   51
               Top             =   240
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.CommandButton CmdCargar 
               Caption         =   "Cargar"
               Height          =   300
               Left            =   9540
               TabIndex        =   52
               Top             =   210
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.TextBox TxtGlosa1 
               Height          =   330
               Left            =   1575
               Locked          =   -1  'True
               TabIndex        =   55
               Text            =   "TxtGlosa"
               Top             =   1095
               Width           =   9780
            End
            Begin VB.TextBox TxtNumDet1 
               Height          =   300
               Left            =   4470
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   54
               Text            =   "TxtNumDet"
               Top             =   720
               Width           =   2610
            End
            Begin VB.Frame Frame5 
               Height          =   975
               Left            =   90
               TabIndex        =   62
               Top             =   5220
               Width           =   11430
               Begin VB.CommandButton CmdDelItemTodos 
                  Caption         =   "&Eliminar &Todos"
                  Enabled         =   0   'False
                  Height          =   495
                  Left            =   2700
                  TabIndex        =   75
                  Top             =   255
                  Width           =   1170
               End
               Begin VB.TextBox TxtTotDet 
                  Alignment       =   1  'Right Justify
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
                  Left            =   10080
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   73
                  TabStop         =   0   'False
                  Text            =   "TxtTotDet"
                  Top             =   570
                  Width           =   1215
               End
               Begin VB.TextBox TxtTotExMN 
                  Alignment       =   1  'Right Justify
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
                  Left            =   8790
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   71
                  TabStop         =   0   'False
                  Text            =   "TxtTotExMN"
                  Top             =   570
                  Width           =   1215
               End
               Begin VB.TextBox TxtTotME 
                  Alignment       =   1  'Right Justify
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
                  Left            =   7500
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   64
                  TabStop         =   0   'False
                  Text            =   "TxtImpME"
                  Top             =   570
                  Width           =   1200
               End
               Begin VB.TextBox TxtTotMN 
                  Alignment       =   1  'Right Justify
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
                  Left            =   6180
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   63
                  TabStop         =   0   'False
                  Text            =   "TxtImpMN"
                  Top             =   570
                  Width           =   1215
               End
               Begin VB.CommandButton CmdAddItem 
                  Caption         =   "&Agregar Item"
                  Enabled         =   0   'False
                  Height          =   495
                  Left            =   120
                  TabIndex        =   56
                  Top             =   255
                  Width           =   1170
               End
               Begin VB.CommandButton CmdDelItem 
                  Caption         =   "&Eliminar Item"
                  Enabled         =   0   'False
                  Height          =   495
                  Left            =   1410
                  TabIndex        =   57
                  Top             =   255
                  Width           =   1170
               End
               Begin VB.Label Label1 
                  Caption         =   "Imp. Total Detracción"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   405
                  Index           =   2
                  Left            =   10080
                  TabIndex        =   74
                  Top             =   150
                  Width           =   1290
               End
               Begin VB.Label Label1 
                  Caption         =   "Imp. Total Exp. MN"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   0
                  Left            =   8790
                  TabIndex        =   72
                  Top             =   150
                  Width           =   1050
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Imp. Total ME"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   4
                  Left            =   7530
                  TabIndex        =   66
                  Top             =   150
                  Width           =   1200
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Imp. Total MN"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   3
                  Left            =   6180
                  TabIndex        =   65
                  Top             =   150
                  Width           =   1215
               End
               Begin VB.Line Line3 
                  BorderColor     =   &H80000003&
                  Index           =   2
                  X1              =   3960
                  X2              =   3960
                  Y1              =   270
                  Y2              =   750
               End
            End
            Begin VSFlex7Ctl.VSFlexGrid Fg1 
               Height          =   3570
               Left            =   150
               TabIndex        =   58
               Top             =   1530
               Width           =   11385
               _cx             =   20082
               _cy             =   6297
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
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   19
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManDetraccion.frx":2C42
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
            Begin VB.TextBox TxtArchivo 
               Appearance      =   0  'Flat
               Height          =   300
               Left            =   1575
               TabIndex        =   50
               Text            =   "TxtArchivo"
               Top             =   210
               Visible         =   0   'False
               Width           =   7755
            End
            Begin VB.Line Line4 
               X1              =   120
               X2              =   11490
               Y1              =   600
               Y2              =   600
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Archivo"
               Height          =   195
               Index           =   5
               Left            =   150
               TabIndex        =   70
               Top             =   330
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Fecha de Pago"
               Height          =   195
               Index           =   14
               Left            =   90
               TabIndex        =   69
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               Caption         =   "Glosa"
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   68
               Top             =   1170
               Width           =   405
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Nº Detracción"
               Height          =   195
               Index           =   13
               Left            =   3330
               TabIndex        =   67
               Top             =   840
               Width           =   1005
            End
         End
         Begin VB.Frame Frame3 
            Height          =   4935
            Left            =   870
            TabIndex        =   21
            Top             =   720
            Width           =   9960
            Begin VB.TextBox TxtImpDoc 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   300
               Left            =   2385
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   10
               Text            =   "TxtImpDoc"
               Top             =   2280
               Width           =   1260
            End
            Begin VB.CommandButton CmdBusMon 
               Height          =   240
               Left            =   3015
               Picture         =   "FrmManDetraccion.frx":2E6F
               Style           =   1  'Graphical
               TabIndex        =   48
               Top             =   1950
               Width           =   240
            End
            Begin VB.TextBox TxtGlosa 
               Height          =   300
               Left            =   2385
               Locked          =   -1  'True
               TabIndex        =   6
               Text            =   "TxtGlosa"
               Top             =   4320
               Width           =   6720
            End
            Begin VB.TextBox TxtNumDet 
               Height          =   300
               Left            =   2385
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   4
               Text            =   "TxtNumDet"
               Top             =   3690
               Width           =   1830
            End
            Begin VB.CommandButton CmdBusDetra 
               Height          =   240
               Left            =   3045
               Picture         =   "FrmManDetraccion.frx":2FA1
               Style           =   1  'Graphical
               TabIndex        =   40
               Top             =   2775
               Width           =   240
            End
            Begin VB.CommandButton CmdBusDoc 
               Height          =   240
               Left            =   3870
               Picture         =   "FrmManDetraccion.frx":30D3
               Style           =   1  'Graphical
               TabIndex        =   39
               Top             =   870
               Width           =   240
            End
            Begin VB.CommandButton CmdBusPro 
               Height          =   240
               Left            =   3870
               Picture         =   "FrmManDetraccion.frx":3205
               Style           =   1  'Graphical
               TabIndex        =   38
               Top             =   495
               Width           =   240
            End
            Begin VB.CommandButton CmdBusTipDoc 
               Height          =   240
               Left            =   3015
               Picture         =   "FrmManDetraccion.frx":3337
               Style           =   1  'Graphical
               TabIndex        =   37
               Top             =   1230
               Width           =   240
            End
            Begin VB.TextBox TxtTasa 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   300
               Left            =   2385
               Locked          =   -1  'True
               MaxLength       =   8
               TabIndex        =   12
               Text            =   "TxtTasa"
               Top             =   3060
               Width           =   915
            End
            Begin VB.TextBox TxtIdDet 
               Height          =   300
               Left            =   2385
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   2
               Text            =   "TxtIdDet"
               Top             =   2745
               Width           =   915
            End
            Begin VB.TextBox TxtImpDocExpMN 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               Height          =   300
               Left            =   6795
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   11
               Text            =   "TxtImpDocExpMN"
               Top             =   2280
               Width           =   1260
            End
            Begin VB.TextBox TxtImpDet 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2385
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   3
               Text            =   "TxtImpDet"
               Top             =   3375
               Width           =   1260
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchPag 
               Height          =   300
               Left            =   2385
               TabIndex        =   5
               Top             =   4005
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
               Left            =   2385
               TabIndex        =   8
               Top             =   1560
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
               BackColor       =   -2147483633
               Locked          =   -1  'True
               Valor           =   "25/03/2008"
            End
            Begin VB.TextBox TxtNumDoc 
               Height          =   300
               Left            =   2385
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   1
               Text            =   "TxtNumDoc"
               Top             =   828
               Width           =   1740
            End
            Begin VB.TextBox TxtNumRuc 
               Height          =   300
               Left            =   2385
               Locked          =   -1  'True
               MaxLength       =   11
               TabIndex        =   0
               Text            =   "TxtNumRuc"
               Top             =   465
               Width           =   1740
            End
            Begin VB.TextBox TxtTipDoc 
               BackColor       =   &H8000000F&
               Height          =   300
               Left            =   2385
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   7
               Text            =   "TxtTipDoc"
               Top             =   1191
               Width           =   915
            End
            Begin VB.TextBox TxtIdMon 
               BackColor       =   &H8000000F&
               Height          =   300
               Left            =   2385
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   9
               Text            =   "TxtIdMon"
               Top             =   1920
               Width           =   915
            End
            Begin VB.Line Line3 
               BorderColor     =   &H80000005&
               Index           =   0
               X1              =   930
               X2              =   9210
               Y1              =   2670
               Y2              =   2670
            End
            Begin VB.Line Line2 
               X1              =   930
               X2              =   9210
               Y1              =   2640
               Y2              =   2640
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Importe Doc."
               Height          =   195
               Index           =   12
               Left            =   915
               TabIndex        =   60
               Top             =   2370
               Width           =   915
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Nº Registro"
               Height          =   195
               Index           =   11
               Left            =   7020
               TabIndex        =   59
               Top             =   930
               Width           =   810
            End
            Begin VB.Label lblRegistro 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblReg"
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
               Left            =   7920
               TabIndex        =   49
               Top             =   810
               Width           =   1230
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               Caption         =   "Glosa"
               Height          =   195
               Index           =   4
               Left            =   915
               TabIndex        =   47
               Top             =   4395
               Width           =   405
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Fecha de Emisión"
               Height          =   195
               Index           =   2
               Left            =   915
               TabIndex        =   44
               Top             =   1635
               Width           =   1260
            End
            Begin VB.Label LblIdDocumento 
               AutoSize        =   -1  'True
               Caption         =   "LblIdDocumento"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   6825
               TabIndex        =   43
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
               TabIndex        =   42
               Top             =   4095
               Width           =   1095
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Nº Detracción"
               Height          =   195
               Index           =   9
               Left            =   915
               TabIndex        =   41
               Top             =   3765
               Width           =   1005
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Tasa"
               Height          =   195
               Index           =   8
               Left            =   915
               TabIndex        =   36
               Top             =   3120
               Width           =   360
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Detracción"
               Height          =   195
               Index           =   6
               Left            =   915
               TabIndex        =   35
               Top             =   2790
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
               Left            =   3330
               TabIndex        =   34
               Top             =   2745
               Width           =   5790
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               Caption         =   "Moneda"
               Height          =   195
               Index           =   0
               Left            =   915
               TabIndex        =   33
               Top             =   2010
               Width           =   585
            End
            Begin VB.Label LblTipoCambio 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00800000&
               Height          =   300
               Left            =   6795
               TabIndex        =   32
               Top             =   1920
               Width           =   1260
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
               Left            =   3330
               TabIndex        =   31
               Top             =   1920
               Width           =   1920
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Proveedor"
               Height          =   195
               Index           =   7
               Left            =   915
               TabIndex        =   30
               Top             =   570
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
               Left            =   4155
               TabIndex        =   29
               Top             =   465
               Width           =   4995
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Documento"
               Height          =   195
               Index           =   1
               Left            =   915
               TabIndex        =   28
               Top             =   1275
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
               Left            =   3330
               TabIndex        =   27
               Top             =   1170
               Width           =   5820
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Nº de Documento"
               Height          =   195
               Index           =   0
               Left            =   915
               TabIndex        =   26
               Top             =   930
               Width           =   1275
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "T.C."
               Height          =   195
               Index           =   5
               Left            =   6330
               TabIndex        =   25
               Top             =   2010
               Width           =   300
            End
            Begin VB.Label LblIdProveedor 
               AutoSize        =   -1  'True
               Caption         =   "LblIdProveedor"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   8070
               TabIndex        =   24
               Top             =   195
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Importe Exp. MN"
               Height          =   195
               Index           =   3
               Left            =   5445
               TabIndex        =   23
               Top             =   2370
               Width           =   1185
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Imp. Detracción"
               Height          =   195
               Index           =   4
               Left            =   915
               TabIndex        =   22
               Top             =   3450
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
            TabIndex        =   46
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
            TabIndex        =   15
            Top             =   30
            Width           =   11565
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
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
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Individual"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Grupal"
               EndProperty
            EndProperty
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
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir "
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Registro Detracciones"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
'*****************************************************************************************************
'* Nombre Archivo   : FrmMandetraccion.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : PERMITE LAS ALTAS Y BAJAS DE LAS DETRACCIONES
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 19/10/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim QueHace As Integer                  ' ESPECIFICA EN QUE ESTADO SE ENCUENTRA EL FORMULARIO
Dim SeEjecuto As Boolean                ' ESPECIFICA SI EL EVENTO ACTIVATE YA SE EJECUTO
Dim RstDetra As New ADODB.Recordset     ' RECORDSET QUE ALMACENA LAS DETRACCIONES REGISTRADAS
Public xTIPO_MOVIMIETO As Integer       ' ESPECIFICA EL TIPO DE MOVIMIENTO QUE SE ESTA OPERANDO 1 = COMPRAS; 2 = VENTAS
Dim xHorIni As Date                     ' ESPECIFICA LA HORA DE INICIO
Dim fOrdenLista As Boolean              ' especfica el orden de la lista de la consulta
Dim mIdRegistro&                        ' identificador del registro
Dim IdMenuActivo As Integer             'INDICA EL CODIGO DEL MENU ACTIVO

Dim Agregando As Boolean

Private Sub CmdBusDetra_Click()
    ' EJECURA LA BUSQUEDA DE UNA DETRACCION
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(3, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripción":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Tasa (%)":       xCampos(1, 1) = "tasa":             xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    xCampos(2, 0) = "Base":           xCampos(2, 1) = "impbase":          xCampos(2, 2) = "1000":         xCampos(2, 3) = "N"
    
    xform.SqlCad = "SELECT mae_detraccion.* From mae_detraccion ORDER BY mae_detraccion.descripcion"
    
    xform.Titulo = "Buscando Detracción"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            
            If NulosN(xRs("impbase")) <> 0 Then
                If NulosN(TxtImpDocExpMN.Text) < NulosN(xRs("impbase")) Then
                    If MsgBox("El importe expresado en moneda nacional es inferior al importe base para el cálculo de detracción" & vbCr & "Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo) = vbNo Then
                        TxtIdDet.Text = ""
                        LblDetraccion.Caption = ""
                        TxtIdDet.SetFocus
                        Set xform = Nothing
                        Set xRs = Nothing
                        Exit Sub
                    End If
                End If
            End If
        
            TxtIdDet.Text = xRs("id")
            LblDetraccion.Caption = NulosC(xRs("descripcion"))
            TxtTasa.Text = Format(NulosN(xRs("tasa")), "0.00")
            TxtImpDet.Text = NulosN(TxtImpDocExpMN.Text) * ((NulosN(TxtTasa.Text) / 100))
            TxtImpDet.Text = Format(TxtImpDet.Text, FORMAT_MONTO)
            
            TxtImpDet.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusDoc_Click()
    ' EJECUTA LA BUSQUEDA DE UN DOCUMENTO
    If QueHace = 3 Then Exit Sub

    ' VERIFICA QUE LOS DATOS NECESARIOS PARA LA BUSQUEDA ESTEN INGRESADOS
    
    If NulosC(TxtNumRuc.Text) = "" Then
        MsgBox "No ha especificado el proveedor", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumRuc.SetFocus
        Exit Sub
    End If
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(8, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Registro":         xCampos(0, 1) = "registro":    xCampos(0, 2) = "1100":        xCampos(0, 3) = "C"
    xCampos(1, 0) = "T.D.":             xCampos(1, 1) = "abrev":       xCampos(1, 2) = "600":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Nº Documento":     xCampos(2, 1) = "numerodoc":   xCampos(2, 2) = "1500":        xCampos(2, 3) = "C"
    xCampos(3, 0) = "Fch. Emisión":     xCampos(3, 1) = "fchdoc":      xCampos(3, 2) = "1200":        xCampos(3, 3) = "C"
    xCampos(4, 0) = "Fch. Venc":        xCampos(4, 1) = "fchven":      xCampos(4, 2) = "1200":        xCampos(4, 3) = "C"
    xCampos(5, 0) = "M":                xCampos(5, 1) = "simbolo":     xCampos(5, 2) = "500":         xCampos(5, 3) = "C"
    xCampos(6, 0) = "T.C.":             xCampos(6, 1) = "tipcam":      xCampos(6, 2) = "600":         xCampos(6, 3) = "N"
    xCampos(7, 0) = "Importe":          xCampos(7, 1) = "imptot":      xCampos(7, 2) = "1100":        xCampos(7, 3) = "N"
    
    Dim nSQLDocDetra As String
    
    If ChkDoc.Value = 0 Then nSQLDocDetra = " and detra.iddoc is null "
    
    If xTIPO_MOVIMIETO = 1 Then
        ' SI ES COMPRA
        'xform.SqlCad = "SELECT [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, com_compras.fchdoc, com_compras.fchven, mae_moneda.simbolo, idmon, " _
            & " com_compras.imptot, com_compras.id, mae_moneda.descripcion AS nommon, Left([com_compras].[numreg],2) & [mae_libros].[codsun] & Right([com_compras].[numreg],4) AS registro, IIf([com_compras].[tc]=0,[con_tc].[impven],[com_compras].[tc]) AS tipcam FROM (mae_moneda RIGHT JOIN com_compras ON mae_moneda.id = com_compras.idmon) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id " _
            & " Where (((com_compras.tipdoc) = " & NulosN(TxtTipDoc.Text) & ") And ((com_compras.idpro) = " & NulosN(LblIdProveedor.Caption) & ")) ORDER BY [com_compras]![numser]+'-'+[com_compras]![numdoc]"
            
        xform.SqlCad = "SELECT [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numerodoc, com_compras.fchdoc, com_compras.fchven, mae_moneda.simbolo, com_compras.idmon, com_compras.imptot, com_compras.id, mae_moneda.descripcion AS nommon, " _
            + vbCr + " Left([com_compras].[numreg],2) & [mae_libros].[codsun] & Right([com_compras].[numreg],4) AS registro,IIf([com_compras].[tc]=0,[con_tc].[impven],[com_compras].[tc]) AS tipcam, IIf([com_compras].[idmon]=1,[com_compras].[imptot],[tipcam]*[com_compras].[imptot]) AS imptotexmn, com_compras.tipdoc,mae_documento.abrev, mae_documento.descripcion as nomtipdoc " _
            + vbCr + " FROM ( mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon " _
            + vbCr + "       ) LEFT JOIN " _
            + vbCr + " (select con_detraccion.iddoc from  con_detraccion where con_detraccion.tipo=1 " _
            + vbCr + "  ) as detra on com_compras.id = detra.iddoc " _
            + vbCr + " Where com_compras.idpro = " & NulosN(LblIdProveedor.Caption) & " " _
            + vbCr + " ORDER BY [com_compras]![numser]+'-'+[com_compras]![numdoc]; "
    
        xform.Titulo = "Buscando Documentos del Proveedor"
    Else
        ' SI ES VENTA
        'xform.SqlCad = "SELECT [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, vta_ventas.fchdoc, vta_ventas.fchven, mae_moneda.simbolo, idmon, " _
            & " vta_ventas.imptotdoc as imptot, vta_ventas.id, mae_moneda.descripcion AS nommon FROM mae_moneda RIGHT JOIN vta_ventas ON mae_moneda.id = vta_ventas.idmon " _
            & " Where (((vta_ventas.tipdoc) = " & NulosN(TxtTipDoc.Text) & ") And ((vta_ventas.idcli) = " & NulosN(LblIdProveedor.Caption) & ")) ORDER BY [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc]"
            
        xform.SqlCad = "SELECT [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numerodoc, vta_ventas.fchdoc, vta_ventas.fchven, mae_moneda.simbolo, vta_ventas.idmon, vta_ventas.imptotdoc AS imptot, vta_ventas.id, mae_moneda.descripcion AS nommon, " _
            + vbCr + " Left([vta_ventas].[numreg],2) & [mae_libros].[codsun] & Right([vta_ventas].[numreg],4) AS registro, IIf([vta_ventas].[tc]=0,[con_tc].[impven],[vta_ventas].[tc]) AS tipcam, IIf([vta_ventas].[idmon]=1,[vta_ventas].[imptotdoc],[tipcam]*[vta_ventas].[imptotdoc]) AS imptotexmn,vta_ventas.tipdoc,mae_documento.abrev, mae_documento.descripcion as nomtipdoc " _
            + vbCr + " FROM ( mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((vta_ventas LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon " _
            + vbCr + "       ) LEFT JOIN " _
            + vbCr + " (select con_detraccion.iddoc from  con_detraccion where con_detraccion.tipo=2 " _
            + vbCr + "  ) as detra on vta_ventas.id = detra.iddoc " _
            + vbCr + " Where vta_ventas.anulado=0 and vta_ventas.idcli = " & NulosN(LblIdProveedor.Caption) & " " _
            + vbCr + " ORDER BY [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc]; "

        xform.Titulo = "Buscando Documentos del Cliente"
    End If
   
    ' EJECUTAMOS LA BUSQUEDA
    xform.FormaBusca = CualquierParte
    xform.Criterio = ""
    xform.Ordenado = "numerodoc"
    xform.CampoBusca = "numerodoc"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            Dim xRst As New ADODB.Recordset
            
            If ChkDoc.Value = 0 Then
                RST_Busq xRst, "SELECT * FROM con_detraccion WHERE iddoc = " & xRs("id") & " AND tipo = " & xTIPO_MOVIMIETO, xCon
                
                If xRst.RecordCount <> 0 Then
                    MsgBox "La detracción del documento especificado ya se generó" & vbCr & "Busque el documento para que lo modifique", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    Exit Sub
                End If
                Set xRst = Nothing
            End If
            
            TxtNumDoc.Text = NulosC(xRs("numerodoc"))
            LblIdDocumento.Caption = NulosN(xRs("id"))
            TxtIdMon.Text = NulosN(xRs("idmon"))
            LblMoneda.Caption = NulosC(xRs("nommon"))
            TxtTipDoc.Text = NulosN(xRs("tipdoc"))
            LblNomDoc.Caption = NulosC(xRs("nomtipdoc"))
            
            lblRegistro.Caption = NulosC(xRs("registro"))
            LblTipoCambio.Caption = Format(NulosN(xRs("tipcam")), "0.000")
            
            If IsDate(xRs("fchdoc")) = True Then TxtFchDoc.Valor = xRs("fchdoc")
            TxtImpDoc.Text = Format(NulosN(xRs("imptot")), FORMAT_MONTO)
            TxtImpDocExpMN.Text = Format(NulosN(xRs("imptotexmn")), FORMAT_MONTO)
            TxtIdDet.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub CmdBusPro_Click()
    ' EJECUTA LA BUSQUEDA DE UN PROVEEDOR O UN CLIENTE
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Proveedor":    xCampos(0, 1) = "nombre":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":    xCampos(1, 1) = "numruc":      xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    If xTIPO_MOVIMIETO = 1 Then
        ' SI ES COMPRA BUSCA UN PROVEEDOR
        xform.SqlCad = "SELECT mae_prov.nombre, mae_prov.numruc, mae_prov.id From mae_prov where mae_prov.id <>0"
        xform.Titulo = "Buscando Proveedor"
    Else
        ' SI ES VENTA BUSCA UN CLIENTE
        xform.SqlCad = "SELECT mae_cliente.nombre, mae_cliente.numruc, mae_cliente.id From mae_cliente where mae_cliente.id <>0"
        xform.Titulo = "Buscando Cliente"
    End If
    
    ' EJECUTA EL PROCESO DE BUSQUEDA
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtNumRuc.Text = NulosC(xRs("numruc"))
            LblNomPro.Caption = NulosC(xRs("nombre"))
            LblIdProveedor.Caption = xRs("id")
            TxtNumDoc.SetFocus
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
        VerMovimientos1 IdMenuActivo, NulosN(RstDetra("id")), xCon
    End If
End Sub

Private Sub Form_Activate()
    'Modificado: 11/01/11 Johan Castro
    '            Agregar linea de codigo para bloquear accesos de usuarios

    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
        
    If SeEjecuto = False Then
    
    Dim nSQL As String '--Sentencia SQL
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        '----------------------------------------------
        
        If xTIPO_MOVIMIETO = 1 Then
            nSQL = "SELECT con_detraccion.*, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numerodoc, mae_prov.nombre, com_compras.fchdoc, mae_documento.abrev, mae_documento.descripcion AS descdoc, mae_moneda.simbolo AS abremon, mae_moneda.descripcion AS descmon, mae_detraccion.descripcion AS descdetra, com_compras.imptot, " _
                + vbCr + " com_compras.tipdoc, mae_prov.numruc, com_compras.idpro, con_detraccion.fchpag & '' AS fchpag1, com_compras.fchdoc & '' AS fchdoc1, com_compras.imptot & '' AS imptot1, con_detraccion.imp & '' AS imp1, Left([com_compras].[numreg],2) & [mae_libros].[codsun] & Right([com_compras].[numreg],4) AS registro, " _
                + vbCr + " IIf([com_compras].[tc]=0,[con_tc].[impven],[com_compras].[tc]) AS tipcam, " _
                + vbCr + " cdbl(format(IIf([com_compras].[idmon]=1,[com_compras].[imptot],[tipcam]*[com_compras].[imptot]),'0.00')) AS imptotexmn, " _
                + vbCr + " (com_compras.impbru + com_compras.impbru2 + com_compras.impbru3 + com_compras.impina) as impbase, " _
                + vbCr + " cdbl(format(IIf([com_compras].[idmon]=1,(com_compras.impbru + com_compras.impbru2 + com_compras.impbru3 + com_compras.impina) ,[tipcam]*(com_compras.impbru + com_compras.impbru2 + com_compras.impbru3 + com_compras.impina) ),'0.00')) AS impbaseexmn " _
                + vbCr + " FROM (mae_prov RIGHT JOIN (mae_detraccion RIGHT JOIN (mae_documento RIGHT JOIN ((com_compras INNER JOIN (con_detraccion LEFT JOIN mae_moneda ON con_detraccion.idmon = mae_moneda.id) ON com_compras.id = con_detraccion.iddoc) " _
                + vbCr + " LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON mae_documento.id = com_compras.tipdoc) ON mae_detraccion.id = con_detraccion.iddet) ON mae_prov.id = com_compras.idpro) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha " _
                + vbCr + " Where (((con_detraccion.Tipo) = 1)) " _
                + vbCr + " ORDER BY com_compras.fchdoc DESC; "

               
                
        Else
            nSQL = "SELECT con_detraccion.*, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numerodoc, mae_cliente.nombre, vta_ventas.fchdoc, mae_documento.abrev, mae_documento.descripcion AS descdoc, mae_moneda.simbolo AS abremon, mae_moneda.descripcion AS descmon, mae_detraccion.descripcion AS descdetra, vta_ventas.imptotdoc as imptot, " _
                + vbCr + " vta_ventas.tipdoc, mae_cliente.numruc, vta_ventas.idcli AS idpro, con_detraccion.fchpag & '' AS fchpag1, vta_ventas.fchdoc & '' AS fchdoc1, vta_ventas.imptotdoc & '' AS imptot1, con_detraccion.imp & '' AS imp1, Left([vta_ventas].[numreg],2) & [mae_libros].[codsun] & Right([vta_ventas].[numreg],4) AS registro, " _
                + vbCr + " IIf([vta_ventas].[tc]=0,[con_tc].[impven],[vta_ventas].[tc]) AS tipcam, " _
                + vbCr + " format(IIf([vta_ventas].[idmon]=1,[vta_ventas].[imptotdoc],[tipcam]*[vta_ventas].[imptotdoc]),'0.00') AS imptotexmn, " _
                + vbCr + " (vta_ventas.impbru + vta_ventas.impbru2 + vta_ventas.impbru3 + vta_ventas.impinaf) as impbase, " _
                + vbCr + " format(IIf([vta_ventas].[idmon]=1,(vta_ventas.impbru + vta_ventas.impbru2 + vta_ventas.impbru3 + vta_ventas.impinaf) ,[tipcam]*(vta_ventas.impbru + vta_ventas.impbru2 + vta_ventas.impbru3 + vta_ventas.impinaf) ),'0.00') AS impbaseexmn " _
                + vbCr + " FROM (mae_cliente RIGHT JOIN (mae_detraccion RIGHT JOIN (mae_documento RIGHT JOIN ((vta_ventas INNER JOIN (con_detraccion LEFT JOIN mae_moneda ON con_detraccion.idmon = mae_moneda.id) ON vta_ventas.id = con_detraccion.iddoc) " _
                + vbCr + " LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) ON mae_documento.id = vta_ventas.tipdoc) ON mae_detraccion.id = con_detraccion.iddet) ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha " _
                + vbCr + " Where (((con_detraccion.Tipo) = 2)) " _
                + vbCr + " ORDER BY vta_ventas.fchdoc DESC; "
        
        End If
        Set RstDetra = Nothing
        RST_Busq RstDetra, nSQL, xCon
        Set Dg1.DataSource = RstDetra
        
    End If
End Sub

Private Sub Form_Load()
    QueHace = 3
    SeEjecuto = False
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    TabOne1.CurrTab = 0
    
    Dg1.Columns("fchpag1").NumberFormat = FORMAT_DATE
    Dg1.Columns("fchdoc1").NumberFormat = FORMAT_DATE
    
    Dg1.Columns("imptot1").NumberFormat = FORMAT_MONTO
    Dg1.Columns("imp1").NumberFormat = FORMAT_MONTO
    
   
    If xTIPO_MOVIMIETO = 1 Then
        ' SI ES UN COMPRA
        Label2.Caption = "Compras"
        Label1(1).Caption = "Compras"
        Dg1.Columns.Item(4).Caption = "Proveedor"
        Label3(7).Caption = "Proveedor"
    Else
        ' SI ES UNA VENTA
        Label2.Caption = "Ventas"
        Label1(1).Caption = "Ventas"
        Dg1.Columns.Item(4).Caption = "Cliente"
        Label3(7).Caption = "Cliente"
    End If
    
    pGridConfigurar
    
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DETALLADA DEL REGISTRO EN LA PESTAÑA DETALLE DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    If QueHace = 1 Then Exit Sub
    
    Blanquea
    
    If (RstDetra.EOF = True Or RstDetra.BOF = True) And RstDetra.RecordCount = 0 Then Exit Sub
    
    TxtTipDoc.Text = RstDetra("tipdoc")
    LblNomDoc.Caption = NulosC(RstDetra("descdoc"))
    TxtNumRuc.Text = NulosC(RstDetra("numruc"))
    LblNomPro.Caption = NulosC(RstDetra("nombre"))
    LblIdProveedor.Caption = RstDetra("idpro")
    TxtNumDoc.Text = NulosC(RstDetra("numerodoc"))
    TxtIdMon.Text = NulosN(RstDetra("idmon"))
    LblMoneda.Caption = NulosC(RstDetra("descmon"))
    If IsDate(RstDetra("fchdoc")) = True Then TxtFchDoc.Valor = RstDetra("fchdoc")
    TxtImpDoc.Text = Format(NulosN(RstDetra("imptot")), FORMAT_MONTO)
    TxtIdDet.Text = RstDetra("iddet")
    LblDetraccion.Caption = NulosC(RstDetra("descdetra"))
    TxtTasa.Text = Format(NulosN(RstDetra("por")), "0.00")
    TxtImpDet.Text = Format(NulosN(RstDetra("imp")), FORMAT_MONTO)
    
    TxtNumDet.Text = NulosC(RstDetra("numdet"))
    TxtFchPag.Valor = NulosC(RstDetra("fchpag"))
    TxtGlosa.Text = NulosC(RstDetra("glosa"))
    LblIdDocumento.Caption = RstDetra("iddoc")
    
    lblRegistro.Caption = NulosC(RstDetra("registro"))
    LblTipoCambio.Caption = NulosN(RstDetra("tipcam"))
    TxtImpDocExpMN.Text = Format(NulosN(RstDetra("imptotexmn")), FORMAT_MONTO)
    
    
    '--verificar si la detraccion proviene de un grupo
    FrmMasivo.Visible = False
    If NulosN(RstDetra("idgr")) <> 0 Then
        Dim nSQL As String
        Dim xRs As New ADODB.Recordset
        
        '--posicionar el frame
        FrmMasivo.Visible = True
        FrmMasivo.Left = 90
        FrmMasivo.Top = 450
    
        TxtNumDet1.Text = NulosC(RstDetra("numdet"))
        TxtFchPag1.Valor = NulosC(RstDetra("fchpag"))
        TxtGlosa1.Text = NulosC(RstDetra("glosa"))
        
        
        If xTIPO_MOVIMIETO = 1 Then           'Compras
                
                
            nSQL = "SELECT 0 AS xsel, Left(com_compras.numreg,2) & mae_libros.codsun & Right(com_compras.numreg,4) AS registro, mae_prov.numruc, mae_prov.nombre, mae_documento.abrev, IIf(com_compras.numser='',com_compras.numdoc,com_compras.numser+'-'+com_compras.numdoc) AS numerodoc, " _
                + vbCr + " com_compras.fchdoc, IIf(com_compras.tc=0,con_tc.impven,com_compras.tc) & '' AS tipcam, mae_moneda.simbolo, com_compras.imptot, IIf(com_compras.idmon=1,com_compras.imptot,com_compras.imptot*tipcam) AS impexmn, com_compras.glosa, " _
                + vbCr + " com_compras.id as iddoc, com_compras.tipdoc, com_compras.idmon, com_compras.idpro AS idclipro, " _
                + vbCr + " con_detraccion.id AS idcab, con_detraccion.iddet, con_detraccion.idgr, con_detraccion.fchmov, con_detraccion.numdet, mae_detraccion.descripcion AS detraccion, con_detraccion.por, con_detraccion.imp, con_detraccion.glosa AS xglosa " _
                + vbCr + " FROM mae_detraccion RIGHT JOIN (mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (((com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) INNER JOIN con_detraccion ON com_compras.id = con_detraccion.iddoc) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro) ON mae_detraccion.id = con_detraccion.iddet " _
                + vbCr + " WHERE (((con_detraccion.tipo)=1) AND ((con_detraccion.idgr)=" & NulosN(RstDetra("idgr")) & ")) " & _
                vbCr + " ORDER BY mae_prov.nombre, IIf(com_compras.numser='',com_compras.numdoc,com_compras.numser+'-'+com_compras.numdoc) "
                
        Else           'Ventas
            
            nSQL = "SELECT 0 AS xsel, Left(vta_ventas.numreg,2) & mae_libros.codsun & Right(vta_ventas.numreg,4) AS registro, mae_cliente.numruc, mae_cliente.nombre, mae_documento.abrev, vta_ventas.numser+'-'+vta_ventas.numdoc AS numerodoc, vta_ventas.fchdoc, " _
                + vbCr + " mae_moneda.simbolo, IIf(vta_ventas.tc=0,con_tc.impven,vta_ventas.tc) & '' AS tipcam, vta_ventas.imptotdoc AS imptot, IIf(vta_ventas.idmon=1,vta_ventas.imptotdoc,vta_ventas.imptotdoc*tipcam) AS impexmn, " _
                + vbCr + " vta_ventas.id as iddoc, vta_ventas.tipdoc, vta_ventas.idmon, vta_ventas.idcli AS idclipro, vta_ventas.glosa, " _
                + vbCr + " con_detraccion.id AS idcab, con_detraccion.iddet, con_detraccion.idgr, con_detraccion.fchmov, con_detraccion.numdet, mae_detraccion.descripcion AS detraccion, con_detraccion.por, con_detraccion.imp, con_detraccion.glosa AS xglosa " _
                + vbCr + " FROM mae_detraccion RIGHT JOIN (mae_libros RIGHT JOIN (mae_documento RIGHT JOIN ((((mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) INNER JOIN con_detraccion ON vta_ventas.id = con_detraccion.iddoc) ON mae_documento.id = vta_ventas.tipdoc) ON mae_libros.id = vta_ventas.idlib) ON mae_detraccion.id = con_detraccion.iddet " _
                + vbCr + " WHERE (((con_detraccion.tipo)=2) AND ((con_detraccion.idgr)=" & NulosN(RstDetra("idgr")) & ")) " _
                + vbCr + " ORDER BY mae_cliente.nombre, vta_ventas.numser+'-'+vta_ventas.numdoc DESC "
        
        End If
        
        RST_Busq xRs, nSQL, xCon
        
        Fg1.Rows = 1

        If xRs.State = 1 Then
            If xRs.RecordCount = 0 Then
                FrmMasivo.Visible = False
                Set xRs = Nothing
                Exit Sub
            End If
            
            Agregando = True
            
            xRs.MoveFirst
            Do While Not xRs.EOF
                With Me.Fg1
                    .AddItem ""
                    .Row = .Rows - 1
                    .TextMatrix(.Row, 1) = NulosC(xRs("nombre"))
                    .TextMatrix(.Row, 2) = NulosC(xRs("registro"))
                    .TextMatrix(.Row, 3) = NulosC(xRs("abrev"))
                    .TextMatrix(.Row, 4) = NulosC(xRs("numerodoc"))
                    .TextMatrix(.Row, 5) = Format(NulosC(xRs("fchdoc")), FORMAT_DATE)
                    .TextMatrix(.Row, 6) = NulosC(xRs("simbolo"))
                    .TextMatrix(.Row, 7) = NulosN(xRs("tipcam"))
                    .TextMatrix(.Row, 8) = Format(NulosN(xRs("imptot")), FORMAT_MONTO)
                    .TextMatrix(.Row, 9) = Format(NulosN(xRs("impexmn")), FORMAT_MONTO)
                    
                    .TextMatrix(.Row, 10) = NulosN(xRs("iddet"))
                    .TextMatrix(.Row, 11) = NulosC(xRs("detraccion"))
                    .TextMatrix(.Row, 12) = NulosC(xRs("por"))
                    .TextMatrix(.Row, 13) = Format(NulosN(xRs("imp")), FORMAT_MONTO)
                    
                    .TextMatrix(.Row, 14) = NulosN(xRs("iddoc"))
                    .TextMatrix(.Row, 15) = NulosN(xRs("idclipro"))
                    .TextMatrix(.Row, 16) = NulosN(xRs("tipdoc"))
                    .TextMatrix(.Row, 17) = NulosN(xRs("idmon"))
                    
                    .TextMatrix(.Row, 18) = NulosN(xRs("idcab"))
                    
                    
                End With
                
                xRs.MoveNext
            Loop
            Agregando = False
        End If
        HallarTotal
        Set xRs = Nothing
        
    End If
    
    
End Sub

'*****************************************************************************************************
'* Nombre           : Bloquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES TextBox DEL FORMULARIO, PARA AGREGAR O
'*                    MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Bloquea()
    TxtNumRuc.Locked = Not TxtNumRuc.Locked
    
    TxtIdDet.Locked = Not TxtIdDet.Locked
    TxtImpDet.Locked = Not TxtImpDet.Locked
    TxtNumDet.Locked = Not TxtNumDet.Locked
    TxtFchPag.Locked = Not TxtFchPag.Locked
    TxtGlosa.Locked = Not TxtGlosa.Locked
    
    '---
    TxtIdMon.Locked = True
    TxtImpDoc.Locked = True
    TxtFchDoc.Locked = True
    CmdBusMon.Enabled = True
    
    TxtNumDet1.Locked = Not TxtNumDet1.Locked
    TxtFchPag1.Locked = Not TxtFchPag1.Locked
    TxtGlosa1.Locked = Not TxtGlosa1.Locked
    CmdAddItem.Enabled = Not CmdAddItem.Enabled
    CmdDelItem.Enabled = Not CmdDelItem.Enabled
    CmdDelItemTodos.Enabled = Not CmdDelItemTodos.Enabled

    ChkDoc.Visible = Not ChkDoc.Visible
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : INICIALIZA LOS CONTROLES TextBox PARA EL INGRESO DE NUEVOS DATOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    TxtTipDoc.Text = ""
    TxtNumRuc.Text = ""
    TxtNumDoc.Text = ""
    TxtIdMon.Text = ""
    TxtImpDoc.Text = ""
    TxtIdDet.Text = ""
    TxtTasa.Text = ""
    TxtImpDet.Text = ""
    
    TxtFchPag.Valor = Date
    TxtFchDoc.Valor = ""
    TxtNumDet.Text = ""
    TxtFchPag.Valor = ""
    
    TxtGlosa.Text = ""
    
    LblNomDoc.Caption = ""
    LblNomPro.Caption = ""
    LblMoneda.Caption = ""
    LblIdDocumento.Caption = ""
    LblDetraccion.Caption = ""
    
    lblRegistro.Caption = ""
    LblTipoCambio.Caption = ""
    TxtImpDoc.Text = "0.00"
    TxtImpDocExpMN.Text = "0.00"
    
    '--------
    TxtFchPag1.Valor = Date
    TxtNumDet1.Text = ""
    TxtFchPag1.Valor = ""
    TxtGlosa1.Text = ""
    
    TxtTotMN.Text = "0.00"
    TxtTotME.Text = "0.00"
    TxtTotExMN.Text = "0.00"
    TxtTotDet.Text = "0.00"
    Fg1.Rows = 1
    
    ChkDoc.Value = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario miestras este agregando o editando una Detracción", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        MuestraSegundoTab
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
    TabOne1.TabEnabled(0) = False
    ActivaTool
    Blanquea
    Bloquea
    Label5.Caption = "Agregando Detracción"
    xHorIni = Time
    
    If FrmMasivo.Visible = False Then
        TxtNumRuc.SetFocus
    Else
        TxtFchPag1.SetFocus
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
    QueHace = 2
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    Bloquea
    Label5.Caption = "Modificando Detracción"
    xHorIni = Time
    
    If FrmMasivo.Visible = False Then
        TxtNumRuc.SetFocus
    Else
        TxtFchPag1.SetFocus
    End If
    
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LA BARRA DE HERRAMIENTAS DEL FORMULARIO
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
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO DE LA TABLA con_detraccion
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    Dim Rpta As Integer
    If (RstDetra.EOF = True Or RstDetra.BOF = True) And RstDetra.RecordCount <> 0 Then
        MsgBox "Seleccione un registro correcto", vbInformation, xTitulo
        Exit Sub
    End If
    If RstDetra.RecordCount = 0 Then
        MsgBox "No hay registros a eliminar", vbInformation, xTitulo
        Exit Sub
    End If
    
    Rpta = MsgBox("Esta seguro de eliminar la detraccion registrada", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM con_detraccion WHERE id =" & RstDetra("id") & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstDetra("id") & " AND idform = " & IdMenuActivo

        
        MsgBox "La detraccion se elimino cone exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstDetra.Requery
        Dg1.Refresh
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : FUNCION
'* Descripcion      : FUNCION QUE GRABA UN REGISOTRO EN LA TABLA con_detraccion, ESTA FUNCION DEVUELVE
'*                    VERDADERO SI TIENE EXITO
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Function Grabar() As Boolean
    ' VERIFICA QUE LOS DATOS NECESARIO HAYAN SIDO INGRESADOS CORRECTAMENTE

    If NulosC(TxtNumRuc.Text) = "" Then
        MsgBox "No ha especificado en Nº R.U.C.", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumRuc.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtNumDoc.Text) = "" Then
        MsgBox "No ha especificado el número de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDoc.SetFocus
        Exit Function
    End If
    
    If NulosN(TxtImpDet.Text) = 0 Then
        MsgBox "No ha especificado el importe retenido", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtImpDet.SetFocus
        Exit Function
    ElseIf NulosN(TxtImpDet.Text) > NulosN(TxtImpDocExpMN.Text) Then
        MsgBox "El importe retenido supera al importe del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtImpDet.Text = ""
        TxtImpDet.SetFocus
        Exit Function
    End If
        
        
    If NulosC(TxtNumDet.Text) = "" Or NulosC(TxtNumDet.Text) = "SIN NUMERO" Then
        MsgBox "No ha especificado el Nº del comprobante de detracción", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDet.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtFchPag.Valor) = "" Then
        MsgBox "No ha especificado la fecha de pago de la detracción", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchPag.SetFocus
        Exit Function
    End If
    
    Dim RstCab As New ADODB.Recordset
    Dim xId As Double
    
    On Error GoTo LaCague
    
    xCon.BeginTrans
    
    If QueHace = 1 Then
        ' SI SE ESTA AGREGANDO UN NUEVO REGISTRO, OBTENEMOS EL ULTIMO ID DE LA TABLA con_detraccion
        xId = HallaCodigoTabla("con_detraccion", xCon, "id")
        RST_Busq RstCab, "SELECT * FROM con_detraccion", xCon
        
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstDetra("id")
        ' SI SE ESTA MODIFICANDO OBTENEMOS EL ID DEL REGISTRO QUE SE ESTA EDITANDO
        RST_Busq RstCab, "SELECT * FROM con_detraccion WHERE id = " & xId & "", xCon
    End If
    
    mIdRegistro = xId
    
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
    
    ' grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    xCon.CommitTrans
    
    Set RstCab = Nothing
    MsgBox "La detracción se grabo con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Exit Function
LaCague:
    Grabar = False
    xCon.RollbackTrans
    Set RstCab = Nothing:
    MsgBox "No se pudo guardar la detracción por el siguiente motivo: " + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
End Function

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL INGRESO O MODIFICACION DE UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    Label5.Caption = "Detalle de Detracciones"
    ActivaTool
    Bloquea
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        FrmMasivo.Visible = False
        Nuevo
    End If

    If Button.Index = 2 Then Modificar

    If Button.Index = 3 Then Eliminar

    If Button.Index = 5 Then
        If FrmMasivo.Visible = False Then
        
            If Grabar = True Then
                Cancelar
                RstDetra.Requery
                Dg1.Refresh
                '-----------------------------------
                If RstDetra.RecordCount <> 0 Then
                    RstDetra.MoveFirst
                    RstDetra.Find "id=" & mIdRegistro
                    If RstDetra.EOF = True Then RstDetra.MoveFirst
                End If
                '-----------------------------------
            End If
        Else
            If GrabarGrupo = True Then
                Cancelar
                RstDetra.Requery
                Dg1.Refresh
                '-----------------------------------
                If RstDetra.RecordCount <> 0 Then
                    RstDetra.MoveFirst
                    RstDetra.Find "id=" & mIdRegistro
                    If RstDetra.EOF = True Then RstDetra.MoveFirst
                End If
                '-----------------------------------
            End If
        
        End If
    End If

    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then
        TDB_Actualizar Me, TabOne1, Dg1, RstDetra
    End If
    
    If Button.Index = 13 Then pExportar
    
    If Button.Index = 16 Then
        
        Unload Me
    End If
    
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.Index = 1 Then
        If ButtonMenu.Index = 1 Then '--Individual
            FrmMasivo.Visible = False
            Nuevo
        End If
        
        If ButtonMenu.Index = 2 Then '--Grupal
            '--posicionar el frame
            FrmMasivo.Visible = True
            FrmMasivo.Left = 90
            FrmMasivo.Top = 450
            Nuevo
            
        End If
    Else
        If ButtonMenu.Index = 1 Then
            Dim xFchIni, xFchFin As String
            
            xFchIni = "01/01/" + Trim(AnoTra)
            
            If Year(Date) = Trim(AnoTra) Then
                xFchFin = Date
            Else
                xFchFin = Format("31/12/" + Trim(AnoTra))
            End If
            
            FrmRegDetraccion.TxtFchIni.Valor = xFchIni
            FrmRegDetraccion.TxtFchFin.Valor = xFchFin
            
            FrmRegDetraccion.Show
            
            
        End If
    End If
End Sub

Private Sub TxtGlosa1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdDet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdDet_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = 116 Then
        CmdBusDetra_Click
    End If
End Sub

Private Sub TxtIdDet_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    Dim xImpBase As Double
    
    If NulosN(TxtIdDet.Text) <> 0 Then
        LblDetraccion.Caption = Busca_Codigo(NulosN(TxtIdDet.Text), "id", "descripcion", "mae_detraccion", "N", xCon)
        
        If LblDetraccion.Caption <> "" Then
            
            xImpBase = Busca_Codigo(NulosN(TxtIdDet.Text), "id", "impbase", "mae_detraccion", "N", xCon)
            
            If xImpBase <> 0 Then
                            
                If NulosN(TxtImpDocExpMN.Text) < xImpBase Then
                    If MsgBox("El importe expresado en moneda nacional es inferior al importe base para el cálculo de detracción" & vbCr & "Importe base: " & Format(xImpBase, FORMAT_MONTO) & vbCr & "Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo) = vbNo Then
                        TxtIdDet.SetFocus
                        GoTo Anular
                        Exit Sub
                    End If
                End If
            
            End If
            
            TxtTasa.Text = Format(NulosN(Busca_Codigo(NulosN(TxtIdDet.Text), "id", "tasa", "mae_detraccion", "N", xCon)), "0.00")
            TxtImpDet.Text = NulosN(TxtImpDocExpMN.Text) * ((NulosN(TxtTasa.Text) / 100))
            TxtImpDet.Text = Format(TxtImpDet.Text, FORMAT_MONTO)
            
        Else
            GoTo Anular
        End If
    Else
        
        GoTo Anular
    End If
Exit Sub
Anular:
    LblDetraccion.Caption = ""
    TxtIdDet.Text = ""
    TxtTasa.Text = ""
    TxtImpDet.Text = ""
    TxtIdDet.SetFocus
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

Private Sub TxtNumDet1_KeyPress(KeyAscii As Integer)
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
    If QueHace = 3 Then Exit Sub
    
    If TxtNumRuc.Text <> "" Then
        If xTIPO_MOVIMIETO = 1 Then
            LblNomPro.Caption = Busca_Codigo(NulosC(TxtNumRuc.Text), "numruc", "nombre", "mae_prov", "C", xCon)
        Else
            LblNomPro.Caption = Busca_Codigo(NulosC(TxtNumRuc.Text), "numruc", "nombre", "mae_cliente", "C", xCon)
        End If
        If LblNomPro.Caption = "" Then
            TxtNumRuc.Text = ""
            LblNomPro.Caption = ""
            LblIdProveedor.Caption = ""
            TxtNumDoc.Text = ""
            TxtIdMon.Text = ""
            TxtFchDoc.Valor = ""
            TxtImpDoc.Text = ""
        Else
            If xTIPO_MOVIMIETO = 1 Then
                LblIdProveedor.Caption = NulosN(Busca_Codigo(NulosC(TxtNumRuc.Text), "numruc", "id", "mae_prov", "C", xCon))
            Else
                LblIdProveedor.Caption = NulosN(Busca_Codigo(NulosC(TxtNumRuc.Text), "numruc", "id", "mae_cliente", "C", xCon))
            End If
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

Private Sub pExportar()
    
    TabOne1.CurrTab = 0

    Dim nSQL As String
    Dim oExport As New SGI2_funciones.formularios
    Dim RstTmp  As New ADODB.Recordset

    Dim xCampos(19, 3) As String
    
    '0::Nombre a Mostrar;
    '1::nombre de Campo del Rst;
    '2::alineacion(0::derecha, 1::centro, 2::izquierda);
    '3::ancho de columna
    '--obs: el rst puede tener mas columnas solo se consideran los campos del array
    xCampos(0, 0) = "IdDoc":        xCampos(0, 1) = "iddoc":        xCampos(0, 2) = 2:   xCampos(0, 3) = "500"
    xCampos(1, 0) = "Nº Reg":       xCampos(1, 1) = "registro":     xCampos(1, 2) = 0:   xCampos(1, 3) = "900"
    xCampos(2, 0) = "R.U.C.":       xCampos(2, 1) = "numruc":       xCampos(2, 2) = 0:   xCampos(2, 3) = "1200"
    xCampos(3, 0) = "Proveedor":    xCampos(3, 1) = "nombre":       xCampos(3, 2) = 0:   xCampos(3, 3) = "3000"
    xCampos(4, 0) = "T.D.":         xCampos(4, 1) = "abrev":        xCampos(4, 2) = 0:   xCampos(4, 3) = "350"
    xCampos(5, 0) = "Num. Doc":     xCampos(5, 1) = "numerodoc":       xCampos(5, 2) = 0:   xCampos(5, 3) = "1600"
    xCampos(6, 0) = "Fch.Doc.":     xCampos(6, 1) = "fchdoc":       xCampos(6, 2) = 1:   xCampos(6, 3) = "1000"
    xCampos(7, 0) = "M":            xCampos(7, 1) = "abremon":      xCampos(7, 2) = 1:   xCampos(7, 3) = "500"
    xCampos(8, 0) = "T.C.":         xCampos(8, 1) = "tipcam":       xCampos(8, 2) = 2:   xCampos(8, 3) = "700"
    
    xCampos(9, 0) = "Imp Base":     xCampos(9, 1) = "impbase":      xCampos(9, 2) = 2:   xCampos(9, 3) = "1000"
    xCampos(10, 0) = "Imp Total":   xCampos(10, 1) = "imptot":      xCampos(10, 2) = 2:  xCampos(10, 3) = "1000"
    
    xCampos(11, 0) = "Imp Base. MN":     xCampos(11, 1) = "impbaseexmn":    xCampos(11, 2) = 2:   xCampos(11, 3) = "1200"
    xCampos(12, 0) = "Imp Exp. MN":      xCampos(12, 1) = "imptotexmn":        xCampos(12, 2) = 2:   xCampos(12, 3) = "1200"
    
    xCampos(13, 0) = "Id":               xCampos(13, 1) = "id":         xCampos(13, 2) = 2:   xCampos(13, 3) = "500"
    xCampos(14, 0) = "Num. Comprobante": xCampos(14, 1) = "numdet":     xCampos(14, 2) = 0:   xCampos(14, 3) = "1700"
    xCampos(15, 0) = "Tipo Detracción":  xCampos(15, 1) = "descdetra":  xCampos(15, 2) = 0:   xCampos(15, 3) = "3300"
    xCampos(16, 0) = "% Detrac":         xCampos(16, 1) = "por":        xCampos(16, 2) = 2:   xCampos(16, 3) = "900"
    xCampos(17, 0) = "Fch.Pag.":         xCampos(17, 1) = "fchpag":     xCampos(17, 2) = 1:   xCampos(17, 3) = "1000"
    xCampos(18, 0) = "Imp. Detra":       xCampos(18, 1) = "imp":        xCampos(18, 2) = 2:   xCampos(18, 3) = "900"
    xCampos(19, 0) = "Glosa":            xCampos(19, 1) = "glosa":      xCampos(19, 2) = 0:   xCampos(19, 3) = "2000"
    
    '--cambiar el nombre de celda para diferenciar al cliente
    If xTIPO_MOVIMIETO = 2 Then xCampos(3, 0) = "Cliente"
  
'    Set RstTmp = RstDetra.Clone
    
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "LISTADO DE DETRACCIONES - " & UCase(Label2), "Periodo  : Todos", "", "Listado de Detracciones", RstDetra, xCampos()
    RstDetra.MoveFirst
    
    Set oExport = Nothing
    Set RstTmp = Nothing
    
End Sub



'----------------------------
Private Sub CmdBusArch_Click()
    Err.Clear
    'CommonDialog1.CancelError = True
    'Especificar las extensiones a usar
    CommonDialog1.DefaultExt = "*.xls"
    CommonDialog1.FileName = ""
    'CommonDialog1.Filter = "Cardfile (*.crd)|*.crd|Textos (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
    CommonDialog1.Filter = "Documentos de Excel (*.xls)|*.xls"
    CommonDialog1.ShowOpen
    If Err Then
        'Cancelada la operación de abrir
        MsgBox Err.Description & vbCr & Err.Source & vbCr & CommonDialog1.FileName, vbInformation, xTitulo
        Err.Clear
        
    Else
        TxtArchivo.Text = CommonDialog1.FileName
        TxtFchPag1.SetFocus
    End If
    Err.Clear
End Sub

Private Sub CmdCargar_Click()
    If QueHace = 1 Then
        If TxtArchivo.Text = "" Then
            MsgBox "No ha especificado el nombre del archivo cabecera ", vbInformation, xTitulo
            TxtArchivo.SetFocus
            Exit Sub
        End If

        CargaDocumentos
    End If

End Sub

Sub CargaDocumentos()
''    '--x implementar
''
''    Dim xNumFilas As Integer
''    Dim A&
''    Dim B As Integer
''    Dim xFilas As Long
''    Dim xFilaIni As Long
''    '-------
''
''    Dim nSQL As String
''
''    On Error GoTo error:
''
''    '---------------------------------------------------------------------
''
''    Dim objExcel As Object
''    Set objExcel = CreateObject("Excel.Application")
''    'Dim objExcel As New Excel.Application
''
''    objExcel.Visible = True
''    objExcel.SheetsInNewWorkbook = 1
''
''    'abre el Libro
''    objExcel.WindowState = 2
''    objExcel.Workbooks.Open Trim(TxtArchivo.Text)
''
''    Frame5.Left = 3090
''    Frame5.Top = 2910
''    Frame5.Visible = True
''
''    '--indica el inicio de lectura de registros
''    xFilaIni = 9
''
''    xNumFilas = 1
''
''    fg1.Rows = 1
''
''    With objExcel.ActiveSheet
''        'DETERMINAMOS EL NUMERO DE FILAS CON DATOS
''        LblBarra.Caption = "Calculando número de registros"
''        DoEvents
''        ProgressBar2.Max = 32000
''        For A = xFilaIni To 32000
''            ProgressBar2.Value = A
''            '--verificar idsolcheque, proveedor
''            If NulosC(.Cells(A, 1)) <> "" Or NulosC(.Cells(A, 8)) <> "" Then
''                xNumFilas = xNumFilas + 1
''            Else
''                Exit For
''            End If
''        Next A
''
''        xNumFilas = xNumFilas + xFilaIni
''        LblBarra.Caption = "Cargando registros - Cabecera de Documentos"
''        DoEvents
''        ProgressBar2.Max = xNumFilas
''
''        For A = xFilaIni To xNumFilas
''            ProgressBar2.Value = A
''
''            DoEvents
''            '--verificar si proveedor es nulo para cancelar
''            If NulosC(.Cells(A, 1)) = "" And NulosC(.Cells(A, 8)) = "" Then Exit For
''
''            fg1.Rows = fg1.Rows + 1
''
''            xFilas = fg1.Rows - 1
''
''            fg1.TextMatrix(xFilas, 1) = NulosC(.Cells(A, 21)) '--Año registro segun cliente
''
''            fg1.TextMatrix(xFilas, 2) = NulosC(.Cells(A, 20)) '--Periodo registro segun cliente
''            fg1.TextMatrix(xFilas, 3) = Format(NulosC(.Cells(A, 2)), "0000") '--Correlativo
''            fg1.TextMatrix(xFilas, 4) = NulosC(.Cells(A, 8)) '--Ruc proveedor
''            fg1.TextMatrix(xFilas, 5) = NulosC(.Cells(A, 9)) '--Razon social del proveedor
''            fg1.TextMatrix(xFilas, 6) = NulosC(.Cells(A, 4)) '--Tipo de Documento
''            fg1.TextMatrix(xFilas, 7) = NulosC(.Cells(A, 5)) '--N° Serie
''            fg1.TextMatrix(xFilas, 8) = NulosC(.Cells(A, 6)) '--N° Documento
''
''            If IsDate(CDate(.Cells(A, 22))) = True Then fg1.TextMatrix(xFilas, 9) = Format(CDate(.Cells(A, 22)), FORMAT_DATE) '--Fecha Documento
''            If IsDate(CDate(.Cells(A, 23))) = True Then fg1.TextMatrix(xFilas, 10) = Format(CDate(.Cells(A, 23)), FORMAT_DATE)  '--Fecha Recepción
''            If IsDate(CDate(.Cells(A, 24))) = True Then fg1.TextMatrix(xFilas, 11) = Format(CDate(.Cells(A, 24)), FORMAT_DATE)  '--Fecha Vencimiento
''            If IsDate(CDate(.Cells(A, 25))) = True Then fg1.TextMatrix(xFilas, 12) = Format(CDate(.Cells(A, 25)), FORMAT_DATE)  '--Fecha Sistema
''
''            fg1.TextMatrix(xFilas, 13) = NulosC(.Cells(A, 17)) '--Moneda
''            fg1.TextMatrix(xFilas, 14) = NulosN(.Cells(A, 18)) '--Tipo de Cambio
''
''            '--Si documento es nota de credito, mostrar los importes en positivo
''            fg1.TextMatrix(xFilas, 15) = Abs(NulosN(.Cells(A, 11))) '--Imp. Afecto
''            fg1.TextMatrix(xFilas, 16) = Abs(NulosN(.Cells(A, 12))) '--Imp Inafecto
''            fg1.TextMatrix(xFilas, 17) = Abs(NulosN(.Cells(A, 13))) '--Imp. Retencion(para honorarios)
''            fg1.TextMatrix(xFilas, 18) = Abs(NulosN(.Cells(A, 14))) '--Imp Igv
''            fg1.TextMatrix(xFilas, 19) = Abs(NulosN(.Cells(A, 10))) '--Imp. Total
''
''            fg1.TextMatrix(xFilas, 20) = NulosC(.Cells(A, 19)) '--Glosa
''            fg1.TextMatrix(xFilas, 21) = NulosC(.Cells(A, 35)) '--Ruc cliente
''            fg1.TextMatrix(xFilas, 22) = NulosC(.Cells(A, 36)) '--Razón Social del Cliente
''            fg1.TextMatrix(xFilas, 23) = NulosC(.Cells(A, 33)) '--Orden de Despacho
''
''
''            If NulosC(.Cells(A, 8)) = "" Then
''                GRID_COLOR_FONDO fg1, xFilas, 0, xFilas, fg1.Cols - 1, vbRed
''            End If
''
''
''        Next A
''    End With
''
''    DoEvents
''
''
''    Frame5.Visible = False
''    MsgBox "El proceso terminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
''    objExcel.WindowState = 2
''    objExcel.Workbooks.Close
''
''    Set objExcel = Nothing
''    Exit Sub
''error:
'''Resume
''    Frame5.Visible = False
''    objExcel.Workbooks.Close
''    If Err.Number = 424 Then
''        MsgBox Err.Description & vbCr & "El archivo fue cerrado antes de terminar de importar, vuelva a importar nuevamente.", vbCritical, xTitulo
''    Else
''        MsgBox Err.Description & vbCr & Err.Source, vbCritical, xTitulo
''    End If
''    fg1.Rows = 1
''    Set objExcel = Nothing

End Sub

'*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    CargarDatosEnDetalle Row, Col
                    
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If Fg1.Row < 0 Then Exit Sub
    
    If Col = 10 Then
        If IsNumeric(Fg1.TextMatrix(Row, 10)) = False Or NulosN(Fg1.TextMatrix(Row, 10)) = 0 Then
            Fg1.TextMatrix(Row, 10) = 0
            Fg1.TextMatrix(Row, 11) = ""
            Fg1.TextMatrix(Row, 12) = 0
            Fg1.TextMatrix(Row, 13) = 0
        Else
            Dim xRs As New ADODB.Recordset
            RST_Busq xRs, "Select * From mae_detraccion where id = " & NulosN(Fg1.TextMatrix(Row, 10)), xCon
            If xRs.State = 0 Then
                Fg1.TextMatrix(Row, 11) = ""
                Fg1.TextMatrix(Row, 12) = 0
                Fg1.TextMatrix(Row, 13) = 0
                Set xRs = Nothing
                Exit Sub
            End If
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Row, 11) = NulosC(xRs("descripcion"))
                Fg1.TextMatrix(Row, 12) = NulosN(xRs("tasa"))
                Fg1_CellChanged Row, 12
            Else
                Fg1.TextMatrix(Row, 10) = 0
                Fg1.TextMatrix(Row, 11) = ""
                Fg1.TextMatrix(Row, 12) = 0
                Fg1.TextMatrix(Row, 13) = 0
            End If
            Set xRs = Nothing
            
        End If
    ElseIf Col = 12 Then
        '--verificar si % es nulo
        Fg1.TextMatrix(Row, 12) = NulosN(Fg1.TextMatrix(Row, 12))
        If IsNumeric(Fg1.TextMatrix(Row, 12)) = False Then
            Fg1.TextMatrix(Row, 12) = 0
            Fg1.TextMatrix(Row, 13) = 0
        Else
            If NulosN(Fg1.TextMatrix(Row, 12)) > 100 Then
                MsgBox "Ingrese un % correcto", vbInformation, xTitulo
                Fg1.TextMatrix(Row, 12) = 0
                Fg1.TextMatrix(Row, 13) = 0
            ElseIf InStr(Fg1.TextMatrix(Row, 12), ".") <> 0 Then
                MsgBox "Ingrese un % entero", vbInformation, xTitulo
                Fg1.TextMatrix(Row, 12) = 0
                Fg1.TextMatrix(Row, 13) = 0
            End If
            
            Fg1.TextMatrix(Row, 13) = Format((NulosN(Fg1.TextMatrix(Row, 9)) * NulosN(Fg1.TextMatrix(Row, 12))) / 100, FORMAT_MONTO)
        
        End If
    ElseIf Col = 13 Then
        Fg1.TextMatrix(Row, 13) = Format(NulosN(Fg1.TextMatrix(Row, 13)), FORMAT_MONTO)
    End If
    
    HallarTotal
    
End Sub


Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
        Fg1.SelectionMode = flexSelectionByRow
        Exit Sub
    Else
        Fg1.SelectionMode = flexSelectionFree
    End If
    
    If Agregando = True Then Exit Sub
    
    If Fg1.Col = 1 Or Fg1.Col = 2 Or Fg1.Col = 11 Then
        Fg1.Editable = flexEDKbdMouse
        Fg1.ColComboList(Fg1.Col) = "|..."
        
    ElseIf Fg1.Col = 10 Or Fg1.Col = 12 Or Fg1.Col = 13 Then
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
        
    End If

End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If QueHace = 3 Then KeyAscii = 0
    If KeyAscii = 13 Then Exit Sub
    
    '--si es seleccionado no hacer nada
    If Col = 10 Then
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
    
End Sub

Private Sub fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    
    If KeyCode = vbKeyF10 Then
        CargarDatosEnDetalle Fg1.Row, Fg1.Col
        Fg1.SetFocus
    End If
    If KeyCode = 45 Then CmdAddItem_Click
    
End Sub


Private Sub CargarDatosEnDetalle(xFil As Long, xCol As Long)
    '===================================================================================================
    'Creado : 12/12/11 Por: Johan Castro
    'Propósito: Mostrar ventana de seleccion
    '
    'Entradas:  xfil= Posicion de la Fila
    '           xCol= Posicion de la Columna
    '
    '===================================================================================================

    Dim xRs As New ADODB.Recordset
    ReDim xCampos(2, 4) As String
    Dim nSQL As String
    Dim nTitulo As String

    'Descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    If xCol = 1 Then
        
        xCampos(0, 0) = "Cliente":      xCampos(0, 1) = "nombre":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Nº R.U.C.":    xCampos(1, 1) = "numruc":      xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
        
        If xTIPO_MOVIMIETO = 1 Then
            nSQL = "SELECT mae_prov.nombre, mae_prov.numruc, mae_prov.id From mae_prov WHERE mae_prov.id <>0 "
            nTitulo = "Buscando Proveedor"
            
            xCampos(0, 0) = "Proveedor":
        
        Else
            nSQL = "SELECT mae_cliente.nombre, mae_cliente.numruc, mae_cliente.id From mae_cliente where mae_cliente.id <>0"
            nTitulo = "Buscando Cliente"
        
        End If
        
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio, ""
        
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 1) = NulosC(xRs("nombre"))
                Fg1.TextMatrix(Fg1.Row, 15) = NulosN(xRs("id"))
            End If
        End If
                        
    
    ElseIf xCol = 2 Then
        
        CargarDocumentos (NulosN(Fg1.TextMatrix(Me.Fg1.Row, 15)))
            
    ElseIf xCol = 11 Then
        ReDim xCampos(3, 4) As String
        xCampos(0, 0) = "Detracción":   xCampos(0, 1) = "nombre":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Tasa":         xCampos(1, 1) = "tasa1":       xCampos(1, 2) = "800":          xCampos(1, 3) = "C"
        xCampos(2, 0) = "Cód.":         xCampos(2, 1) = "id":          xCampos(2, 2) = "500":          xCampos(2, 3) = "N"
        
        nSQL = "SELECT mae_detraccion.descripcion as nombre, mae_detraccion.tasa & ' %' as tasa1, mae_detraccion.tasa, mae_detraccion.id FROM mae_detraccion"
        nTitulo = "Buscando Detracciones"
        
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio, ""
        
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 10) = NulosN(xRs("id"))
                Fg1.TextMatrix(Fg1.Row, 11) = NulosC(xRs("nombre"))
                Fg1.TextMatrix(Fg1.Row, 12) = NulosC(xRs("tasa"))
                
            End If
        End If
    End If

    Set xRs = Nothing
End Sub


Private Sub pGridConfigurar()
    Fg1.ColWidth(9) = 800 '
    Fg1.ColWidth(10) = 450 '
    Fg1.ColWidth(13) = 800 '
    
    Fg1.ColWidth(14) = 0 'iddoc
    Fg1.ColWidth(15) = 0 'idclipro
    Fg1.ColWidth(16) = 0 'iddoc
    Fg1.ColWidth(17) = 0 'idmon
    Fg1.ColWidth(18) = 0 'idcab(tabla con_detraccion)
    
    Fg1.RowHeight(0) = 500
    Fg1.WordWrap = True
    Fg1.SelectionMode = flexSelectionByRow
    
End Sub


Sub CargarDocumentos(idclipro As Long)
    
    Dim xCampos(9, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLIdPer As String
    Dim nSQLDoc As String
    Dim X As Integer
    Dim nSQLDocDetra As String
    
    '1 Compras
    '2 Ventas
    
    xCampos(0, 0) = "Cliente":    xCampos(0, 1) = "nombre":       xCampos(0, 2) = "2500":   xCampos(0, 3) = "C":     xCampos(0, 4) = "N"
    xCampos(1, 0) = "TD":         xCampos(1, 1) = "abrev":        xCampos(1, 2) = "600":    xCampos(1, 3) = "C":     xCampos(2, 4) = "N"
    xCampos(2, 0) = "Nro Doc":    xCampos(2, 1) = "numerodoc":    xCampos(2, 2) = "1500":   xCampos(2, 3) = "C":     xCampos(3, 4) = "S"
    xCampos(3, 0) = "Fch. Doc":   xCampos(3, 1) = "fchdoc1":      xCampos(3, 2) = "1000":   xCampos(3, 3) = "C":     xCampos(1, 4) = "N"
    xCampos(4, 0) = "M":          xCampos(4, 1) = "simbolo":      xCampos(4, 2) = "500":    xCampos(4, 3) = "C":     xCampos(4, 4) = "N"
    xCampos(5, 0) = "T.C.":       xCampos(5, 1) = "tipcam":       xCampos(5, 2) = "600":    xCampos(5, 3) = "N":     xCampos(5, 4) = "N"
    xCampos(6, 0) = "Importe":    xCampos(6, 1) = "imptot":       xCampos(6, 2) = "1200":   xCampos(6, 3) = "N":     xCampos(6, 4) = "N"
    xCampos(7, 0) = "Nro Detra":  xCampos(7, 1) = "numdet":       xCampos(7, 2) = "1100":   xCampos(7, 3) = "C":     xCampos(7, 4) = "N"
    xCampos(8, 0) = "Fch. Detra": xCampos(8, 1) = "fchmov1":      xCampos(8, 2) = "1000":    xCampos(8, 3) = "C":     xCampos(8, 4) = "N"
    
    If ChkDoc.Value = 1 Then nSQLDocDetra = " or (detra.idgr is not null) "
    
    If xTIPO_MOVIMIETO = 1 Then 'Compras
        
        nSQLDoc = GRID_GENERAR_SQL_ID(Fg1, 14, " AND com_compras.id", "NOT IN", True)
                
        If idclipro <> 0 Then nSQLIdPer = " and  com_compras.idpro = " & idclipro & " "
        
        nSQL = "SELECT 0 AS xsel, Left(com_compras.numreg,2) & mae_libros.codsun & Right(com_compras.numreg,4) AS registro, mae_prov.numruc, mae_prov.nombre, mae_documento.abrev, IIf(com_compras.numser='',com_compras.numdoc,com_compras.numser+'-'+com_compras.numdoc) AS numerodoc, " _
            + vbCr + " com_compras.fchdoc & '' as fchdoc1, IIf(com_compras.tc=0,con_tc.impven,com_compras.tc) & '' AS tipcam, mae_moneda.simbolo, com_compras.imptot, IIf(com_compras.idmon=1,com_compras.imptot,com_compras.imptot*tipcam) AS impexmn, com_compras.glosa, " _
            + vbCr + " com_compras.id as iddoc, com_compras.tipdoc, com_compras.idmon, com_compras.idpro AS idclipro, " _
            + vbCr + " detra.* " _
            + vbCr + " FROM ( mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro " _
            + vbCr + "      ) LEFT JOIN" _
            + vbCr + " (SELECT con_detraccion.id AS idcab, con_detraccion.iddoc as iddoc1, con_detraccion.iddet, con_detraccion.idgr, con_detraccion.fchmov & '' as fchmov1, con_detraccion.numdet, mae_detraccion.descripcion AS detraccion, con_detraccion.por, con_detraccion.[imp], con_detraccion.glosa AS xglosa " _
            + vbCr + "    FROM mae_detraccion RIGHT JOIN con_detraccion ON mae_detraccion.id = con_detraccion.iddet " _
            + vbCr + "    WHERE con_detraccion.tipo=1 " _
            + vbCr + " ) as detra ON com_compras.id = detra.iddoc1 " _
            + vbCr + " Where ( (detra.idgr =0 or detra.idgr is null) " & nSQLDocDetra & " ) " & nSQLIdPer & nSQLDoc _
            + vbCr + " ORDER BY mae_prov.nombre, IIf(com_compras.numser='',com_compras.numdoc,com_compras.numser+'-'+com_compras.numdoc) "

            xCampos(0, 0) = "Proveedor":
        
    Else        'Ventas
    
        nSQLDoc = GRID_GENERAR_SQL_ID(Fg1, 14, " AND vta_ventas.id", "NOT IN", True)
        
        If idclipro <> 0 Then nSQLIdPer = " and vta_ventas.idcli = " & idclipro & " "
        
        nSQL = "SELECT 0 AS xsel, Left(vta_ventas.numreg,2) & mae_libros.codsun & Right(vta_ventas.numreg,4) AS registro, mae_cliente.numruc, mae_cliente.nombre, mae_documento.abrev, vta_ventas.numser+'-'+vta_ventas.numdoc AS numerodoc, " _
            + vbCr + " vta_ventas.fchdoc & '' as fchdoc1,mae_moneda.simbolo, IIf(vta_ventas.tc=0,con_tc.impven,vta_ventas.tc) & '' AS tipcam, vta_ventas.imptotdoc AS imptot, IIf(vta_ventas.idmon=1,vta_ventas.imptotdoc,vta_ventas.imptotdoc*tipcam) AS impexmn, " _
            + vbCr + " vta_ventas.id as iddoc, vta_ventas.tipdoc, vta_ventas.idmon, vta_ventas.idcli AS idclipro, vta_ventas.glosa, " _
            + vbCr + " detra.* " _
            + vbCr + " FROM (mae_libros RIGHT JOIN (mae_documento RIGHT JOIN (((mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) ON mae_documento.id = vta_ventas.tipdoc) ON mae_libros.id = vta_ventas.idlib " _
            + vbCr + "      ) LEFT JOIN" _
            + vbCr + " (SELECT con_detraccion.id AS idcab, con_detraccion.iddoc as iddoc1, con_detraccion.iddet, con_detraccion.idgr, con_detraccion.fchmov & '' as fchmov1, con_detraccion.numdet, mae_detraccion.descripcion AS detraccion, con_detraccion.por, con_detraccion.[imp], con_detraccion.glosa AS xglosa " _
            + vbCr + "    FROM mae_detraccion RIGHT JOIN con_detraccion ON mae_detraccion.id = con_detraccion.iddet " _
            + vbCr + "    WHERE con_detraccion.tipo=2 " _
            + vbCr + " ) as detra ON vta_ventas.id = detra.iddoc1 " _
            + vbCr + " Where vta_ventas.anulado=0 and ( (detra.idgr =0 or detra.idgr is null) " & nSQLDocDetra & " )" & nSQLIdPer & nSQLDoc _
            + vbCr + " ORDER BY mae_cliente.nombre, vta_ventas.numser+'-'+vta_ventas.numdoc DESC "
    
    End If
    '--------------------------------------
    
    CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), "Buscando Documentos"
    
    If xRs.State = 1 Then
        If xRs.RecordCount = 0 Then
            Set xRs = Nothing
            Exit Sub
        End If
        
        Agregando = True
        
        xRs.MoveFirst
        Do While Not xRs.EOF
            With Me.Fg1
                                           
                If X >= 1 Then
                    .AddItem ""
                End If
                .Row = .Rows - 1
                
                .TextMatrix(.Row, 1) = NulosC(xRs("nombre"))
                .TextMatrix(.Row, 2) = NulosC(xRs("registro"))
                .TextMatrix(.Row, 3) = NulosC(xRs("abrev"))
                .TextMatrix(.Row, 4) = NulosC(xRs("numerodoc"))
                .TextMatrix(.Row, 5) = Format(NulosC(xRs("fchdoc1")), FORMAT_DATE)
                .TextMatrix(.Row, 6) = NulosC(xRs("simbolo"))
                .TextMatrix(.Row, 7) = NulosN(xRs("tipcam"))
                .TextMatrix(.Row, 8) = Format(NulosN(xRs("imptot")), FORMAT_MONTO)
                .TextMatrix(.Row, 9) = Format(NulosN(xRs("impexmn")), FORMAT_MONTO)
                
                .TextMatrix(.Row, 10) = NulosN(xRs("iddet"))
                .TextMatrix(.Row, 11) = NulosC(xRs("detraccion"))
                
                .TextMatrix(.Row, 12) = NulosC(xRs("por"))
                .TextMatrix(.Row, 13) = Format(NulosN(xRs("imp")), FORMAT_MONTO)
                
                .TextMatrix(.Row, 14) = NulosN(xRs("iddoc"))
                .TextMatrix(.Row, 15) = NulosN(xRs("idclipro"))
                .TextMatrix(.Row, 16) = NulosN(xRs("tipdoc"))
                .TextMatrix(.Row, 17) = NulosN(xRs("idmon"))
                If ChkDoc.Value = 1 Then
                    .TextMatrix(.Row, 18) = 0
                Else
                    .TextMatrix(.Row, 18) = NulosN(xRs("idcab"))
                End If
                
                X = X + 1
                
            End With
            
            xRs.MoveNext
        Loop
        Agregando = False
    End If
    
    Set xRs = Nothing
        
    HallarTotal
    
End Sub

Private Sub HallarTotal()
    Dim xFila As Long
    Dim xCol As Long
    Dim A As Long
    
    xFila = Fg1.Row
    xCol = Fg1.Col
    
    If xFila < 1 Then xFila = 1
    If xCol < 1 Then xCol = 1
    
    TxtTotMN.Text = "0.00"
    TxtTotME.Text = "0.00"
    TxtTotExMN.Text = "0.00"
    TxtTotDet.Text = "0.00"
    
    If Fg1.Rows = Fg1.FixedRows Then Exit Sub
    
    For A = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(A, 17)) = 1 Then
            TxtTotMN.Text = NulosN(TxtTotMN.Text) + NulosN(Fg1.TextMatrix(A, 8))
        Else
            TxtTotME.Text = NulosN(TxtTotME.Text) + NulosN(Fg1.TextMatrix(A, 8))
        End If
        TxtTotExMN.Text = NulosN(TxtTotExMN.Text) + NulosN(Fg1.TextMatrix(A, 9))
        TxtTotDet.Text = NulosN(TxtTotDet.Text) + NulosN(Fg1.TextMatrix(A, 13))
        
    Next A

    TxtTotMN.Text = Format(NulosN(TxtTotMN.Text), FORMAT_MONTO)
    TxtTotME.Text = Format(NulosN(TxtTotME.Text), FORMAT_MONTO)
    TxtTotExMN.Text = Format(NulosN(TxtTotExMN.Text), FORMAT_MONTO)
    TxtTotDet.Text = Format(NulosN(TxtTotDet.Text), FORMAT_MONTO)

    Fg1.Row = xFila
    Fg1.Col = xCol
    
End Sub

Function GrabarGrupo() As Boolean
             
    If NulosC(TxtNumDet1.Text) = "" Or NulosC(TxtNumDet1.Text) = "SIN NUMERO" Then
        MsgBox "No ha especificado el Nº del comprobante de detracción", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDet.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtFchPag1.Valor) = "" Then
        MsgBox "No ha especificado la fecha de pago de la detracción", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchPag1.SetFocus
        Exit Function
    End If
    
    Dim RstCab As New ADODB.Recordset
    Dim xId As Double
    Dim xIdGr As Double
    Dim xFila  As Long
    Dim xQuehace As Integer
    On Error GoTo LaCague
    
    xCon.BeginTrans
    
    If QueHace = 1 Then
        ' SI SE ESTA AGREGANDO UN NUEVO REGISTRO, OBTENEMOS EL ULTIMO ID DE LA TABLA con_detraccion
        xIdGr = HallaCodigoTabla("con_detraccion", xCon, "idgr")
        
        RST_Busq RstCab, "SELECT TOP 1 * FROM con_detraccion", xCon
                
    Else
        xId = RstDetra("id")
        xIdGr = RstDetra("idgr")
       
    End If
    
    For xFila = 1 To Fg1.Rows - 1
        '--obtener codigo de detraccion si hubiera
        xId = NulosN(Fg1.TextMatrix(xFila, 18))
        
        If xId = 0 Then
            '--si detraccion no hubiera se procede a crear
            xId = HallaCodigoTabla("con_detraccion", xCon, "id")
            xQuehace = 1
            RstCab.AddNew
        Else
            Set RstCab = Nothing
            RST_Busq RstCab, "SELECT * FROM con_detraccion WHERE id = " & xId & "", xCon
            xQuehace = 2
            If RstCab.RecordCount = 0 Then
                RstCab.AddNew
                xQuehace = 1
            End If
        End If
        mIdRegistro = xId
        '----------------------
        
        RstCab("id") = xId
        RstCab("iddet") = NulosN(Fg1.TextMatrix(xFila, 10))
        RstCab("por") = NulosN(Fg1.TextMatrix(xFila, 12))
        RstCab("iddoc") = NulosN(Fg1.TextMatrix(xFila, 14))
        RstCab("idmon") = NulosN(Fg1.TextMatrix(xFila, 17))
        RstCab("tipo") = xTIPO_MOVIMIETO
        If IsDate(NulosC(Fg1.TextMatrix(xFila, 5))) = True Then
            RstCab("fchmov") = NulosC(Fg1.TextMatrix(xFila, 5))
        Else
            RstCab("fchmov") = Null
        End If
        RstCab("imp") = NulosN(Fg1.TextMatrix(xFila, 13))
        RstCab("numdet") = NulosC(TxtNumDet1.Text)
        RstCab("fchpag") = CDate(TxtFchPag1.Valor)
        RstCab("glosa") = NulosC(TxtGlosa1.Text)
        RstCab("idgr") = xIdGr
        RstCab.Update
        
        '----------------------
        ' grabamos el movimiento en la tabla var_edicion
        GrabarOperacion xIdUsuario, IdMenuActivo, xQuehace, xHorIni, Time, Date, xCon, xId
    
    Next xFila
    
    xCon.CommitTrans
    
    Set RstCab = Nothing
    MsgBox "La detracción se grabo con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    GrabarGrupo = True
    Exit Function
LaCague:
''    Resume
    GrabarGrupo = False
    xCon.RollbackTrans
    Set RstCab = Nothing:
    MsgBox "No se pudo guardar la detracción por el siguiente motivo: " + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
End Function

Private Sub CmdAddItem_Click()
    
    If QueHace = 3 Then Exit Sub
    If NulosC(Fg1.TextMatrix(Fg1.Rows - 1, 1)) = "" Then
        Fg1.Col = 1
        Fg1.Row = Fg1.Rows - 1
        Fg1_CellButtonClick Fg1.Rows - 1, 1
        Fg1.SetFocus
        Exit Sub
    End If
    
    Fg1.Rows = Fg1.Rows + 1
    
    With Fg1
        .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, 1
    End With
  
    Fg1.SetFocus
End Sub


Private Sub CmdDelItem_Click()
    If QueHace = 3 Then Exit Sub
    If Fg1.Row < 1 Or Fg1.Rows < 1 Then Exit Sub
    
    Fg1.RemoveItem Fg1.Row
    HallarTotal
    If Fg1.Rows <> 1 Then Fg1.Select Fg1.Rows - 1, 1
End Sub

Private Sub CmdDelItemTodos_Click()
    If QueHace = 3 Then Exit Sub
    If Fg1.Row < 1 Or Fg1.Rows < 1 Then Exit Sub
    Fg1.Rows = Fg1.FixedRows
    HallarTotal
    CmdAddItem.SetFocus
End Sub

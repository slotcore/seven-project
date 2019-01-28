VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmRetencion 
   Caption         =   "Contabilidad - Retenciones"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7590
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
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRetencion.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRetencion.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRetencion.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRetencion.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRetencion.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRetencion.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRetencion.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRetencion.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRetencion.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRetencion.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRetencion.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRetencion.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRetencion.frx":277E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7215
      Left            =   0
      TabIndex        =   12
      Top             =   375
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
         TabIndex        =   15
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6420
            Left            =   30
            TabIndex        =   16
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
            Columns(1).Caption=   "Nº Reg."
            Columns(1).DataField=   "registro"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nº R.U.C."
            Columns(2).DataField=   "numruc"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Cliente / Proveedor"
            Columns(3).DataField=   "nombre"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "T.D."
            Columns(4).DataField=   "docabrev"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Fch. Emi."
            Columns(5).DataField=   "fchemi"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Nº Documento"
            Columns(6).DataField=   "numedoc"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "M"
            Columns(7).DataField=   "simbolo"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Importe Ret."
            Columns(8).DataField=   "imp1"
            Columns(8).NumberFormat=   "0.00"
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
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1588"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1508"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=2619"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2540"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=6720"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=6641"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=512"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=820"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=741"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=1852"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1773"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=2752"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=2672"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=900"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=820"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=513"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(50)=   "Column(8).Width=2011"
            Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=1931"
            Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=514"
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=66,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=50,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=43,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=44,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=45,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=51,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=52,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=53,.parent=17"
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
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   300
            Left            =   8835
            TabIndex        =   46
            Top             =   30
            Width           =   1275
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Registro de Retenciones"
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
            TabIndex        =   17
            Top             =   30
            Width           =   11610
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Detalle de la Cuenta"
         Height          =   6795
         Left            =   45
         TabIndex        =   13
         Top             =   375
         Width           =   11790
         Begin VB.Frame Frame6 
            Caption         =   "( Periodo )"
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
            Height          =   720
            Left            =   9570
            TabIndex        =   49
            Top             =   630
            Width           =   2010
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
               Left            =   120
               TabIndex        =   50
               Top             =   330
               Width           =   1740
            End
         End
         Begin VB.CommandButton CmdBusDoc 
            Height          =   240
            Left            =   2160
            Picture         =   "FrmRetencion.frx":2B10
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   1230
            Width           =   240
         End
         Begin VB.Frame Frame5 
            Height          =   585
            Left            =   165
            TabIndex        =   39
            Top             =   6165
            Width           =   11505
            Begin VB.TextBox TxtImpRet 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
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
               Left            =   7875
               Locked          =   -1  'True
               TabIndex        =   41
               Text            =   "TxtImpRet"
               Top             =   195
               Width           =   1155
            End
            Begin VB.TextBox TxtImpPag 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
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
               Left            =   6000
               Locked          =   -1  'True
               TabIndex        =   40
               Text            =   "TxtImpPag"
               Top             =   195
               Width           =   1095
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total ==>"
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
               Left            =   4995
               TabIndex        =   42
               Top             =   240
               Width           =   825
            End
         End
         Begin VB.CommandButton CmdIdRet 
            Height          =   240
            Left            =   2160
            Picture         =   "FrmRetencion.frx":2C42
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   2520
            Width           =   240
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   5
            Text            =   "TxtNumSer"
            Top             =   2175
            Width           =   900
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2685
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   6
            Text            =   "TxtNumDoc"
            Top             =   2175
            Width           =   1950
         End
         Begin VB.Frame Frame3 
            Height          =   3045
            Left            =   10110
            TabIndex        =   29
            Top             =   3135
            Width           =   1560
            Begin VB.CommandButton CmdDel 
               Caption         =   "Eliminar Documento"
               Height          =   690
               Left            =   135
               TabIndex        =   30
               Top             =   1380
               Width           =   1305
            End
            Begin VB.CommandButton CmdAdd 
               Caption         =   "Agregar Documentos"
               Height          =   690
               Left            =   135
               TabIndex        =   10
               Top             =   660
               Width           =   1305
            End
         End
         Begin VB.CommandButton CmdBusMon 
            Height          =   240
            Left            =   7140
            Picture         =   "FrmRetencion.frx":2D74
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   2205
            Width           =   240
         End
         Begin VB.CommandButton CmdBusProv 
            Height          =   240
            Left            =   2880
            Picture         =   "FrmRetencion.frx":2EA6
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   1545
            Width           =   240
         End
         Begin VB.Frame FraTipo 
            Caption         =   "[ Tipo de Movimiento ]"
            Enabled         =   0   'False
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
            Height          =   600
            Left            =   165
            TabIndex        =   25
            Top             =   510
            Width           =   3090
            Begin VB.OptionButton Opt2 
               Caption         =   "Venta"
               Height          =   195
               Left            =   495
               TabIndex        =   0
               Top             =   240
               Width           =   1110
            End
            Begin VB.OptionButton Opt1 
               Caption         =   "Compra"
               Height          =   195
               Left            =   1650
               TabIndex        =   1
               Top             =   240
               Width           =   1110
            End
         End
         Begin VB.TextBox TxtIdMoneda 
            Height          =   300
            Left            =   6735
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   7
            Text            =   "TxtIdMoneda"
            Top             =   2175
            Width           =   675
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmi 
            Height          =   300
            Left            =   1530
            TabIndex        =   4
            Top             =   1845
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
         End
         Begin VB.TextBox TxtGlosa 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   9
            Text            =   "TxtGlosa"
            Top             =   2820
            Width           =   7845
         End
         Begin VB.TextBox TxtRucPro 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   3
            Text            =   "TxtRucPro"
            Top             =   1515
            Width           =   1620
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   2670
            Left            =   165
            TabIndex        =   11
            Top             =   3480
            Width           =   9840
            _cx             =   17357
            _cy             =   4710
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
            Cols            =   14
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmRetencion.frx":2FD8
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
         Begin VB.TextBox TxtIdRet 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   8
            Text            =   "TxtIdRet"
            Top             =   2490
            Width           =   900
         End
         Begin VB.TextBox TxtIdDoc 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   2
            Text            =   "TxtIdDoc"
            Top             =   1200
            Width           =   900
         End
         Begin VSFlex7Ctl.VSFlexGrid VSFlexGrid1 
            Height          =   270
            Left            =   165
            TabIndex        =   47
            Top             =   3225
            Width           =   9840
            _cx             =   17357
            _cy             =   476
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
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmRetencion.frx":3184
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
         Begin VB.Label lblReg 
            Caption         =   "lblReg"
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
            Height          =   270
            Left            =   9510
            TabIndex        =   48
            Top             =   300
            Width           =   2250
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Documento"
            Height          =   195
            Index           =   7
            Left            =   165
            TabIndex        =   45
            Top             =   1320
            Width           =   825
         End
         Begin VB.Label LblDocumento 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDocumento"
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
            Left            =   2445
            TabIndex        =   44
            Top             =   1200
            Width           =   4020
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio"
            Height          =   195
            Index           =   5
            Left            =   9495
            TabIndex        =   38
            Top             =   2265
            Width           =   885
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
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   10455
            TabIndex        =   37
            Top             =   2175
            Width           =   1155
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Tasa"
            Height          =   195
            Index           =   2
            Left            =   7710
            TabIndex        =   36
            Top             =   2580
            Width           =   360
         End
         Begin VB.Label LblTasa 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTasa"
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
            Left            =   8235
            TabIndex        =   35
            Top             =   2490
            Width           =   1155
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Reteción"
            Height          =   195
            Index           =   1
            Left            =   165
            TabIndex        =   34
            Top             =   2580
            Width           =   645
         End
         Begin VB.Label LblRetencion 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblRetencion"
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
            Left            =   2445
            TabIndex        =   33
            Top             =   2490
            Width           =   4920
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Nº Documento"
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   31
            Top             =   2265
            Width           =   1050
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00800000&
            BackStyle       =   1  'Opaque
            Height          =   75
            Left            =   2475
            Top             =   2295
            Width           =   135
         End
         Begin VB.Label LblIdProveedor 
            AutoSize        =   -1  'True
            Caption         =   "LblIdProveedor"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   4560
            TabIndex        =   27
            Top             =   975
            Visible         =   0   'False
            Width           =   1080
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
            Left            =   7425
            TabIndex        =   24
            Top             =   2175
            Width           =   1965
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   6
            Left            =   6090
            TabIndex        =   23
            Top             =   2265
            Width           =   585
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Glosa"
            Height          =   195
            Index           =   4
            Left            =   165
            TabIndex        =   22
            Top             =   2895
            Width           =   405
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Emisión"
            Height          =   195
            Index           =   3
            Left            =   165
            TabIndex        =   21
            Top             =   1950
            Width           =   1290
         End
         Begin VB.Label LblTitulo 
            AutoSize        =   -1  'True
            Caption         =   "LblTitulo"
            Height          =   195
            Left            =   165
            TabIndex        =   20
            Top             =   1635
            Width           =   600
         End
         Begin VB.Label LblProveedor 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblProveedor"
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
            Left            =   3165
            TabIndex        =   19
            Top             =   1515
            Width           =   6195
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Retención"
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
            Left            =   60
            TabIndex        =   14
            Top             =   30
            Width           =   11610
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   18
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
            Object.ToolTipText     =   "Cambiar mes de trabajo"
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
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Comprobante de Retencion"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Registro de Retenciones"
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
Attribute VB_Name = "FrmRetencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RstRet As New ADODB.Recordset
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim xIdCuenRet As Integer
Dim Agregando As Boolean
Dim xHorIni As Date

Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta
Dim mMesActivo As Integer '--indica el mes activo
Dim mIdRegistro& '--identificador del registro
Dim fCierrePeriodo As Boolean '--indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO


Private Sub CambiarMes()
    mMesActivo = SeleccionaMes(xCon)
    If mMesActivo = 0 Or mMesActivo = 13 Then
        MsgBox "Seleccione un periodo correcto", vbExclamation, xTitulo
        CambiarMes
        Exit Sub
    End If
    OpcionesPeriodo
    pCargarGrid
End Sub

Sub Cancelar()
    ActivaTool
    TabOne1.TabEnabled(0) = True
    Bloquea
    QueHace = 3
    Label5.Caption = "Detalle de la Retención"
    TabOne1.CurrTab = 0
End Sub

Private Sub CmdAdd_Click()
    If QueHace = 3 Then Exit Sub
    
    If TxtRucPro.Text = "" Then
        If Opt1.Value = True Then
            MsgBox "No ha especificado el proveedor al que se le aplicará la retención", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Else
            MsgBox "No ha especificado el cliente que se aplicara la retención", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
        Exit Sub
    End If
    
    If NulosC(TxtIdMoneda.Text) = "" Then
        MsgBox "No ha especificado la moneda para la retencion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMoneda.SetFocus
        Exit Sub
    End If
    
    If NulosC(TxtIdRet.Text) = "" Then
        MsgBox "No ha especificado la retencion que se aplicara", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdRet.SetFocus
        Exit Sub
    End If
    
    If NulosC(TxtFchEmi.Valor) = "" Then
        MsgBox "No ha especificado la fecha de emision de la retencion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchEmi.SetFocus
        Exit Sub
    End If
    
    'Dim xfrm As New EPS_Buscar.Seleccion
    Dim xCampos(6, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim xRs1 As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLNotId As String '--almacenara los documentos ya seleccionados para que no se vuelva a agregar
    Dim nTitulo  As String
    
    xCampos(0, 0) = "T.D.":             xCampos(0, 1) = "abrev":     xCampos(0, 2) = "450":    xCampos(0, 3) = "C":     xCampos(0, 4) = "N"
    xCampos(1, 0) = "Nº Documento":     xCampos(1, 1) = "numdoc":    xCampos(1, 2) = "2000":    xCampos(1, 3) = "C":     xCampos(1, 4) = "S"
    xCampos(2, 0) = "M":                xCampos(2, 1) = "simbolo":   xCampos(2, 2) = "450":    xCampos(2, 3) = "C":     xCampos(2, 4) = "N"
    xCampos(3, 0) = "Fch. Emision":     xCampos(3, 1) = "fchdoc1":    xCampos(3, 2) = "1200":    xCampos(3, 3) = "C":     xCampos(3, 4) = "N"
    xCampos(4, 0) = "Importe":          xCampos(4, 1) = "imptotdoc": xCampos(4, 2) = "1200":    xCampos(4, 3) = "N":     xCampos(4, 4) = "N"
    xCampos(5, 0) = "Saldo":            xCampos(5, 1) = "impsal":    xCampos(5, 2) = "1200":    xCampos(5, 3) = "N":     xCampos(5, 4) = "N"

    If Opt1.Value = True Then
        'tipoope = 0 Compras
        'generando el codigo para que no se repita los documentos ya seleccionado
        nSQLNotId = GRID_GENERAR_SQL_ID(Fg1, 10, " AND com_compras.id", " NOT IN", True)
        
        nSQL = "SELECT 0 as xsel,mae_documento.abrev, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, com_compras.impsal, " _
            & " com_compras.fchdoc & '' as fchdoc1, com_compras.idpro, com_compras.imptot as imptotdoc, com_compras.id, mae_documentocta.idcuen " _
            & " FROM (mae_documento INNER JOIN com_compras ON mae_documento.id = com_compras.tipdoc) LEFT JOIN mae_documentocta " _
            & " ON mae_documento.id = mae_documentocta.iddoc Where (((com_compras.idpro) = " & NulosN(LblIdProveedor.Caption) & ") And ((com_compras.impsal) <> 0) " _
            & " And ((com_compras.idmon) = " & NulosN(TxtIdMoneda.Text) & ") And ((mae_documentocta.tipope) = 0) And " _
            & " ((mae_documentocta.idmon) = " & NulosN(TxtIdMoneda.Text) & ")) " & nSQLNotId & " ORDER BY com_compras.fchdoc"

        nTitulo = "Buscando Documentos del Proveedor"
    Else
        'tipoope = -1 ventas
        'generando el codigo para que no se repita los documentos ya seleccionado
        nSQLNotId = GRID_GENERAR_SQL_ID(Fg1, 10, " AND vta_ventas.id", " NOT IN", True)
        
        nSQL = "SELECT 0 as xsel, IIf(IsNull([vta_ventas]![numser])=-1,[vta_ventas]![numdoc],[vta_ventas]![numser]+'-'+[vta_ventas]![numdoc]) AS numdoc," _
            & " mae_documento.abrev, mae_cliente.nombre, vta_ventas.tipdoc, mae_moneda.simbolo, vta_ventas.fchdoc & '' as fchdoc1, vta_ventas.imptotdoc, " _
            & " vta_ventas.impsal, vta_ventas.idmon, vta_ventas.id, mae_documentocta.idcuen, mae_documentocta.tipope" _
            & " FROM (((mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) " _
            & " LEFT JOIN mae_documentocta ON (vta_ventas.tipdoc = mae_documentocta.iddoc) AND (vta_ventas.idmon = mae_documentocta.idmon)) LEFT JOIN mae_documento " _
            & " ON vta_ventas.tipdoc = mae_documento.id Where (((vta_ventas.impsal) <> 0) And ((mae_documentocta.tipope) = -1) " _
            & " And ((vta_ventas.idcli) = " & NulosN(LblIdProveedor.Caption) & ")) " & nSQLNotId & " ORDER BY IIf(IsNull([vta_ventas]![numser])=-1,[vta_ventas]![numdoc],[vta_ventas]![numser]+'-'+[vta_ventas]![numdoc])"

        
        nTitulo = "Buscando Documentos del Cliente"
    End If
    RST_Busq xRs, nSQL, xCon
    
    CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), nTitulo
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            Dim A As Integer
           
            xRs.MoveFirst
            
            Agregando = True
            For A = 1 To xRs.RecordCount
                Fg1.Rows = Fg1.Rows + 1
                
                Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(xRs("numdoc"))
                Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(xRs("simbolo"))
                Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(xRs("abrev"))
                Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(xRs("fchdoc1"))
                Fg1.TextMatrix(Fg1.Rows - 1, 12) = NulosN(xRs("idmon"))
                Fg1.TextMatrix(Fg1.Rows - 1, 13) = NulosC(xRs("tipdoc"))
                
                If xRs("tipdoc") = "7" Then
                    Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(NulosN(xRs("imptotdoc")), "-0.00")
                    Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(NulosN(xRs("imptotdoc")), "-0.00")
                    Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(NulosN(LblTasa.Caption), "-0.00")
                    
                    If NulosN(TxtIdMoneda.Text) = 1 Then
                        If xRs("idmon") = 1 Then
                            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(NulosN(xRs("imptotdoc")), "-0.00")
                        Else
                            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(NulosN(xRs("imptotdoc")) * NulosN(LblTipoCambio.Caption), "-0.00")
                        End If
                    End If
                Else
                    Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(NulosN(xRs("imptotdoc")), FORMAT_MONTO)
                    Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(NulosN(xRs("imptotdoc")), FORMAT_MONTO)
                    Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(NulosN(LblTasa.Caption), FORMAT_MONTO)
                    
                    If NulosN(TxtIdMoneda.Text) = 1 Then
                        If xRs("idmon") = 1 Then
                            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(NulosN(xRs("imptotdoc")), FORMAT_MONTO)
                        Else
                            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(NulosN(xRs("imptotdoc")) * NulosN(LblTipoCambio.Caption), FORMAT_MONTO)
                        End If
                    End If
                End If

                Fg1.TextMatrix(Fg1.Rows - 1, 10) = NulosN(xRs("id"))
                Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosN(xRs("idcuen"))
                
'                Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(NULOSN(LblTasa.Caption), "0.00")
                Fg1.TextMatrix(Fg1.Rows - 1, 9) = (NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 7)) * (NulosN(LblTasa.Caption) / 100))
                Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(Fg1.TextMatrix(Fg1.Rows - 1, 9), FORMAT_MONTO)
                
                Set xRs1 = Nothing
                
                xRs.MoveNext
                If xRs.EOF = True Then
                    Exit For
                End If
            Next A
            Agregando = False
        End If
    End If
    HallarTotales
    Set xRs = Nothing
End Sub

Private Sub CmdBusDoc_Click()
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos2(3, 4) As String
    
    xCampos2(0, 0) = "Documento":    xCampos2(0, 1) = "descripcion":    xCampos2(0, 2) = "5000":         xCampos2(0, 3) = "C"
    xCampos2(1, 0) = "Sigla":        xCampos2(1, 1) = "abrev":          xCampos2(1, 2) = "1000":         xCampos2(1, 3) = "N"
    xCampos2(2, 0) = "Codigo":       xCampos2(2, 1) = "id":             xCampos2(2, 2) = "1000":         xCampos2(2, 3) = "N"

    xform.SqlCad = "SELECT * FROM mae_documento"
    xform.Titulo = "Buscando Monedas"
        
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos2)
    
    If xRs.State = 1 Then
        TxtIdDoc.Text = xRs("id")
        LblDocumento.Caption = xRs("descripcion")
        TxtRucPro.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMon_Click()
    If QueHace = 3 Then Exit Sub
    Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos2(2, 4) As String
    
    xCampos2(0, 0) = "Moneda":        xCampos2(0, 1) = "descripcion":      xCampos2(0, 2) = "4000":         xCampos2(0, 3) = "C"
    xCampos2(1, 0) = "Abreviatura":   xCampos2(1, 1) = "simbolo":          xCampos2(1, 2) = "1500":         xCampos2(1, 3) = "C"
    
    xform.SqlCad = "SELECT mae_moneda.* FROM mae_moneda ORDER BY descripcion"
    
    xform.Titulo = "Buscando Monedas"
    
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos2)
    
    If xRs.State = 1 Then
        TxtIdMoneda.Text = xRs("id")
        LblMoneda.Caption = xRs("descripcion")
        If TxtIdMoneda.Text = "1" Then
            LblTipoCambio.Caption = ""
            LblTipCam(5).Visible = False
            LblTipoCambio.Visible = False
        Else
            If TxtFchEmi.Valor = "" Then
                TxtIdMoneda.Text = ""
                LblTipoCambio.Caption = ""
                Exit Sub
            End If
            LblTipoCambio.Caption = HallaTipoCambio(TxtFchEmi.Valor, NulosN(TxtIdMoneda.Text), Venta, xCon)
            LblTipCam(5).Visible = True
            LblTipoCambio.Visible = True
        End If
        TxtIdMoneda.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusProv_Click()
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos2(2, 4) As String
    
    If Opt1.Value = True Then
        xCampos2(0, 0) = "Proveedor":   xCampos2(0, 1) = "nombre":       xCampos2(0, 2) = "5000":         xCampos2(0, 3) = "C"
        xform.SqlCad = "SELECT mae_prov.* From mae_prov where ageret=-1 ORDER BY mae_prov.nombre"
        xform.Titulo = "Buscando Proveedores"
    Else
        xCampos2(0, 0) = "Cliente":   xCampos2(0, 1) = "nombre":       xCampos2(0, 2) = "5000":         xCampos2(0, 3) = "C"
        xform.SqlCad = "SELECT mae_cliente.* From mae_cliente where ageret=-1 ORDER BY mae_cliente.nombre"
        xform.Titulo = "Buscando Clientes"
    End If
    xCampos2(1, 0) = "Nº R.U.C.":   xCampos2(1, 1) = "numruc":       xCampos2(1, 2) = "1500":         xCampos2(1, 3) = "C"
        
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos2)
    
    If xRs.State = 1 Then
        TxtRucPro.Text = xRs("numruc")
        LblProveedor.Caption = xRs("nombre")
        LblIdProveedor.Caption = xRs("id")
        TxtRucPro.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdDel_Click()
    If QueHace = 3 Then Exit Sub
    If Fg1.Row < 1 Then Exit Sub
    Fg1.RemoveItem Fg1.Row
    HallarTotales
End Sub

Private Sub CmdIdRet_Click()
    If QueHace = 3 Then Exit Sub
    
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim xCampos2(5, 4) As String
    
    xCampos2(0, 0) = "Descripción":     xCampos2(0, 1) = "descripcion": xCampos2(0, 2) = "3200":         xCampos2(0, 3) = "C"
    xCampos2(1, 0) = "Tasa":            xCampos2(1, 1) = "tasa1":       xCampos2(1, 2) = "800":          xCampos2(1, 3) = "C"
    xCampos2(2, 0) = "Cta Número.":     xCampos2(2, 1) = "ctanum":      xCampos2(2, 2) = "1200":         xCampos2(2, 3) = "C"
    xCampos2(3, 0) = "Cta Descripción": xCampos2(3, 1) = "ctadesc":     xCampos2(3, 2) = "2500":         xCampos2(3, 3) = "C"
    xCampos2(4, 0) = "Id":              xCampos2(4, 1) = "id":          xCampos2(4, 2) = "500":          xCampos2(4, 3) = "N"
    
    If Opt1.Value = True Then '--compra
        nSQL = "SELECT mae_retencion.id, mae_retencion.descripcion, mae_retencion.tasa,mae_retencion.idcuencom, mae_retencion.idcuenven, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, Format([mae_retencion].[tasa],'0.00') & '%' AS tasa1 " _
                & " FROM mae_retencion LEFT JOIN con_planctas ON mae_retencion.idcuencom = con_planctas.id "
    Else '--venta
        nSQL = "SELECT mae_retencion.id, mae_retencion.descripcion,mae_retencion.tasa, mae_retencion.idcuencom, mae_retencion.idcuenven, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, Format([mae_retencion].[tasa],'0.00') & '%' AS tasa1 " _
                & " FROM mae_retencion LEFT JOIN con_planctas ON mae_retencion.idcuenven = con_planctas.id;  "
    End If
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos2(), "Buscando Retenciones", "descripcion", "descripcion", Principio

    
    If xRs.State = 1 Then
        TxtIdRet.Text = xRs("id")
        LblRetencion.Caption = xRs("descripcion")
        LblTasa.Caption = xRs("tasa")
        TxtIdRet.SetFocus
        
        If Opt1.Value = True Then
            'CUENTA CONTABLE DE LA RETENCION CUANDO SEA COMPRA
            xIdCuenRet = xRs("idcuencom")
        Else
            'CUENTA CONTABLE DE LA RETENCION CUANDO SEA VENTA
            xIdCuenRet = xRs("idcuenven")
        End If
    End If
    
    Set xRs = Nothing
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstRet
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstRet.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 Then
        VerMovimientos1 IdMenuActivo, RstRet("id"), xCon
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If QueHace = 3 Then Exit Sub

    If Fg1.Col = 1 Then
        CmdAdd_Click
    End If
End Sub

Sub HallarTotales()
   
    TxtImpPag.Text = Format(GRID_SUMAR_COL(Fg1, 7), FORMAT_MONTO)
    TxtImpRet.Text = Format(GRID_SUMAR_COL(Fg1, 9), FORMAT_MONTO)
    
End Sub

Function ExisteRetencion(NumSer As String, NumDoc As String, IdCliente As Integer, Tipo As Integer) As Boolean
    'Tipo  1 = Compras     2 = Ventas
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT * FROM con_retencion WHERE numdoc = '" & NumDoc & "'  AND numser = '" & NumSer & "' AND idpro = " & IdCliente & " AND tipo = " & Tipo & "", xCon
    If Rst.RecordCount <> 0 Then
        ExisteRetencion = True
    Else
        ExisteRetencion = False
    End If
    Set Rst = Nothing
End Function

Function Grabar() As Boolean
    Dim xTipo As Integer
    If Opt1.Value = True Then
        xTipo = 1
    End If
    If Opt2.Value = True Then
        xTipo = 2
    End If
    If QueHace = 1 Then
        If ExisteRetencion(TxtNumSer.Text, TxtNumDoc.Text, LblIdProveedor.Caption, xTipo) = True Then
            MsgBox "El número de retención especificado ya fue registrado, verifique los datos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
    End If
    
    If TxtRucPro.Text = "" Then
        If Opt1.Value = True Then
            MsgBox "No ha especificado el proveedor", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Else
            MsgBox "No ha especificado el cliente", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
        TxtRucPro.SetFocus
        Exit Function
    End If
    
    If IsDate(TxtFchEmi.Valor) = False Then
        MsgBox "No ha especificado la fecha de emisión del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchEmi.SetFocus
        Exit Function
    End If
    
    If TxtNumSer.Text = "" Then
        MsgBox "No ha especificado el número de serie del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumSer.SetFocus
        Exit Function
    End If

    If TxtNumSer.Text = "" Then
        MsgBox "No ha especificado el número de serie del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumSer.SetFocus
        Exit Function
    End If

    If TxtNumDoc.Text = "" Then
        MsgBox "No ha especificado el número del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDoc.SetFocus
        Exit Function
    End If

    If TxtIdMoneda.Text = "" Then
        MsgBox "No ha especificado la moneda", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMoneda.SetFocus
        Exit Function
    End If
    
    If TxtIdRet.Text = "" Then
        MsgBox "No ha especificado la retención que se aplicara", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdRet.SetFocus
        Exit Function
    End If
    
    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado a que documentos se le aplicara la retención", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.SetFocus
        Exit Function
    End If
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
'''    Dim RstDia As New ADODB.Recordset
    Dim A As Integer
    Dim xId As Double
    Dim xNumAsiento As String
    
On Error GoTo LaCague

    xCon.BeginTrans
    If QueHace = 1 Then
        xId = HallaCodigoTabla("con_retencion", xCon, "id")
        xNumAsiento = NuevoNumAsiento(5, mMesActivo, xCon)
        
        RST_Busq RstCab, "SELECT TOP 1 * FROM con_retencion", xCon
        
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstRet("id")
        RST_Busq RstCab, "SELECT * FROM con_retencion WHERE id = " & xId & "", xCon
        
        'actualizamos el saldo de los documentos involucrados en la retencion
        For A = 1 To Fg1.Rows - 1
            If Opt1.Value = True Then
                'actualizasmos el saldo de los documentos de compra
                xCon.Execute "UPDATE com_compras SET com_compras.impsal = [com_compras]![impsal]+" & NulosN(Fg1.TextMatrix(A, 7)) & " " _
                    & " WHERE (((com_compras.id)=" & NulosN(Fg1.TextMatrix(A, 8)) & "))"
            Else
                'actualizasmos el saldo de los documentos de venta
                xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = [vta_ventas]![impsal]+" & NulosN(Fg1.TextMatrix(A, 7)) & " " _
                    & " WHERE (((vta_ventas.id)=" & NulosN(Fg1.TextMatrix(A, 8)) & "))"
            End If
        Next A
        
        'eliminamos el detalle de la retencion
        xCon.Execute "DELETE * FROM con_retenciondet WHERE id = " & xId & ""
        
        'elimiminamos el libro diario
''''        xNumAsiento = DevuelveNumAsiento(5, RstRet("id"), mMesActivo, xCon)
''''        If xNumAsiento = "" Then
''''            xNumAsiento = NuevoNumAsiento(4, mMesActivo, xCon)
''''        End If
''''        xCon.Execute "DELETE * FROM con_diario WHERE ((idmes = " & mMesActivo & ") and (idlib = 5) AND (idmov = " & xId & "))"
        
    End If
    '-----------------------------------------------------------------
    mIdRegistro = xId
    RST_Busq RstDet, "SELECT TOP 1 * FROM con_retenciondet", xCon
''''    RST_Busq RstDia, "SELECT TOP 1 * FROM con_diario", xCon
    '-----------------------------------------------------------------
   
    RstCab("idret") = NulosN(TxtIdRet.Text)
    If Opt1.Value = True Then
        RstCab("tipo") = 1
    Else
        RstCab("tipo") = 2
    End If
    
    RstCab("idlib") = 5
    RstCab("idpro") = NulosN(LblIdProveedor.Caption)
    RstCab("iddoc") = NulosN(TxtIdDoc.Text)
    RstCab("numser") = TxtNumSer.Text
    RstCab("numdoc") = TxtNumDoc.Text
    RstCab("fchemi") = TxtFchEmi.Valor
    RstCab("idmon") = NulosN(TxtIdMoneda.Text)
    RstCab("imp") = NulosN(TxtImpRet.Text)
    RstCab("fchreg") = "01/" + Format(mMesActivo, "00") + "/" + Trim(Str(AnoTra))
'''    RstCab("numreg") = Format(mMesActivo, "00") + xNumAsiento
    RstCab.Update
    
    For A = 1 To Fg1.Rows - 1
        RstDet.AddNew
        RstDet("id") = xId
        RstDet("iddoc") = Fg1.TextMatrix(A, 10)
        RstDet("impcob") = Fg1.TextMatrix(A, 6)
        RstDet("impret") = Fg1.TextMatrix(A, 9)
       
        If Opt1.Value = True Then
            xCon.Execute "UPDATE com_compras SET com_compras.impsal = [com_compras]![impsal]-" & NulosN(Fg1.TextMatrix(A, 7)) & " " _
                & " WHERE (((com_compras.id)=" & NulosN(Fg1.TextMatrix(A, 10)) & "))"
        Else
            If NulosN(Fg1.TextMatrix(A, 12)) = 1 Then
                xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = [vta_ventas]![impsal]-" & NulosN(Fg1.TextMatrix(A, 7)) & " " _
                    & " WHERE (((vta_ventas.id)=" & NulosN(Fg1.TextMatrix(A, 10)) & "))"
            Else
                xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = [vta_ventas]![impsal]-" & NulosN(Fg1.TextMatrix(A, 6)) & " " _
                    & " WHERE (((vta_ventas.id)=" & NulosN(Fg1.TextMatrix(A, 10)) & "))"
            End If
        End If
        
        RstDet.Update
    Next A
    
    
    'GRABAMOS EL LIBRO DIARIO
    Dim Cambio As Boolean
   
'''''    'GRABAMOS LA CUENTA DEBE CON EL CODIGO DE LA CUENTA DE LA RETENCION
'''''    RstDia.AddNew
'''''    RstDia("año") = AnoTra
'''''    RstDia("idmes") = mMesActivo
'''''    RstDia("idlib") = 5
'''''    RstDia("idmov") = xId
'''''    RstDia("numasi") = xNumAsiento
'''''    RstDia("tc") = NulosN(LblTipoCambio.Caption)
'''''    RstDia("fchasi") = "01/" + Format(mMesActivo, "00") + "/" + Trim(Str(AnoTra))
'''''    RstDia("fchdoc") = NulosC(TxtFchEmi.Valor)
'''''    RstDia("idcue") = xIdCuenRet
'''''    If Opt1.Value = True Then
'''''        If TxtIdMoneda.Text = "1" Then
'''''            RstDia("imphabsol") = NulosN(TxtImpRet.Text)
'''''            RstDia("imphabdol") = 0
'''''        Else
'''''            RstDia("imphabsol") = NulosN(TxtImpRet.Text) * NulosN(LblTipoCambio.Caption)
'''''            RstDia("imphabdol") = NulosN(TxtImpRet.Text)
'''''        End If
'''''    Else
'''''        If TxtIdMoneda.Text = "1" Then
'''''            RstDia("impdebsol") = NulosN(TxtImpRet.Text)
'''''            RstDia("impdebdol") = 0
'''''        Else
'''''            RstDia("impdebsol") = NulosN(TxtImpRet.Text) * NulosN(LblTipoCambio.Caption)
'''''            RstDia("impdebdol") = NulosN(TxtImpRet.Text)
'''''        End If
'''''    End If
'''''    RstDia.Update
'''''
'''''
'''''    'GRABAMOS LA CUENTA HABER CON EL CODIGO DE LA CUENTA DE LOS DOCUMENTOS INVOLUCRADOS EN LA RETENCION
'''''    Dim xIdCuen As Integer
'''''    Dim xTotal As Double
'''''    A = 1
'''''    xIdCuen = NULOSN(Fg1.TextMatrix(A, 11))
'''''    For A = 1 To Fg1.Rows - 1
'''''        'If xIdCuen = NULOSN(Fg1.TextMatrix(A, 9)) Then
'''''        '    xTotal = xTotal + NULOSN(Fg1.TextMatrix(A, 7))
'''''        '    Cambio = False
'''''        'Else
'''''            RstDia.AddNew
'''''            RstDia("año") = AnoTra
'''''            RstDia("idmes") = mMesActivo               'LLAVE - CODIGO DEL MES
'''''            RstDia("idlib") = 5                  'LLAVE - CODIGO DEL LIBRO
'''''            RstDia("idmov") = xId                'LLAVE - CODIGO DEL MOVIMIENTO
'''''            RstDia("numasi") = xNumAsiento       'LLAVE - NUMERO DE ASIENTO
'''''
'''''            RstDia("tc") = NulosN(LblTipoCambio.Caption)
'''''            RstDia("idcue") = xIdCuen
'''''            RstDia("iddocpro") = Fg1.TextMatrix(A, 10)
'''''            RstDia("fchasi") = "01/" + Format(mMesActivo, "00") + "/" + Trim(Str(AnoTra))
'''''            RstDia("fchdoc") = NulosC(TxtFchEmi.Valor)
'''''            If Opt1.Value = True Then
'''''                If TxtIdMoneda.Text = "1" Then
'''''                    RstDia("impdebsol") = NulosN(Fg1.TextMatrix(A, 9))
'''''                    RstDia("impdebdol") = 0
'''''                Else
'''''                    RstDia("impdebsol") = NulosN(Fg1.TextMatrix(A, 9)) * NulosN(LblTipoCambio.Caption)
'''''                    RstDia("impdebdol") = NulosN(Fg1.TextMatrix(A, 9))
'''''                End If
'''''            Else
'''''                If TxtIdMoneda.Text = "1" Then
'''''                    RstDia("imphabsol") = NulosN(Fg1.TextMatrix(A, 9))
'''''                    RstDia("imphabdol") = 0
'''''                Else
'''''                    RstDia("imphabsol") = NulosN(Fg1.TextMatrix(A, 9)) * NulosN(LblTipoCambio.Caption)
'''''                    RstDia("imphabdol") = NulosN(Fg1.TextMatrix(A, 9))
'''''                End If
'''''            End If
'''''            RstDia.Update
'''''            Cambio = True
'''''        'End If
'''''    Next A
'''''
'''''
'''''
    '----------------------------------------------------------------------------------
    '---generar asiento
    xNumAsiento = GenerarAsiento(xCon, 5, xId, AnoTra, mMesActivo, 0)
    If xNumAsiento = "" Then GoTo LaCague
    '----------------------------------------------------------------------------------

    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId

    
    xCon.CommitTrans
    Grabar = True
    Set RstCab = Nothing
    Set RstDet = Nothing
    
    MsgBox "La Retención se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito" + vbCr + "Num.Reg. :" & xNumAsiento, vbInformation, xTitulo

    Exit Function
    
LaCague:
    xCon.RollbackTrans
    MsgBox "No se pudo guardar la retención por el siguiente motivo: " + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set RstCab = Nothing
    Set RstDet = Nothing
''    Set RstDia = Nothing
End Function

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    
    If Fg1.Col = 6 Or Fg1.Col = 7 Then
        Fg1.TextMatrix(Fg1.Row, 6) = Format(Fg1.TextMatrix(Fg1.Row, 6), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Row, 7) = Format(Fg1.TextMatrix(Fg1.Row, 7), FORMAT_MONTO)
        
        If Fg1.Col = 6 Then
            If Fg1.TextMatrix(Fg1.Row, 12) = 1 Then
                Fg1.TextMatrix(Fg1.Row, 7) = NulosN(Fg1.TextMatrix(Fg1.Row, 6))
            Else
                Fg1.TextMatrix(Fg1.Row, 7) = NulosN(Fg1.TextMatrix(Fg1.Row, 6)) * NulosN(LblTipoCambio.Caption)
                Fg1.TextMatrix(Fg1.Row, 7) = Format(Fg1.TextMatrix(Fg1.Row, 7), FORMAT_MONTO)
            End If
        End If
        
        If Fg1.Col = 7 Then
            If Fg1.TextMatrix(Fg1.Row, 12) = 1 Then
                Fg1.TextMatrix(Fg1.Row, 6) = NulosN(Fg1.TextMatrix(Fg1.Row, 7))
            Else
                Fg1.TextMatrix(Fg1.Row, 6) = NulosN(Fg1.TextMatrix(Fg1.Row, 7)) / NulosN(LblTipoCambio.Caption)
                Fg1.TextMatrix(Fg1.Row, 6) = Format(Fg1.TextMatrix(Fg1.Row, 6), FORMAT_MONTO)
            End If
        End If
        
        Fg1.TextMatrix(Fg1.Row, 9) = (NulosN(Fg1.TextMatrix(Fg1.Row, 7)) * (NulosN(LblTasa.Caption) / 100))
        Fg1.TextMatrix(Fg1.Row, 9) = Format(Fg1.TextMatrix(Fg1.Row, 9), FORMAT_MONTO)
        
'        If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 13)) = 7 Then
'            If NulosN(TxtIdMoneda.Text) = 1 Then
'                If Fg1.TextMatrix(Fg1.Rows - 1, 12) = 1 Then
'                    Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(Fg1.TextMatrix(Fg1.Rows - 1, 6), "-0.00")
'                Else
'                    Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 6)) * NulosN(LblTipoCambio.Caption), "-0.00")
'                End If
'            End If
'        Else
'            If NulosN(TxtIdMoneda.Text) = 1 Then
'                If Fg1.TextMatrix(Fg1.Rows - 1, 12) = 1 Then
'                    Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(Fg1.TextMatrix(Fg1.Rows - 1, 6), FORMAT_MONTO)
'                Else
'                    Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 6)) * NulosN(LblTipoCambio.Caption), FORMAT_MONTO)
'                End If
'            End If
'        End If
    End If
    HallarTotales
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
        Exit Sub
    End If
    If Fg1.Col = 1 Or Fg1.Col = 6 Or Fg1.Col = 7 Then
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 45 Then
        CmdAdd_Click
    End If
    
    If KeyCode = 46 Then
        CmdDel_Click
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
    
        SeEjecuto = True
                
        mMesActivo = xMes
    
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        pCargarGrid
    
    End If
End Sub

Private Sub Form_Load()
    QueHace = 3
    SeEjecuto = False
    
    Dg1.Columns("fchemi").NumberFormat = FORMAT_DATE
    Dg1.Columns("imp1").NumberFormat = FORMAT_MONTO

    
    Fg1.ColWidth(10) = 0
    Fg1.ColWidth(11) = 0
    Fg1.ColWidth(12) = 0
    Fg1.ColWidth(13) = 0
    
    TabOne1.CurrTab = 0
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    Fg1.SelectionMode = flexSelectionByRow
    
    Fg1.ColFormat(4) = FORMAT_DATE

End Sub

Sub ActivaTool()

    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Sub Nuevo()
    QueHace = 1
    Label5.Caption = "Agregando Retención"
    Bloquea
    Blanquea
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Fg1.ColComboList(1) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    Fg1.Rows = 1
    
    Fg1.Editable = flexEDKbdMouse
    Opt2.Value = True
    Opt2_Click
    'LblTituloDoc.Caption = "Documentos de Compra"
    TxtIdMoneda.Text = "1"
    TxtIdMoneda_Validate True
    
    TxtIdDoc.Text = "20"
    TxtIdDoc_Validate True
    xHorIni = Time
    TxtRucPro.SetFocus
End Sub

Sub Bloquea()

    FraTipo.Enabled = Not FraTipo.Enabled
    
    TxtRucPro.Locked = Not TxtRucPro.Locked
    TxtFchEmi.Locked = Not TxtFchEmi.Locked
    TxtNumSer.Locked = Not TxtNumSer.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    'TxtIdMoneda.Locked = Not TxtIdMoneda.Locked
    TxtIdDoc.Locked = Not TxtIdDoc.Locked
    TxtIdRet.Locked = Not TxtIdRet.Locked
    TxtGlosa.Locked = Not TxtGlosa.Locked
    
End Sub

Sub Modificar()
    If RstRet.RecordCount = 0 Then
        MsgBox "No hay Registros para Modificar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If

    QueHace = 2
    Label5.Caption = "Modificando Retención"
    Bloquea
    ActivaTool
    If TabOne1.CurrTab = 0 Then
        TabOne1.CurrTab = 1
        MuestraSegundoTab
    End If
    TabOne1.TabEnabled(0) = False
    Fg1.ColComboList(1) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    Fg1.Editable = flexEDKbdMouse
    xHorIni = Time
    TxtRucPro.SetFocus
End Sub

Sub MuestraSegundoTab()
    Blanquea
    If RstRet.State = 0 Then Exit Sub
    If RstRet.EOF = True Or RstRet.BOF = True Or RstRet.RecordCount = 0 Then
        MsgBox "No hay Registros de Retenciones", vbExclamation, xTitulo
        Exit Sub
    End If
    lblReg.Caption = "Nº Reg. " & NulosC(RstRet("registro"))
    TxtIdDoc.Text = RstRet("iddoc")
    LblDocumento.Caption = Busca_Codigo(RstRet("iddoc"), "id", "descripcion", "mae_documento", "N", xCon)
    
    If RstRet("tipo") = 1 Then
        'LblTituloDoc.Caption = "Documentos de Compra"
        Opt1.Value = True
        Opt2.Value = False
    Else
        'LblTituloDoc.Caption = "Documentos de Venta"
        Opt1.Value = False
        Opt2.Value = True
    End If
    
    TxtRucPro.Text = RstRet("numruc")
    LblProveedor.Caption = NulosC(RstRet("nombre"))
    LblIdProveedor.Caption = RstRet("id")
    If IsDate(RstRet("fchemi")) = True Then
        TxtFchEmi.Valor = RstRet("fchemi")
        TxtFchEmi_Validate True
    End If
    
    TxtNumSer.Text = NulosC(RstRet("numser"))
    TxtNumDoc.Text = NulosC(RstRet("numdoc"))
    TxtIdMoneda.Text = NulosN(RstRet("idmon"))
    TxtIdRet.Text = NulosN(RstRet("idret"))
    TxtGlosa.Text = NulosC(RstRet("glosa"))
    
    
    
    If NulosC(TxtIdMoneda.Text) = "2" Then
        LblTipoCambio.Caption = HallaTipoCambio(TxtFchEmi.Valor, NulosN(TxtIdMoneda.Text), Venta, xCon)
        LblTipoCambio.Visible = True
        LblTipCam(5).Visible = True
    End If
    
    LblTasa.Caption = Format(Busca_Codigo(NulosN(TxtIdRet.Text), "id", "tasa", "mae_retencion", "N", xCon), "0.00")
    LblMoneda.Caption = Busca_Codigo(NulosN(TxtIdMoneda.Text), "id", "descripcion", "mae_moneda", "N", xCon)
    LblRetencion.Caption = Busca_Codigo(NulosN(TxtIdRet.Text), "id", "descripcion", "mae_retencion", "N", xCon)
    
    If Opt1.Value = True Then
        'CUENTA CONTABLE DE LA RETENCION CUANDO SEA COMPRA
        xIdCuenRet = Busca_Codigo(TxtIdRet.Text, "id", "idcuencom", "mae_retencion", "N", xCon)
    Else
        'CUENTA CONTABLE DE LA RETENCION CUANDO SEA VENTA
        xIdCuenRet = Busca_Codigo(TxtIdRet.Text, "id", "idcuenven", "mae_retencion", "N", xCon)
    End If
    
    Fg1.Rows = Fg1.Rows - 1
    Dim A As Integer
    Dim RstDet As New ADODB.Recordset
    Dim SqlCad As String
    
    If Opt1.Value = True Then
        SqlCad = "SELECT con_retencion.id, con_retenciondet.iddoc, com_compras!numser+'-'+com_compras!numdoc AS numdoc, mae_documento.abrev, " _
            & " com_compras.fchdoc, com_compras.imptot, con_retenciondet.impcob, con_retenciondet.impret, mae_documentocta.idmon, " _
            & " mae_documentocta.idcuen, mae_documentocta.tipope FROM (mae_documento RIGHT JOIN (con_retencion LEFT JOIN " _
            & " (con_retenciondet LEFT JOIN com_compras ON con_retenciondet.iddoc = com_compras.id) ON " _
            & " con_retencion.id = con_retenciondet.id) ON mae_documento.id = com_compras.tipdoc) LEFT JOIN mae_documentocta " _
            & " ON mae_documento.id = mae_documentocta.iddoc WHERE (((con_retencion.id)=" & NulosN(RstRet("id")) & ") " _
            & " AND ((mae_documentocta.idmon)=" & NulosN(TxtIdMoneda.Text) & ") AND ((mae_documentocta.tipope)=0))"
    Else
        SqlCad = "SELECT con_retencion.id, con_retenciondet.iddoc, IIf(IsNull([vta_ventas]![numser])=-1,[vta_ventas]![numdoc],[vta_ventas]![numser]+'-'+[vta_ventas]![numdoc]) AS numdoc, " _
            & " mae_moneda.simbolo, mae_documento.abrev, vta_ventas.fchdoc, vta_ventas.imptotdoc AS imptot, con_retenciondet.impcob, con_retenciondet.impret, " _
            & " mae_documentocta.idcuen, vta_ventas.idmon, vta_ventas.tipdoc FROM ((con_retencion LEFT JOIN ((con_retenciondet LEFT JOIN vta_ventas ON con_retenciondet.iddoc = vta_ventas.id) " _
            & " LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) ON con_retencion.id = con_retenciondet.id) LEFT JOIN mae_documentocta " _
            & " ON (vta_ventas.idmon = mae_documentocta.idmon) AND (vta_ventas.tipdoc = mae_documentocta.iddoc)) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id " _
            & " WHERE (((con_retencion.id)=" & NulosN(RstRet("id")) & ") AND ((mae_documentocta.tipope)=-1 Or (mae_documentocta.tipope) Is Null))  "
    End If
    
    RST_Busq RstDet, SqlCad, xCon
    Fg1.Rows = 1
    Agregando = True
    If RstDet.RecordCount <> 0 Then
        RstDet.MoveFirst
        For A = 1 To RstDet.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = NulosC(RstDet("numdoc"))
            Fg1.TextMatrix(A, 2) = NulosC(RstDet("simbolo"))
            Fg1.TextMatrix(A, 3) = NulosC(RstDet("abrev"))
            Fg1.TextMatrix(A, 4) = NulosC(RstDet("fchdoc"))
            Fg1.TextMatrix(A, 5) = Format(RstDet("imptot"), FORMAT_MONTO)
            
            If RstDet("idmon") = 1 Then
                Fg1.TextMatrix(A, 6) = Format(NulosN(RstDet("impcob")), FORMAT_MONTO)
                Fg1.TextMatrix(A, 7) = Format(NulosN(RstDet("impcob")), FORMAT_MONTO)
            Else
                Fg1.TextMatrix(A, 6) = Format(NulosN(RstDet("impcob")), FORMAT_MONTO)
                Fg1.TextMatrix(A, 7) = Format(NulosN(RstDet("impcob")) * NulosN(LblTipoCambio.Caption), FORMAT_MONTO)
            End If
            
            Fg1.TextMatrix(A, 8) = Format(LblTasa.Caption, "0.00")
            Fg1.TextMatrix(A, 9) = Format(NulosN(RstDet("impret")), FORMAT_MONTO)
            Fg1.TextMatrix(A, 10) = NulosN(RstDet("iddoc"))
            Fg1.TextMatrix(A, 11) = NulosN(RstDet("idcuen"))
            Fg1.TextMatrix(A, 12) = NulosN(RstDet("idmon"))
            Fg1.TextMatrix(A, 13) = NulosN(RstDet("tipdoc"))
            RstDet.MoveNext
            If RstDet.EOF = True Then Exit For
        Next A
    End If
    HallarTotales
    Agregando = False
    
End Sub

Sub Eliminar()
    TabOne1.CurrTab = 0
    If RstRet.RecordCount = 0 Then
        MsgBox "No hay registros para eliminar", vbExclamation, xTitulo
        Exit Sub
    End If
    
    
    Dim Rpta, A As Integer

    Rpta = MsgBox("Esta seguro de eliminar la retención seleccionada", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        'actualizamos el saldo de los documentos involucrados en la retencion
        For A = 1 To Fg1.Rows - 1
            If Opt1.Value = True Then
                'actualizasmos el saldo de los documentos de compra
                xCon.Execute "UPDATE com_compras SET com_compras.impsal = [com_compras]![impsal]+" & NulosN(Fg1.TextMatrix(A, 7)) & " " _
                    & " WHERE (((com_compras.id)=" & NulosN(Fg1.TextMatrix(A, 8)) & "))"
            Else
                'actualizasmos el saldo de los documentos de venta
                xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = [vta_ventas]![impsal]+" & NulosN(Fg1.TextMatrix(A, 7)) & " " _
                    & " WHERE (((vta_ventas.id)=" & NulosN(Fg1.TextMatrix(A, 8)) & "))"
            End If
        Next A
        
        'eliminamos el detalle de la retencion y su detalle
        xCon.Execute "DELETE * FROM con_retencion WHERE id = " & RstRet("id") & ""
        'xCon.Execute "DELETE * FROM con_retenciondet WHERE id = " & RstRet("id") & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstRet("id") & " AND idform = " & IdMenuActivo
        
        
        'elimiminamos el libro diario
        xCon.Execute "DELETE * FROM con_diario WHERE ((idmes = " & mMesActivo & ") and (idlib = 5) AND (idmov = " & RstRet("id") & "))"
        
        MsgBox "La Retencion se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        
        TabOne1.CurrTab = 0
        RstRet.Requery
        Dg1.Refresh
    End If
End Sub

Sub Blanquea()
    TxtRucPro.Text = ""
    TxtFchEmi.Valor = ""
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    TxtIdMoneda.Text = ""
    TxtIdRet.Text = ""
    TxtGlosa.Text = ""
    TxtIdDoc.Text = ""
    
    LblMoneda.Caption = ""
    LblProveedor.Caption = ""
    LblRetencion.Caption = ""
    LblDocumento.Caption = ""
    
    TxtImpRet.Text = ""
    TxtImpPag.Text = ""
    LblTipoCambio.Caption = ""
    LblTasa.Caption = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario miestras este agregando o editando una Retención", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub Opt1_Click()
    If Opt1.Value = True Then
        LblTitulo.Caption = "Proveedor"
        'LblTituloDoc.Caption = "Documentos de Compra"
        TxtRucPro.Text = ""
        LblProveedor.Caption = ""
    End If
End Sub

Private Sub Opt2_Click()
    If Opt2.Value = True Then
        LblTitulo.Caption = "Cliente"
        'LblTituloDoc.Caption = "Documentos de Venta"
        TxtRucPro.Text = ""
        LblProveedor.Caption = ""
    End If
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If RstRet.RecordCount = 0 And QueHace = 3 Then
            Cancel = 1
            Exit Sub
        End If
        If QueHace <> 1 Then MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstRet.Requery
            Dg1.Refresh
            '-----------------------------------
            If RstRet.RecordCount <> 0 Then
                RstRet.MoveFirst
                RstRet.Find "id=" & mIdRegistro
                If RstRet.EOF = True Then RstRet.MoveFirst
            End If
            '-----------------------------------
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 8 Then Filtrar
    
    If Button.Index = 9 Then
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstRet.Filter = ""
    End If
    If Button.Index = 10 Then CambiarMes
    If Button.Index = 11 Then Buscar
    
    If Button.Index = 13 Then pExportar
    If Button.Index = 14 Then
        If TabOne1.CurrTab = 1 Then IMPRIMIR 1
        If TabOne1.CurrTab = 0 Then IMPRIMIR 0
    End If
    
    If Button.Index = 16 Then
        Set RstRet = Nothing
        Unload Me
    End If
End Sub

Sub OpcionesPeriodo()
    
    '------------------------------------------------------------------------------------------
    '--bloqueamos los botones del toolbar
    CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
    '------------------------------------------------------------------------------------------
    
End Sub

Sub Filtrar()
    Dim xform As New eps_librerias.FormFiltrar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(4, 4) As String
   
    xCampos(0, 0) = "Tipo":               xCampos(0, 1) = "tipo2":         xCampos(0, 2) = "C":         xCampos(0, 3) = "4200"
    xCampos(1, 0) = "Cliente Proveedor":  xCampos(1, 1) = "nombre":        xCampos(1, 2) = "C":         xCampos(1, 3) = "1500"
    xCampos(2, 0) = "Fch. Emision":       xCampos(2, 1) = "fchemi":        xCampos(2, 2) = "F":         xCampos(2, 3) = "1500"
    xCampos(3, 0) = "Nº R.U.C.":          xCampos(3, 1) = "numruc":        xCampos(3, 2) = "C":         xCampos(3, 3) = "1500"
    
    Set xform.Coneccion = xCon
    Set xform.Rst = RstRet       'recorset que llena el grid
    Set RstRet = xform.FiltrarReg(xCampos)
    Set Dg1.DataSource = RstRet
    Dg1.Refresh
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index = 2 Then
        Dim xFchIni, xFchFin As String
        
        xFchIni = "01/" + Format(mMesActivo, "00") + "/" + Trim(AnoTra)
        xFchFin = Format(HallaDiasMes(CDate(xFchIni)), "00") + "/" + Format(mMesActivo, "00") + "/" + Trim(AnoTra)
        
        FrmRepRetenciones.TxtFchIni.Valor = xFchIni
        FrmRepRetenciones.TxtFchFin.Valor = xFchFin
        FrmRepRetenciones.Show
    End If
End Sub

Private Sub TxtFchEmi_Validate(Cancel As Boolean)
    If NulosC(TxtFchEmi.Valor) <> "" Then
        LblTipoCambio.Caption = Format(HallaTipoCambio(TxtFchEmi.Valor, 2, Venta, xCon), "0.000")
    End If
End Sub

Private Sub TxtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdDoc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusDoc_Click
    End If
End Sub

Private Sub TxtIdDoc_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If TxtIdDoc.Text = " " Then Exit Sub
    LblDocumento.Caption = Busca_Codigo(NulosN(TxtIdDoc.Text), "id", "descripcion", "mae_documento", "N", xCon)
    If LblDocumento.Caption = "" Then
        TxtIdDoc.Text = ""
    End If
End Sub

Private Sub TxtIdMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdMoneda_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 116 Then
'        CmdBusMon_Click
'    End If
End Sub

Private Sub TxtIdMoneda_Validate(Cancel As Boolean)
    If TxtIdMoneda.Text = "" Then Exit Sub
    LblMoneda.Caption = Busca_Codigo(TxtIdMoneda.Text, "id", "descripcion", "mae_moneda", "N", xCon)
    If LblMoneda.Caption = "" Then
        TxtIdMoneda.Text = ""
    End If
'    Else
'        If TxtIdMoneda.Text = "1" Then
'            LblTipoCambio.Caption = ""
'            'LblTipCam(5).Visible = False
'            'LblTipoCambio.Visible = False
'        Else
'            If TxtFchEmi.Valor = "" Then
'                TxtIdMoneda.Text = ""
'                LblTipoCambio.Caption = ""
'                Exit Sub
'            End If
'            LblTipoCambio.Caption = HallaTipoCambio(TxtFchEmi.Valor, NULOSN(TxtIdMoneda.Text), Venta, xCon)
'            'LblTipCam(5).Visible = True
'            'LblTipoCambio.Visible = True
'        End If
'    End If
End Sub

Private Sub TxtIdRet_Change()
    If NulosN(TxtIdRet.Text) = 0 Then LblRetencion.Caption = ""
End Sub

Private Sub TxtIdRet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdRet_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdIdRet_Click
    End If
End Sub

Private Sub TxtIdRet_Validate(Cancel As Boolean)
    If TxtIdRet.Text = "" Then Exit Sub
    LblRetencion.Caption = Busca_Codigo(TxtIdRet.Text, "id", "descripcion", "mae_retencion", "N", xCon)
    If LblRetencion.Caption = "" Then
        TxtIdRet.Text = ""
    Else
        LblTasa.Caption = Busca_Codigo(TxtIdRet.Text, "id", "tasa", "mae_retencion", "N", xCon)
        If Opt1.Value = True Then
            'CUENTA CONTABLE DE LA RETENCION CUANDO SEA COMPRA
            xIdCuenRet = Busca_Codigo(TxtIdRet.Text, "id", "idcuencom", "mae_retencion", "N", xCon)
        Else
            'CUENTA CONTABLE DE LA RETENCION CUANDO SEA VENTA
            xIdCuenRet = Busca_Codigo(TxtIdRet.Text, "id", "idcuenven", "mae_retencion", "N", xCon)
        End If
    End If
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumDoc_Validate(Cancel As Boolean)
    If TxtNumDoc.Text <> "" Then
        TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000000")
    End If
End Sub

Private Sub TxtNumSer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumSer_Validate(Cancel As Boolean)
    If TxtNumSer.Text <> "" Then
        TxtNumSer.Text = Format(TxtNumSer.Text, "0000")
    End If
End Sub

Private Sub TxtRucPro_Change()
    
    If NulosC(TxtRucPro.Text) = "" Then LblProveedor.Caption = ""
    
End Sub

Private Sub TxtRucPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        TxtRucPro_Validate False
    End If
End Sub

Private Sub TxtRucPro_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusProv_Click
    End If
End Sub

Private Sub TxtRucPro_Validate(Cancel As Boolean)
    If TxtRucPro.Text <> "" Then
        Dim Rst As New ADODB.Recordset
        If Opt1.Value = True Then
            RST_Busq Rst, "SELECT * FROM mae_prov WHERE numruc LIKE '" & Trim(TxtRucPro.Text) & "%'", xCon
            If Rst.RecordCount <> 0 Then
                Rst.MoveFirst
                TxtRucPro.Text = Rst("numruc")
                LblProveedor.Caption = Rst("nombre")
                LblIdProveedor.Caption = Rst("id")
            End If
        Else
            RST_Busq Rst, "SELECT * FROM mae_cliente WHERE numruc LIKE '" & Trim(TxtRucPro.Text) & "%'", xCon
            If Rst.RecordCount <> 0 Then
                Rst.MoveFirst
                TxtRucPro.Text = Rst("numruc")
                LblProveedor.Caption = Rst("nombre")
                LblIdProveedor.Caption = Rst("id")
            End If
        End If
        Set Rst = Nothing
    End If
End Sub


Private Sub pCargarGrid()
    On Error GoTo error
    Dim nSQL  As String
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    lblperiodo(1).Caption = LblMes.Caption
    
    OpcionesPeriodo
    
    TDB_FiltroLimpiar Dg1
    Set RstRet = Nothing
    
    DoEvents
    TabOne1.CurrTab = 0
    
    nSQL = "SELECT con_retencion.*, mae_prov.nombre, con_retencion!numser+'-'+con_retencion!numdoc AS numedoc, mae_moneda.simbolo, 'Compra' AS tipo2, mae_prov.numruc, Mid([con_retencion]![numreg],1,2)+[mae_libros].[codsun]+Mid([con_retencion]![numreg],3,4) AS registro, mae_documento.abrev AS docabrev,con_retencion.imp & '' as imp1 " _
        & " FROM (mae_moneda RIGHT JOIN (mae_retencion RIGHT JOIN ((con_retencion LEFT JOIN mae_prov ON con_retencion.idpro = mae_prov.id) LEFT JOIN mae_libros ON con_retencion.idlib = mae_libros.id) ON mae_retencion.id = con_retencion.idret) ON mae_moneda.id = con_retencion.idmon) LEFT JOIN mae_documento ON con_retencion.iddoc = mae_documento.id " _
        & " WHERE con_retencion.tipo=1 and year(con_retencion.fchreg)= " & AnoTra & " and month(con_retencion.fchreg)= " & mMesActivo & " ; " _
        & " UNION " _
        & " SELECT DISTINCT con_retencion.*, mae_cliente.nombre, con_retencion!numser+'-'+con_retencion!numdoc AS numedoc, mae_moneda.simbolo, 'Venta' AS tipo2, mae_cliente.numruc, Mid([con_retencion]![numreg],1,2)+[mae_libros].[codsun]+Mid([con_retencion]![numreg],3,4) AS registro, mae_documento.abrev as docabrev,con_retencion.imp & '' as imp1 " _
        & " FROM (mae_moneda RIGHT JOIN (mae_retencion RIGHT JOIN ((con_retencion LEFT JOIN mae_cliente ON con_retencion.idpro = mae_cliente.id) LEFT JOIN mae_libros ON con_retencion.idlib = mae_libros.id) ON mae_retencion.id = con_retencion.idret) ON mae_moneda.id = con_retencion.idmon) LEFT JOIN mae_documento ON con_retencion.iddoc = mae_documento.id " _
        & " WHERE con_retencion.tipo=2 and year(con_retencion.fchreg)=" & AnoTra & " and month(con_retencion.fchreg)=" & mMesActivo & " ; "

    Me.MousePointer = vbHourglass
    RST_Busq RstRet, nSQL, xCon

    Set Dg1.DataSource = RstRet
    Me.MousePointer = vbDefault
Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pCargarGrid"
End Sub

Private Sub pExportar()
    If TabOne1.CurrTab = 0 Then
        Dim oExport As New SGI2_funciones.formularios
        Dim RstTmp  As New ADODB.Recordset
        Dim xCampos(9, 3) As String

        '0::Nombre a Mostrar;
        '1::nombre de Campo del Rst;
        '2::alineacion(0::derecha, 1::centro, 2::izquierda);
        '3::ancho de columna
        '--obs: el rst puede tener mas columnas solo se consideran los campos del array
        xCampos(0, 0) = "Nº. Reg":              xCampos(0, 1) = "registro":   xCampos(0, 2) = 1:    xCampos(0, 3) = "900"
        xCampos(1, 0) = "Tipo":                 xCampos(1, 1) = "tipo2":     xCampos(1, 2) = 0:    xCampos(1, 3) = "743"
        xCampos(2, 0) = "Nº RUC":               xCampos(2, 1) = "numruc":     xCampos(2, 2) = 0:    xCampos(2, 3) = "1229"
        xCampos(3, 0) = "Cliente/Proveedor":    xCampos(3, 1) = "nombre":     xCampos(3, 2) = 0:    xCampos(3, 3) = "3500"
        xCampos(4, 0) = "T.D.":                 xCampos(4, 1) = "docabrev":   xCampos(4, 2) = 1:    xCampos(4, 3) = "443"
        xCampos(5, 0) = "Fch.Emi":              xCampos(5, 1) = "fchemi":     xCampos(5, 2) = 1:    xCampos(5, 3) = "1014"
        xCampos(6, 0) = "Nº.Documento":         xCampos(6, 1) = "numedoc":    xCampos(6, 2) = 0:    xCampos(6, 3) = "1700"
        xCampos(7, 0) = "Glosa":                xCampos(7, 1) = "glosa":      xCampos(7, 2) = 0:    xCampos(7, 3) = "3000"
        xCampos(8, 0) = "M":                    xCampos(8, 1) = "simbolo":   xCampos(8, 2) = 1:    xCampos(8, 3) = "386"
        xCampos(9, 0) = "Imp. Ret":              xCampos(9, 1) = "imp":        xCampos(9, 2) = 2:    xCampos(9, 3) = "943"
        
        Set RstTmp = RstRet.Clone
        oExport.VSFlexGrid_Exportar_MSExcel xCon, , "Registro de Retenciones", LblMes.Caption & " - " & AnoTra, "", "Registro de Retenciones", RstTmp, xCampos
        Set oExport = Nothing
        Set RstTmp = Nothing

        Exit Sub
    End If

'

    Dim A&, B&
    Dim xFilas&
    On Error GoTo error
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    'determina el numero de hojas que se mostrara en el Excel
    objExcel.SheetsInNewWorkbook = 1
    'abre el Libro
    objExcel.WindowState = 1
    objExcel.Workbooks.Add
    With objExcel.ActiveSheet

        .Cells(1, 2) = NomEmp
        .Cells(1, 9) = Date
        .Cells(2, 2) = "Nº R.U.C. : " + NumRUC
        
        .Columns(1).ColumnWidth = 1.2
        .Columns(2).ColumnWidth = Fg1.ColWidth(1) / 100
        .Columns(3).ColumnWidth = Fg1.ColWidth(2) / 100
        .Columns(4).ColumnWidth = Fg1.ColWidth(3) / 100
        .Columns(5).ColumnWidth = Fg1.ColWidth(4) / 100
        .Columns(6).ColumnWidth = Fg1.ColWidth(5) / 100
        .Columns(7).ColumnWidth = Fg1.ColWidth(6) / 100
        .Columns(8).ColumnWidth = Fg1.ColWidth(7) / 100
        .Columns(9).ColumnWidth = Fg1.ColWidth(8) / 100
        .Columns(10).ColumnWidth = Fg1.ColWidth(9) / 100

        '-----encabezado
        xFilas = 4
        .Cells(xFilas, 2) = "Registro de Retenciones"
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "Tipo Mov."
        .Cells(xFilas, 3) = IIf(Opt1.Value = True, "Compra", "Venta")
        .Cells(xFilas, 8) = "Periodo"
        .Cells(xFilas, 9) = LblMes.Caption
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = LblTitulo.Caption
        .Cells(xFilas, 3) = LblProveedor.Caption

        .Cells(xFilas, 8) = "RUC"
        .Cells(xFilas, 9) = "'" + TxtRucPro.Text
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "Fch.Emi."
        .Cells(xFilas, 3) = "'" & TxtFchEmi.Valor
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "Documento"
        .Cells(xFilas, 3) = LblDocumento.Caption
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "N°.Doc."
        .Cells(xFilas, 3) = "'" & TxtNumSer.Text & "-" & TxtNumDoc.Text

        .Cells(xFilas, 8) = "T.C."
        .Cells(xFilas, 9) = NulosN(LblTipoCambio.Caption)

        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "Moneda"
        .Cells(xFilas, 3) = "'" & LblMoneda.Caption
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "Retención"
        .Cells(xFilas, 3) = "'" & LblRetencion.Caption

        .Cells(xFilas, 2) = "Tasa"
        .Cells(xFilas, 3) = NulosN(LblTasa.Caption) & "%"
        '--titulo
        xFilas = xFilas + 2
        .Cells(xFilas, 2) = "Datos de la Operación"
        .Cells(xFilas, 7) = "Datos de Retención"
        xFilas = xFilas + 1
        For A = 1 To 9
            .Cells(xFilas, A + 1) = Fg1.TextMatrix(0, A)
        Next A
       '--detalle
        xFilas = xFilas + 1
        For A = 1 To Fg1.Rows - 1
            For B = 1 To 9
                If B < 5 Then
                    .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(A, B)
                Else
                    .Cells(xFilas, B + 1) = NulosN(Fg1.TextMatrix(A, B))
                End If
            Next B
            xFilas = xFilas + 1
        Next A

        .Cells(xFilas, 4) = "Total =>"
        
        .Cells(xFilas, 8) = NulosN(TxtImpPag.Text)
        .Cells(xFilas, 10) = NulosN(TxtImpRet.Text)

    End With

    MsgBox "El Registro se exportó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    objExcel.WindowState = 1
    objExcel.Visible = True
    Set objExcel = Nothing
    Exit Sub
error:
    Set objExcel = Nothing
    SHOW_ERROR Me.Name, "Exportar", , IIf(Err.Number <> 50290, "", "No manipule el archivo hasta que termine de exportar!!!!")
    
    
End Sub





Private Sub Buscar()
    On Error GoTo error
    Dim xRs As New ADODB.Recordset
    Dim RstTmp  As New ADODB.Recordset
    
    Dim xSQL As String
    ReDim xCampos(6, 4) As String
    
    xCampos(0, 0) = "Num.Reg.":             xCampos(0, 1) = "registro":  xCampos(0, 2) = "900":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Cliente / Proveedor":  xCampos(1, 1) = "nombre":    xCampos(1, 2) = "3200":  xCampos(1, 3) = "C"
    xCampos(2, 0) = "Fch. Doc.":            xCampos(2, 1) = "fchemi":    xCampos(2, 2) = "1000":  xCampos(2, 3) = "F"
    xCampos(3, 0) = "Nº Documento":         xCampos(3, 1) = "numedoc":   xCampos(3, 2) = "1500":  xCampos(3, 3) = "C"
    xCampos(4, 0) = "M":                    xCampos(4, 1) = "simbolo":   xCampos(4, 2) = "450":   xCampos(4, 3) = "C"
    xCampos(5, 0) = "Imp.Per.":             xCampos(5, 1) = "imp":       xCampos(5, 2) = "1000":  xCampos(5, 3) = "N"
    Set RstTmp = RstRet.Clone
    CARGAR_DLL_EPSBUSCAR xCon, xRs, "", xCampos(), "Buscando Retención", "registro", "registro", CualquierParte, , RstTmp
    
    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    
    RstRet.MoveFirst
    RstRet.Find "id = " & xRs("id") & ""
SALIR:
    Set RstTmp = Nothing
    Set xRs = Nothing
error:
    Set RstTmp = Nothing
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub

Sub CrearCabeceraVS(numPag As Integer)
    Dim xCad As String

    FrmVsPrinter.Vs.TextAlign = taLeftTop
    FrmVsPrinter.Vs.FontName = "Courier New"
    FrmVsPrinter.Vs.FontBold = True
    FrmVsPrinter.Vs.FontSize = 9

    FrmVsPrinter.Vs.CurrentX = 1000:      FrmVsPrinter.Vs.CurrentY = 1200
    FrmVsPrinter.Vs.Paragraph = "EMPRESA   : " & NomEmp

    FrmVsPrinter.Vs.CurrentX = 7950:      FrmVsPrinter.Vs.CurrentY = 1200
    FrmVsPrinter.Vs.Paragraph = "FECHA        : " & Format(Date, "dd/mm/yy")

    FrmVsPrinter.Vs.CurrentX = 1000:      FrmVsPrinter.Vs.CurrentY = 1400
    FrmVsPrinter.Vs.Paragraph = "Nº R.U.C. : " & NumRUC

    FrmVsPrinter.Vs.CurrentX = 7950:      FrmVsPrinter.Vs.CurrentY = 1400
    FrmVsPrinter.Vs.Paragraph = "Nº Pagina    : " & Format(numPag, "0000")

    FrmVsPrinter.Vs.DrawLine 1000, 1650, 11000, 1650
End Sub

Private Sub IMPRIMIR(Tipo As Integer)
    Dim xLinea As Integer
    Dim numeroPag As Integer
    Dim xform As New eps_librerias.FormSeleccion
    Dim xRs As New ADODB.Recordset
    Dim nSQLFiltro As String '--Almacenara el filtro por movimiento
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(8, 5) As String
    
    numeroPag = 1
    Select Case Tipo
        Case 0
            Dim nSQL As String
            
            xCampos(0, 0) = "Nº Registro":              xCampos(0, 1) = "registro":       xCampos(0, 2) = "900":      xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
            xCampos(1, 0) = "Nº R.U.C":                 xCampos(1, 1) = "numruc":       xCampos(1, 2) = "1200":     xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
            xCampos(2, 0) = "Cliente/Proveedor":        xCampos(2, 1) = "nombre":       xCampos(2, 2) = "3500":      xCampos(2, 3) = "C":    xCampos(2, 4) = "N"
            xCampos(3, 0) = "TD":                       xCampos(3, 1) = "docabrev":     xCampos(3, 2) = "500":     xCampos(3, 3) = "C":    xCampos(3, 4) = "N"
            xCampos(4, 0) = "Fech. Emi.":               xCampos(4, 1) = "fchemi":       xCampos(4, 2) = "1000":     xCampos(4, 3) = "D":    xCampos(4, 4) = "N"
            xCampos(5, 0) = "Nº Documento":             xCampos(5, 1) = "numedoc":      xCampos(5, 2) = "1500":     xCampos(5, 3) = "C":    xCampos(5, 4) = "N"
            xCampos(6, 0) = "M":                        xCampos(6, 1) = "simbolo":      xCampos(6, 2) = "500":     xCampos(6, 3) = "C":    xCampos(6, 4) = "N"
            xCampos(7, 0) = "Importe Ret.":             xCampos(7, 1) = "imp1":         xCampos(7, 2) = "1000":     xCampos(7, 3) = "N":    xCampos(7, 4) = "N"
                        
            'consulta para obtener listado de
            nSQL = "SELECT 0 as xsel, con_retencion.*, mae_prov.nombre, con_retencion!numser+'-'+con_retencion!numdoc AS numedoc, mae_moneda.simbolo, 'Compra' AS tipo2, mae_prov.numruc, Mid([con_retencion]![numreg],1,2)+[mae_libros].[codsun]+Mid([con_retencion]![numreg],3,4) AS registro, mae_documento.abrev AS docabrev,con_retencion.imp & '' as imp1 " _
                & " FROM (mae_moneda RIGHT JOIN (mae_retencion RIGHT JOIN ((con_retencion LEFT JOIN mae_prov ON con_retencion.idpro = mae_prov.id) LEFT JOIN mae_libros ON con_retencion.idlib = mae_libros.id) ON mae_retencion.id = con_retencion.idret) ON mae_moneda.id = con_retencion.idmon) LEFT JOIN mae_documento ON con_retencion.iddoc = mae_documento.id " _
                & " WHERE con_retencion.tipo=1 and year(con_retencion.fchreg)= " & AnoTra & " and month(con_retencion.fchreg)= " & mMesActivo & " ; " _
                & " UNION " _
                & " SELECT DISTINCT 0 as xsel, con_retencion.*, mae_cliente.nombre, con_retencion!numser+'-'+con_retencion!numdoc AS numedoc, mae_moneda.simbolo, 'Venta' AS tipo2, mae_cliente.numruc, Mid([con_retencion]![numreg],1,2)+[mae_libros].[codsun]+Mid([con_retencion]![numreg],3,4) AS registro, mae_documento.abrev as docabrev,con_retencion.imp & '' as imp1 " _
                & " FROM (mae_moneda RIGHT JOIN (mae_retencion RIGHT JOIN ((con_retencion LEFT JOIN mae_cliente ON con_retencion.idpro = mae_cliente.id) LEFT JOIN mae_libros ON con_retencion.idlib = mae_libros.id) ON mae_retencion.id = con_retencion.idret) ON mae_moneda.id = con_retencion.idmon) LEFT JOIN mae_documento ON con_retencion.iddoc = mae_documento.id " _
                & " WHERE con_retencion.tipo=2 and year(con_retencion.fchreg)=" & AnoTra & " and month(con_retencion.fchreg)=" & mMesActivo & " ; "
            
            xform.SqlCad = nSQL
                
            xform.Titulo = "Operaciones a Imprimir"
            Set xform.Coneccion = xCon
            Set xRs = Nothing
            Set xRs = xform.Seleccionar(xCampos)
            
            If xRs.State = 1 Then
                If xRs.RecordCount <> 0 Then
                    xRs.MoveFirst
                    With FrmVsPrinter.Vs
                        .StartDoc
                        .BrushColor = &H80000005
                        xLinea = 1700
                        Me.MousePointer = vbHourglass
                        Dim A As Integer
                        For A = 1 To xRs.RecordCount
                            .BrushColor = &H80000005
                            xLinea = 1700
                            CrearCabeceraVS numeroPag
                            llenarRetencion xRs, xLinea, numeroPag, FrmVsPrinter.Vs
                            xRs.MoveNext
                            If xRs.EOF Then Exit For
                            .NewPage
                            numeroPag = numeroPag + 1
                        Next A
                        Me.MousePointer = vbDefault
                        .EndDoc
                    End With
                End If
            Else
                Exit Sub
            End If
        Case 1
            With FrmVsPrinter.Vs
                .StartDoc
                .BrushColor = &H80000005
                xLinea = 1700
                
                CrearCabeceraVS numeroPag
                llenarRetencion RstRet, xLinea, numeroPag, FrmVsPrinter.Vs
                
                .EndDoc
            End With
    End Select
    
    FrmVsPrinter.Show
End Sub

Private Sub llenarRetencion(Rst As ADODB.Recordset, ByRef xLinea As Integer, ByRef numeroPag As Integer, ByRef Vs As VSPrinter)
    Dim A As Integer
    
    With FrmVsPrinter.Vs
        '-----Encabezado
        .FontSize = 15
        .TextAlign = taCenterMiddle
        
        If Rst("tipo") = 1 Then
            .TextBox "Registro de Retenciones - Compra", 1000, xLinea, 7500, 500, True, False, True
        Else
            .TextBox "Registro de Retenciones - Venta", 1000, xLinea, 7500, 500, True, False, True
        End If
            
        .FontSize = 10
        .TextBox "N° Registro", 8600, xLinea, 2375, 250, True, False, True
        .TextBox NulosC(Rst("registro")), 8600, xLinea + 250, 2375, 250, True, False, True
        
        '-----Descripcion
        .TextAlign = taLeftMiddle
        xLinea = xLinea + 500
        .TextBox "R.U.C :  ", 1000, xLinea, 2375, 250, True, False, False
        .TextBox NulosN(Rst("numruc")), 3000, xLinea, 2375, 250, True, False, False
        xLinea = xLinea + 250
        
        If Rst("tipo") = 1 Then
            .TextBox "Proveedor : ", 1000, xLinea, 2375, 250, True, False, False
        Else
            .TextBox "Cliente : ", 1000, xLinea, 2375, 250, True, False, False
        End If
        .TextBox NulosC(Rst("nombre")), 3000, xLinea, 2375, 250, True, False, False
        xLinea = xLinea + 250
        
        .TextBox "Fch.Doc : ", 1000, xLinea, 2375, 250, True, False, False
        .TextBox Rst("fchemi"), 3000, xLinea, 2375, 250, True, False, False
        
        Dim tipoC As Double
        If NulosN(Rst("idmon")) = "2" Then
            tipoC = HallaTipoCambio(Rst("fchemi"), NulosN(Rst("idmon")), Venta, xCon)
        Else
            tipoC = Format(HallaTipoCambio(TxtFchEmi.Valor, 2, Venta, xCon), "0.000")
        End If
        
        .TextBox "T.C. : ", 7000, xLinea, 2375, 250, True, False, False
        .TextBox NulosN(tipoC), 9000, xLinea, 5000, 250, True, False, False
        xLinea = xLinea + 250
        
        Dim documento As String
        documento = Busca_Codigo(Rst("iddoc"), "id", "descripcion", "mae_documento", "N", xCon)
        .TextBox "Documento : ", 1000, xLinea, 2375, 250, True, False, False
        .TextBox documento, 3000, xLinea, 5000, 250, True, False, False
        .TextBox "N° Documento : ", 7000, xLinea, 2375, 250, True, False, False
        .TextBox NulosC(Rst("numser")) & "-" & NulosC(Rst("numdoc")), 9000, xLinea, 5000, 250, True, False, False
        xLinea = xLinea + 250
        
        Dim moneda As String
        moneda = Busca_Codigo(NulosN(Rst("idmon")), "id", "descripcion", "mae_moneda", "N", xCon)
        .TextBox "Moneda : ", 1000, xLinea, 2375, 250, True, False, False
        .TextBox NulosC(moneda), 3000, xLinea, 5000, 250, True, False, False
        xLinea = xLinea + 250
        
        Dim tasa As String
        tasa = Format(Busca_Codigo(NulosN(Rst("idret")), "id", "tasa", "mae_retencion", "N", xCon), "0.00")
        .TextBox "Tasa : ", 1000, xLinea, 2375, 250, True, False, False
        .TextBox NulosC(tasa) & " %", 3000, xLinea, 5000, 250, True, False, False
        xLinea = xLinea + 250
        
        .TextBox "Glosa : ", 1000, xLinea, 2375, 250, True, False, False
        .TextBox NulosC(Rst("glosa")), 3000, xLinea, 9000, 250, True, False, False
        xLinea = xLinea + 500
        
        .FontSize = 9
        .TextAlign = taCenterMiddle
        
        '-----Contenido
        .TextBox "DATOS DE LA OPERACION", 1000, xLinea, 5600, 300, True, False, True
        .TextBox "DATOS DE LA RETENCION", 6600, xLinea, 4400, 300, True, False, True
        xLinea = xLinea + 300
        
        .TextBox "Nº Registro", 1000, xLinea, 1000, 600, True, False, True
        .TextBox "Nº Documento", 2000, xLinea, 1700, 600, True, False, True
        
        .TextBox "M", 3700, xLinea, 400, 600, True, False, True
        .TextBox "TD", 4100, xLinea, 400, 600, True, False, True
        .TextBox "Fech. Emision", 4500, xLinea, 1000, 600, True, False, True
        .TextBox "Imp. Total", 5500, xLinea, 1100, 600, True, False, True
        
        .TextBox "Imp. Cobrado", 6600, xLinea, 1100, 600, True, False, True
        .TextBox "Imp. Cob. MN", 7700, xLinea, 1100, 600, True, False, True
        .TextBox "Tasa Retenida", 8800, xLinea, 1100, 600, True, False, True
        .TextBox "Imp. Retenido", 9900, xLinea, 1100, 600, True, False, True
        
        xLinea = xLinea + 600
        .FontSize = 7
        .TextAlign = taLeftMiddle
        Dim RstAux As New ADODB.Recordset
        Dim cSQL As String
        
        If Opt1.Value = True Then
        
            cSQL = "SELECT con_retencion.id, con_retenciondet.iddoc, IIf(IsNull([com_compras]![numser])=-1,[com_compras]![numdoc],[com_compras]![numser]+'-'+[com_compras]![numdoc]) AS numdoc, mae_moneda.simbolo, mae_documento.abrev, com_compras.fchdoc, com_compras.imptot, con_retenciondet.impcob, con_retenciondet.impret, mae_documentocta.idcuen, com_compras.idmon, com_compras.tipdoc, mid(com_compras.numreg,1,2) & mae_libros.codsun & mid(com_compras.numreg,3,4) AS numeroReg " _
                + vbCr + "FROM (((con_retencion LEFT JOIN ((con_retenciondet LEFT JOIN com_compras ON con_retenciondet.iddoc = com_compras.id) LEFT JOIN mae_documento ON com_compras.tipdoc = mae_documento.id) ON con_retencion.id = con_retenciondet.id) LEFT JOIN mae_documentocta ON (com_compras.tipdoc = mae_documentocta.iddoc) AND (com_compras.idmon = mae_documentocta.idmon)) LEFT JOIN mae_moneda ON com_compras.idmon = mae_moneda.id) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id " _
                + vbCr + "WHERE (((con_retencion.id)=" & NulosN(Rst("id")) & ") AND ((mae_documentocta.tipope)=-1 Or (mae_documentocta.tipope) Is Null));"
        
'            cSQL = "SELECT con_retencion.id, con_retenciondet.iddoc, com_compras!numser+'-'+com_compras!numdoc AS numdoc, mae_documento.abrev, com_compras.fchdoc, com_compras.imptot, com_compras.numreg, con_retenciondet.impcob, con_retenciondet.impret, mae_documentocta.idmon, mae_documentocta.idcuen, mae_documentocta.tipope " _
'                + vbCr + "FROM (mae_documento RIGHT JOIN (con_retencion LEFT JOIN (con_retenciondet LEFT JOIN com_compras ON con_retenciondet.iddoc = com_compras.id) ON con_retencion.id = con_retenciondet.id) ON mae_documento.id = com_compras.tipdoc) LEFT JOIN mae_documentocta ON mae_documento.id = mae_documentocta.iddoc " _
'                + vbCr + "WHERE (((con_retencion.id)=" & NulosN(RstRet("id")) & ") AND ((mae_documentocta.idmon)=" & NulosN(TxtIdMoneda.Text) & ") AND ((mae_documentocta.tipope)=0));"
        Else
            cSQL = "SELECT con_retencion.id, con_retenciondet.iddoc, IIf(IsNull([vta_ventas]![numser])=-1,[vta_ventas]![numdoc],[vta_ventas]![numser]+'-'+[vta_ventas]![numdoc]) AS numdoc, mae_moneda.simbolo, mae_documento.abrev, vta_ventas.fchdoc, vta_ventas.imptotdoc AS imptot, con_retenciondet.impcob, con_retenciondet.impret, mae_documentocta.idcuen, vta_ventas.idmon, vta_ventas.tipdoc, mid(vta_ventas.numreg,1,2) & mae_libros.codsun & mid(vta_ventas.numreg,3,4) AS numeroReg " _
                + vbCr + "FROM (((con_retencion LEFT JOIN ((con_retenciondet LEFT JOIN vta_ventas ON con_retenciondet.iddoc = vta_ventas.id) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) ON con_retencion.id = con_retenciondet.id) LEFT JOIN mae_documentocta ON (vta_ventas.tipdoc = mae_documentocta.iddoc) AND (vta_ventas.idmon = mae_documentocta.idmon)) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
                + vbCr + "WHERE (((con_retencion.id)=" & NulosN(Rst("id")) & ") AND ((mae_documentocta.tipope)=-1 Or (mae_documentocta.tipope) Is Null));"
        End If
        RST_Busq RstAux, cSQL, xCon
        
        Dim imptotAux As Double
        Dim impcobAux As Double
        Dim impcobMNAux As Double
        Dim impretAux As Double
        
        imptotAux = 0
        impcobAux = 0
        impcobMNAux = 0
        impretAux = 0
        If Not RstAux.EOF Then
            RstAux.MoveFirst
            For A = 1 To RstAux.RecordCount
                .TextAlign = taLeftMiddle
                .TextBox " " & RstAux("numeroreg"), 1000, xLinea, 1200, 300, True, False, True
                .TextBox " " & NulosC(RstAux("numdoc")), 2000, xLinea, 1700, 300, True, False, True
                .TextAlign = taCenterMiddle
                .TextBox NulosC(RstAux("simbolo")), 3700, xLinea, 400, 300, True, False, True
                .TextBox NulosC(RstAux("abrev")), 4100, xLinea, 400, 300, True, False, True
                .TextBox NulosC(RstAux("fchdoc")), 4500, xLinea, 1000, 300, True, False, True
                
                .TextAlign = taRightMiddle
                .TextBox Format(RstAux("imptot"), FORMAT_MONTO) & " ", 5500, xLinea, 1100, 300, True, False, True
                imptotAux = imptotAux + NulosN(RstAux("imptot"))
                
                If RstAux("idmon") = 1 Then
                    .TextBox Format(NulosN(RstAux("impcob")), FORMAT_MONTO) & " ", 6600, xLinea, 1100, 300, True, False, True
                    impcobAux = impcobAux + NulosN(RstAux("impcob"))
                    .TextBox Format(NulosN(RstAux("impcob")), FORMAT_MONTO) & " ", 7700, xLinea, 1100, 300, True, False, True
                    impcobMNAux = impcobMNAux + NulosN(RstAux("impcob"))
                Else
                    .TextBox Format(NulosN(RstAux("impcob")), FORMAT_MONTO) & " ", 6600, xLinea, 1100, 300, True, False, True
                    impcobAux = impcobAux + NulosN(RstAux("impcob"))
                    .TextBox Format(NulosN(RstAux("impcob")) * NulosN(LblTipoCambio.Caption), FORMAT_MONTO) & " ", 7700, xLinea, 1100, 300, True, False, True
                    impcobMNAux = impcobMNAux + (NulosN(RstAux("impcob")) * NulosN(LblTipoCambio.Caption))
                End If
                
                .TextBox Format(LblTasa.Caption, "0.00") & " % ", 8800, xLinea, 1100, 300, True, False, True
                
                .TextBox Format(NulosN(RstAux("impret")), FORMAT_MONTO) & " ", 9900, xLinea, 1100, 300, True, False, True
                impretAux = impretAux + NulosN(RstAux("impret"))
                xLinea = xLinea + 300
                
                RstAux.MoveNext
                If RstAux.EOF Then Exit For
            Next A
            
            .FontSize = 9
            .TextAlign = taRightMiddle
            .TextBox "Total ", 1000, xLinea, 4500, 300, True, False, True
            
            .FontSize = 7
            .TextBox Format(imptotAux, FORMAT_MONTO) & " ", 5500, xLinea, 1100, 300, True, False, True
            .TextBox Format(impcobAux, FORMAT_MONTO) & " ", 6600, xLinea, 1100, 300, True, False, True
            .TextBox Format(impcobMNAux, FORMAT_MONTO) & " ", 7700, xLinea, 1100, 300, True, False, True
            .TextBox Format(impretAux, FORMAT_MONTO) & " ", 9900, xLinea, 1100, 300, True, False, True
            
            xLinea = xLinea + 600
            llenarAsiento Rst("id"), xLinea, numeroPag, FrmVsPrinter.Vs
        End If
    End With
End Sub

Private Sub llenarAsiento(idretencion As Integer, ByRef xLinea As Integer, ByRef numeroPag As Integer, ByRef Vs As VSPrinter)

    With Vs
        '-----Contenido
        If xLinea >= 15500 Then
            .NewPage
            numeroPag = numeroPag + 1
            CrearCabeceraVS numeroPag
            xLinea = 1700
        End If
        
        .TextAlign = taCenterMiddle
        .FontSize = 9
        
        .TextBox "ASIENTO CONTABLE", 1000, xLinea, 10000, 300, True, False, True
        xLinea = xLinea + 350
        
        .TextBox "EXPRESADO EN MN", 6600, xLinea, 2200, 300, True, False, True
        .TextBox "EXPRESADO EN ME", 8800, xLinea, 2200, 300, True, False, True
        
        .TextBox "N° Cuenta", 1000, xLinea, 800, 600, True, False, True
        .TextBox "Nombre de la Cuenta", 1800, xLinea, 2500, 600, True, False, True
        
        .TextBox "T.D.", 4300, xLinea, 500, 600, True, False, True
        .TextBox "Nº Documento ", 4800, xLinea, 1300, 600, True, False, True
        .TextBox "T.C.", 6100, xLinea, 500, 600, True, False, True
        
        xLinea = xLinea + 300
        .TextBox "Debe", 6600, xLinea, 1100, 300, True, False, True
        .TextBox "Haber", 7700, xLinea, 1100, 300, True, False, True
        .TextBox "Debe", 8800, xLinea, 1100, 300, True, False, True
        .TextBox "Haber", 9900, xLinea, 1100, 300, True, False, True
        
    
        xLinea = xLinea + 300
        .FontSize = 7
        .TextAlign = taLeftMiddle
        Dim RstDet As New ADODB.Recordset
        Dim cSQL As String
        
        cSQL = "SELECT * " _
        + vbCr + "FROM [SELECT con_retencion.numreg, con_planctas.cuenta, con_planctas.descripcion, con_tc.impven AS tc, IIf(con_retencion.idmon=1,con_retencion.imp,IIf(con_tc.impven Is Null,0,con_retencion.imp*con_tc.impven)) AS debe, 0 AS haber, mae_documento.abrev, mae_moneda.simbolo, IIf(con_retencion.numser Is Null Or con_retencion.numser='','',con_retencion.numser & '-' & con_retencion.numdoc) AS numerodoc " _
        + vbCr + "FROM mae_moneda INNER JOIN ((mae_retencion LEFT JOIN con_planctas ON mae_retencion.idcuenven = con_planctas.id) RIGHT JOIN (((con_retencion LEFT JOIN con_tc ON con_retencion.fchemi = con_tc.fecha) LEFT JOIN mae_libros ON con_retencion.idlib = mae_libros.id) LEFT JOIN mae_documento ON con_retencion.iddoc = mae_documento.id) ON mae_retencion.id = con_retencion.idret) ON mae_moneda.id = con_retencion.idmon " _
        + vbCr + "Where (((con_retencion.tipo) = 2) And ((con_retencion.id) = " & idretencion & ")) " _
        + vbCr + "Union " _
        + vbCr + "SELECT vta_ventas.numreg, con_planctas.cuenta, con_planctas.descripcion, con_tc.impven AS tc, Abs(IIf(vta_ventas.tipdoc<>7,0,IIf(con_retencion.idmon=1,con_retenciondet.impret,IIf(con_tc.impven Is Null,0,con_retenciondet.impret*con_tc.impven)))) AS debe, IIf(vta_ventas.tipdoc=7,0,IIf(con_retencion.idmon=1,con_retenciondet.impret,IIf(con_tc.impven Is Null,0,con_retenciondet.impret*con_tc.impven))) AS haber, mae_documento.abrev, mae_moneda.simbolo, IIf(vta_ventas.numser Is Null Or vta_ventas.numser='','',vta_ventas.numser & '-' & vta_ventas.numdoc) AS numerodoc " _
        + vbCr + "FROM (mae_moneda INNER JOIN ((con_retencion LEFT JOIN con_tc ON con_retencion.fchemi = con_tc.fecha) INNER JOIN (((con_retenciondet INNER JOIN vta_ventas ON con_retenciondet.iddoc = vta_ventas.id) INNER JOIN (con_planctas INNER JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON (vta_ventas.tipdoc = mae_documentocta.iddoc) AND (vta_ventas.idmon = mae_documentocta.idmon)) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) ON con_retencion.id = con_retenciondet.id) ON mae_moneda.id = con_retencion.idmon) LEFT JOIN mae_documento ON mae_documentocta.iddoc = mae_documento.id " _
        + vbCr + "WHERE (((con_retencion.tipo)=2) AND ((mae_documentocta.tipope)=-1) AND ((con_retencion.id)=" & idretencion & "))]. AS vista;"
            
        RST_Busq RstDet, cSQL, xCon
        
        Agregando = True
        Dim sumaDeb As Double
        Dim sumaHab As Double
        Dim tipoC As Double
        tipoC = RstDet("tc")
        If RstDet.RecordCount <> 0 Then
            RstDet.MoveFirst
            sumaDeb = 0
            sumaHab = 0
            Dim B As Integer
            For B = 1 To RstDet.RecordCount
                If xLinea >= 15500 Then
                    .FontSize = 9
                    .TextAlign = taRightMiddle
                    .TextBox "VAN ", 1000, xLinea, 5600, 300, True, False, False
                    .FontSize = 7
                    'van debe MN
                    .TextBox Format(NulosN(sumaDeb), FORMAT_MONTO), 6600, xLinea, 1100, 300, True, False, True
                    'van haber MN
                    .TextBox Format(NulosN(sumaHab), FORMAT_MONTO), 7700, xLinea, 1100, 300, True, False, True
                    'van debe ME
                    .TextBox Format(NulosN(sumaDeb) / NulosN(RstDet("tc")), FORMAT_MONTO), 8800, xLinea, 1100, 300, True, False, True
                    'van haber ME
                    .TextBox Format(NulosN(sumaHab) / NulosN(RstDet("tc")), FORMAT_MONTO), 9900, xLinea, 1100, 300, True, False, True
                
                    .NewPage
                    numeroPag = numeroPag + 1
                    CrearCabeceraVS numeroPag
                    xLinea = 1700
                    
                    .FontSize = 9
                    .TextAlign = taRightMiddle
                    .TextBox "VIENEN ", 1000, xLinea, 5600, 300, True, False, False
                    .FontSize = 7
                    'vienen debe MN
                    .TextBox Format(NulosN(sumaDeb), FORMAT_MONTO), 6600, xLinea, 1100, 300, True, False, True
                    'vienen haber MN
                    .TextBox Format(NulosN(sumaHab), FORMAT_MONTO), 7700, xLinea, 1100, 300, True, False, True
                    'vienen debe ME
                    .TextBox Format(NulosN(sumaDeb) / NulosN(RstDet("tc")), FORMAT_MONTO), 8800, xLinea, 1100, 300, True, False, True
                    'vienen haber ME
                    .TextBox Format(NulosN(sumaHab) / NulosN(RstDet("tc")), FORMAT_MONTO), 9900, xLinea, 1100, 300, True, False, True
                    
                    xLinea = xLinea + 300
                End If
            
                .TextAlign = taLeftMiddle
                .TextBox " " & NulosC(RstDet("cuenta")), 1000, xLinea, 800, 300, True, False, True
                .TextBox " " & NulosC(RstDet("descripcion")), 1800, xLinea, 2500, 300, True, False, True
                
                .TextBox " " & NulosC(RstDet("abrev")), 4300, xLinea, 500, 300, True, False, True
                .FontSize = 6
                .TextBox " " & NulosC(RstDet("numerodoc")), 4800, xLinea, 1300, 300, True, False, True
                .FontSize = 7
                .TextBox " " & NulosC(RstDet("tc")), 6100, xLinea, 500, 300, True, False, True
    '
                .TextAlign = taRightMiddle
                'deben MN
                .TextBox Format(NulosN(RstDet("debe")), FORMAT_MONTO) & " ", 6600, xLinea, 1100, 300, True, False, True
                sumaDeb = sumaDeb + NulosN(RstDet("debe"))
                'haber MN
                .TextBox Format(NulosN(RstDet("haber")), FORMAT_MONTO) & " ", 7700, xLinea, 1100, 300, True, False, True
                sumaHab = sumaHab + NulosN(RstDet("haber"))
                
                'deben ME
                .TextBox Format(NulosN(RstDet("debe")) / NulosN(RstDet("tc")), FORMAT_MONTO) & " ", 8800, xLinea, 1100, 300, True, False, True
                'haber ME
                .TextBox Format(NulosN(RstDet("haber")) / NulosN(RstDet("tc")), FORMAT_MONTO) & " ", 9900, xLinea, 1100, 300, True, False, True
                xLinea = xLinea + 300
            
                RstDet.MoveNext
                If RstDet.EOF = True Then Exit For
            Next B
        End If
            
        .FontSize = 9
        .TextAlign = taRightMiddle
        .TextBox "Total ", 1000, xLinea, 5600, 300, True, False, True
    '
        .FontSize = 7
        .TextBox Format(NulosN(sumaDeb), FORMAT_MONTO) & " ", 6600, xLinea, 1100, 300, True, False, True
        .TextBox Format(NulosN(sumaHab), FORMAT_MONTO) & " ", 7700, xLinea, 1100, 300, True, False, True
        
        .TextBox Format(NulosN(sumaDeb) / NulosN(tipoC), FORMAT_MONTO) & " ", 8800, xLinea, 1100, 300, True, False, True
        .TextBox Format(NulosN(sumaHab) / NulosN(tipoC), FORMAT_MONTO) & " ", 9900, xLinea, 1100, 300, True, False, True
    End With
End Sub

VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SIZERONE.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCajaBancos 
   Caption         =   "Caja y Bancos - Egreso de Caja y Bancos"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7605
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
            Picture         =   "FrmCajaBancos.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCajaBancos.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCajaBancos.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCajaBancos.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCajaBancos.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCajaBancos.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCajaBancos.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCajaBancos.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCajaBancos.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCajaBancos.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCajaBancos.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCajaBancos.frx":2236
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7230
      Left            =   0
      TabIndex        =   10
      Top             =   375
      Width           =   11880
      _cx             =   20955
      _cy             =   12753
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
         Height          =   6810
         Left            =   -12435
         TabIndex        =   25
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6420
            Left            =   30
            TabIndex        =   26
            Top             =   345
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11324
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "T.D."
            Columns(0).DataField=   "abrev"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Moneda"
            Columns(1).DataField=   "simbolo"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nº Documento"
            Columns(2).DataField=   "numerodoc"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Fch. Emi"
            Columns(3).DataField=   "fchdoc"
            Columns(3).NumberFormat=   "Short Date"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Cliente"
            Columns(4).DataField=   "nombre"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Forma Pago"
            Columns(5).DataField=   "desccond"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Importe"
            Columns(6).DataField=   "imptot"
            Columns(6).NumberFormat=   "0.00"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Fch. Ven."
            Columns(7).DataField=   "fchven"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Saldo"
            Columns(8).DataField=   "impsal"
            Columns(8).NumberFormat=   "0.00"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   9
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=9"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=900"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=820"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=1402"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1323"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2566"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2487"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=1773"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1693"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=5398"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=5318"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=2672"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2593"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=1640"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1561"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=514"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(43)=   "Column(7).Width=1773"
            Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=1693"
            Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=516"
            Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(49)=   "Column(8).Width=1667"
            Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=1588"
            Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=514"
            Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=67,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=68,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=69,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=0"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.alignment=1"
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
            AutoSize        =   -1  'True
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
            Left            =   8235
            TabIndex        =   29
            Top             =   30
            Width           =   765
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta Compras"
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
            TabIndex        =   28
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
            TabIndex        =   27
            Top             =   0
            Width           =   1860
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6810
         Left            =   45
         TabIndex        =   11
         Top             =   375
         Width           =   11790
         Begin VB.OptionButton OptBanco 
            Caption         =   "Banco"
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
            Left            =   3045
            TabIndex        =   2
            Top             =   885
            Width           =   1170
         End
         Begin VB.OptionButton OptCaja 
            Caption         =   "Caja"
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
            Left            =   1815
            TabIndex        =   1
            Top             =   885
            Width           =   1170
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   360
            Left            =   3750
            TabIndex        =   43
            Top             =   2100
            Visible         =   0   'False
            Width           =   8295
            Begin VB.CommandButton CmdBusMedPag 
               Height          =   240
               Left            =   2190
               Picture         =   "FrmCajaBancos.frx":277E
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   60
               Width           =   240
            End
            Begin VB.TextBox TxtIdMedioPago 
               Height          =   300
               Left            =   1545
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   45
               Text            =   "TxtIdMedioPago"
               Top             =   30
               Width           =   915
            End
            Begin VB.Label LblDesMedPag 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblDesMedPag"
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
               Left            =   2505
               TabIndex        =   47
               Top             =   30
               Width           =   5760
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Medio Pago"
               Height          =   195
               Index           =   6
               Left            =   0
               TabIndex        =   46
               Top             =   60
               Width           =   855
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   360
            Left            =   255
            TabIndex        =   38
            Top             =   1170
            Visible         =   0   'False
            Width           =   8295
            Begin VB.CommandButton CmdBusCueBan 
               Height          =   240
               Left            =   3255
               Picture         =   "FrmCajaBancos.frx":28B0
               Style           =   1  'Graphical
               TabIndex        =   40
               Top             =   60
               Width           =   240
            End
            Begin VB.TextBox TxtNumCue 
               Height          =   300
               Left            =   1545
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   39
               Text            =   "TxtNumCue"
               Top             =   30
               Width           =   1980
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Nº Cuenta"
               Height          =   195
               Index           =   1
               Left            =   0
               TabIndex        =   42
               Top             =   60
               Width           =   735
            End
            Begin VB.Label LblBanco 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblBanco"
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
               Left            =   3570
               TabIndex        =   41
               Top             =   30
               Width           =   4695
            End
         End
         Begin VB.CommandButton CmdAgregar 
            Caption         =   "&Agregar"
            Height          =   570
            Left            =   9330
            TabIndex        =   35
            Top             =   3150
            Width           =   1380
         End
         Begin VB.CommandButton CmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   570
            Left            =   9270
            TabIndex        =   34
            Top             =   5505
            Width           =   1380
         End
         Begin VB.CommandButton CmdBusPro 
            Height          =   240
            Left            =   6165
            Picture         =   "FrmCajaBancos.frx":29E2
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   2550
            Width           =   240
         End
         Begin VB.TextBox TxtProv 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   7
            Text            =   "TxtProv"
            Top             =   2520
            Width           =   4635
         End
         Begin VB.CommandButton CmdNumDoc 
            Height          =   240
            Left            =   3525
            Picture         =   "FrmCajaBancos.frx":2B14
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   1890
            Width           =   240
         End
         Begin VB.TextBox TxtImporte 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   6
            Text            =   "TxtImporte"
            Top             =   2190
            Width           =   1125
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   5
            Text            =   "TxtNumDoc"
            Top             =   1860
            Width           =   1995
         End
         Begin VB.CommandButton CmdBusMon 
            Height          =   240
            Left            =   8250
            Picture         =   "FrmCajaBancos.frx":2C46
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   900
            Width           =   240
         End
         Begin VB.Frame Frame5 
            Caption         =   "( Periodo )"
            Height          =   750
            Left            =   9405
            TabIndex        =   12
            Top             =   1320
            Width           =   2115
            Begin VB.Label LblPeriodo2 
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
               Left            =   120
               TabIndex        =   13
               Top             =   255
               Width           =   1860
            End
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchMov 
            Height          =   300
            Left            =   1800
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
         End
         Begin VB.TextBox TxtIdMon 
            Height          =   300
            Left            =   7605
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   3
            Text            =   "TxtIdMon"
            Top             =   870
            Width           =   915
         End
         Begin VB.CommandButton CmdBusMovi 
            Height          =   240
            Left            =   2445
            Picture         =   "FrmCajaBancos.frx":2D78
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   1560
            Width           =   240
         End
         Begin VB.TextBox TxtIdMov 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   4
            Text            =   "TxtIdMov"
            Top             =   1530
            Width           =   915
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   1470
            Left            =   255
            TabIndex        =   8
            Top             =   3150
            Width           =   8985
            _cx             =   15849
            _cy             =   2593
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
            Rows            =   50
            Cols            =   9
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmCajaBancos.frx":2EAA
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
         Begin VSFlex7Ctl.VSFlexGrid Fg2 
            Height          =   1470
            Left            =   255
            TabIndex        =   9
            Top             =   4935
            Width           =   8985
            _cx             =   15849
            _cy             =   2593
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
            Rows            =   50
            Cols            =   11
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmCajaBancos.frx":2FAD
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
         Begin VB.Label aaaaaa 
            Caption         =   "aaaaaa"
            Height          =   240
            Left            =   10470
            TabIndex        =   51
            Top             =   585
            Width           =   990
         End
         Begin VB.Label xxxxxx 
            Caption         =   "xxxxxx"
            Height          =   240
            Left            =   10470
            TabIndex        =   50
            Top             =   315
            Width           =   990
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Haber"
            Height          =   240
            Index           =   1
            Left            =   9375
            TabIndex        =   49
            Top             =   585
            Width           =   990
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Debe"
            Height          =   240
            Index           =   0
            Left            =   9375
            TabIndex        =   48
            Top             =   315
            Width           =   990
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Documentos Pendientes"
            Height          =   195
            Index           =   9
            Left            =   255
            TabIndex        =   37
            Top             =   2895
            Width           =   1740
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Documentos x Pagar"
            Height          =   195
            Index           =   2
            Left            =   255
            TabIndex        =   36
            Top             =   4680
            Width           =   1485
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            Height          =   195
            Index           =   3
            Left            =   255
            TabIndex        =   32
            Top             =   2550
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   5
            Left            =   6870
            TabIndex        =   24
            Top             =   915
            Width           =   585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Importe"
            Height          =   195
            Index           =   4
            Left            =   255
            TabIndex        =   23
            Top             =   2235
            Width           =   525
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Documento"
            Height          =   195
            Index           =   0
            Left            =   255
            TabIndex        =   22
            Top             =   1890
            Width           =   1050
         End
         Begin VB.Label LblDescMov 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDescMov"
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
            Top             =   1530
            Width           =   5760
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Movimiento"
            Height          =   195
            Index           =   7
            Left            =   255
            TabIndex        =   20
            Top             =   1545
            Width           =   810
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Operacion"
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
            TabIndex        =   19
            Top             =   30
            Width           =   11610
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
            Left            =   8565
            TabIndex        =   18
            Top             =   870
            Width           =   2115
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Emision"
            Height          =   195
            Index           =   8
            Left            =   255
            TabIndex        =   17
            Top             =   570
            Width           =   1260
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Operacion"
            Height          =   195
            Index           =   10
            Left            =   255
            TabIndex        =   16
            Top             =   870
            Width           =   1095
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   30
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
Attribute VB_Name = "FrmCajaBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim xCuentaDebe As Integer
Dim xCuentaHaber As Integer

'Private Sub CmdBusTipMov_Click()
'    If QueHace = 3 Then Exit Sub
'
'    Dim xForm As New EPS_Buscar.Buscar
'    Dim xRs As New ADODB.Recordset
'    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
'
'    Dim xCampos(3, 4) As String
'
'    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "desban":        xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
'    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "numcue":        xCampos(1, 2) = "1000":         xCampos(1, 3) = "C"
'    xCampos(2, 0) = "Moneda":     xCampos(2, 1) = "desmon":        xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
'
'    xForm.SQLCad = "SELECT mae_bancos.descripcion AS desban, con_bancocuenta.*, mae_moneda.descripcion AS desmon" _
'        & " FROM (mae_bancos INNER JOIN con_bancocuenta ON mae_bancos.id = con_bancocuenta.idban) " _
'        & " LEFT JOIN mae_moneda ON con_bancocuenta.idmon = mae_moneda.id ORDER BY mae_bancos.descripcion"
'
'    xForm.Titulo = "Buscando Cuentas de Banco"
'    xForm.FormaBusca = Principio
'    xForm.Criterio = ""
'    xForm.Ordenado = "desban"
'    xForm.CampoBusca = "desban"
'    Set xForm.Coneccion = xCon
'    Set xRs = xForm.BuscarReg(xCampos)
'    If xRs.State = 1 Then
'        TxtIdCuenta.Text = xRs("id")
'        LblCuentaBan.Caption = Trim(xRs("desban")) + "   Cuenta Nº " & xRs("numcue")
'        TxtTipMov.SetFocus
'    End If
'    Set xForm = Nothing
'    Set xRs = Nothing
'End Sub

Private Sub CmdBusCueBan_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New EPS_Buscar.Buscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(3, 4) As String
    
    xCampos(0, 0) = "Banco":      xCampos(0, 1) = "desban":        xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº Cuenta":  xCampos(1, 1) = "numcue":        xCampos(1, 2) = "1200":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Moneda":     xCampos(2, 1) = "desmon":        xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
    
    xForm.SQLCad = "SELECT mae_bancos.descripcion AS desban, con_bancocuenta.*, mae_moneda.descripcion AS desmon, " _
        & " con_planctas.cuenta FROM con_planctas RIGHT JOIN (mae_bancos INNER JOIN (con_bancocuenta LEFT JOIN " _
        & " mae_moneda ON con_bancocuenta.idmon = mae_moneda.id) ON mae_bancos.id = con_bancocuenta.idban) ON " _
        & " con_planctas.id = con_bancocuenta.idcuen ORDER BY mae_bancos.descripcion"
    
    xForm.Titulo = "Buscando Cuentas de Banco"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "desban"
    xForm.CampoBusca = "desban"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtNumCue.Text = xRs("id")
        LblBanco.Caption = Trim(xRs("desban")) + "   Cuenta Nº " & xRs("numcue")
        xCuentaHaber = xRs("idcuen")
        aaaaaa.Caption = xRs("cuenta")
        TxtIdMov.SetFocus
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMedPag_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New EPS_Buscar.Buscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Codigo":        xCampos(0, 1) = "id":            xCampos(0, 2) = "1000":         xCampos(0, 3) = "N"
    xCampos(1, 0) = "Descripcion":   xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Cuenta":        xCampos(2, 1) = "descta":        xCampos(2, 2) = "3000":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Nº Cuenta":     xCampos(3, 1) = "cuenta":        xCampos(3, 2) = "1000":         xCampos(3, 3) = "C"
    
    xForm.SQLCad = "SELECT con_cajabanmedpag.*, con_planctas.descripcion AS descta, con_planctas.cuenta " _
        & " FROM con_planctas RIGHT JOIN con_cajabanmedpag ON con_planctas.id = con_cajabanmedpag.idcuenta"
    
    xForm.Titulo = "Buscando Medios de Pago"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "id"
    xForm.CampoBusca = "id"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        LblDesMedPag.Caption = xRs("descripcion")
        TxtIdMedioPago.Text = xRs("id")
        xCuentaHaber = xRs("idcuenta")
        aaaaaa.Caption = xRs("cuenta")
        
        TxtIdMov.SetFocus
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub Command1_Click()

End Sub

Private Sub CmdBusMovi_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New EPS_Buscar.Buscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(3, 4) As String
    
    xCampos(0, 0) = "Codigo":       xCampos(0, 1) = "id":            xCampos(0, 2) = "1200":         xCampos(0, 3) = "N"
    xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Cuenta":       xCampos(2, 1) = "descuen":       xCampos(2, 2) = "3000":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Nº Cuenta":    xCampos(3, 1) = "cuenta":        xCampos(3, 2) = "1200":         xCampos(3, 3) = "C"
    
    'filtramos por tipo de operacion  = 1 (caja)
    'filtramos por tipo de movimiento  = 2 (salida)
    If OptCaja.Value = True Then
        xForm.SQLCad = "SELECT con_planctas.cuenta, con_planctas.descripcion AS descuen, con_cajabanmovi.id, " _
            & " con_cajabanmovi.descripcion, con_cajabanmovi.idcue, con_cajabanmovi.tipmov FROM con_planctas INNER JOIN " _
            & " con_cajabanmovi ON con_planctas.id = con_cajabanmovi.idcue WHERE (con_cajabanmovi.tipmov = 2)"
    Else
        xForm.SQLCad = "SELECT con_cajabanmovi.*, con_planctas.cuenta, con_planctas.descripcion AS descuen " _
            & " FROM con_cajabanmovi LEFT JOIN con_planctas ON con_cajabanmovi.idcue = con_planctas.id " _
            & " WHERE (((con_cajabanmovi.tipope)=2) AND ((con_cajabanmovi.tipmov)=2))"
    End If

    xForm.Titulo = "Buscando Movimiento"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "id"
    xForm.CampoBusca = "id"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdMov.Text = xRs("id")
        LblDescMov.Caption = xRs("descripcion")
        xxxxxx.Caption = xRs("cuenta")
        TxtNumDoc.SetFocus
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
    End If
End Sub

Private Sub Form_Load()
    QueHace = 3
    SeEjecuto = False
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Frame3.BackColor = &H8000000F
    Frame4.BackColor = &H8000000F
End Sub

Private Sub OptBanco_Click()
    Frame4.Visible = False
    Frame3.Visible = True
    Frame3.Left = 255
    Frame3.Top = 1170
End Sub

Private Sub OptCaja_Click()
    Frame3.Visible = False
    Frame4.Visible = True
    Frame4.Left = 255
    Frame4.Top = 1170
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
End Sub

Sub Nuevo()
    QueHace = 1
    TabOne1.CurrTab = 1
    
    Label5.Caption = "Agregando Operacion"
    Blanquea
    Bloquea
    OptCaja.Value = True
    OptCaja_Click
    TxtFchMov.SetFocus
End Sub

Sub Bloquea()
    TxtFchMov.Locked = Not TxtFchMov.Locked
    TxtNumCue.Locked = Not TxtNumCue.Locked
    TxtIdMon.Locked = Not TxtIdMon.Locked
    TxtIdMov.Locked = Not TxtIdMov.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    TxtImporte.Locked = Not TxtImporte.Locked
    TxtIdMedioPago.Locked = Not TxtIdMedioPago.Locked
    TxtProv.Locked = Not TxtProv.Locked
End Sub

Sub Blanquea()
    TxtFchMov.Valor = ""
    TxtIdMon.Text = ""
    TxtNumCue.Text = ""
    TxtIdMedioPago.Text = ""
    TxtProv.Text = ""
    TxtImporte.Text = ""
    TxtIdMov.Text = ""
    TxtNumDoc.Text = ""
    
    LblBanco.Caption = ""
    LblDescMov.Caption = ""
    LblDesMedPag.Caption = ""
    LblMoneda.Caption = ""
End Sub


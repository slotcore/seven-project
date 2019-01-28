VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SIZERONE.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "aspatextboxfecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmIngCajBan 
   Caption         =   "Caja y Bancos - Ingreso"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   11865
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
            Picture         =   "FrmIngCajBan.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngCajBan.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngCajBan.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngCajBan.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngCajBan.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngCajBan.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngCajBan.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngCajBan.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngCajBan.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngCajBan.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngCajBan.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngCajBan.frx":2236
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7260
      Left            =   0
      TabIndex        =   9
      Top             =   375
      Width           =   11880
      _cx             =   20955
      _cy             =   12806
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
         Height          =   6840
         Left            =   -12435
         TabIndex        =   10
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6420
            Left            =   30
            TabIndex        =   11
            Top             =   345
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11324
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nº Reg."
            Columns(0).DataField=   "numreg"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "T. M."
            Columns(1).DataField=   "motmov"
            Columns(1).NumberFormat=   "0.00"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Mov."
            Columns(2).DataField=   "fchope"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "M"
            Columns(3).DataField=   "simbolo"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Origen"
            Columns(4).DataField=   "descori"
            Columns(4).NumberFormat=   "Short Date"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "T.D."
            Columns(5).DataField=   "abredoc"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Nº Documento"
            Columns(6).DataField=   "numdoc"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Banco"
            Columns(7).DataField=   "descban"
            Columns(7).NumberFormat=   "0.00"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Nº Cuenta"
            Columns(8).DataField=   "numcue"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Importe"
            Columns(9).DataField=   "importe"
            Columns(9).NumberFormat=   "0.00"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   10
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=10"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1429"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1349"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=1191"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1111"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1693"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1614"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=556"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=476"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=512"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=4471"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=4392"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=1005"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=926"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=2646"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2566"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=512"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(43)=   "Column(7).Width=2910"
            Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=2831"
            Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=512"
            Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(49)=   "Column(8).Width=2408"
            Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=2328"
            Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=516"
            Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(55)=   "Column(9).Width=1588"
            Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=1508"
            Splits(0)._ColumnProps(58)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(59)=   "Column(9)._ColStyle=514"
            Splits(0)._ColumnProps(60)=   "Column(9).Order=10"
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=0"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=74,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=71,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=72,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=73,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=0"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=62,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=28,.parent=13,.alignment=1"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=25,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=26,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=27,.parent=17"
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
         Begin VB.Label LblMes 
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
            TabIndex        =   13
            Top             =   0
            Width           =   1860
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta Ingresos"
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
            Left            =   90
            TabIndex        =   12
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
         Caption         =   "LblidDocumento"
         Height          =   6840
         Left            =   45
         TabIndex        =   14
         Top             =   375
         Width           =   11790
         Begin VB.Frame Frame9 
            Height          =   1095
            Left            =   8460
            TabIndex        =   77
            Top             =   1425
            Width           =   3240
            Begin VB.CommandButton CmdAddCon 
               Caption         =   "&Agregar Destino"
               Height          =   315
               Left            =   645
               TabIndex        =   79
               Top             =   270
               Width           =   1860
            End
            Begin VB.CommandButton CmdDelCon 
               Caption         =   "Eliminar Destino"
               Height          =   315
               Left            =   645
               TabIndex        =   78
               Top             =   600
               Width           =   1860
            End
         End
         Begin VB.CommandButton CmdBusMedPag 
            Height          =   240
            Left            =   2310
            Picture         =   "FrmIngCajBan.frx":277E
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   2850
            Width           =   240
         End
         Begin VB.CommandButton CmdNumDoc 
            Height          =   240
            Left            =   10080
            Picture         =   "FrmIngCajBan.frx":28B0
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   3165
            Width           =   240
         End
         Begin VB.Frame Frame5 
            Caption         =   "[ Datos del Movimiento ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   4875
            Left            =   12675
            TabIndex        =   52
            Top             =   2100
            Visible         =   0   'False
            Width           =   11565
            Begin VB.Frame Frame7 
               BackColor       =   &H00E0FEE7&
               BorderStyle     =   0  'None
               Caption         =   "Frame7"
               Height          =   1020
               Index           =   0
               Left            =   135
               TabIndex        =   60
               Top             =   500
               Visible         =   0   'False
               Width           =   8985
               Begin VB.CommandButton Command1 
                  Height          =   240
                  Left            =   5640
                  Picture         =   "FrmIngCajBan.frx":29E2
                  Style           =   1  'Graphical
                  TabIndex        =   62
                  Top             =   150
                  Width           =   240
               End
               Begin VB.TextBox TxtImportePer 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   1275
                  TabIndex        =   61
                  Text            =   "TxtImportePer"
                  Top             =   435
                  Width           =   1200
               End
               Begin VB.TextBox TxtPersonal 
                  Height          =   300
                  Left            =   1275
                  TabIndex        =   63
                  Text            =   "TxtPersonal"
                  Top             =   120
                  Width           =   4635
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Personal"
                  Height          =   195
                  Index           =   12
                  Left            =   0
                  TabIndex        =   66
                  Top             =   150
                  Width           =   615
               End
               Begin VB.Label LblIdPersonal 
                  AutoSize        =   -1  'True
                  Caption         =   "LblIdPersonal"
                  ForeColor       =   &H00000080&
                  Height          =   195
                  Left            =   6270
                  TabIndex        =   65
                  Top             =   180
                  Visible         =   0   'False
                  Width           =   960
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Importe"
                  Height          =   195
                  Index           =   13
                  Left            =   0
                  TabIndex        =   64
                  Top             =   465
                  Width           =   525
               End
            End
            Begin VB.OptionButton Option2 
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
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   1710
               TabIndex        =   68
               Top             =   375
               Width           =   915
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Personal"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   195
               Left            =   375
               TabIndex        =   67
               Top             =   375
               Width           =   1110
            End
            Begin VB.Frame Frame8 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Caption         =   "Frame7"
               Height          =   1005
               Index           =   0
               Left            =   135
               TabIndex        =   53
               Top             =   705
               Visible         =   0   'False
               Width           =   8985
               Begin VB.TextBox TxtImporteBan 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   1275
                  TabIndex        =   55
                  Text            =   "TxtImporteBan"
                  Top             =   435
                  Width           =   1200
               End
               Begin VB.CommandButton CmdBusBan 
                  Height          =   240
                  Left            =   5640
                  Picture         =   "FrmIngCajBan.frx":2B14
                  Style           =   1  'Graphical
                  TabIndex        =   54
                  Top             =   150
                  Width           =   240
               End
               Begin VB.TextBox TxtBanco 
                  Height          =   300
                  Left            =   1275
                  TabIndex        =   56
                  Text            =   "TxtBanco"
                  Top             =   120
                  Width           =   4635
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Importe"
                  Height          =   195
                  Index           =   14
                  Left            =   0
                  TabIndex        =   59
                  Top             =   465
                  Width           =   525
               End
               Begin VB.Label LblIdBanco 
                  AutoSize        =   -1  'True
                  Caption         =   "LblIdBanco"
                  ForeColor       =   &H00000080&
                  Height          =   195
                  Left            =   6270
                  TabIndex        =   58
                  Top             =   180
                  Visible         =   0   'False
                  Width           =   810
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Banco"
                  Height          =   195
                  Index           =   15
                  Left            =   0
                  TabIndex        =   57
                  Top             =   150
                  Width           =   465
               End
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Height          =   3315
            Left            =   105
            TabIndex        =   42
            Top             =   3495
            Width           =   11610
            Begin VB.Frame Frame3 
               Height          =   2295
               Left            =   10365
               TabIndex        =   80
               Top             =   480
               Width           =   1230
               Begin VB.CommandButton CmdEliminar 
                  Caption         =   "&Eliminar"
                  Enabled         =   0   'False
                  Height          =   570
                  Left            =   90
                  TabIndex        =   82
                  Top             =   1185
                  Width           =   1050
               End
               Begin VB.CommandButton CmdAgregar 
                  Caption         =   "&Agregar Documentos"
                  Enabled         =   0   'False
                  Height          =   570
                  Left            =   90
                  Style           =   1  'Graphical
                  TabIndex        =   81
                  Top             =   585
                  Width           =   1050
               End
            End
            Begin VB.TextBox TxtTotal5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   4830
               Locked          =   -1  'True
               TabIndex        =   72
               Text            =   "TxtTotal5"
               Top             =   3660
               Width           =   990
            End
            Begin VB.CommandButton CmdBusCliente 
               Height          =   240
               Left            =   5925
               Picture         =   "FrmIngCajBan.frx":2C46
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   75
               Width           =   240
            End
            Begin VB.TextBox TxtProv 
               Height          =   300
               Left            =   1560
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   6
               Text            =   "TxtProv"
               Top             =   45
               Width           =   4635
            End
            Begin VB.TextBox TxtTotal4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   9030
               Locked          =   -1  'True
               TabIndex        =   46
               Text            =   "TxtTotal4"
               Top             =   3015
               Width           =   1035
            End
            Begin VB.TextBox TxtTotal3 
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
               Height          =   285
               Left            =   8070
               Locked          =   -1  'True
               TabIndex        =   45
               Text            =   "TxtTotal3"
               Top             =   3015
               Width           =   975
            End
            Begin VB.TextBox TxtTotal2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   7110
               Locked          =   -1  'True
               TabIndex        =   44
               Text            =   "TxtTotal2"
               Top             =   3015
               Width           =   975
            End
            Begin VB.TextBox TxtTotal1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   6150
               Locked          =   -1  'True
               TabIndex        =   43
               Text            =   "TxtTotal1"
               Top             =   3015
               Width           =   975
            End
            Begin VSFlex7Ctl.VSFlexGrid Fg2 
               Height          =   2430
               Left            =   0
               TabIndex        =   8
               Top             =   570
               Width           =   10335
               _cx             =   18230
               _cy             =   4286
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
               Cols            =   14
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmIngCajBan.frx":2D78
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
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "TOTAL==>"
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
               Index           =   18
               Left            =   3555
               TabIndex        =   74
               Top             =   3705
               Width           =   930
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Total Haber ==>"
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
               Index           =   17
               Left            =   4650
               TabIndex        =   73
               Top             =   3045
               Width           =   1395
            End
            Begin VB.Label LblIdCliente 
               AutoSize        =   -1  'True
               Caption         =   "LblIdCliente"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   6255
               TabIndex        =   50
               Top             =   105
               Visible         =   0   'False
               Width           =   825
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Cliente"
               Height          =   195
               Index           =   3
               Left            =   0
               TabIndex        =   49
               Top             =   75
               Width           =   480
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Documentos x Pagar"
               Height          =   195
               Index           =   2
               Left            =   0
               TabIndex        =   48
               Top             =   390
               Width           =   1485
            End
         End
         Begin VB.CommandButton CmdBusDoc 
            Height          =   240
            Left            =   2310
            Picture         =   "FrmIngCajBan.frx":2F13
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   3165
            Width           =   240
         End
         Begin VB.CommandButton CmdBusMovi 
            Height          =   240
            Left            =   2310
            Picture         =   "FrmIngCajBan.frx":3045
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   1185
            Width           =   240
         End
         Begin VB.CommandButton CmdBusMon 
            Height          =   240
            Left            =   8145
            Picture         =   "FrmIngCajBan.frx":3177
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   870
            Width           =   240
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   8355
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   5
            Text            =   "TxtNumDoc"
            Top             =   3135
            Width           =   1995
         End
         Begin VB.TextBox TxtImporte 
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
            Left            =   7080
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   7
            Text            =   "TxtImporte"
            Top             =   2505
            Width           =   975
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
            Left            =   1665
            TabIndex        =   16
            Top             =   885
            Width           =   1170
         End
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
            Left            =   2895
            TabIndex        =   15
            Top             =   885
            Width           =   1170
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchMov 
            Height          =   300
            Left            =   1665
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
            Valor           =   "20/08/2007"
         End
         Begin VB.TextBox TxtIdMov 
            Height          =   300
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   2
            Text            =   "TxtIdMov"
            Top             =   1155
            Width           =   915
         End
         Begin VB.TextBox TxtIdDoc 
            Height          =   300
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   4
            Text            =   "TxtIdDoc"
            Top             =   3135
            Width           =   915
         End
         Begin VB.TextBox TxtIdMon 
            Height          =   300
            Left            =   7500
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   1
            Text            =   "TxtIdMon"
            Top             =   840
            Width           =   915
         End
         Begin VB.TextBox TxtMedPag 
            Height          =   300
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   3
            Text            =   "TxtMedPag"
            Top             =   2820
            Width           =   915
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   990
            Left            =   1665
            TabIndex        =   76
            Top             =   1500
            Width           =   6720
            _cx             =   11853
            _cy             =   1746
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
            Rows            =   50
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmIngCajBan.frx":32A9
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Destino del Ingreso"
            Height          =   195
            Index           =   19
            Left            =   120
            TabIndex        =   75
            Top             =   1515
            Width           =   1365
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Medio de Pago"
            Height          =   195
            Index           =   16
            Left            =   120
            TabIndex        =   70
            Top             =   2835
            Width           =   1080
         End
         Begin VB.Label LblMedPag 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblMedPag"
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
            Left            =   2625
            TabIndex        =   69
            Top             =   2820
            Width           =   5760
         End
         Begin VB.Label LblMes1 
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
            Left            =   8460
            TabIndex        =   51
            Top             =   510
            Width           =   2100
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
            Left            =   9840
            TabIndex        =   41
            Top             =   2550
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cambio"
            Height          =   195
            Left            =   8460
            TabIndex        =   40
            Top             =   2595
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label LblIdCtaHaber 
            Caption         =   "LblIdCtaHaber"
            Height          =   180
            Left            =   4740
            TabIndex        =   39
            Top             =   555
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label LblIdCtaDebe 
            Caption         =   "LblIdCtaDebe"
            Height          =   180
            Left            =   4740
            TabIndex        =   38
            Top             =   330
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label LblIdCueBan 
            Caption         =   "LblIdCueBan "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Left            =   4890
            TabIndex        =   37
            Top             =   840
            Visible         =   0   'False
            Width           =   1395
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
            Left            =   2625
            TabIndex        =   36
            Top             =   3135
            Width           =   4065
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Documento"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   34
            Top             =   3165
            Width           =   825
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Operacion"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   32
            Top             =   870
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Emision"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   31
            Top             =   570
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
            Left            =   8460
            TabIndex        =   30
            Top             =   840
            Width           =   2115
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Origen del Ingreso"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   28
            Top             =   1170
            Width           =   1290
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
            Left            =   2655
            TabIndex        =   27
            Top             =   1155
            Width           =   5760
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Documento"
            Height          =   195
            Index           =   0
            Left            =   7125
            TabIndex        =   26
            Top             =   3165
            Width           =   1050
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Total Debe ==>"
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
            Index           =   4
            Left            =   5685
            TabIndex        =   25
            Top             =   2550
            Width           =   1335
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   5
            Left            =   6765
            TabIndex        =   24
            Top             =   885
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Debe"
            Height          =   240
            Index           =   0
            Left            =   6075
            TabIndex        =   23
            Top             =   300
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Haber"
            Height          =   240
            Index           =   1
            Left            =   6075
            TabIndex        =   22
            Top             =   570
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label xxxxxx 
            Caption         =   "xxxxxx"
            Height          =   240
            Left            =   7170
            TabIndex        =   21
            Top             =   300
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label aaaaaa 
            Caption         =   "aaaaaa"
            Height          =   240
            Left            =   7170
            TabIndex        =   20
            Top             =   570
            Visible         =   0   'False
            Width           =   990
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
            Left            =   90
            TabIndex        =   29
            Top             =   60
            Width           =   11610
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
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
Attribute VB_Name = "FrmIngCajban"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SeEjecuto As Boolean
Dim QueHace As Integer
Dim RstMov As New ADODB.Recordset
Dim xCuentaHaber As Integer        'para almacenar el codigo de la cuenta haber de la operacion
Dim xCuentaDebe As Integer
Dim Rst As New ADODB.Recordset
Dim xSQL As String
Dim IdEntGen As Integer
Dim xFchPer As String
Dim xOrigenOpera As Integer        'para especificar el origen de la operacion ver tabla con_origenes
Dim Agregando As Boolean

Sub Eliminar()
    Dim Rpta As Integer
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    Rpta = MsgBox("¿ Esta seguro de eliminar el movimiento ?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        Set Rst = BuscaConCriterio("SELECT * FROM con_cajabancodet WHERE id = " & RstMov("id") & "", xCon)
        If Rst.RecordCount <> 0 Then
            'actualizamos el saldo de los documentos cancelados
            Rst.MoveFirst
            For A = 1 To Rst.RecordCount
                ' si la entidad generadora es 1 osea caja y bancos
                If Rst("identgen") = 1 Then
                    xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = ([vta_ventas]![impsal]+" & Rst("impabo") & ") " _
                        & " WHERE (((vta_ventas.id)=" & Rst("iddoc") & "))"
                End If
                Rst.MoveNext
                If Rst.EOF = True Then
                    Exit For
                End If
            Next A
            
            'eliminamos el libro diario
            xCon.Execute "DELETE * FROM con_diario WHERE idlib = 6 AND idmov = " & RstMov("id") & ""
            
            'eliminamos el movimiento de caja y banco
            xCon.Execute "DELETE * FROM con_cajabanco WHERE id = " & RstMov("id") & ""
            
            MsgBox "El movimiento se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Set Rst = Nothing
            RstMov.Requery
            If RstMov.RecordCount = 0 Then
                Rpta = MsgBox("No se ha registrado movimientos, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
                If Rpta = vbYes Then
                    Nuevo
                Else
                    Set RstMov = Nothing
                    Unload Me
                    Exit Sub
                End If
            End If
        End If
    End If
End Sub

Sub MuestraSegundoTab()
    Dim RstDet As New ADODB.Recordset
    Dim A As Integer
    
    Blanquea
    
    TxtFchMov.Valor = RstMov("fchope")
    
    If RstMov("tipope") = 1 Then
        OptCaja.Value = True
    Else
        OptBanco.Value = True
    End If
    
    TxtIdMon.Text = RstMov("idmon")
    LblMoneda.Caption = Busca_Codigo(RstMov("idmon"), "id", "descripcion", "mae_moneda", "N", xCon)
    
    TxtIdMov.Text = RstMov("idori")
    TxtIdMov_Validate True
    
    TxtIdDoc.Text = RstMov("iddoc")
    TxtNumDoc.Text = NulosC(RstMov("numdoc"))
    TxtImporte.Text = Format(RstMov("importe"), "0.00")
    'LblDescMov.Caption = NulosC(RstMov("descori"))
    TxtIdDoc.Text = RstMov("iddoc")
    LblDescDoc.Caption = NulosC(RstMov("descdoc"))
    
    'Mostramos los destinos del egreso
    RST_Busq RstDet, "SELECT con_cajabancoorides.idorides, con_destino.descripcion AS descdest, con_destino.idcuen, con_planctas.cuenta, con_cajabancoorides.importe, " _
        & " con_planctas.descripcion AS desccta, con_destino.entgen FROM con_cajabanco LEFT JOIN (con_planctas RIGHT JOIN (con_cajabancoorides LEFT JOIN con_destino " _
        & " ON con_cajabancoorides.idorides = con_destino.id) ON con_planctas.id = con_destino.idcuen) ON con_cajabanco.id = con_cajabancoorides.id " _
        & " WHERE (((con_cajabanco.id)=" & RstMov("id") & ") AND ((con_cajabanco.tipmov)=1))", xCon

    If RstDet.RecordCount <> 0 Then
        RstDet.MoveFirst
        Fg1.Rows = 1
        For A = 1 To RstDet.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = NulosC(RstDet("descdest"))
            Fg1.TextMatrix(A, 2) = Format(NulosN(RstDet("importe")), "0.00")
            Fg1.TextMatrix(A, 3) = NulosN(RstDet("idorides"))
            Fg1.TextMatrix(A, 4) = NulosN(RstDet("idcuen"))
            Fg1.TextMatrix(A, 5) = NulosN(RstDet("entgen"))
            
            RstDet.MoveNext
            If RstDet.EOF = True Then
                Exit For
            End If
        Next A
        SumarDestinos
    End If
    Set RstDet = Nothing
    
    'Mostramos detalles de caja y bancos
    RST_Busq RstDet, "SELECT con_cajabancodet.id, con_cajabancodet.iddoc, mae_cliente.nombre, " _
        & " [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, mae_documento.abrev AS nomdoc, vta_ventas.fchdoc, " _
        & " vta_ventas.fchven, vta_ventas.imptotdoc, con_cajabancodet.salant, con_cajabancodet.impabo, vta_ventas.idmon, " _
        & " mae_moneda.simbolo, con_cajabancodet.idori FROM mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN (con_cajabancodet " _
        & " LEFT JOIN vta_ventas ON con_cajabancodet.iddoc = vta_ventas.id) ON mae_cliente.id = vta_ventas.idcli) " _
        & " ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon WHERE (((con_cajabancodet.id)=" & RstMov("id") & "))", xCon

    Fg2.Rows = 1
    If RstDet.RecordCount <> 0 Then
        RstDet.MoveFirst
        For A = 1 To RstDet.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(A, 1) = NulosC(RstDet("nombre"))
            Fg2.TextMatrix(A, 2) = NulosC(RstDet("nomdoc"))
            Fg2.TextMatrix(A, 3) = Format(RstDet("fchdoc"), "dd/mm/yy")
            Fg2.TextMatrix(A, 4) = NulosC(RstDet("simbolo"))
            Fg2.TextMatrix(A, 5) = NulosC(RstDet("numdoc"))
            Fg2.TextMatrix(A, 6) = Format(RstDet("imptotdoc"), "0.00")
            Fg2.TextMatrix(A, 7) = Format(RstDet("salant"), "0.00")
            Fg2.TextMatrix(A, 8) = Format(RstDet("impabo"), "0.00")
            Fg2.TextMatrix(A, 9) = Format((RstDet("salant") - RstDet("impabo")), "0.00")
            Fg2.TextMatrix(A, 10) = RstDet("iddoc")
            Fg2.TextMatrix(A, 12) = NulosN(RstDet("idori"))
            
            RstDet.MoveNext
            If RstDet.EOF = True Then
                Exit For
            End If
        Next A
    End If
    HallarTotal
End Sub

Sub SumarDestinos()
    Dim A As Integer
    Dim xTot As Double
    For A = 1 To Fg1.Rows - 1
        xTot = xTot + NulosN(Fg1.TextMatrix(A, 2))
    Next A
    
    TxtImporte.Text = Format(xTot, "0.00")
End Sub

Private Sub CmdAddCon_Click()
    If Fg1.TextMatrix(Fg1.Rows - 1, 1) = "" Then Exit Sub
    
    Fg1.Rows = Fg1.Rows + 1
    Fg1_CellButtonClick Fg1.Rows - 1, 1
End Sub

Private Sub CmdAgregar_Click()
    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado el detino del ingreso", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.SetFocus
        Exit Sub
    End If
    
    Dim xfrm As New eps_librerias.FormSeleccion
    Dim xCampos(6, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim xRs1 As New ADODB.Recordset
    Dim A As Integer
    
    xCampos(0, 0) = "Nº Documento":    xCampos(0, 1) = "numdoc":        xCampos(0, 2) = "1500":   xCampos(0, 3) = "C":     xCampos(0, 4) = "S"
    xCampos(1, 0) = "Fch. Giro":       xCampos(1, 1) = "fchdoc":        xCampos(1, 2) = "1000":   xCampos(1, 3) = "C":     xCampos(1, 4) = "N"
    xCampos(2, 0) = "Tip. Doc.":       xCampos(2, 1) = "abrev":         xCampos(2, 2) = "1200":   xCampos(2, 3) = "C":     xCampos(2, 4) = "N"
    xCampos(3, 0) = "Moneda":          xCampos(3, 1) = "simbolo":       xCampos(3, 2) = "1500":   xCampos(3, 3) = "C":     xCampos(3, 4) = "N"
    xCampos(4, 0) = "Importe":         xCampos(4, 1) = "imptotdoc":     xCampos(4, 2) = "1500":   xCampos(4, 3) = "N":     xCampos(4, 4) = "N"
    xCampos(5, 0) = "Saldo":           xCampos(5, 1) = "impsal":        xCampos(5, 2) = "1500":   xCampos(5, 3) = "N":     xCampos(5, 4) = "N"
    
    xfrm.SQLCad = "SELECT mae_cliente.numruc, mae_cliente.nombre, mae_documento.abrev, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, " _
        & " vta_ventas.fchdoc, mae_moneda.simbolo, vta_ventas.impsal, vta_ventas.fchven, vta_ventas.imptotdoc, vta_ventas.idcli, vta_ventas.id " _
        & " FROM mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) " _
        & " ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon Where (((vta_ventas.impsal) > 0) And ((vta_ventas.idcli) = " & NulosN(LblIdCliente.Caption) & ")) " _
        & " ORDER BY [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc]"

    xfrm.Titulo = "Buscando Documentos del Cliente"
    
    Set xfrm.Coneccion = xCon
    Set xRs = xfrm.Seleccionar(xCampos)
    
    If xRs.State = 1 Then
        Agregando = True
        'CARGAMOS LOS DOCUMENTOS ADJUNTOS Y LO MOSTRAMOS EN LA LISTA DE "DOCUMENTOS ADJUNTOS"
        For A = 1 To xRs.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(A, 1) = xRs("nombre")
            Fg2.TextMatrix(A, 2) = xRs("abrev")
            Fg2.TextMatrix(A, 3) = xRs("fchdoc")
            Fg2.TextMatrix(A, 4) = xRs("simbolo")
            Fg2.TextMatrix(A, 5) = xRs("numdoc")
            Fg2.TextMatrix(A, 6) = xRs("imptotdoc")
            Fg2.TextMatrix(A, 7) = xRs("impsal")
            Fg2.TextMatrix(A, 10) = xRs("id")
            Fg2.TextMatrix(A, 11) = xRs("idcli")
            Fg2.TextMatrix(A, 12) = Fg1.TextMatrix(Fg1.Row, 3)
            Fg2.TextMatrix(A, 13) = Fg1.TextMatrix(Fg1.Row, 4)
            
            xRs.MoveNext
            If xRs.EOF = True Then Exit For
        Next A
        Agregando = False
    End If
    
    'HallarTotal
    Set xfrm = Nothing
End Sub

Sub HallarTotal()
    If Fg2.Rows = 1 Then Exit Sub
    Dim A, B As Integer
    Dim xTotal1, xTotal2, xTotal3, xTotal4 As Double
    
    For A = 1 To Fg2.Rows - 1
        xTotal1 = xTotal1 + Val(Fg2.TextMatrix(A, 6))
        xTotal2 = xTotal2 + Val(Fg2.TextMatrix(A, 7))
        xTotal3 = xTotal3 + Val(Fg2.TextMatrix(A, 8))
        xTotal4 = xTotal4 + Val(Fg2.TextMatrix(A, 9))
    Next A
    
    TxtTotal1.Text = Format(xTotal1, "0.00")
    TxtTotal2.Text = Format(xTotal2, "0.00")
    TxtTotal3.Text = Format(xTotal3, "0.00")
    TxtTotal4.Text = Format(xTotal4, "0.00")
    
    'hallamos los totales para cada destino del ingreso
    For A = 1 To Fg1.Rows - 1
        xTotal1 = 0
        For B = 1 To Fg2.Rows - 1
            If NulosN(Fg1.TextMatrix(A, 3)) = NulosN(Fg2.TextMatrix(B, 12)) Then
                xTotal1 = xTotal1 + NulosN(Fg2.TextMatrix(B, 8))
            End If
        Next B
        Fg1.TextMatrix(A, 2) = xTotal1
    Next A
    
    xTotal1 = 0
    For A = 1 To Fg1.Rows - 1
        xTotal1 = xTotal1 + NulosN(Fg1.TextMatrix(A, 2))
    Next A
    TxtImporte = xTotal1
End Sub

Private Sub CmdBusBan_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Codigo":       xCampos(0, 1) = "id":            xCampos(0, 2) = "1200":         xCampos(0, 3) = "N"
    xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
    
    xForm.SQLCad = "SELECT * FROM  mae_bancos ORDER BY descripcion"

    xForm.Titulo = "Buscando Bancos"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "id"
    xForm.CampoBusca = "id"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtBanco.Text = xRs("descripcion")
        LblIdBanco.Caption = xRs("id")
        TxtImporteBan.SetFocus
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusCliente_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Nombre":      xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":   xCampos(1, 1) = "numruc":        xCampos(1, 2) = "1400":         xCampos(1, 3) = "C"
    
    xForm.SQLCad = "SELECT mae_cliente.numruc, mae_cliente.id, mae_cliente.nombre, mae_cliente.ageret FROM mae_cliente"
    
    xForm.Titulo = "Buscando Clientes"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "nombre"
    xForm.CampoBusca = "nombre"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtProv.Text = xRs("nombre")
        LblIdCliente.Caption = xRs("id")
'        If xRs("ageret") = -1 Then
'            LblEsAgente.Left = 6705
'            LblEsAgente.Top = 2940
'            LblEsAgente.Visible = True
'        Else
'            LblEsAgente.Visible = False
'        End If
'        CargarFacturasPorCobrar
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

'Sub CargarFacturasPorCobrar()
'    Dim Rst As New ADODB.Recordset
'    Dim A As Integer
'
'    RST_Busq Rst, "SELECT mae_cliente.numruc, mae_cliente.nombre, mae_documento.abrev, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, " _
'        & " vta_ventas.fchdoc, mae_moneda.simbolo, vta_ventas.impsal, vta_ventas.fchven, vta_ventas.imptotdoc, vta_ventas.idcli, vta_ventas.id " _
'        & " FROM mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) " _
'        & " ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon " _
'        & " WHERE (((vta_ventas.impsal)>0) AND ((vta_ventas.idcli)=" & Val(LblIdCliente.Caption) & ")) ORDER BY  [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc]", xCon
'
'    Fg1.Rows = 1
'    If Rst.RecordCount <> 0 Then
'        Rst.MoveFirst
'        For A = 1 To Rst.RecordCount
'            Fg1.Rows = Fg1.Rows + 1
'
'            Fg1.TextMatrix(A, 1) = Rst("abrev")
'            Fg1.TextMatrix(A, 2) = Rst("fchdoc")
'            Fg1.TextMatrix(A, 3) = Rst("fchven")
'            Fg1.TextMatrix(A, 4) = Rst("simbolo")
'            Fg1.TextMatrix(A, 5) = Rst("numdoc")
'            Fg1.TextMatrix(A, 6) = Format(Rst("imptotdoc"), "0.00")
'            Fg1.TextMatrix(A, 7) = Format(Rst("impsal"), "0.00")
'            Fg1.TextMatrix(A, 8) = Rst("id")
'            'Fg1.TextMatrix(A, 9) = Rst("idprov")
'            Rst.MoveNext
'            If Rst.EOF = True Then
'                Exit For
'            End If
'        Next A
'    End If
'End Sub

'Private Sub CmdBusCueBan_Click()
'    If QueHace = 3 Then Exit Sub
'
'    If NulosC(TxtIdMon.Text) = "" Then
'        MsgBox "No ha seleccionado la moneda para la operacion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        TxtIdMon.SetFocus
'        Exit Sub
'    End If
'
'    Dim xForm As New eps_librerias.FormBuscar
'    Dim xRs As New ADODB.Recordset
'    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
'
'    Dim xCampos(4, 4) As String
'
'    xCampos(0, 0) = "Banco":           xCampos(0, 1) = "desban":        xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
'    xCampos(1, 0) = "Nº Cuenta":       xCampos(1, 1) = "numcue":        xCampos(1, 2) = "1200":         xCampos(1, 3) = "C"
'    xCampos(2, 0) = "Moneda":          xCampos(2, 1) = "desmon":        xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
'    xCampos(3, 0) = "Nº Cta Contable": xCampos(3, 1) = "cuenta":        xCampos(3, 2) = "1500":         xCampos(3, 3) = "C"
'
'    xForm.SQLCad = "SELECT mae_bancos.descripcion AS desban, con_bancocuenta.*, mae_moneda.descripcion AS desmon, " _
'        & " con_planctas.cuenta FROM mae_bancos INNER JOIN (con_planctas RIGHT JOIN (con_bancocuenta LEFT JOIN mae_moneda " _
'        & " ON con_bancocuenta.idmon = mae_moneda.id) ON con_planctas.id = con_bancocuenta.idcuen) ON " _
'        & " mae_bancos.id = con_bancocuenta.idban Where (((con_bancocuenta.idmon) = " & Val(Val(TxtIdMon.Text)) & ")) " _
'        & " ORDER BY mae_bancos.descripcion"
'
'    xForm.Titulo = "Buscando Cuentas de Banco"
'    xForm.FormaBusca = Principio
'    xForm.Criterio = ""
'    xForm.Ordenado = "desban"
'    xForm.CampoBusca = "desban"
'    Set xForm.Coneccion = xCon
'    Set xRs = xForm.BuscarReg(xCampos)
'    If xRs.State = 1 Then
'        TxtNumCue.Text = xRs("numcue")
'        LblIdCueBan = xRs("id")
'        LblBanco.Caption = Trim(xRs("desban")) '"   Cuenta Nº " & xRs("numcue")
'        xCuentaHaber = xRs("idcuen")
'
'        aaaaaa.Caption = xRs("cuenta")
'        LblIdCtaHaber.Caption = xCuentaHaber
'
'        TxtIdDoc.SetFocus
'    End If
'    Set xForm = Nothing
'    Set xRs = Nothing
'End Sub

Private Sub CmdBusDoc_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "id":           xCampos(1, 1) = "id":            xCampos(1, 2) = "1200":         xCampos(1, 3) = "N"
    
    If OptCaja.Value = True Then
        xForm.SQLCad = "SELECT mae_doccajaban.tipo, * From mae_doccajaban Where (((mae_doccajaban.tipo) = 1)) " _
            & " ORDER BY mae_doccajaban.descripcion"
    Else
        xForm.SQLCad = "SELECT mae_doccajaban.tipo, * From mae_doccajaban Where (((mae_doccajaban.tipo) = 2)) " _
            & " ORDER BY mae_doccajaban.descripcion"
    End If
    
    xForm.Titulo = "Buscando Documentos"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "descripcion"
    xForm.CampoBusca = "descripcion"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdDoc.Text = xRs("id")
        LblDescDoc.Caption = xRs("descripcion")
        TxtNumDoc.SetFocus
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMedPag_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Codigo":       xCampos(0, 1) = "id":            xCampos(0, 2) = "1000":         xCampos(0, 3) = "N"
    xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "7000":         xCampos(1, 3) = "C"
    
    'filtramos por tipo de movimiento  = 1 (Ingreso)
    xForm.SQLCad = "SELECT * FROM  con_mediopago ORDER BY descripcion"

    xForm.Titulo = "Buscando Medio de Pago"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "id"
    xForm.CampoBusca = "id"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtMedPag.Text = xRs("id")
        LblMedPag.Caption = xRs("descripcion")
        TxtIdDoc.SetFocus
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMon_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Codigo":       xCampos(0, 1) = "id":            xCampos(0, 2) = "1200":         xCampos(0, 3) = "N"
    xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
    
    'filtramos por tipo de movimiento  = 1 (Ingreso)
    xForm.SQLCad = "SELECT * FROM  mae_moneda ORDER BY descripcion"

    xForm.Titulo = "Buscando Moneda"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "id"
    xForm.CampoBusca = "id"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdMon.Text = xRs("id")
        LblMoneda.Caption = xRs("descripcion")
        TxtIdMov.SetFocus
        
        If TxtIdMon.Text <> 1 Then
            LblTipoCambio.Caption = HallaTipoCambio(CDate(TxtFchMov.Valor), Val(TxtIdMon.Text), 2, xCon)
            LblTipoCambio.Caption = Format(LblTipoCambio.Caption, "0.00")
            
            LblTipoCambio.Visible = True
            LblTipCam.Visible = True
        Else
            LblTipoCambio.Visible = False
            LblTipCam.Visible = False
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMovi_Click()
    If QueHace = 3 Then Exit Sub
    
    If NulosC(TxtIdMon.Text) = "" Then
        MsgBox "No ha especificado la moneda para la operacion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Codigo":        xCampos(0, 1) = "id":            xCampos(0, 2) = "1000":         xCampos(0, 3) = "N"
    xCampos(1, 0) = "Descripcion":   xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Cuenta":        xCampos(2, 1) = "desccuen":      xCampos(2, 2) = "3000":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Nº Cuenta":     xCampos(3, 1) = "cuenta":        xCampos(3, 2) = "1000":         xCampos(3, 3) = "C"
    
    xForm.SQLCad = "SELECT con_origen.*, con_planctas.cuenta, con_planctas.descripcion AS desccuen, con_origen.idmon FROM con_planctas RIGHT JOIN con_origen " _
        & " ON con_planctas.id = con_origen.idcue WHERE (((con_origen.idmon)=" & Val(TxtIdMon.Text) & ") AND ((con_origen.tipmov)=1))"
    
    xForm.Titulo = "Buscando Origen del Egreso"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "id"
    xForm.CampoBusca = "id"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        LblDescMov.Caption = xRs("descripcion")
        TxtIdMov.Text = xRs("id")
        xCuentaDebe = xRs("idcue")
        If xRs("entgen") = 1 Then
            'Frame6.Visible = True
            'CmdBusCliente.Enabled = True
            'CmdAgregar.Enabled = True
            'CmdEliminar.Enabled = True
        Else
            'CmdBusCliente.Enabled = False
            'CmdAgregar.Enabled = False
            'CmdEliminar.Enabled = False
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
    
End Sub


Function HallarNumeroDocumentoCaja(CodigoDocumento As Integer) As String
    Dim Rst  As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT con_cajabanco.iddoc, con_cajabanco.numdoc From con_cajabanco " _
        & " WHERE (((con_cajabanco.iddoc)=" & CodigoDocumento & ")) ORDER BY numdoc", xCon

    If Rst.RecordCount = 0 Then
        HallarNumeroDocumentoCaja = "000001"
    Else
        Rst.MoveLast
        HallarNumeroDocumentoCaja = Format(Val(Rst("numdoc")) + 1, "000000")
    End If
End Function

Private Sub CmdDelCon_Click()
    If Fg1.Rows = 1 Then Exit Sub
    Fg1.RemoveItem Fg1.Row
End Sub

Private Sub CmdEliminar_Click()
    Fg2.RemoveItem Fg2.Row
    HallarTotal
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 1 Then
        Dim xForm As New eps_librerias.FormBuscar
        Dim xRs As New ADODB.Recordset
        Dim xCampos(3, 4) As String
        
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "3000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Cuenta":       xCampos(1, 1) = "descuen":       xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
        xCampos(2, 0) = "Nº Cuenta":    xCampos(2, 1) = "cuenta":        xCampos(2, 2) = "1200":         xCampos(2, 3) = "C"
        xForm.Ordenado = "descripcion"
        xForm.CampoBusca = "descripcion"
        
        xForm.SQLCad = "SELECT con_planctas.cuenta, con_planctas.descripcion AS descuen, con_destino.id, con_destino.idmon, con_destino.descripcion, con_destino.idcuen, " _
            & " con_destino.tipmov, con_destino.entgen FROM con_planctas RIGHT JOIN con_destino ON con_planctas.id = con_destino.idcuen " _
            & " WHERE (((con_destino.idmon)=" & Val(TxtIdMon.Text) & ") AND ((con_destino.tipmov)=1))"

        xForm.Titulo = "Buscando Destino del Ingreso"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        If xRs.State = 1 Then
            Fg1.TextMatrix(Row, 1) = xRs("descripcion")
            Fg1.TextMatrix(Row, 3) = xRs("id")
            Fg1.TextMatrix(Row, 4) = xRs("idcuen")
            Fg1.TextMatrix(Row, 5) = NulosN(xRs("entgen"))
        End If
        Set xForm = Nothing
        Set xRs = Nothing
    End If
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 2 Then
        Fg1.TextMatrix(Fg1.Row, 2) = Format(Fg1.TextMatrix(Fg1.Row, 2), "0.00")
        HallarTotalFG1
    End If
End Sub

Sub HallarTotalFG1()
    Dim A As Integer
    Dim xTotal As Double
    
    For A = 1 To Fg1.Rows - 1
        xTotal = xTotal + NulosN(Fg1.TextMatrix(A, 2))
    Next A
    TxtImporte.Text = Format(xTotal, "0.00")
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        Fg1.Rows = Fg1.Rows + 1
    End If
    If KeyCode = 46 Then
        Fg1.RemoveItem Fg1.Row
    End If
End Sub

Private Sub Fg2_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Fg2.Col = 8 Then
        Fg2.TextMatrix(Fg2.Row, 8) = Format(Fg2.TextMatrix(Fg2.Row, 8), "0.00")
        Fg2.TextMatrix(Fg2.Row, 9) = Val(Fg2.TextMatrix(Fg2.Row, 7)) - Val(Fg2.TextMatrix(Fg2.Row, 8))
        Fg2.TextMatrix(Fg2.Row, 9) = Format(Fg2.TextMatrix(Fg2.Row, 9), "0.00")
        HallarTotal
    End If
    If Fg1.Rows <> 1 Then
        If Fg1.TextMatrix(Fg1.Row, 4) = "5" Then
            Fg1.TextMatrix(Fg1.Row, 2) = TxtTotal4.Text
        End If
    End If
End Sub

Private Sub Fg2_EnterCell()
    If Fg2.Col = 8 Then
        Fg2.Editable = flexEDKbdMouse
    Else
        Fg2.Editable = flexEDNone
    End If
End Sub

Private Sub Fg2_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = 45 Then
        CmdAgregar_Click
    End If
    If KeyCode = 46 Then
        Fg2.RemoveItem Fg2.Row
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        Dim Rpta As Integer
        
        xFchPer = "01/" + Trim(Str(xMes)) + "/" + Trim(Str(AnoTra))
        RST_Busq RstMov, "SELECT con_cajabanco.*, mae_moneda.simbolo, mae_doccajaban.descripcion AS descdoc, con_origen.descripcion AS descori, " _
            & " mae_doccajaban.abrev AS abredoc, IIf([con_cajabanco]![tipope]=1,'Caja','Banco') AS motmov, con_cajabanco.fchreg, con_origen.idcue " _
            & " FROM (mae_moneda RIGHT JOIN (con_cajabanco LEFT JOIN mae_doccajaban ON con_cajabanco.iddoc = mae_doccajaban.id) ON mae_moneda.id = con_cajabanco.idmon) " _
            & " LEFT JOIN con_origen ON con_cajabanco.idori = con_origen.id " _
            & " WHERE (((con_cajabanco.fchreg)=CDate('" & xFchPer & "')) AND ((con_cajabanco.tipmov)=1))", xCon
        
        Set Dg1.DataSource = RstMov
        OpcionesPeriodo
        If RstMov.RecordCount = 0 Then
            Rpta = MsgBox("No se ha registrado ninguna cobranza, ¿Desea agregar una ahora?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
            If Rpta = vbYes Then
                Nuevo
            Else
                xMes = SeleccionaMes(xCon)
                OpcionesPeriodo
                CargarRSTCom
                If RstMov.RecordCount = 0 Then
                    Set RstMov = Nothing
                    Unload Me
                End If
            End If
        Else
            OpcionesPeriodo
            Dg1.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
    QueHace = 3
    SeEjecuto = False
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    
    Fg1.ColWidth(3) = 0
    Fg1.ColWidth(4) = 0
    Fg1.ColWidth(5) = 0
    Fg1.SelectionMode = flexSelectionByRow
    
    Fg2.ColWidth(10) = 0
    Fg2.ColWidth(11) = 0
    Fg2.ColWidth(12) = 0
    Fg2.ColWidth(13) = 0
    
    Frame6.BackColor = &H8000000F
    Frame6.Left = 105
    Frame6.Top = 3495
    TxtTotal5.Text = ""
End Sub

Private Sub OptBanco_Click()
    If QueHace = 3 Then Exit Sub
    TxtMedPag.Enabled = True
    CmdBusMedPag.Enabled = True
    
End Sub

Private Sub OptCaja_Click()
    If QueHace = 3 Then Exit Sub
    TxtMedPag.Enabled = False
    CmdBusMedPag.Enabled = False
    TxtMedPag.Text = ""
    LblMedPag.Caption = ""
End Sub

'Private Sub Option1_Click()
'    Frame8.Visible = False
'    Frame7.Left = 120
'    Frame7.Top = 705
'    Frame7.Visible = True
'End Sub
'
'Private Sub Option2_Click()
'    Frame7.Visible = False
'    Frame8.Top = 705
'    Frame8.Left = 120
'    Frame8.Visible = False
'End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then
            MuestraSegundoTab
        End If
    End If
End Sub

Sub OpcionesPeriodo()
    Dim NomMes As String
    Dim Cerrado As Boolean
    Dim xFechaMes As String
    Dim xFchIni, xFchFin As Date
    Dim Rpta As Integer
    
    LblMes.Caption = Busca_Codigo(xMes, "id", "descripcion", "con_meses", "N", xCon)
    Cerrado = Busca_Codigo(xMes, "id", "cerrado", "con_meses", "N", xCon)
    
    If Cerrado = True Then
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = False
        Toolbar1.Buttons(3).Visible = False
        Toolbar1.Buttons(4).Visible = False
    Else
        Toolbar1.Buttons(1).Visible = True
        Toolbar1.Buttons(2).Visible = True
        Toolbar1.Buttons(3).Visible = True
        Toolbar1.Buttons(4).Visible = True
    End If
    If xMes <> 0 Then
        xFechaMes = "01/" + Trim(Format(xMes, "00")) + "/" + Trim(Format(Year(Date), "0000"))
        xFchIni = xFechaMes
        xFchFin = Format(HallaDiasMes(CDate(xFechaMes)), "00") + "/" + Mid(xFechaMes, 4, 7)
        LblMes.Caption = Busca_Codigo(xMes, "id", "descripcion", "con_meses", "N", xCon)
        LblMes1.Caption = LblMes.Caption
    End If
End Sub

Sub CargarRSTCom()
    Dim Rpta As Integer
    
    xFchPer = "01/" + Format(Trim(Str(xMes)), "00") + "/" + Trim(Str(AnoTra))
    
    RST_Busq RstMov, "SELECT con_cajabanco.*, mae_moneda.simbolo, mae_doccajaban.descripcion AS descdoc, con_origen.descripcion AS descori, " _
        & " mae_doccajaban.abrev AS abredoc, IIf([con_cajabanco]![tipope]=1,'Caja','Banco') AS motmov, con_cajabanco.fchreg, con_origen.idcue " _
        & " FROM (mae_moneda RIGHT JOIN (con_cajabanco LEFT JOIN mae_doccajaban ON con_cajabanco.iddoc = mae_doccajaban.id) ON mae_moneda.id = con_cajabanco.idmon) " _
        & " LEFT JOIN con_origen ON con_cajabanco.idori = con_origen.id " _
        & " WHERE (((con_cajabanco.fchreg)=CDate('" & xFchPer & "')) AND ((con_cajabanco.tipmov)=1))", xCon

    Set Dg1.DataSource = RstMov

    If RstMov.RecordCount = 0 Then
        Rpta = MsgBox("No se ha registrado ningun movimiento, ¿Desea agregar una ahora?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
        If Rpta = vbYes Then
            Nuevo
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then Modificar
    
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstMov.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then
        Cancelar
        If RstMov.RecordCount = 0 Then
            Unload Me
            Exit Sub
        End If
    End If
    
    If Button.Index = 11 Then
        TabOne1.CurrTab = 0
        xMes = SeleccionaMes(xCon)
        OpcionesPeriodo
        If xMes = 0 Then
            Set RstMov = Nothing
            Unload Me
            Exit Sub
        End If
        CargarRSTCom
        'If RstMov.RecordCount = 0 Then
        '    Set RstMov = Nothing
        '    Unload Me
        'End If
    End If
    
    If Button.Index = 15 Then
        Set RstMov = Nothing
        Unload Me
    End If
End Sub

Sub Cancelar()
    Bloquea
    Label5.Caption = "Detalle de la operacion"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    ActivaTool
    QueHace = 3
    Dg1.SetFocus
End Sub

Function Grabar() As Boolean
    If NulosC(TxtFchMov.Valor) = "" Then
        MsgBox "No ha especificado la fecha de movimiento del documento", vbInformation + vbOKOnly + vbDefaultButton1
        TxtFchMov.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtIdMon.Text) = "" Then
        MsgBox "No ha especificado la moneda", vbInformation + vbOKOnly + vbDefaultButton1
        TxtIdMon.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtIdMov.Text) = "" Then
        MsgBox "No ha especificado el origen del movimiento", vbInformation + vbOKOnly + vbDefaultButton1
        TxtIdMov.SetFocus
        Exit Function
    End If

    If IdEntGen = 1 Then
        If Fg2.Rows = 1 Then
            MsgBox "No ha especificado que documentos se estan cancelando con el movimiento", vbInformation + vbOKOnly + vbDefaultButton1
            TxtProv.SetFocus
            Exit Function
        End If
    End If
    If IdEntGen = 2 Then
        If NulosC(TxtBanco.Text) = "" Then
            MsgBox "No ha especificado el banco", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtBanco.SetFocus
            Exit Function
        End If
        
        If NulosC(TxtImporteBan.Text) = "" Then
            MsgBox "No ha especificado el importe", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtImporteBan.SetFocus
            Exit Function
        End If
    End If
    
    If OptBanco.Value = True Then
        If NulosC(TxtMedPag.Text) = "" Then
            MsgBox "No ha especificado el medio de pago pago para la operacion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtMedPag.SetFocus
            Exit Function
        End If
    End If
    
    If NulosC(TxtIdDoc.Text) = "" Then
        MsgBox "No ha especificado el tipo de documento para registrar la operacion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdDoc.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtNumDoc.Text) = "" Then
        MsgBox "No ha especificado el numero del documento para registrar la operacion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDoc.SetFocus
        Exit Function
    End If
    
    If NulosN(TxtImporte.Text) <> NulosN(TxtTotal3.Text) Then
        MsgBox "El total debe no cuadra con el total haber", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.SetFocus
        Exit Function
    End If
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstOri As New ADODB.Recordset
    Dim RstDia As New ADODB.Recordset
    Dim xNumAsiento As String
    Dim xId As Integer
    
On Error GoTo LaCague

    xCon.BeginTrans
    
    If QueHace = 1 Then
        xNumAsiento = NuevoNumAsiento(6, xMes, xCon)
        xId = HallaCodigoTabla("con_cajabanco", xCon, "id")
        
        RST_Busq RstCab, "SELECT * FROM con_cajabanco", xCon
        RST_Busq RstDet, "SELECT * FROM con_cajabancodet", xCon
        RST_Busq RstOri, "SELECT * FROM con_cajabancoorides", xCon
        RST_Busq RstDia, "SELECT * FROM con_diario", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xNumAsiento = DevuelveNumAsiento(6, RstMov("id"), xMes, xCon)
        xId = RstMov("id")
        RST_Busq RstCab, "SELECT * FROM con_cajabanco WHERE id = " & RstMov("id") & "", xCon
        
        xCon.Execute "DELETE * FROM con_cajabancodet WHERE id = " & RstMov("id") & ""
        xCon.Execute "DELETE * FROM con_diario WHERE idmov = " & RstMov("id") & " AND idlib = 6"
        xCon.Execute "DELETE * FROM con_cajabancoorides WHERE id = " & RstMov("id") & ""
        
        RST_Busq RstDet, "SELECT * FROM con_cajabancodet", xCon
        RST_Busq RstDia, "SELECT * FROM con_diario", xCon
        RST_Busq RstOri, "SELECT * FROM con_cajabancoorides", xCon
    End If
    
    RstCab("iddoc") = Val(TxtIdDoc.Text)
    RstCab("numdoc") = TxtNumDoc.Text
    RstCab("idmedpag") = NulosN(TxtMedPag.Text)
    RstCab("idmon") = Val(TxtIdMon.Text)
    RstCab("idori") = Val(TxtIdMov.Text)
    RstCab("fchope") = TxtFchMov.Valor
    RstCab("importe") = Val(TxtTotal3.Text)
    RstCab("saldo") = Val(TxtTotal3.Text)
    RstCab("fchreg") = CDate("01/" + Format(xMes, "00") + "/" + AnoTra)
    RstCab("numreg") = Format(xMes, "00") + Trim(xNumAsiento)
    RstCab("tipmov") = 1
    RstCab.Update
    
    Dim A As Integer
    'SI ES UNA OPERACION CON CLIENTE
    If Frame6.Visible = True Then
        For A = 1 To Fg2.Rows - 1
            RstDet.AddNew
            RstDet("id") = xId
            RstDet("identgen") = 1
            RstDet("idori") = Fg2.TextMatrix(A, 12)
            RstDet("iddoc") = Fg2.TextMatrix(A, 10)
            RstDet("salant") = Fg2.TextMatrix(A, 7)
            RstDet("impabo") = Fg2.TextMatrix(A, 8)
            RstDet("idorigen") = xOrigenOpera
            RstDet.Update
            xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = " & Val(Fg2.TextMatrix(A, 9)) & " " _
                & " WHERE (((vta_ventas.id) = " & Val(Fg2.TextMatrix(A, 10)) & "))"
        Next A
    End If
    
    'GRABAMOS EL DESTINO DE LA OPERACION
    For A = 1 To Fg1.Rows - 1
        RstOri.AddNew
        RstOri("id") = xId
        RstOri("idorides") = Fg1.TextMatrix(A, 3)
        RstOri("importe") = NulosN(Fg1.TextMatrix(A, 2))
        RstOri.Update
    Next A
        
    'GRABAMOS EL HABER DE LA OPERACION
    For A = 1 To Fg1.Rows - 1
        RstDia.AddNew
        RstDia("año") = AnoTra
        RstDia("idmes") = xMes
        RstDia("idlib") = 6
        RstDia("idmov") = xId
        RstDia("numasi") = xNumAsiento
        RstDia("tc") = Val(LblTipoCambio.Caption)
        RstDia("idcue") = NulosN(Fg1.TextMatrix(A, 4))
        RstDia("fchasi") = CDate("01/" + Format(xMes, "00") + "/" + AnoTra)
        RstDia("fchdoc") = CDate(TxtFchMov.Valor)
        
        If TxtIdMon.Text = "1" Then
            RstDia("impdebsol") = NulosN(Fg1.TextMatrix(A, 2))
            RstDia("impdebdol") = 0
        Else
            RstDia("impdebsol") = Val(TxtTotal3.Text) * NulosN(Fg1.TextMatrix(A, 2))
            RstDia("impdebdol") = Val(TxtTotal3.Text)
        End If
        RstDia.Update
    Next A
    
    'GRABAMOS EL HABER DE LA OPERACION
    If Frame6.Visible = True Then
        'grabamos la cuenta haber de la operacion
        For A = 1 To Fg2.Rows - 1
            RstDia.AddNew
            RstDia("año") = AnoTra
            RstDia("idmes") = xMes
            RstDia("idlib") = 6
            RstDia("idmov") = xId
            RstDia("numasi") = xNumAsiento
            RstDia("tc") = Val(LblTipoCambio.Caption)
            RstDia("idcue") = xCuentaDebe
            RstDia("iddocpro") = Fg2.TextMatrix(A, 10)
            RstDia("fchasi") = CDate("01/" + Format(xMes, "00") + "/" + AnoTra)
            RstDia("fchdoc") = CDate(TxtFchMov.Valor)
            
            If TxtIdMon.Text = "1" Then
                RstDia("imphabsol") = Val(Fg2.TextMatrix(A, 8))
                RstDia("imphabdol") = 0
            Else
                RstDia("imphabsol") = Val(Fg2.TextMatrix(A, 8)) * Val(LblTipoCambio.Caption)
                RstDia("imphabdol") = Val(Fg2.TextMatrix(A, 8))
            End If
            RstDia.Update
        Next A
    End If
    
    xCon.CommitTrans
    MsgBox "La operacion se grabo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Exit Function
    
LaCague:
    Resume
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDia = Nothing
    
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function

Sub Modificar()
    QueHace = 2
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    Label5.Caption = "Modificando Ingreso"
    Bloquea
    MuestraSegundoTab
    TxtFchMov.SetFocus
End Sub

Sub Nuevo()
    QueHace = 1
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    
    Label5.Caption = "Agregando Operacion"
    Blanquea
    Bloquea
    
    Fg1.ColComboList(1) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    Fg1.Editable = flexEDKbdMouse
    
    OptCaja.Value = True
    OptCaja_Click
    Fg1.Rows = 1
    Fg2.Rows = 1
    
    TxtFchMov.SetFocus
End Sub

Sub Bloquea()
    TxtFchMov.Locked = Not TxtFchMov.Locked
    'TxtNumCue.Locked = Not TxtNumCue.Locked
    TxtIdMon.Locked = Not TxtIdMon.Locked
    TxtIdMov.Locked = Not TxtIdMov.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    TxtImporte.Locked = Not TxtImporte.Locked
    'TxtIdMedioPago.Locked = Not TxtIdMedioPago.Locked
    TxtIdDoc.Locked = Not TxtIdDoc.Locked
    TxtMedPag.Locked = Not TxtMedPag.Locked
    CmdAgregar.Enabled = Not CmdAgregar.Enabled
    CmdEliminar.Enabled = Not CmdEliminar.Enabled
    
End Sub

Sub Blanquea()
    TxtFchMov.Valor = ""
    TxtIdMon.Text = ""
    'TxtNumCue.Text = ""
    'TxtIdMedioPago.Text = ""
    TxtProv.Text = ""
    TxtImporte.Text = ""
    TxtIdMov.Text = ""
    TxtIdDoc.Text = ""
    TxtNumDoc.Text = ""
    TxtMedPag.Text = ""
    
'    LblBanco.Caption = ""
    LblDescMov.Caption = ""
'    LblDesMedPag.Caption = ""
    LblMoneda.Caption = ""
    LblDescDoc.Caption = ""
    LblMedPag.Caption = ""
    
    TxtTotal1.Text = ""
    TxtTotal2.Text = ""
    TxtTotal3.Text = ""
    TxtTotal4.Text = ""
    
    TxtPersonal.Text = ""
    TxtImportePer.Text = ""
    TxtBanco.Text = ""
    TxtImporteBan.Text = ""
    LblIdPersonal.Caption = ""
    LblIdBanco.Caption = ""
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

Private Sub TxtBanco_KeyPress(KeyAscii As Integer)
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

Private Sub TxtIdMedioPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

'Private Sub TxtIdMedioPago_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 116 Then
'        CmdBusPro_Click
'    End If
'End Sub

'Private Sub TxtIdMedioPago_Validate(Cancel As Boolean)
'    If TxtIdMedioPago.Text <> "" And TxtIdMon.Text <> "" Then
'        If OptCaja.Value = True Then
'            xSQL = "SELECT con_destino.*, con_planctas.descripcion AS descta, con_planctas.cuenta " _
'                & " FROM con_planctas RIGHT JOIN con_destino ON con_planctas.id = con_destino.idcuenta " _
'                & " WHERE (((con_destino.idmon) = " & Val(TxtIdMon.Text) & "))"
'        Else
'            xSQL = "SELECT con_destino.*, con_planctas.descripcion AS descta, con_planctas.cuenta " _
'                & " FROM con_planctas RIGHT JOIN con_destino ON con_planctas.id = con_destino.idcuenta " _
'                & " WHERE (((con_destino.idmon) = " & TxtIdMon.Text & "))"
'        End If
'        Set Rst = BuscaConCriterio(xSQL, xCon)
'        If Rst.RecordCount <> 0 Then
'            LblDesMedPag.Caption = Rst("descripcion")
'            xCuentaHaber = Rst("idcuenta")
'
'            LblIdCtaHaber.Caption = Rst("idcuenta")
'            aaaaaa.Caption = Rst("cuenta")
'
'            If NulosN(Rst("iddoc")) <> 0 Then
'                TxtIdDoc.Text = Rst("iddoc")
'                LblDescDoc.Caption = Busca_Codigo(Rst("iddoc"), "id", "descripcion", "mae_doccajaban", "N", xCon)
'                TxtNumDoc.Text = HallarNumeroDocumentoCaja(Val(TxtIdDoc.Text))
'            End If
'        Else
'            TxtIdMedioPago.Text = ""
'            LblDesMedPag.Caption = ""
'        End If
'        Set Rst = Nothing
'        'TxtIdDoc.SetFocus
'    End If
'End Sub

Private Sub TxtIdDoc_Validate(Cancel As Boolean)
    If NulosC(TxtIdDoc.Text) <> "" Then
        If OptCaja.Value = True Then
            xSQL = "SELECT mae_doccajaban.tipo, * From mae_doccajaban Where ((mae_doccajaban.tipo = 1) AND (id = " & Val(TxtIdDoc.Text) & ")) " _
                & " ORDER BY mae_doccajaban.descripcion"
        Else
            xSQL = "SELECT mae_doccajaban.tipo, * From mae_doccajaban Where ((mae_doccajaban.tipo = 2) AND (id = " & Val(TxtIdDoc.Text) & ")) " _
                & " ORDER BY mae_doccajaban.descripcion"
        End If
        
        Set Rst = BuscaConCriterio(xSQL, xCon)
        
        If Rst.RecordCount <> 0 Then
            LblDescDoc.Caption = Rst("descripcion")
        End If
        Set Rst = Nothing
    End If
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdMon_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMon_Click
    End If
End Sub

Private Sub TxtIdMon_Validate(Cancel As Boolean)
    If TxtIdMon.Text <> "" Then
        If TxtIdMon.Text <> "1" And TxtFchMov.Valor = "" Then
            TxtIdMon.Text = ""
            LblMoneda.Caption = ""
            Exit Sub
        End If
        
        LblMoneda.Caption = Busca_Codigo(Val(TxtIdMon.Text), "id", "descripcion", "mae_moneda", "N", xCon)
        
        If LblMoneda.Caption = "" Then
            TxtIdMon.Text = ""
        End If
        
        If TxtIdMon.Text <> 1 Then
            LblTipoCambio.Caption = HallaTipoCambio(CDate(TxtFchMov.Valor), Val(TxtIdMon.Text), 2, xCon)
            LblTipoCambio.Caption = Format(LblTipoCambio.Caption, "0.00")
            
            LblTipoCambio.Visible = True
            LblTipCam.Visible = True
        Else
            LblTipoCambio.Visible = False
            LblTipCam.Visible = False
        End If
        
        TxtIdMov.SetFocus
    End If
End Sub

Private Sub TxtIdMov_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdMov_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMovi_Click
    End If
End Sub

Function HallaCuentaDebe(Moneda As Integer, CodigoMovimiento As Integer) As ADODB.Recordset
    Dim xSQL As String
    Dim Rst1 As New ADODB.Recordset
    
    xSQL = "SELECT con_planctas.cuenta, con_planctas.descripcion AS descuen, con_origen.id, con_origen.idmon, " _
        & " con_origen.descripcion, con_origen.idcue, con_origen.tipmov, con_origen.id, con_origen.entgen " _
        & " FROM con_planctas INNER JOIN con_origen ON con_planctas.id = con_origen.idcue " _
        & " WHERE (((con_origen.idmon)=" & Val(TxtIdMon.Text) & ") AND ((con_origen.tipmov)=1) AND ((con_origen.id)=" & Val(TxtIdMov.Text) & "))"
    
    RST_Busq Rst1, xSQL, xCon
    Set HallaCuentaDebe = Rst1
End Function

Private Sub TxtIdMov_Validate(Cancel As Boolean)
    If TxtIdMov.Text <> "" And TxtIdMon.Text <> "" Then

        Set Rst = HallaCuentaDebe(Val(TxtIdMon.Text), Val(TxtIdMov.Text)) 'BuscaConCriterio(xSql, xCon)
        If Rst.RecordCount <> 0 Then
            LblDescMov.Caption = Rst("descripcion")
            xCuentaDebe = Rst("idcue")
            
            LblIdCtaDebe.Caption = xCuentaDebe
            xxxxxx.Caption = Rst("cuenta")
            
            IdEntGen = Rst("entgen")
            'If Rst("entgen") = 1 Then
            '    Frame6.Visible = True
            '    CmdBusCliente.Enabled = True
            '    CmdAgregar.Enabled = True
            '    CmdEliminar.Enabled = True
            'Else
            '    CmdBusCliente.Enabled = False
            '    CmdAgregar.Enabled = False
            '    CmdEliminar.Enabled = False
            'End If
            ''If Rst("entgen") = 2 Then Frame8.Left = 150: Frame8.Top = 2865: Frame8.Visible = True
            'If Rst("entgen") = 3 Or Rst("entgen") = 4 Then Frame7.Left = 150: Frame7.Top = 2865: Frame7.Visible = True
            
            If OptCaja.Value = True Then
                'TxtIdMedioPago.SetFocus
            Else
                'TxtMedPag.SetFocus
            End If
        Else
            TxtIdMov.Text = ""
            LblDescMov.Caption = ""
        End If
    End If
End Sub

Private Sub TxtImporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtImporteBan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtImporteBan.Text = Format(TxtImporteBan.Text, "0.00")
        SendKeys vbTab
    End If
End Sub

Private Sub TxtMedPag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtMedPag_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMedPag_Click
    End If
End Sub

Private Sub TxtMedPag_Validate(Cancel As Boolean)
    If TxtMedPag.Text <> "" Then
        LblMedPag.Caption = Busca_Codigo(TxtMedPag.Text, "id", "descripcion", "con_mediopago", "N", xCon)
        If LblMedPag.Caption = "" Then
            TxtMedPag.Text = ""
        End If
    End If
End Sub

Private Sub TxtNumCue_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

'Private Sub TxtNumCue_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 116 Then
'        CmdBusCueBan_Click
'    End If
'End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtProv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtProv_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCliente_Click
    End If

End Sub

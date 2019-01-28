VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmManDevolucion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Almacén - Devolución"
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
            Picture         =   "FrmManDevolucion.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDevolucion.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDevolucion.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDevolucion.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDevolucion.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDevolucion.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDevolucion.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDevolucion.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDevolucion.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDevolucion.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDevolucion.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDevolucion.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDevolucion.frx":277E
            Key             =   "IMG13"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7020
      Left            =   0
      TabIndex        =   9
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
         TabIndex        =   15
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6090
            Left            =   0
            TabIndex        =   32
            Top             =   450
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
            Columns(1).Caption=   "Fch. Dev."
            Columns(1).DataField=   "fching"
            Columns(1).NumberFormat=   "Short Date"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "T.D."
            Columns(2).DataField=   "destipdoc"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Nº Documento"
            Columns(3).DataField=   "numdoc"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Ítem"
            Columns(4).DataField=   "desitem"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "T.D. Ref."
            Columns(5).DataField=   "destipdocref"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Nª Doc: Ref."
            Columns(6).DataField=   "numdocref"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1879"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1799"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=1323"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1244"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=3122"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=3043"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=7938"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=7858"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=1402"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1323"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=3149"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=3069"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=66,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=74,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=17"
            _StyleDefs(64)  =   "Named:id=33:Normal"
            _StyleDefs(65)  =   ":id=33,.parent=0"
            _StyleDefs(66)  =   "Named:id=34:Heading"
            _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(68)  =   ":id=34,.wraptext=-1"
            _StyleDefs(69)  =   "Named:id=35:Footing"
            _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(71)  =   "Named:id=36:Selected"
            _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=37:Caption"
            _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(75)  =   "Named:id=38:HighlightRow"
            _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(77)  =   "Named:id=39:EvenRow"
            _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(79)  =   "Named:id=40:OddRow"
            _StyleDefs(80)  =   ":id=40,.parent=33"
            _StyleDefs(81)  =   "Named:id=41:RecordSelector"
            _StyleDefs(82)  =   ":id=41,.parent=34"
            _StyleDefs(83)  =   "Named:id=42:FilterBar"
            _StyleDefs(84)  =   ":id=42,.parent=33"
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
            TabIndex        =   16
            Top             =   0
            Width           =   1860
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Consulta de Devolución"
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
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6600
         Left            =   12525
         TabIndex        =   10
         Top             =   375
         Width           =   11790
         Begin VB.CommandButton cmd 
            Height          =   240
            Index           =   4
            Left            =   7710
            Picture         =   "FrmManDevolucion.frx":2B10
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   735
            Width           =   240
         End
         Begin VB.CommandButton cmd 
            Height          =   240
            Index           =   6
            Left            =   11430
            Picture         =   "FrmManDevolucion.frx":2C42
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   1080
            Width           =   240
         End
         Begin VB.CommandButton cmd 
            Height          =   240
            Index           =   0
            Left            =   7710
            Picture         =   "FrmManDevolucion.frx":2D74
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   390
            Width           =   240
         End
         Begin VB.CommandButton CmdBusRes 
            Height          =   240
            Left            =   1770
            Picture         =   "FrmManDevolucion.frx":2EA6
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   720
            Width           =   240
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1110
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   3
            Text            =   "TxtNumSer"
            Top             =   1380
            Width           =   915
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2205
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   4
            Text            =   "TxtNumDoc"
            Top             =   1380
            Width           =   1440
         End
         Begin VB.CommandButton CmdBusTipDoc 
            Height          =   240
            Left            =   1755
            Picture         =   "FrmManDevolucion.frx":2FD8
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1065
            Width           =   240
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   4635
            Left            =   90
            TabIndex        =   8
            Top             =   1830
            Width           =   11595
            _cx             =   20452
            _cy             =   8176
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
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmManDevolucion.frx":310A
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
            Left            =   1110
            TabIndex        =   0
            Top             =   360
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
            Left            =   1110
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   1
            Text            =   "TxtIdRes"
            Top             =   690
            Width           =   915
         End
         Begin VB.TextBox TxtIdAlm 
            Height          =   300
            Left            =   7065
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   5
            Text            =   "TxtIdAlm"
            Top             =   345
            Width           =   915
         End
         Begin VB.TextBox TxtTipDoc 
            Height          =   300
            Left            =   1110
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   2
            Text            =   "TxtTipDoc"
            Top             =   1035
            Width           =   915
         End
         Begin VB.TextBox TxtIdTipDocRef 
            Height          =   300
            Left            =   7065
            Locked          =   -1  'True
            TabIndex        =   6
            Text            =   "TxtIdTipDocRef"
            Top             =   690
            Width           =   915
         End
         Begin VB.TextBox txtNumDocRef 
            Height          =   300
            Left            =   7065
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   7
            Text            =   "txtNumDocRef"
            Top             =   1035
            Width           =   4635
         End
         Begin VB.Label lbliddocref 
            AutoSize        =   -1  'True
            Caption         =   "lbliddocref"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   10770
            TabIndex        =   33
            Top             =   750
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tip. Doc. Ref."
            Height          =   195
            Index           =   5
            Left            =   6000
            TabIndex        =   31
            Top             =   735
            Width           =   1005
         End
         Begin VB.Label LblTipDocRef 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipDocRef"
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
            Left            =   8025
            TabIndex        =   30
            Top             =   705
            Width           =   3690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Doc. Ref."
            Height          =   195
            Index           =   7
            Left            =   6000
            TabIndex        =   28
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Almacen"
            Height          =   195
            Index           =   8
            Left            =   6000
            TabIndex        =   25
            Top             =   390
            Width           =   615
         End
         Begin VB.Label LblAlmacen 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblAlmacen"
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
            Left            =   8025
            TabIndex        =   24
            Top             =   360
            Width           =   3690
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
            Left            =   2040
            TabIndex        =   22
            Top             =   1050
            Width           =   3690
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
            Left            =   2055
            TabIndex        =   21
            Top             =   705
            Width           =   3690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Responsable"
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   20
            Top             =   735
            Width           =   930
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000001&
            BackStyle       =   1  'Opaque
            Height          =   90
            Left            =   2055
            Top             =   1500
            Width           =   105
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Doc."
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   18
            Top             =   1425
            Width           =   795
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Doc."
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   14
            Top             =   1080
            Width           =   930
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Devolución"
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
            TabIndex        =   13
            Top             =   30
            Width           =   11670
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Reg."
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   12
            Top             =   405
            Width           =   705
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   26
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
Attribute VB_Name = "FrmManDevolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMINGRESOALMACEN.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : FORMULARIO PARA EL INGRESO DE DOCUMENTOS NO CONTABLES DE INGRESO O SALIDA,
'* DISEÑADO POR     : JOSE CHACON - 13/04/12
'* ULTIMA REVISION  :
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim RstDev As New ADODB.Recordset                  ' RECORDSET PRINCIPAL QUE CARGARA TODAS LAS OPERACIONES REGISTRADAS
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

Private Enum COLUMNA_
    COLUMNASEL_ = 1
    COLUMNAITEM_
    COLUMNAUNIMED_
    COLUMNACANTIDAD_
    COLUMNAMOTIVO_
    COLUMNAHORA_
    COLUMNAIDITEM_
End Enum

Const ESTADOPENDIENTE_ = 1
Const ESTADOPROCESADO_ = 2
Const ESTADOAPROBADO_ = 3
Const ESTADOANULADO_ = 4
Dim ESTADOANTERIOR_ As Double

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

Private Sub cmd_Click(Index As Integer)
    Dim xRs As New ADODB.Recordset
    Dim xRsAux As New ADODB.Recordset
    Dim nTitulo As String
    Dim xCampos() As String
    Dim nSQLId As String
        
    If QueHace = 3 Then Exit Sub
    
    Select Case Index
        Case 0 ' -------------------------------ALMACEN
            If QueHace = 3 Then Exit Sub
            
            Dim xform As New eps_librerias.FormBuscar
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
            
            TxtIdAlm.Text = xRs("id")
            LblAlmacen.Caption = xRs("descripcion")
            LblAlmacen.ToolTipText = xRs("descripcion")
            TxtIdTipDocRef.SetFocus
            Set xRs = Nothing
        
        Case 4 ' Tipo de documento de referencia' BUSCA EL TIPO DE DOCUMENTO
            ReDim xCampos(2, 4) As String
            
            If QueHace = 3 Then Exit Sub
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
                  
            nTitulo = "Buscando Tipo de Documento de Referencia"
            
            cSQL = "SELECT mae_documento.* FROM mae_documento WHERE (id = 1 OR id = 9)"
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "descripcion", "descripcion", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            TxtIdTipDocRef.Text = NulosN(xRs("id"))
            LblTipDocRef.Caption = NulosC(xRs("descripcion"))
            txtNumDocRef.SetFocus
            Set xRs = Nothing
            
        Case 6 ' Numero de Referencia
            ReDim xCampos(3, 4) As String
            
            If QueHace = 3 Then Exit Sub
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Fecha":        xCampos(0, 1) = "fecha":        xCampos(0, 2) = "1000":         xCampos(0, 3) = "D"
            xCampos(1, 0) = "Documento":    xCampos(1, 1) = "numdoc":       xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
            xCampos(2, 0) = "Cliente":      xCampos(2, 1) = "descli":       xCampos(2, 2) = "5000":         xCampos(2, 3) = "C"
                  
            nTitulo = "Buscando " & NulosC(LblTipDocRef.Caption)
            
            Select Case NulosN(TxtIdTipDocRef)
                Case 1 ' factura
                    cSQL = "SELECT vta_ventas.id, vta_ventas.fchreg AS fecha, [vta_ventas].[numser] & '-' & [vta_ventas].[numdoc] AS numdoc, mae_cliente.nombre AS descli " _
                        + vbCr + "FROM vta_ventas LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id;"
                
                Case 9 ' Guia de remision
                    cSQL = "SELECT vta_guia.id, vta_guia.fecgiro AS fecha, [vta_guia].[numser] & '-' & [vta_guia].[numdoc] AS numdoc, mae_cliente.nombre AS descli " _
                        + vbCr + "FROM vta_guia LEFT JOIN mae_cliente ON vta_guia.idcli = mae_cliente.id;"
                
            End Select
                        
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "fecha", "fecha", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            lbliddocref.Caption = NulosN(xRs("id"))
            txtNumDocRef.Text = NulosC(xRs("numdoc"))
            
            ' Se carga el detalle
            Select Case NulosN(TxtIdTipDocRef)
                Case 1 ' factura
                    cSQL = "SELECT vta_ventasdet.iditem, alm_inventario.descripcion AS desitem, vta_ventasdet.idunimed, mae_unidades.abrev AS desunimed " _
                        + vbCr + "FROM (vta_ventasdet LEFT JOIN alm_inventario ON vta_ventasdet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON vta_ventasdet.idunimed = mae_unidades.id " _
                        + vbCr + "WHERE (((vta_ventasdet.idvta)=" & NulosN(xRs("id")) & "));"
                
                Case 9 ' Guia de remision
                    cSQL = "SELECT vta_guiadet.iditem, alm_inventario.descripcion AS desitem, vta_guiadet.idunimed, mae_unidades.abrev AS desunimed " _
                        + vbCr + "FROM (vta_guiadet LEFT JOIN alm_inventario ON vta_guiadet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON vta_guiadet.idunimed = mae_unidades.id " _
                        + vbCr + "WHERE (((vta_guiadet.idgui)=" & NulosN(xRs("id")) & "));"
                
            End Select
            
            Set xRs = Nothing
            RST_Busq xRs, cSQL, xCon
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            xRs.MoveFirst
            While Not xRs.EOF
                Fg1.Rows = Fg1.Rows + 1
                Fg1.TextMatrix(Fg1.Rows - 1, COLUMNASEL_) = -1
                Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAITEM_) = NulosC(xRs("desitem"))
                Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAUNIMED_) = NulosN(xRs("idunimed"))
                Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAIDITEM_) = NulosN(xRs("iditem"))
                xRs.MoveNext
            Wend
            
    End Select
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
    
    cSQL = "SELECT pla_empleados.nombre AS apenom, pla_empleados.id " _
        + vbCr + "FROM pla_empleados " _
        + vbCr + "ORDER BY pla_empleados.nombre;"
    
    xform.SQLCad = cSQL
    
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
        TxtTipDoc.SetFocus
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
        TxtNumSer.SetFocus
        
'        If NulosN(TxtIdAlm.Text) <> 0 Then
'            Dim Rst As New ADODB.Recordset
'            Set Rst = BuscaConCriterio("SELECT * FROM alm_numseries WHERE idtipdoc = " & NulosN(TxtTipDoc.Text) & " AND idalm = " & NulosN(TxtIdAlm.Text) & "", xCon)
'            If Rst.RecordCount <> 0 Then
'                TxtNumSer.Text = NulosC(Rst("numser"))
'                TxtNumSer_Validate True
'            Else
'                TxtNumSer.Text = ""
'                TxtNumDoc.Text = ""
'            End If
'            Set Rst = Nothing
'        End If

    End If
    
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstDev
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA LA COLUMNA SELECCIONADA DEL CONTROL DataGrid Dg1
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstDev.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
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
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then VerMovimientos1 IdMenuActivo, NulosN(RstDev("id")), xCon
End Sub

Private Sub Fg1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
        Case COLUMNASEL_
            Cancel = False
        
        Case Else
            If NulosN(Fg1.TextMatrix(Row, COLUMNASEL_)) = 0 Then
                Cancel = True
            Else
                Cancel = False
            End If
        
    End Select
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim NUEVO_ As Boolean
    
    If Agregando = True Then Exit Sub
    If Row <= 0 Then Exit Sub
    
    Select Case Col
        Case COLUMNACANTIDAD_
            Fg1.TextMatrix(Row, Col) = Format(NulosN(Fg1.TextMatrix(Row, Col)), FORMAT_CANTIDAD)
        
        Case COLUMNAHORA_
            If IsDate(Fg1.TextMatrix(Row, Col)) Then
                Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), FORMAT_HORA_SIN_SEGUNDO)
            Else
                MsgBox "Ingrese una hora correcta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Fg1.TextMatrix(Row, Col) = ""
                Fg1.Col = Col
            End If
            
    End Select
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then Fg1.Editable = flexEDNone: Exit Sub
    If Agregando Then Exit Sub
    If Fg1.Rows - 1 <= Fg1.FixedRows Then Exit Sub
    If Fg1.Row = Fg1.Rows - 1 Then Exit Sub
    
    Select Case Fg1.Col
        Case COLUMNAMOTIVO_, COLUMNACANTIDAD_, COLUMNAUNIMED_, _
                                COLUMNASEL_, COLUMNAHORA_
            Fg1.Editable = flexEDKbdMouse

        Case Else
            Fg1.Editable = flexEDNone
    End Select
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case Col
        Case COLUMNAUNIMED_, COLUMNACANTIDAD_
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
            
        Case COLUMNASEL_, COLUMNAITEM_
            KeyAscii = 0
    End Select
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    
    If KeyCode = 45 Then
        cmd_Click 3
    End If
    If KeyCode = 46 Then
        cmd_Click 2
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
        
    Fg1.ColWidth(COLUMNAIDITEM_) = 0
    Fg1.ColEditMask(COLUMNAHORA_) = "##:##"
    
    Fg1.AllowUserResizing = flexResizeColumns
    Fg1.ExplorerBar = flexExSortShow
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.ForeColorSel = &H80000005
    Fg1.BackColorSel = &H80&
    Fg1.WordWrap = True
    
    ESTADOANTERIOR_ = 1
    llenarMotivosUnidades Fg1, COLUMNAMOTIVO_, COLUMNAUNIMED_
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
    Fg1.SelectionMode = flexSelectionFree
    
    xHorIni = Time
    TxtFchIng.Valor = Date
    TxtFchIng.SetFocus
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A AJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    QueHace = 3
    iniciarCampos
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario miestras este agregando o editando un registro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If RstDev.State = 0 Then Exit Sub
        If RstDev.RecordCount = 0 And QueHace <> 1 Then
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
            RstDev.Requery
            Dg1.Refresh
            Cancelar
            
            If RstDev.RecordCount <> 0 Then
                RstDev.MoveFirst
                RstDev.Find "id=" & mIdRegistro
                If RstDev.EOF = True Then RstDev.MoveFirst
            End If
            
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then
        RstDev.Filter = ""
        TDB_FiltroLimpiar Dg1
    End If
    
    If Button.Index = 11 Then
        mMesActivo = SeleccionaMes(xCon)
        pCargarDatos
    End If
    
    If Button.Index = 16 Then
        Unload Me
        Set RstDev = Nothing
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

Private Function crearLote(iDITEM_ As Double, FCHING_ As String) As String
    Dim LOTE_ As String
    Dim A As Integer
    Dim MAXLIS_ As Integer
    Dim MAXBD_ As Integer
    Dim NUMERO_ As Integer
    Dim xRs As New ADODB.Recordset
    
    LOTE_ = Format(iDITEM_, "0000") & Format(CDate(FCHING_), "yy") & Format(Month(CDate(FCHING_)), "00") & Format(Day(CDate(FCHING_)), "00")
        
    ' Se verifica el mayor en la base de datos
    cSQL = "SELECT Max(Mid([alm_inventariolote].[descripcion],11,2)) AS orden, alm_inventariolote.iditem, alm_inventariolote.fching " _
        + vbCr + "FROM alm_inventariolote " _
        + vbCr + "GROUP BY alm_inventariolote.iditem, alm_inventariolote.fching " _
        + vbCr + "HAVING (((alm_inventariolote.iditem)=" & iDITEM_ & ") AND ((alm_inventariolote.fching)=CDate('" & FCHING_ & "')))"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    MAXBD_ = 0
    If xRs.State = 0 Then GoTo SALIR_
    If xRs.RecordCount = 0 Then GoTo SALIR_
    
    MAXBD_ = NulosN(xRs("orden"))
SALIR_:

    NUMERO_ = 0
    NUMERO_ = MAXBD_ + 1
        
    LOTE_ = LOTE_ & Format(NUMERO_, "00")
    
    crearLote = LOTE_
End Function

Private Sub crearNotaCredito(TIPDOCREF_ As Integer, IDDOCREF As Integer, FECHA_ As String, _
                                                    IMPOTBRU_ As Double)
    Dim xRs As New ADODB.Recordset
    If TIPDOCREF_ = 9 Then ' ----------------------------Guia
        ' Se verifica que la guia este facturada y a su vez con NC
        cSQL = "SELECT vta_guia.iddocven AS idfac, vta_ventas.id AS idnc " _
            + vbCr + "FROM vta_guia LEFT JOIN vta_ventas ON vta_guia.iddocven = vta_ventas.iddocref " _
            + vbCr + "WHERE (((vta_guia.iddocven) Is Not Null And (vta_guia.iddocven)<>0) AND ((vta_guia.id)=" & IDDOCREF & "));"
        
        Set xRs = Nothing
        RST_Busq xRs, cSQL, xCon
        
        If xRs.State = 0 Then
            MsgBox "Ha ocurrido un Error al crear la NC"
            Exit Sub
        End If
        If xRs.RecordCount = 0 Then
            MsgBox "Ha ocurrido un Error al crear la NC"
            Exit Sub
        End If
        
        ' Si es facturado
        ' Se verfica si tiene nota de credito
        If NulosN(xRs("idnc")) = 0 Then
            If MsgBox("Desea crear la Nota de Crédito para el Ítem ingresado?", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
            If Not grabarNotaCredito(NulosN(xRs("idfac")), FECHA_, IMPOTBRU_) Then
                MsgBox "Ha ocurrido un Error al crear la NC"
            End If
        Else
            Exit Sub
        End If
    ElseIf TIPDOCREF_ = 1 Then '---------------------------Factura
        ' Se busca que la Factura tenga o no notas de credito
        cSQL = "SELECT vta_ventas.id AS idnc " _
            + vbCr + "FROM vta_ventas " _
            + vbCr + "WHERE (((vta_ventas.iddocref)=" & IDDOCREF & "));"
        
        Set xRs = Nothing
        RST_Busq xRs, cSQL, xCon
        
        If xRs.State = 0 Then
            MsgBox "Ha ocurrido un Error al crear la NC"
            Exit Sub
        End If
        
        If xRs.RecordCount = 0 Then ' No tiene NC
            If MsgBox("Desea crear la Nota de Crédito para el Ítem ingresado?", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
            If Not grabarNotaCredito(IDDOCREF, FECHA_, IMPOTBRU_) Then
                MsgBox "Ha ocurrido un Error al crear la NC"
            End If
        Else
            Exit Sub
        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : Function
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA Alm_ingreso, DEVUELVE VERDADERO SI TIENE EXITO
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Function Grabar() As Boolean
    Dim xId As Double
    Dim xIdDet As Double
    Dim A As Integer
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    Dim IDGUIA_ As Integer
    
    ' VALIDAMOS QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
    If Year(TxtFchIng.Valor) <> AnoTra Then
        MsgBox "El año ingresado en la " & Label3(3).Caption & " no coincide con el Ejercicio" & vbCr & "Corrija la fecha o registre en su año que corresponde", vbInformation, xTitulo
        TxtFchIng.Valor = ""
        TxtFchIng.SetFocus
        Exit Function
    End If
    
    If Not IsDate(TxtFchIng.Valor) Then
        MsgBox "No ha especificado la fecha de ingreso", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIng.SetFocus
        Exit Function
    End If
        
    If NulosC(TxtTipDoc.Text) = "" Then
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc.SetFocus
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
       
    If Fg1.Rows = 2 Then
        If NulosN(Fg1.TextMatrix(1, COLUMNASEL_)) = 0 And _
                                        NulosN(Fg1.TextMatrix(1, COLUMNAIDITEM_)) = 0 Then
            MsgBox "Falta ingresar el detalle...!", vbExclamation, "Mensaje...!"
            Exit Function
        End If
    ElseIf Fg1.Rows = 1 Then
        MsgBox "Falta ingresar el detalle...!", vbExclamation, "Mensaje...!"
        Exit Function
    End If
    
'On Error GoTo LaCague
    xCon.BeginTrans
    
    If QueHace = 1 Then
        xId = HallaCodigoTabla("alm_devolucion", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM alm_devolucion", xCon
        RST_Busq RstDet, "SELECT TOP 1 * FROM alm_devoluciondet", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = NulosN(RstDev("id"))
        RST_Busq RstCab, "SELECT * FROM alm_devolucion WHERE id = " & RstDev("id") & "", xCon
        xCon.Execute "DELETE * FROM alm_devoluciondet WHERE iddev = " & RstDev("id") & " "
        RST_Busq RstDet, "SELECT * FROM alm_devoluciondet", xCon
    End If
    
    mIdRegistro = xId
    xIdDet = HallaCodigoTabla("alm_devoluciondet", xCon, "id")
        
    RstCab("fching") = TxtFchIng.Valor
    RstCab("tipdoc") = NulosN(TxtTipDoc.Text)
    RstCab("numser") = NulosC(TxtNumSer.Text)
    RstCab("numdoc") = NulosC(TxtNumDoc.Text)
    RstCab("idresp") = NulosN(TxtIdRes.Text)
    RstCab("idtipdocref") = NulosN(TxtIdTipDocRef.Text)
    RstCab("iddocref") = NulosN(lbliddocref.Caption)
    RstCab("idalm") = NulosN(TxtIdAlm.Text)
    RstCab("ano") = AnoTra
    RstCab("idmes") = mMesActivo
    RstCab.Update
    
    Dim IMPOTBRU_ As Double
    IMPOTBRU_ = 0
    For A = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(A, COLUMNASEL_)) = 0 Then GoTo SIGUIENTE_
        RstDet.AddNew
        RstDet("id") = xIdDet
        RstDet("iddev") = xId
        RstDet("iditem") = NulosN(Fg1.TextMatrix(A, COLUMNAIDITEM_))
        RstDet("idmotdev") = NulosN(Fg1.TextMatrix(A, COLUMNAMOTIVO_))
        RstDet("idunimed") = NulosN(Fg1.TextMatrix(A, COLUMNAUNIMED_))
        RstDet("cantidad") = NulosN(Fg1.TextMatrix(A, COLUMNACANTIDAD_))
        RstDet("hora") = NulosC(Fg1.TextMatrix(A, COLUMNAHORA_))
        xIdDet = xIdDet + 1
        IMPOTBRU_ = IMPOTBRU_ + NulosN(RstDet("cantidad"))
        RstDet.Update
SIGUIENTE_:
    Next A
    
    Dim IDDOCREF As Integer
    Dim FECHA_ As String
    Dim TIPDOCREF_ As Integer
        
    IDDOCREF = NulosN(lbliddocref.Caption)
    FECHA_ = TxtFchIng.Valor
    TIPDOCREF_ = NulosN(TxtIdTipDocRef.Text)
    
    ' Se crea el Ingreso en Almacen
    If GrabarIngreso(114, CInt(xId)) Then
        ' Se crea la NC
        crearNotaCredito TIPDOCREF_, IDDOCREF, FECHA_, IMPOTBRU_
    Else
        GoTo LaCague
    End If
    
SALIR_:
    ' grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    xCon.CommitTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    
    Grabar = True
    MsgBox "El registro se guardo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Exit Function

LaCague:
    'Resume
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function

Sub preparaRST(ByRef RST_ As ADODB.Recordset)
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(11, 3) As String
    ' N: Numerico
    ' D: Double
    ' F: Fecha
    ' C: Caracter
    ' L: Logico
    xCampos(0, 0) = "iditem":           xCampos(0, 1) = "N":      xCampos(0, 2) = ""
    xCampos(1, 0) = "cantidad":         xCampos(1, 1) = "D":      xCampos(1, 2) = ""
    xCampos(2, 0) = "cantteo":          xCampos(2, 1) = "D":      xCampos(2, 2) = ""
    xCampos(3, 0) = "idtipo":           xCampos(3, 1) = "N":      xCampos(3, 2) = ""
    xCampos(4, 0) = "idlote":           xCampos(4, 1) = "N":      xCampos(4, 2) = ""
    xCampos(5, 0) = "idlotedet":        xCampos(5, 1) = "N":      xCampos(5, 2) = ""
    xCampos(6, 0) = "canant":           xCampos(6, 1) = "D":      xCampos(6, 2) = ""
    xCampos(7, 0) = "idalm":            xCampos(7, 1) = "N":      xCampos(7, 2) = ""
    xCampos(8, 0) = "idloteant":        xCampos(8, 1) = "N":      xCampos(8, 2) = ""
    xCampos(9, 0) = "idlotedetant":     xCampos(9, 1) = "N":      xCampos(9, 2) = ""
    xCampos(10, 0) = "hora":            xCampos(10, 1) = "F":     xCampos(10, 2) = ""
    
    Set RST_ = xFun.CrearRstTMP(xCampos)
    RST_.Open
End Sub

Private Function GrabarIngreso(IDTIPDOCREFING_ As Integer, IDDOCREFING_ As Integer) As Boolean
    Dim FCHMOV_ As String
    Dim TIPDOC_ As Integer
    Dim NUMSER_ As String
    Dim IDRESP_ As Integer
    Dim IDPROV_ As Integer
    Dim DESPROV_ As String
    Dim IDESTADO_ As Integer
    Dim IDTIPMOV_ As Integer
    Dim IDTIPDOCREF_ As Integer
    Dim IDDOCREF_ As Integer
    Dim IDING_ As Integer
    Dim NUMDOC_ As String
    Dim iDITEM_ As Integer
    Dim IDALM_ As Integer
    Dim xRs As New ADODB.Recordset
    Dim xRsAux As New ADODB.Recordset
    Dim A As Integer
            
    ' Se busca si esq hay datos que modificar de una ingreso anterior
    cSQL = "SELECT alm_ingreso.id, alm_ingreso.numdoc, alm_ingresodet.iditem, alm_inventariolotedet.idlote, alm_ingresodet.idlotedet, alm_ingresodet.cantidad " _
        + vbCr + "FROM (alm_ingreso LEFT JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN alm_inventariolotedet ON alm_ingresodet.idlotedet = alm_inventariolotedet.id " _
        + vbCr + "WHERE (((alm_ingreso.idtipdocref)=" & IDTIPDOCREFING_ & ") AND ((alm_ingreso.iddocref)=" & IDDOCREFING_ & "));"
    
    RST_Busq xRsAux, cSQL, xCon
    If xRsAux.State = 0 Then Exit Function
    ' Se llenan los detalles
    FCHMOV_ = Format(TxtFchIng.Valor, "dd/mm/yyyy")
    TIPDOC_ = 71
    NUMSER_ = "0001"
    IDRESP_ = NulosN(TxtIdRes.Text)
    IDPROV_ = 0
    DESPROV_ = ""
    IDESTADO_ = ESTADOPROCESADO_
    IDTIPMOV_ = -1
    IDTIPDOCREF_ = IDTIPDOCREFING_
    IDDOCREF_ = IDDOCREFING_
    IDALM_ = 4
    ' Se prepara el Recordset
    If xRs.State = 0 Then preparaRST xRs
    limpiarRST xRs
    ' Se llena el recordset
    For A = 1 To Fg1.Rows - 1
        iDITEM_ = NulosN(Fg1.TextMatrix(A, COLUMNAIDITEM_))
        xRsAux.Filter = adFilterNone
        xRsAux.Filter = "iditem=" & iDITEM_
        xRs.AddNew
        xRs("iditem") = iDITEM_
        'xRs("idtipo") = NulosN(Busca_Codigo(iDITEM_, "id", "tippro", "alm_inventario", "N", xCon))
        'xRs("idalm") = 4
        xRs("cantidad") = NulosN(Fg1.TextMatrix(A, COLUMNACANTIDAD_))
        xRs("hora") = NulosC(Fg1.TextMatrix(A, COLUMNAHORA_))
        If xRsAux.RecordCount = 0 Then
            IDING_ = 0
            xRs("idlote") = 0
            xRs("idlotedet") = 0
            xRs("idloteant") = 0
            xRs("idlotedetant") = 0
            xRs("canant") = 0
            NUMDOC_ = ""
        Else
            IDING_ = NulosN(xRsAux("id"))
            xRs("idlote") = NulosN(xRsAux("idlote"))
            xRs("idlotedet") = NulosN(xRsAux("idlotedet"))
            xRs("idloteant") = NulosN(xRsAux("idlote"))
            xRs("idlotedetant") = NulosN(xRsAux("idlotedet"))
            xRs("canant") = NulosN(xRsAux("cantidad"))
            NUMDOC_ = NulosC(xRsAux("numdoc"))
        End If
        xRs.Update
    Next A
    
    ' Se graba el movimiento
    GrabarIngreso = grabarMovimiento(FCHMOV_, TIPDOC_, NUMSER_, "", IDRESP_, IDPROV_, DESPROV_, IDESTADO_, _
                                IDTIPMOV_, IDTIPDOCREF_, IDDOCREF_, IDALM_, xRs, IDING_, NUMDOC_, 6)
End Function

Sub ActualizaSaldoDoc(idDocumento As Double, Tabla As Integer, ImporteRestar As Double)
    '1 = compras
    '2 = Ventas
    '3 = honorarios
    
    Dim Rst As New ADODB.Recordset
    Dim Total As Double
    
    If Tabla = 2 Then
        RST_Busq Rst, "SELECT Sum(tes_cajadestinodet.acuenta) AS total FROM tes_caja LEFT JOIN tes_cajadestinodet ON tes_caja.id = tes_cajadestinodet.idtes " _
            & " GROUP BY tes_cajadestinodet.iddoc, tes_caja.tipmov HAVING (((tes_cajadestinodet.iddoc)=" & idDocumento & ") AND ((tes_caja.tipmov)=1))", xCon
            
        Total = BuscaImporteDocumento(idDocumento, 1)
        
    End If
    
    If Rst.RecordCount <> 0 Then
        Total = ((Total - Rst("total")) - ImporteRestar)
    Else
        Total = (Total - ImporteRestar)
    End If
    
    xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = " & Total & " WHERE (((vta_ventas.id)=" & idDocumento & "))"
    
    Set Rst = Nothing
End Sub

Function BuscaImporteDocumento(idDocumento As Double, Tabla As Integer) As Double
    '1 = compras
    '2 = Ventas
    '3 = honorarios
    Dim Rst As New ADODB.Recordset
    
    'compras
    If Tabla = 1 Then RST_Busq Rst, "SELECT * FROM vta_ventas WHERE id = " & idDocumento & "", xCon
    
    If Rst.RecordCount <> 0 Then
        BuscaImporteDocumento = Rst("imptotdoc")
    Else
        BuscaImporteDocumento = 0
    End If
    
    Set Rst = Nothing
End Function

Private Function grabarNotaCredito(IDFACT_ As Integer, FECHA_ As String, _
                                            IMPORTBRU_ As Double) As Boolean
    Dim A As Integer
    Dim nSQL As String
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xIdCuen As Integer
    Dim xTotal As Double
    Dim xSaldo As Double
    Dim xidtipven As String
    Dim xNumAsiento As String
    Dim xId As Double
    Dim X As Integer
    Dim P As Integer
    Dim xRs As New ADODB.Recordset
    Dim iDITEM_ As Integer
    Dim IDUNIMED_ As Integer
    Dim CANTIDAD_ As Double
    Dim NUMDOC_ As Double
    Dim IDMOTDEV_ As Integer
    
On Error GoTo LaCague
    xCon.BeginTrans
    
    xId = HallaCodigoTabla("vta_ventas", xCon, "id")
    xNumAsiento = NuevoNumAsiento(2, mMesActivo, xCon)
    RST_Busq RstCab, "SELECT TOP 1 * FROM vta_ventas", xCon
    
    RstCab.AddNew
    RstCab("id") = xId
    
    RST_Busq RstDet, "SELECT TOP 1 * FROM vta_ventasdet", xCon
    
    mIdRegistro = xId
    
    ' ESCRIBIMOS LA CABECERA DEL REGISTRO
    '*********************************************
    ' Se hallan los datos de la factura
    cSQL = "SELECT vta_ventas.* " _
        + vbCr + "From vta_ventas " _
        + vbCr + "WHERE (((vta_ventas.id)=" & IDFACT_ & "));"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    If xRs.State = 0 Then GoTo LaCague
    If xRs.RecordCount = 0 Then GoTo LaCague
    
    NUMDOC_ = HallaCodigoTabla("vta_ventas", xCon, "numdoc")
    '*********************************************
    
    RstCab("idlib") = 2
    RstCab("idtipo") = 3
    RstCab("tipdoc") = 7
    RstCab("idcli") = NulosN(xRs("idcli"))
    RstCab("numser") = NulosC(xRs("numser"))
    RstCab("numdoc") = Format(NUMDOC_, "0000000000")
    RstCab("fchreg") = FECHA_
    RstCab("fchdoc") = FECHA_
    RstCab("fchven") = Format(CDate(FECHA_) + (CDate(xRs("fchven")) - CDate(xRs("fchreg"))), FORMAT_DATE)
    RstCab("idconpag") = NulosN(xRs("idconpag"))
    RstCab("idmon") = NulosN(xRs("idmon"))
    RstCab("impbru") = IMPORTBRU_
    RstCab("impigv") = IMPORTBRU_ * NulosN(xRs("tasaigv")) / 100
    RstCab("imptotdoc") = NulosN(RstCab("impbru")) + NulosN(RstCab("impigv"))
    RstCab("idalm") = NulosN(xRs("idalm"))
    RstCab("tipdes") = 1
    RstCab("iddocref") = IDFACT_
    RstCab("idmotnotcre") = 4 ' Devolucion
    RstCab("idmotdev") = IDMOTDEV_
    RstCab("anulado") = 0
    RstCab("idtipven") = 0
    RstCab("oriitem") = 1
    RstCab("tc") = NulosN(xRs("tc"))
    RstCab("tasaigv") = NulosN(xRs("tasaigv"))
        
    ' Actualizamos el saldo del documento
    ActualizaSaldoDoc CDbl(IDFACT_), 2, NulosN(RstCab("imptotdoc"))
    RstCab.Update
    
    ' GRABAMOS EL DETALLE DEL REGISTRO
    '*********************************************
    ' Se hallan los datos del detalle de la factura
    cSQL = "SELECT vta_ventasdet.* " _
        + vbCr + "FROM vta_ventasdet " _
        + vbCr + "WHERE (((vta_ventasdet.idvta)=" & IDFACT_ & "));"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    If xRs.State = 0 Then GoTo LaCague
    If xRs.RecordCount = 0 Then GoTo LaCague
    '*********************************************
    For A = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(A, COLUMNASEL_)) = 0 Then GoTo SIGUIENTE_
        iDITEM_ = NulosN(Fg1.TextMatrix(A, COLUMNAIDITEM_))
        IDUNIMED_ = NulosN(Fg1.TextMatrix(A, COLUMNAUNIMED_))
        CANTIDAD_ = NulosN(Fg1.TextMatrix(A, COLUMNACANTIDAD_))
        IDMOTDEV_ = NulosN(Fg1.TextMatrix(A, COLUMNAMOTIVO_))
        xRs.Filter = adFilterNone
        xRs.Filter = "iditem=" & iDITEM_
        If xRs.RecordCount = 0 Then GoTo LaCague
        
        RstDet.AddNew
        RstDet("idvta") = xId
        RstDet("iditem") = iDITEM_
        RstDet("idunimed") = IDUNIMED_
        RstDet("preuni") = NulosN(xRs("preuni"))
        RstDet("canpro") = CANTIDAD_
        RstDet("imptot") = NulosN(RstDet("canpro")) * NulosN(RstDet("preuni"))
        RstDet("preunibru") = NulosN(xRs("preuni"))
        RstDet("idmotdev") = IDMOTDEV_
        RstDet.Update
        
        ' Actualizamos el Stock
        xCon.Execute "UPDATE alm_inventario SET alm_inventario.stckact = ( [alm_inventario]![stckact]-" & CANTIDAD_ & ")" _
            & " WHERE (((alm_inventario.id)=" & iDITEM_ & "))"
        
SIGUIENTE_:
    Next A
    
    '---generar asiento
    xNumAsiento = GenerarAsiento(xCon, 2, xId, AnoTra, mMesActivo, 1)
    If xNumAsiento = "" Then GoTo LaCague
        
    ' Grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, 18, 6, xHorIni, Time, Date, xCon, xId
    
    GrabarOperacionCtaCte 2, xId, xCon
    
    xCon.CommitTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    grabarNotaCredito = True
    Exit Function
    
LaCague:
    'Resume
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    grabarNotaCredito = False
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
    Dim xRs As New ADODB.Recordset
    Dim xId As Integer
    
    TabOne1.CurrTab = 0
    If RstDev.State = 0 Then Exit Sub
    If RstDev.RecordCount = 0 Then
        MsgBox "No hay Registros de devolución de Almacén para eliminar", vbInformation, xTitulo
        Exit Sub
    End If
    
    Rpta = MsgBox("¿Esta seguro de eliminar la Devolución Nº " + Trim(RstDev("numdoc")) + "?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        ' Se elimina los registros relacionados
        ' LOTES
        cSQL = "SELECT alm_ingresodet.id, alm_ingresodet.idlotedet, alm_inventariolotedet.idlote, alm_inventariolote.cantidad " _
            + vbCr + "FROM alm_inventariolote RIGHT JOIN ((alm_ingresodet RIGHT JOIN alm_ingreso ON alm_ingresodet.id = alm_ingreso.id) LEFT JOIN alm_inventariolotedet ON alm_ingresodet.idlotedet = alm_inventariolotedet.id) ON alm_inventariolote.id = alm_inventariolotedet.idlote " _
            + vbCr + "WHERE (((alm_ingreso.idtipdocref)=114) AND ((alm_ingreso.iddocref)=" & NulosN(RstDev("id")) & "));"
            
        Set xRs = Nothing
        RST_Busq xRs, cSQL, xCon
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then GoTo SIGUIENTE_
        xRs.MoveFirst
        xId = NulosN(xRs("id"))
        While Not xRs.EOF
            ' SE LIMPIAN LOTES
            xCon.Execute "DELETE * FROM alm_inventariolotedet WHERE id = " & NulosN(xRs("idlotedet"))
            xCon.Execute "UPDATE alm_inventariolote " _
                + vbCr + "SET alm_inventariolote.cantidad = alm_inventariolote.cantidad-" & NulosN(xRs("cantidad")) & " " _
                + vbCr + "WHERE (((alm_inventariolote.id)=" & NulosN(xRs("idlote")) & "));"
            xRs.MoveNext
        Wend
        
        ' NOTAS DE CREDITO
        ' --------------------------------------------------------------------------SI ES GUIA
        If NulosN(RstDev("idtipdocref")) = 9 Then
            cSQL = "SELECT vta_guia.iddocven AS idfac, vta_ventas.id AS idnc " _
                + vbCr + "FROM vta_guia LEFT JOIN vta_ventas ON vta_guia.iddocven = vta_ventas.iddocref " _
                + vbCr + "WHERE (((vta_guia.iddocven) Is Not Null And (vta_guia.iddocven)<>0) AND ((vta_guia.id)=" & NulosN(RstDev("iddocref")) & "));"

        ElseIf NulosN(RstDev("idtipdocref")) = 1 Then ' -------------------------------------SI ES FACTURA
            cSQL = "SELECT vta_ventas.id As idnc" _
                + vbCr + "FROM vta_ventas " _
                + vbCr + "WHERE (((vta_ventas.tipdoc)=7) AND ((vta_ventas.iddocref)=" & NulosN(RstDev("iddocref")) & "));"
        End If

        Set xRs = Nothing
        RST_Busq xRs, cSQL, xCon
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then GoTo SIGUIENTE_
        
        ' ELIMINAMOS EL ASIENTO CONTABLE DE LA TABLA con_diario
        xCon.Execute "DELETE * FROM con_diario WHERE idlib = 2 AND idmov = " & xRs("idnc") & ""
        '--eliminamos el registro del analisis de cta cte
        xCon.Execute "DELETE * FROM var_analisisctacte WHERE idlib = 2 AND idope = " & xRs("idnc") & ""
        ' ELIMINAMOS EL REGISTRO
        xCon.Execute "DELETE * FROM vta_ventasdetitems WHERE idventa = " & xRs("idnc") & ""
        xCon.Execute "DELETE * FROM vta_ventasdet WHERE idvta = " & xRs("idnc") & ""
        xCon.Execute "DELETE * FROM vta_ventas WHERE id = " & xRs("idnc") & ""
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & xRs("idnc") & " AND idform = " & 18
        
        ' MOVIMIENTOS
        xCon.Execute "DELETE * FROM alm_ingresodet WHERE id = " & xId
        xCon.Execute "DELETE * FROM alm_ingreso WHERE id = " & xId
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & xId & " AND idform = " & 8
        
SIGUIENTE_:
        ' Se elimina la devolucion
        xCon.Execute "DELETE * FROM alm_devoluciondet WHERE iddev = " & NulosN(RstDev("id"))
        xCon.Execute "DELETE * FROM alm_devolucion WHERE id = " & NulosN(RstDev("id"))
        
        ' Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & NulosN(RstDev("id")) & " AND idform = " & IdMenuActivo

        RstDev.Requery
        Dg1.Refresh
        MsgBox "La devolución se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
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
    TxtFchIng.Valor = ""
    TxtTipDoc.Text = ""
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    TxtIdRes.Text = ""
    LblResp.Caption = ""
    LblTipDoc.Caption = ""
    TxtIdAlm.Text = ""
    LblAlmacen.Caption = ""
    TxtIdTipDocRef.Text = ""
    LblTipDocRef.Caption = ""
    lbliddocref.Caption = ""
    txtNumDocRef.Text = ""
    TxtTipDoc.Text = "71"
    TxtTipDoc_Validate True
    Fg1.Rows = Fg1.FixedRows
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
    TxtTipDoc.Locked = Not TxtTipDoc.Locked
    TxtNumSer.Locked = Not TxtNumSer.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    TxtIdRes.Locked = Not TxtIdRes.Locked
    TxtIdAlm.Locked = Not TxtIdAlm.Locked
    TxtIdTipDocRef.Locked = Not TxtIdTipDocRef.Locked
    txtNumDocRef.Locked = Not txtNumDocRef.Locked
End Sub

Private Sub txtIdAlm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        LblAlmacen.Caption = ""
    End If
End Sub

Private Sub txtIdAlm_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        cmd_Click 0
    End If
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
        CmdBusRes_Click
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
        LblResp.Caption = ""
    Else
        LblResp.Caption = NulosC(xRs("nomsolic"))
    End If
    
    Set xRs = Nothing
End Sub

Private Sub TxtIdTipDocRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        LblTipDocRef.Caption = ""
    End If
End Sub


Private Sub TxtIdTipDocRef_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        cmd_Click 4
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
    If NulosC(TxtNumDoc.Text) = "" Then Exit Sub
    
    TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000000")
End Sub

Private Sub txtNumDocRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub txtNumDocRef_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        cmd_Click 6
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
    If NulosC(TxtNumSer.Text) <> "" Then
        TxtNumSer.Text = Format(TxtNumSer.Text, "0000")
        If TxtNumDoc.Text = "" Then
            TxtNumDoc.Text = HallarNumIngresoAlmacen("alm_recepcion", TxtNumSer.Text)
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
Function HallarNumIngresoAlmacen(TABLA_ As String, NumSerie As String) As String
    Dim vFiltro As String
    Dim Rst As New ADODB.Recordset
    Dim xNum As Double
      
    cSQL = "SELECT * FROM " & TABLA_ & " " _
        + vbCr + "WHERE numser = '" & NulosC(NumSerie) & "' AND tipdoc = " & NulosN(TxtTipDoc.Text) & "" _
        + vbCr + "ORDER BY numdoc"
    
    RST_Busq Rst, cSQL, xCon
    
    If Rst.RecordCount = 0 Then
        HallarNumIngresoAlmacen = "0000000001"
    Else
        Rst.MoveLast
        xNum = NulosN(Rst("numdoc")) + 1
        HallarNumIngresoAlmacen = Format(xNum, "0000000000")
    End If
    
    Set Rst = Nothing
End Function

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
End Sub

Private Sub llenarMotivosUnidades(ByRef FGGRID As VSFlexGrid, COLUMNAMOTIVO_ As Integer, _
                                                                        COLUMNAUNIMED_ As Integer)
    Dim CAMPOS As String
    Dim xRs As New ADODB.Recordset
    
    ' Se llenan las Unidades
    CAMPOS = ""
    cSQL = "SELECT * FROM mae_motivodevolucion ORDER BY id"
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then
        MsgBox "No se ha encontrado Motivos de Devolución, haga clic en Aceptar y agreguelas", _
                                                                        vbInformation, xTitulo
        Exit Sub
    End If
    
    xRs.MoveFirst
    CAMPOS = "#" & NulosN(xRs("id")) & ";" & UCase(NulosC(xRs("descripcion")))
    xRs.MoveNext
    While Not xRs.EOF
        CAMPOS = CAMPOS & "|#" & NulosN(xRs("id")) & ";" & UCase(NulosC(xRs("descripcion")))
        xRs.MoveNext
    Wend
    FGGRID.ColComboList(COLUMNAMOTIVO_) = CAMPOS
    
    ' Se llenan los Motivos
    CAMPOS = ""
    cSQL = "SELECT * FROM mae_unidades ORDER BY id"
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then
        MsgBox "No se ha encontrado Unidades, haga clic en Aceptar y agreguelas", _
                                                                    vbInformation, xTitulo
        Exit Sub
    End If
    
    xRs.MoveFirst
    CAMPOS = "#" & NulosN(xRs("id")) & ";" & UCase(NulosC(xRs("abrev")))
    xRs.MoveNext
    While Not xRs.EOF
        CAMPOS = CAMPOS & "|#" & NulosN(xRs("id")) & ";" & UCase(NulosC(xRs("abrev")))
        xRs.MoveNext
    Wend
    FGGRID.ColComboList(COLUMNAUNIMED_) = CAMPOS
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EL DETALLE DEL REGISTRO ACTUAL EN LA PESTAÑA DETALLE DEL FORMULARIO
'* Parametros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    Dim A As Integer
    Dim xRs As New ADODB.Recordset
    
    Blanquea
    
    If RstDev.RecordCount = 0 Then Exit Sub
    If RstDev.BOF = True Or RstDev.EOF = True Then Exit Sub
    
    TxtFchIng.Valor = RstDev("fching")
    TxtNumSer.Text = NulosC(RstDev("numser"))
    TxtNumDoc.Text = NulosC(RstDev("numdoc"))
    TxtIdRes.Text = NulosN(RstDev("idresp"))
    LblResp.Caption = Busca_Codigo(NulosN(RstDev("idresp")), "id", "nombre", "pla_empleados", "N", xCon)
    TxtIdAlm.Text = NulosN(RstDev("idalm"))
    LblAlmacen.Caption = Busca_Codigo(NulosN(RstDev("idalm")), "id", "descripcion", "alm_almacenes", "N", xCon)
    TxtTipDoc.Text = NulosN(RstDev("tipdoc"))
    LblTipDoc.Caption = Busca_Codigo(NulosN(RstDev("tipdoc")), "id", "descripcion", "mae_documento", "N", xCon)
    TxtIdTipDocRef.Text = NulosN(RstDev("idtipdocref"))
    LblTipDocRef.Caption = Busca_Codigo(NulosN(RstDev("idtipdocref")), "id", "descripcion", "mae_documento", "N", xCon)
    txtNumDocRef.Text = NulosC(RstDev("numdocref"))
    lbliddocref.Caption = NulosN(RstDev("iddocref"))
    
    cSQL = "SELECT alm_devoluciondet.*, alm_inventario.descripcion AS desitem, mae_unidades.abrev AS desunimed, mae_motivodevolucion.descripcion AS desmotdev " _
        + vbCr + "FROM ((alm_devoluciondet LEFT JOIN alm_inventario ON alm_devoluciondet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON alm_devoluciondet.idunimed = mae_unidades.id) LEFT JOIN mae_motivodevolucion ON alm_devoluciondet.idmotdev = mae_motivodevolucion.id " _
        + vbCr + "WHERE (((alm_devoluciondet.iddev)=" & NulosN(RstDev("id")) & "));"
    
    RST_Busq xRs, cSQL, xCon
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    
    Agregando = True
    xRs.MoveFirst
    For A = 1 To xRs.RecordCount
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(A, COLUMNASEL_) = -1
        Fg1.TextMatrix(A, COLUMNAITEM_) = NulosC(xRs("desitem"))
        Fg1.TextMatrix(A, COLUMNAUNIMED_) = NulosN(xRs("idunimed"))
        Fg1.TextMatrix(A, COLUMNACANTIDAD_) = Format(NulosN(xRs("cantidad")), FORMAT_CANTIDAD)
        Fg1.TextMatrix(A, COLUMNAMOTIVO_) = NulosN(xRs("idmotdev"))
        Fg1.TextMatrix(A, COLUMNAIDITEM_) = NulosN(xRs("iditem"))
        Fg1.TextMatrix(A, COLUMNAHORA_) = Format(xRs("hora"), FORMAT_HORA_SIN_SEGUNDO)
        
        xRs.MoveNext
    Next A
    
    Agregando = False
End Sub

Sub pCargarDatos()
     Dim NomMes As String
     Dim Cerrado As Boolean
     
    LblPeriodo.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
    
    TDB_FiltroLimpiar Dg1
    Set RstDev = Nothing

    cSQL = "SELECT [alm_devolucion].[id] & '' AS id, [alm_devolucion].[fching] & '' AS fching, alm_devolucion.tipdoc, mae_documento_1.abrev AS destipdoc, alm_devoluciondet.iditem, alm_inventario.descripcion AS desitem, [alm_devolucion].[numser] & '-' & [alm_devolucion].[numdoc] AS numdoctot, alm_devolucion.idtipdocref, mae_documento.abrev AS destipdocref, alm_devolucion.iddocref, IIf([alm_devolucion].[idtipdocref]=9,[vta_guia].[numser] & '-' & [vta_guia].[numdoc],[vta_ventas].[numser] & '-' & [vta_ventas].[numdoc]) AS numdocref, alm_devolucion.idresp, alm_devolucion.numser, alm_devolucion.numdoc, alm_devolucion.idalm " _
        + vbCr + "FROM (((((alm_devolucion LEFT JOIN alm_devoluciondet ON alm_devolucion.id = alm_devoluciondet.iddev) LEFT JOIN alm_inventario ON alm_devoluciondet.iditem = alm_inventario.id) LEFT JOIN mae_documento ON alm_devolucion.idtipdocref = mae_documento.id) LEFT JOIN mae_documento AS mae_documento_1 ON alm_devolucion.tipdoc = mae_documento_1.id) LEFT JOIN vta_guia ON alm_devolucion.iddocref = vta_guia.id) LEFT JOIN vta_ventas ON alm_devolucion.iddocref = vta_ventas.id " _
        + vbCr + "WHERE (((alm_devolucion.ano) = " & AnoTra & ") And ((alm_devolucion.idmes) = " & mMesActivo & ")) " _
        + vbCr + "ORDER BY [alm_devolucion].[fching] & '' DESC;"
        
    RST_Busq RstDev, cSQL, xCon
        
    Set Dg1.DataSource = RstDev
End Sub


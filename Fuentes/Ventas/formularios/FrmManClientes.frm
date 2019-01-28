VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmManClientes 
   Caption         =   "Ventas  - Mantenimiento de Clientes"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7575
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
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManClientes.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManClientes.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManClientes.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManClientes.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManClientes.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManClientes.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManClientes.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManClientes.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManClientes.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManClientes.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManClientes.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManClientes.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManClientes.frx":277E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7215
      Left            =   0
      TabIndex        =   18
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
         TabIndex        =   31
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6420
            Left            =   30
            TabIndex        =   32
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
            Columns(1).Caption=   "Nº R.U.C."
            Columns(1).DataField=   "numruc"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nombre"
            Columns(2).DataField=   "nombre"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Tipo Empresa"
            Columns(3).DataField=   "tipemp"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Teléfono"
            Columns(4).DataField=   "tel"
            Columns(4).NumberFormat=   "Short Date"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Fax"
            Columns(5).DataField=   "fax"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Contacto"
            Columns(6).DataField=   "nomcon"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   4
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Activo"
            Columns(7).DataField=   "activo"
            Columns(7).NumberFormat=   "General Number"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=8"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2117"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2037"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=5821"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=5741"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=2223"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2143"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=512"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=1799"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1720"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=1773"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1693"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=4815"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=4736"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=512"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=1164"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=1085"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=513"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=70,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=0"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=0"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=2"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(68)  =   "Named:id=33:Normal"
            _StyleDefs(69)  =   ":id=33,.parent=0"
            _StyleDefs(70)  =   "Named:id=34:Heading"
            _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(72)  =   ":id=34,.wraptext=-1"
            _StyleDefs(73)  =   "Named:id=35:Footing"
            _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(75)  =   "Named:id=36:Selected"
            _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(77)  =   "Named:id=37:Caption"
            _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(79)  =   "Named:id=38:HighlightRow"
            _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(81)  =   "Named:id=39:EvenRow"
            _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(83)  =   "Named:id=40:OddRow"
            _StyleDefs(84)  =   ":id=40,.parent=33"
            _StyleDefs(85)  =   "Named:id=41:RecordSelector"
            _StyleDefs(86)  =   ":id=41,.parent=34"
            _StyleDefs(87)  =   "Named:id=42:FilterBar"
            _StyleDefs(88)  =   ":id=42,.parent=33"
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Clientes"
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
            TabIndex        =   34
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
            TabIndex        =   33
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
         TabIndex        =   19
         Top             =   375
         Width           =   11790
         Begin VB.Frame Frame5 
            Caption         =   "[ Datos para Emisión de Letras]"
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
            Height          =   1215
            Left            =   255
            TabIndex        =   62
            Top             =   4800
            Width           =   11250
            Begin VB.TextBox TxtLetNomGir 
               Height          =   300
               Left            =   1635
               Locked          =   -1  'True
               MaxLength       =   25
               TabIndex        =   66
               Text            =   "TxtLetNomGir"
               Top             =   210
               Width           =   4140
            End
            Begin VB.TextBox TxtLetDirGir 
               Height          =   300
               Left            =   1635
               Locked          =   -1  'True
               MaxLength       =   25
               TabIndex        =   65
               Text            =   "TxtLetDirGir"
               Top             =   510
               Width           =   4140
            End
            Begin VB.TextBox TxtLetTel 
               Height          =   300
               Left            =   1635
               Locked          =   -1  'True
               MaxLength       =   25
               TabIndex        =   64
               Text            =   "TxtLetTel"
               Top             =   810
               Width           =   1470
            End
            Begin VB.TextBox TxtLetNumDoc 
               Height          =   300
               Left            =   7185
               Locked          =   -1  'True
               MaxLength       =   25
               TabIndex        =   63
               Text            =   "TxtLetNumDoc"
               Top             =   510
               Width           =   1470
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Nombre"
               Height          =   195
               Index           =   17
               Left            =   90
               TabIndex        =   70
               Top             =   255
               Width           =   555
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Dirección"
               Height          =   195
               Index           =   18
               Left            =   90
               TabIndex        =   69
               Top             =   555
               Width           =   675
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Teléfono"
               Height          =   195
               Index           =   19
               Left            =   90
               TabIndex        =   68
               Top             =   855
               Width           =   630
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Nº Documento"
               Height          =   195
               Left            =   5925
               TabIndex        =   67
               Top             =   555
               Width           =   1050
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "[ Seleccionar ]"
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
            Height          =   645
            Left            =   255
            TabIndex        =   60
            Top             =   6090
            Width           =   11250
            Begin VB.CheckBox ChkAgeRet 
               Caption         =   "Agente de Retención"
               Height          =   195
               Left            =   255
               TabIndex        =   61
               Top             =   285
               Width           =   1815
            End
         End
         Begin VB.CommandButton CmdBusDoc 
            Height          =   240
            Left            =   2445
            Picture         =   "FrmManClientes.frx":2B10
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   3870
            Width           =   240
         End
         Begin VB.CommandButton CmdBusTipDoc2 
            Height          =   240
            Left            =   2445
            Picture         =   "FrmManClientes.frx":2C42
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   885
            Width           =   240
         End
         Begin VB.CommandButton CmdBusVen 
            Height          =   240
            Left            =   2445
            Picture         =   "FrmManClientes.frx":2D74
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   4470
            Width           =   240
         End
         Begin VB.TextBox TxtContac 
            Height          =   300
            Left            =   7350
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   16
            Text            =   "TxtContac"
            Top             =   3840
            Width           =   4140
         End
         Begin VB.TextBox TxtPagWeb 
            Height          =   300
            Left            =   7350
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   15
            Text            =   "TxtPagWeb"
            Top             =   4140
            Width           =   4140
         End
         Begin VB.CommandButton CmdBusDep 
            Height          =   240
            Left            =   10380
            Picture         =   "FrmManClientes.frx":2EA6
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   3285
            Width           =   240
         End
         Begin VB.TextBox TxtDepartamento 
            Height          =   300
            Left            =   7350
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   10
            Text            =   "TxtDepartamento"
            Top             =   3240
            Width           =   3300
         End
         Begin VB.CommandButton CmdBusDis 
            Height          =   240
            Left            =   4830
            Picture         =   "FrmManClientes.frx":2FD8
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   3270
            Width           =   240
         End
         Begin VB.TextBox TxtDistrito 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   9
            Text            =   "TxtDistrito"
            Top             =   3240
            Width           =   3300
         End
         Begin VB.TextBox TxtFax 
            Height          =   300
            Left            =   7350
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   12
            Text            =   "TxtFax"
            Top             =   3540
            Width           =   1470
         End
         Begin VB.TextBox TxtDir 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   8
            Text            =   "TxtDir"
            Top             =   2940
            Width           =   6720
         End
         Begin VB.TextBox TxtNombre 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   3
            Text            =   "TxtNombre"
            Top             =   1455
            Width           =   6705
         End
         Begin VB.Frame Frame3 
            Caption         =   "( Datos Persona Natural )"
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
            Height          =   1035
            Left            =   255
            TabIndex        =   36
            Top             =   1830
            Width           =   11250
            Begin VB.TextBox TxtApe2 
               Height          =   300
               Left            =   5415
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   7
               Text            =   "TxtApe2"
               Top             =   615
               Width           =   2205
            End
            Begin VB.TextBox TxtApe1 
               Height          =   300
               Left            =   1560
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   6
               Text            =   "TxtApe1"
               Top             =   615
               Width           =   2205
            End
            Begin VB.TextBox TxtNom2 
               Height          =   300
               Left            =   5415
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   5
               Text            =   "TxtNom2"
               Top             =   300
               Width           =   2205
            End
            Begin VB.TextBox TxtNom1 
               Height          =   300
               Left            =   1560
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   4
               Text            =   "TxtNom1"
               Top             =   300
               Width           =   2205
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Apellido 2"
               Height          =   195
               Index           =   11
               Left            =   4515
               TabIndex        =   41
               Top             =   660
               Width           =   690
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Apellido 1"
               Height          =   195
               Index           =   10
               Left            =   225
               TabIndex        =   40
               Top             =   660
               Width           =   690
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Nombre 2"
               Height          =   195
               Index           =   9
               Left            =   4515
               TabIndex        =   39
               Top             =   345
               Width           =   690
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Nombre 1"
               Height          =   195
               Index           =   8
               Left            =   225
               TabIndex        =   38
               Top             =   345
               Width           =   690
            End
         End
         Begin VB.CommandButton CmdBusProv 
            Height          =   240
            Left            =   2445
            Picture         =   "FrmManClientes.frx":310A
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   585
            Width           =   240
         End
         Begin VB.TextBox TxtTele 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   11
            Text            =   "TxtTele"
            Top             =   3540
            Width           =   1470
         End
         Begin VB.TextBox TxtEmail 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   14
            Text            =   "TxtEmail"
            Top             =   4140
            Width           =   4140
         End
         Begin VB.TextBox TxtNumRuc 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   2
            Text            =   "TxtNumRuc"
            Top             =   1155
            Width           =   1770
         End
         Begin VB.TextBox TxtTipPer 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   0
            Text            =   "TxtTipPer"
            Top             =   555
            Width           =   915
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   2010
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   22
            Text            =   "TxtNumSer"
            Top             =   2070
            Width           =   915
         End
         Begin VB.CommandButton CmdBusTipDoc 
            Height          =   240
            Left            =   2445
            Picture         =   "FrmManClientes.frx":323C
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   585
            Width           =   240
         End
         Begin VB.TextBox TxtIdVen 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   17
            Text            =   "TxtIdVen"
            Top             =   4440
            Width           =   915
         End
         Begin VB.TextBox TxtIdTipDoc2 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   1
            Text            =   "Tx"
            Top             =   855
            Width           =   915
         End
         Begin VB.TextBox TxtidCondPag 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   13
            Text            =   "TxtidCondPag"
            Top             =   3840
            Width           =   915
         End
         Begin VB.Label LblCondPag 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblCondPag"
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
            Left            =   2730
            TabIndex        =   59
            Top             =   3840
            Width           =   3210
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Condición de Pago"
            Height          =   195
            Index           =   3
            Left            =   255
            TabIndex        =   58
            Top             =   3900
            Width           =   1350
         End
         Begin VB.Label LblDecTipDoc2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDecTipDoc2"
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
            TabIndex        =   56
            Top             =   855
            Width           =   5760
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento"
            Height          =   195
            Index           =   16
            Left            =   255
            TabIndex        =   55
            Top             =   885
            Width           =   1185
         End
         Begin VB.Label LblVendedor 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblVendedor"
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
            TabIndex        =   53
            Top             =   4440
            Width           =   5745
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor"
            Height          =   195
            Index           =   2
            Left            =   255
            TabIndex        =   52
            Top             =   4485
            Width           =   690
         End
         Begin VB.Label LblIdDep 
            AutoSize        =   -1  'True
            Caption         =   "LblIdDep"
            Height          =   195
            Left            =   9705
            TabIndex        =   50
            Top             =   1515
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label LblIdDis 
            AutoSize        =   -1  'True
            Caption         =   "LblIdDis"
            Height          =   195
            Left            =   8670
            TabIndex        =   49
            Top             =   1515
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Contacto"
            Height          =   195
            Left            =   6495
            TabIndex        =   48
            Top             =   3900
            Width           =   645
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
            Height          =   195
            Index           =   15
            Left            =   6135
            TabIndex        =   47
            Top             =   3285
            Width           =   1005
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Distrito"
            Height          =   195
            Index           =   14
            Left            =   255
            TabIndex        =   45
            Top             =   3285
            Width           =   480
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fax"
            Height          =   195
            Index           =   13
            Left            =   6885
            TabIndex        =   43
            Top             =   3585
            Width           =   255
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Dirección"
            Height          =   195
            Index           =   12
            Left            =   255
            TabIndex        =   42
            Top             =   2970
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Razón Social"
            Height          =   195
            Index           =   6
            Left            =   255
            TabIndex        =   37
            Top             =   1485
            Width           =   945
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "E - mail"
            Height          =   195
            Index           =   5
            Left            =   255
            TabIndex        =   30
            Top             =   4185
            Width           =   510
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono"
            Height          =   195
            Index           =   4
            Left            =   255
            TabIndex        =   29
            Top             =   3585
            Width           =   630
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Documento"
            Height          =   195
            Index           =   0
            Left            =   465
            TabIndex        =   28
            Top             =   2100
            Width           =   1275
         End
         Begin VB.Label LblTipoPersona 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipoPersona"
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
            TabIndex        =   27
            Top             =   555
            Width           =   5745
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Persona"
            Height          =   195
            Index           =   1
            Left            =   255
            TabIndex        =   26
            Top             =   585
            Width           =   945
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº R.U.C."
            Height          =   195
            Index           =   7
            Left            =   255
            TabIndex        =   25
            Top             =   1185
            Width           =   705
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Cliente"
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
            TabIndex        =   24
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Pág. Web"
            Height          =   195
            Left            =   6420
            TabIndex        =   23
            Top             =   4200
            Width           =   720
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   35
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar cliente"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Activar cliente"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Eliminar cliente"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Retirar cliente"
               EndProperty
            EndProperty
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
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
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
Attribute VB_Name = "FrmManClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RstPro As New ADODB.Recordset
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim CaracteresNumericos As String
Dim xHorIni As Date

Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO
Dim mIdRegistro&                     ' identificador del registro

Sub Cancelar()
    QueHace = 3
    Bloquea
    Label5.Caption = "Detalle cliente"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    ActivaTool
    Dg1.SetFocus
End Sub

Sub Nuevo()
    QueHace = 1
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Agregando cliente"
    ActivaTool
    Blanquea
    Bloquea
    xHorIni = Time
    TxtTipPer.SetFocus
End Sub

Sub Modificar()
    QueHace = 2
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Modificando cliente"
    ActivaTool
    Blanquea
    Bloquea
    MuestraSegundoTab
    xHorIni = Time
    TxtTipPer.SetFocus
End Sub

Private Sub CmdBusDep_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_departamento"
    
    xform.Titulo = "Buscando Departamentos"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        LblIdDep.Caption = xRs("id")
        txtDepartamento.Text = xRs("descripcion")
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusDis_Click()
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_distrito"
    
    xform.Titulo = "Buscando Distritos"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        LblIdDis.Caption = xRs("id")
        txtDistrito.Text = xRs("descripcion")
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusDoc_Click()
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_condpago"
    
    xform.Titulo = "Buscando condicion de pago"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtidCondPag.Text = xRs("id")
        LblCondPag.Caption = xRs("descripcion")
        TxtEmail.SetFocus
    Else
        TxtidCondPag.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusProv_Click()
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_tipoempresa"
    
    xform.Titulo = "Buscando Tipo de Persona"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtTipPer.Text = xRs("id")
        LblTipoPersona.Caption = xRs("descripcion")
        
        If NulosC(TxtTipPer.Text) = 1 Then
            TxtNombre.Text = ""
            Frame3.Enabled = True
            TxtNombre.Enabled = False
        Else
            TxtNom1.Text = ""
            TxtNom2.Text = ""
            TxtApe1.Text = ""
            TxtApe2.Text = ""
            Frame3.Enabled = False
            TxtNombre.Enabled = True
        End If
        TxtNumRuc.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipDoc2_Click()
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_dociden"
    
    xform.Titulo = "Buscando Documentos de Indentidad"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdTipDoc2.Text = xRs("id")
        LblDecTipDoc2.Caption = xRs("descripcion")
        TxtNumRuc.SetFocus
        
'        If NulosC(Busca_Documento(TxtIdTipDoc2.Text, TxtNumRuc.Text)) <> "" Then
'            MsgBox "El numero de documento registrado ya existe en el maestro de proveedores, esta registrado a" & Chr(13) _
'                & "nombre de: " + NulosC(Busca_Documento(TxtIdTipDoc2.Text, TxtNumRuc.Text)), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'            TxtNumRuc.Text = ""
'        End If
        
    Else
        TxtNumRuc.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusVen_Click()
    If QueHace = 3 Then Exit Sub
    If xDeDonde = 2 Then Exit Sub '--unificado salir
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "apenom":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":          xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT vta_vendedores.id, UCase([pla_empleados]![apepat] & ' ' & [pla_empleados]![apemat] +', '+[pla_empleados]![nom]) AS apenom FROM vta_vendedores LEFT JOIN pla_empleados " _
        & " ON vta_vendedores.idper = pla_empleados.id ORDER BY [pla_empleados]![apepat] & ' ' & [pla_empleados]![apemat] +', '+[pla_empleados]![nom] "

    xform.Titulo = "Buscando Vendedor"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "apenom"
    xform.CampoBusca = "apenom"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdVen.Text = xRs("id")
        Lblvendedor.Caption = xRs("apenom")
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstPro
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstPro.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstPro("id")), xCon
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
        RST_Busq RstPro, "SELECT mae_cliente.*, mae_distrito.descripcion AS nomdis, mae_departamento.descripcion AS nomdep, " _
            & " mae_tipoempresa.descripcion AS tipemp " _
            & " FROM ((mae_tipoempresa RIGHT JOIN mae_cliente ON mae_tipoempresa.id = mae_cliente.tipper) LEFT JOIN mae_distrito " _
            & " ON mae_cliente.iddis = mae_distrito.id) LEFT JOIN mae_departamento ON mae_cliente.iddep = mae_departamento.id" _
            & " WHERE mae_cliente.id<>0 ORDER BY nombre", xCon

        Set Dg1.DataSource = RstPro

    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then '--F3 Nuevo
        If QueHace <> 3 Then Exit Sub
        Nuevo
    End If
    
    If KeyCode = 115 Then '--F4 Modificar
        If QueHace <> 3 Then Exit Sub
        If RstPro.RecordCount = 0 Then Exit Sub
        Modificar
    End If
    
    If KeyCode = 113 Then '--F2 Grabar
        If QueHace = 3 Then Exit Sub
        If Grabar = True Then
            Cancelar
            RstPro.Requery
            Dg1.Refresh
        End If
    End If
    
    If KeyCode = 116 Then '--F5 actualizar
        
    
    End If
    
    If KeyCode = 117 Then '--F6 '--cancelar
        If QueHace = 3 Then Exit Sub
        Cancelar
    End If
    
End Sub

Private Sub Form_Load()
    QueHace = 3
    SeEjecuto = False
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    CaracteresNumericos = "0123456789." & Chr(8)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario mientras este agregando o modificando un cliente", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then MuestraSegundoTab
    End If
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    Dim Rst As New ADODB.Recordset
    
    If xDeDonde = 2 Then Exit Sub '--es unificado
    
    'RST_Busq Rst, "SELECT com_compras.idpro, com_compras.numser, com_compras.numdoc From com_compras " _
    '    & " WHERE (((com_compras.idpro)=" & RstPro("id") & "))", xCon
    
   ' If Rst.RecordCount <> 0 Then
   '     MsgBox "El cliente que intenta eliminar tiene documentos de compra registrados, " & Chr(13) _
   '         & "No se puede eliminar al cliente seleccionado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
   '     Set Rst = Nothing
   '     Exit Sub
   ' End If
   
   
    
    Rpta = MsgBox("¿ Esta seguro de eliminar el registro seleccionado ?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM  mae_cliente WHERE id = " & RstPro("id") & ""
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstPro("id") & " AND idform = " & IdMenuActivo
        
        RstPro.Requery
        Dg1.Refresh
    End If
End Sub

Sub Buscar()
    TabOne1.CurrTab = 0
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "nombre":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":     xCampos(1, 1) = "numruc":      xCampos(1, 2) = "1500":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_cliente.numruc, mae_cliente.nombre, mae_cliente.id From mae_cliente ORDER BY mae_cliente.nombre"
    
    xform.Titulo = "Buscando Clientes"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        RstPro.MoveFirst
        RstPro.Find "id = " & xRs("id") & ""
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Sub Imprimir()
'    Dim RsRep As New ADODB.Recordset
'
'    RST_Busq RsRep, "SELECT mae_cliente.numruc, mae_cliente.nombre, mae_tipoempresa.descripcion, mae_cliente.tel, " _
'        & " mae_cliente.fax, mae_cliente.nomcon FROM mae_tipoempresa RIGHT JOIN mae_cliente " _
'        & " ON mae_tipoempresa.id = mae_cliente.tipper WHERE mae_cliente.activo = -1 ORDER BY nombre", xCon
'
'    rptCliente.Sections("Sección4").Controls("lblEmp").Caption = NomEmp
'    rptCliente.Sections("Sección4").Controls("lblruc").Caption = NumRUC
'
'    Set rptCliente.DataSource = RsRep
'    Set RsRep = Nothing
'    rptCliente.Width = 11865
'    rptCliente.Height = 7980
'    rptCliente.Show vbModal
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstPro.Requery
            Dg1.Refresh
            If RstPro.RecordCount <> 0 Then
                RstPro.MoveFirst
                RstPro.Find "id=" & mIdRegistro
                If RstPro.EOF = True Then RstPro.MoveFirst
            End If
        End If
    End If
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 8 Then
        Filtrar
    End If
    
    If Button.Index = 9 Then
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstPro.Filter = ""
    End If
    
    If Button.Index = 10 Then
        Buscar
    End If
    
    If Button.Index = 12 Then pExportar
    If Button.Index = 13 Then Imprimir
    
    
    If Button.Index = 15 Then
        Set RstPro = Nothing
        Unload Me
    End If
End Sub

Sub Filtrar()
    TabOne1.CurrTab = 0
    
    'Dim xform As New eps_librerias.FormFiltrar
    Dim xform As New eps_librerias.FormFiltrar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(3, 4) As String
    
    xCampos(0, 0) = "Nombre":    xCampos(0, 1) = "nombre":   xCampos(0, 2) = "C":         xCampos(0, 3) = "4200"
    xCampos(1, 0) = "Nº Ruc":    xCampos(1, 1) = "numruc":   xCampos(1, 2) = "C":         xCampos(1, 3) = "1500"
    xCampos(2, 0) = "Tipo":      xCampos(2, 1) = "tipemp":   xCampos(2, 2) = "C":         xCampos(2, 3) = "1500"
    
    Set xform.Coneccion = xCon
    Set xform.Rst = RstPro        'recorset que llena el grid
    Set RstPro = xform.FiltrarReg(xCampos)
    Set Dg1.DataSource = RstPro
    Dg1.Refresh
End Sub

Sub Activar()
    TabOne1.CurrTab = 0
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "nombre":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":     xCampos(1, 1) = "numruc":      xCampos(1, 2) = "1500":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_cliente.numruc, mae_cliente.nombre, mae_cliente.id From mae_cliente  WHERE activo =0 ORDER BY mae_cliente.nombre"
    
    xform.Titulo = "Buscando Clientes"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        Dim Rpta As Integer
        Rpta = MsgBox("¿Esta seguro de activar al cliente " + Trim(xRs("nombre")) + "?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
        If Rpta = vbYes Then
            xCon.Execute "UPDATE mae_cliente SET mae_cliente.activo = -1 WHERE (((mae_cliente.id)=" & NulosN(xRs("id")) & "))"
            
            'grabamos el movimiento en la tabla var_edicion
            GrabarOperacion xIdUsuario, IdMenuActivo, 2, Time, Time, Date, xCon, NulosN(xRs("id"))
            
            RstPro.Requery
            
            
            Dg1.Refresh
            MsgBox "El cliente se activo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Sub Retirar()
    Dim Rpta As Integer
    If xDeDonde = 2 Then Exit Sub '--es unificado
    Rpta = MsgBox("¿Esta seguro de retirar al cliente " + Trim(RstPro("nombre")) + "?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "UPDATE mae_cliente SET mae_cliente.activo = 0 WHERE (((mae_cliente.id)=" & RstPro("id") & "))"
        
        'grabamos el movimiento en la tabla var_edicion
        GrabarOperacion xIdUsuario, IdMenuActivo, 2, Time, Time, Date, xCon, NulosN(RstPro("id"))
        
        RstPro.Requery

        Dg1.Refresh
        MsgBox "El cliente se retiro con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.Index = 2 Then
        If ButtonMenu.Index = 1 Then
            Modificar
        End If
        If ButtonMenu.Index = 2 Then
            Activar
        End If
    End If
    
    If ButtonMenu.Parent.Index = 3 Then
        If ButtonMenu.Index = 1 Then
            Eliminar
        End If
        If ButtonMenu.Index = 2 Then
            Retirar
        End If
    End If
End Sub

Private Sub TxtApe1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtApe1_Validate(Cancel As Boolean)
    If TxtApe1.Text <> "" Then
        TxtNombre.Text = ""
        TxtNombre.Text = Trim(UCase(TxtApe1.Text)) + " " + Trim(UCase(TxtApe2.Text)) + ", " + Trim(TxtNom1.Text) + " " + Trim(TxtNom2.Text)
    End If
End Sub

Private Sub TxtApe2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtApe2_Validate(Cancel As Boolean)
    If TxtApe2.Text <> "" Then
        TxtNombre.Text = ""
        TxtNombre.Text = Trim(UCase(TxtApe1.Text)) + " " + Trim(UCase(TxtApe2.Text)) + ", " + Trim(TxtNom1.Text) + " " + Trim(TxtNom2.Text)
    End If
End Sub

Private Sub TxtContac_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtDepartamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtDepartamento_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusDep_Click
    End If
End Sub

Private Sub TxtDir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtDistrito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtDistrito_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusDis_Click
    End If
End Sub

Private Sub TxtEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtFax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtidCondPag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtidCondPag_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipDoc2_Click
    End If
End Sub

Private Sub TxtidCondPag_Validate(Cancel As Boolean)
    If NulosC(TxtidCondPag.Text) = "" Then
        LblCondPag.Caption = ""
        Exit Sub
    End If
        
    
    LblCondPag.Caption = Busca_Codigo(TxtidCondPag.Text, "id", "descripcion", "mae_condpago", "N", xCon)
    If LblCondPag.Caption = "" Then
        TxtidCondPag.Text = ""
        TxtEmail.SetFocus
    End If
End Sub

Private Sub TxtIdTipDoc2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdTipDoc2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipDoc2_Click
    End If
End Sub

Private Sub TxtIdTipDoc2_Validate(Cancel As Boolean)
    If NulosC(TxtIdTipDoc2.Text) = "" Then Exit Sub
    
    LblDecTipDoc2.Caption = Busca_Codigo(TxtIdTipDoc2.Text, "id", "descripcion", "mae_dociden", "N", xCon)
    If LblDecTipDoc2.Caption = "" Then
        TxtIdTipDoc2.Text = ""
        TxtIdTipDoc2.SetFocus
    End If

    
    If NulosC(Busca_Documento(NulosN(TxtIdTipDoc2.Text), TxtNumRuc.Text)) <> "" Then
        MsgBox "El numero de documento registrado ya existe en el maestro de Clientes, esta registrado a" & Chr(13) _
            & "nombre de: " + NulosC(Busca_Documento(TxtIdTipDoc2.Text, TxtNumRuc.Text)), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumRuc.Text = ""
    End If
End Sub

Function Busca_Documento(IdtipDoc As Integer, NumDoc As String) As String
    Dim Rst As New ADODB.Recordset
    
    If QueHace = 1 Then
        RST_Busq Rst, "SELECT mae_cliente.idtipdoc, mae_cliente.numruc, mae_cliente.nombre From mae_cliente WHERE (((mae_cliente.idtipdoc)=" & IdtipDoc & ") " _
            & " AND ((mae_cliente.numruc)='" & NulosC(TxtNumRuc.Text) & "'))", xCon
    Else
        If NulosC(TxtNumRuc.Text) <> NulosC(RstPro("numruc")) Then
        
            RST_Busq Rst, "SELECT mae_cliente.idtipdoc, mae_cliente.numruc, mae_cliente.nombre From mae_cliente WHERE (((mae_cliente.idtipdoc)=" & IdtipDoc & ") " _
                & " AND ((mae_cliente.numruc)='" & NulosC(TxtNumRuc.Text) & "'))", xCon
        Else
            Busca_Documento = ""
            Exit Function
        End If
    End If
    

    If Rst.RecordCount = 0 Then
        Busca_Documento = ""
    Else
        Busca_Documento = Rst("nombre")
    End If
    Set Rst = Nothing
End Function

Private Sub TxtIdVen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdVen_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusVen_Click
    End If
End Sub

Private Sub TxtIdVen_Validate(Cancel As Boolean)
    If NulosN(TxtIdVen.Text) <> 0 Then
        Dim Rst As New ADODB.Recordset
        Dim SQL As String
        SQL = "SELECT vta_vendedores.id, UCase([pla_empleados]![apepat] & ' ' & [pla_empleados]![apemat] +', '+[pla_empleados]![nom]) AS apenom FROM vta_vendedores LEFT JOIN pla_empleados " _
            & " ON vta_vendedores.idper = pla_empleados.id Where (((vta_vendedores.id) = " & Val(TxtIdVen.Text) & ")) ORDER BY UCase([pla_empleados]![apepat] & ' ' & [pla_empleados]![apemat] +', '+[pla_empleados]![nom])"

        Set Rst = BuscaConCriterio(SQL, xCon)
        If Rst.RecordCount <> 0 Then
            Lblvendedor.Caption = Rst("apenom")
        Else
            TxtIdVen.Text = ""
            Lblvendedor.Caption = ""
        End If
    Else
        TxtIdVen.Text = ""
        Lblvendedor.Caption = ""
    End If
End Sub

Private Sub TxtLetNomGir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtLetNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtLetTel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNom1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtNom1_Validate(Cancel As Boolean)
    If TxtNom1.Text <> "" Then
        TxtNombre.Text = ""
        TxtNombre.Text = Trim(UCase(TxtApe1.Text)) + " " + Trim(UCase(TxtApe2.Text)) + ", " + Trim(TxtNom1.Text) + " " + Trim(TxtNom2.Text)
    End If
End Sub

Private Sub TxtNom2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtNom2_Validate(Cancel As Boolean)
    If TxtNom1.Text <> "" Then
        TxtNombre.Text = ""
        TxtNombre.Text = Trim(UCase(TxtApe1.Text)) + " " + Trim(UCase(TxtApe2.Text)) + ", " + Trim(TxtNom1.Text) + " " + Trim(TxtNom2.Text)
    End If
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumRuc_KeyPress(KeyAscii As Integer)
    Dim cruc As String
    Dim xRs As New ADODB.Recordset
    If QueHace = 1 Then
        RST_Busq xRs, "Select numruc from mae_cliente where numruc ='" & NulosC(TxtNumRuc.Text) & "'", xCon
    Else
        If NulosC(TxtNumRuc.Text) <> NulosC(RstPro("numruc")) Then
        
            RST_Busq xRs, "Select numruc from mae_cliente where numruc ='" & NulosC(TxtNumRuc.Text) & "'", xCon
        Else
            Set xRs = Nothing
            SendKeys vbTab
            Exit Sub
        End If
    End If
    
    If KeyAscii = 13 Then
        If xRs.RecordCount > 0 Then
            MsgBox "Numero de Ruc se encuentra registrado", vbInformation, Me.Caption
            Me.Cancelar
            Exit Sub
            Else
            SendKeys vbTab
        End If
    End If
    Set xRs = Nothing
End Sub

Private Sub TxtNumRuc_Validate(Cancel As Boolean)
    If NulosC(TxtNumRuc.Text) <> "" Then
        '--validar ingreso de documento solo cuando se a ruc
        If NulosN(TxtIdTipDoc2.Text) = 5 Then
            If Len(NulosC(TxtNumRuc.Text)) <> 11 Then
                MsgBox "El numero de digitos del R.U.C. tiene que ser igual a 11", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                TxtNumRuc.SetFocus
                Cancel = True
                Exit Sub
            End If
            If TxtTipPer.Text = 1 Then
                'persona natural
                If Mid(TxtNumRuc.Text, 1, 1) <> "1" Then
                    MsgBox "El primer digito del Nº R.U.C. no corresponde al de una persona natural", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    TxtNumRuc.SetFocus
                    Cancel = True
                    Exit Sub
                End If
            Else
                'persona juridica
                If Mid(TxtNumRuc.Text, 1, 1) <> "2" Then
                    MsgBox "El primer digito del Nº R.U.C. no corresponde al de una persona juridica", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    TxtNumRuc.SetFocus
                    Cancel = True
                    Exit Sub
                End If
            End If
        End If
    End If
End Sub

Private Sub TxtPagWeb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtTele_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtTipPer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        If NulosC(TxtTipPer.Text) = "" Then Exit Sub
        
        LblTipoPersona.Caption = Busca_Codigo(TxtTipPer.Text, "id", "descripcion", "mae_tipoempresa", "N", xCon)
        If LblTipoPersona.Caption = "" Then
            TxtTipPer.Text = ""
            TxtTipPer.SetFocus
        Else
            If NulosC(TxtTipPer.Text) = 1 Then
                TxtNombre.Text = ""
                Frame3.Enabled = True
                TxtNombre.Enabled = False
            Else
                TxtNom1.Text = ""
                TxtNom2.Text = ""
                TxtApe1.Text = ""
                TxtApe2.Text = ""
                Frame3.Enabled = False
                TxtNombre.Enabled = True
            End If
        End If
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Sub Blanquea()
    TxtTipPer.Text = ""
    TxtIdTipDoc2.Text = ""
    TxtNumRuc.Text = ""
    TxtNombre.Text = ""
    TxtNom1.Text = ""
    TxtNom2.Text = ""
    TxtApe1.Text = ""
    TxtApe2.Text = ""
    TxtDir.Text = ""
    txtDistrito.Text = ""
    txtDepartamento.Text = ""
    TxtTele.Text = ""
    TxtFax.Text = ""
    TxtEmail.Text = ""
    TxtPagWeb.Text = ""
    TxtContac.Text = ""
    TxtIdVen.Text = ""
    TxtidCondPag.Text = ""
    
    TxtLetNomGir.Text = ""
    TxtLetDirGir.Text = ""
    TxtLetNumDoc.Text = ""
    TxtLetTel.Text = ""
    
    LblTipoPersona.Caption = ""
    LblIdDep.Caption = ""
    LblIdDis.Caption = ""
    Lblvendedor.Caption = ""
    LblDecTipDoc2.Caption = ""
    LblCondPag.Caption = ""
End Sub

Sub Bloquea()
    TxtTipPer.Locked = Not TxtTipPer.Locked
    TxtNumRuc.Locked = Not TxtNumRuc.Locked
    TxtNombre.Locked = Not TxtNombre.Locked
    TxtNom1.Locked = Not TxtNom1.Locked
    TxtNom2.Locked = Not TxtNom2.Locked
    TxtApe1.Locked = Not TxtApe1.Locked
    TxtApe2.Locked = Not TxtApe2.Locked
    TxtDir.Locked = Not TxtDir.Locked
    'TxtDistrito.Locked = Not TxtDistrito.Locked
    'TxtDepartamento.Locked = Not TxtDepartamento.Locked
    TxtTele.Locked = Not TxtTele.Locked
    TxtFax.Locked = Not TxtFax.Locked
    TxtEmail.Locked = Not TxtEmail.Locked
    TxtPagWeb.Locked = Not TxtPagWeb.Locked
    TxtContac.Locked = Not TxtContac.Locked
    TxtIdVen.Locked = Not TxtIdVen.Locked
    TxtIdTipDoc2.Locked = Not TxtIdTipDoc2.Locked
    TxtidCondPag.Locked = Not TxtidCondPag.Locked
    
    TxtLetNomGir.Locked = Not TxtLetNomGir.Locked
    TxtLetDirGir.Locked = Not TxtLetDirGir.Locked
    TxtLetNumDoc.Locked = Not TxtLetNumDoc.Locked
    TxtLetTel.Locked = Not TxtLetTel.Locked
    
    Frame4.Enabled = Not Frame4.Enabled
    
End Sub

Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A

End Sub

Sub MuestraSegundoTab()
    Blanquea
    
    If RstPro("tipper") <> 0 Then
        TxtTipPer.Text = NulosN(RstPro("tipper"))
        LblTipoPersona.Caption = NulosC(RstPro("tipemp"))
    End If
    
    TxtNumRuc.Text = NulosC(RstPro("numruc"))
    TxtIdTipDoc2.Text = NulosN(RstPro("idtipdoc"))
    LblDecTipDoc2.Caption = Busca_Codigo(NulosN(TxtIdTipDoc2.Text), "id", "descripcion", "mae_dociden", "N", xCon)
    
    TxtNom1.Text = NulosC(RstPro("nomcli1"))
    TxtNom2.Text = NulosC(RstPro("nomcli2"))
    TxtApe1.Text = NulosC(RstPro("apecli1"))
    TxtApe2.Text = NulosC(RstPro("apecli2"))
    
    TxtDir.Text = NulosC(RstPro("dir"))
    txtDistrito.Text = NulosC(RstPro("nomdis"))
    txtDepartamento.Text = NulosC(RstPro("nomdep"))
    TxtTele.Text = NulosC(RstPro("tel"))
    TxtFax.Text = NulosC(RstPro("fax"))
    TxtEmail.Text = NulosC(RstPro("email"))
    TxtPagWeb.Text = NulosC(RstPro("pagweb"))
    TxtContac.Text = NulosC(RstPro("nomcon"))

    If NulosN(RstPro("idcondpag")) <> 0 Then
        TxtidCondPag.Text = NulosN(RstPro("idcondpag"))
        LblCondPag.Caption = Busca_Codigo(NulosN(TxtidCondPag.Text), "id", "descripcion", "mae_condpago", "N", xCon)
    End If

    If NulosC(TxtTipPer.Text) = 1 Then
        TxtNombre.Text = ""
        Frame3.Enabled = True
        TxtNombre.Enabled = False
    Else
        TxtNom1.Text = ""
        TxtNom2.Text = ""
        TxtApe1.Text = ""
        TxtApe2.Text = ""
        Frame3.Enabled = False
        TxtNombre.Enabled = True
    End If
    If RstPro("ageret") = -1 Then
        ChkAgeRet.Value = 1
    Else
        ChkAgeRet.Value = 0
    End If
    
    TxtIdVen.Text = NulosN(RstPro("idven"))
    TxtIdVen_Validate True
    TxtNombre.Text = NulosC(RstPro("nombre"))
    
    TxtLetNomGir.Text = NulosC(RstPro("letnomgir"))
    TxtLetDirGir.Text = NulosC(RstPro("letgirdir"))
    TxtLetNumDoc.Text = NulosC(RstPro("letnumdoc"))
    TxtLetTel.Text = NulosC(RstPro("lettel"))
End Sub

Function Grabar() As Boolean

  
    If TxtTipPer.Text = "" Then
        MsgBox "No ha especificado el tipo de persona", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipPer.SetFocus
        Exit Function
    End If
    
    If TxtNumRuc.Text = "" Then
        MsgBox "No ha especificado el ruc del cliente", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumRuc.SetFocus
        Exit Function
    End If
    
    If NulosN(TxtIdTipDoc2.Text) = 0 Then
        MsgBox "No ha especificado el tipo de documento de identidad para el cliente", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdTipDoc2.SetFocus
        Exit Function
    End If
    
    If TxtNombre.Text = "" Then
        MsgBox "No ha especificado el nombre del cliente", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        If NulosC(TxtTipPer.Text) = "1" Then
            TxtNom1.SetFocus
        Else
            TxtNombre.SetFocus
        End If
        Exit Function
    End If
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "Grabar", "Modificar") + " el Cliente", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo) = vbNo Then Exit Function
    
    Dim RstCab As New ADODB.Recordset
    Dim xId As Double
    
    On Error GoTo LaCague
    
    xCon.BeginTrans
    
    If QueHace = 1 Then
        xId = HallaCodigoTabla("mae_cliente", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM mae_cliente", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstPro("id")
        RST_Busq RstCab, "SELECT * FROM mae_cliente WHERE id = " & RstPro("id") & "", xCon
    End If
    
    mIdRegistro = xId
    
    RstCab("tipper") = NulosN(TxtTipPer.Text)
    RstCab("idtipdoc") = NulosN(TxtIdTipDoc2.Text)
    RstCab("numruc") = TxtNumRuc.Text
    RstCab("nombre") = TxtNombre.Text
    RstCab("nomcli1") = NulosC(TxtNom1.Text)
    RstCab("nomcli2") = NulosC(TxtNom2.Text)
    RstCab("apecli1") = NulosC(TxtApe1.Text)
    RstCab("apecli2") = NulosC(TxtApe2.Text)
    RstCab("dir") = NulosC(TxtDir.Text)
    If NulosN(LblIdDis.Caption) <> 0 Then RstCab("iddis") = NulosN(LblIdDis.Caption)
    If NulosN(LblIdDep.Caption) <> 0 Then RstCab("iddep") = NulosN(LblIdDep.Caption)
    RstCab("tel") = NulosC(TxtTele.Text)
    RstCab("fax") = NulosC(TxtFax.Text)
    RstCab("email") = NulosC(TxtEmail.Text)
    RstCab("pagweb") = NulosC(TxtPagWeb.Text)
    RstCab("idven") = NulosN(TxtIdVen.Text)
    If NulosN(TxtidCondPag.Text) <> 0 Then
        RstCab("idcondpag") = NulosN(TxtidCondPag.Text)
    End If
    If ChkAgeRet.Value = 1 Then RstCab("ageret") = -1
    If ChkAgeRet.Value = 0 Then RstCab("ageret") = 0
    
    RstCab("letnomgir") = NulosC(TxtLetNomGir.Text)
    RstCab("letgirdir") = NulosC(TxtLetDirGir.Text)
    RstCab("letnumdoc") = NulosC(TxtLetNumDoc.Text)
    RstCab("lettel") = NulosC(TxtLetTel.Text)

    
    RstCab.Update
    
    '*************************************************************************************
    '*** SINCRONIZAR BASE DE DATOS - mae_cliente ***'
    If xDeDonde = 2 Then SincronizarBD xCon, "mae_cliente", xId, QueHace
    '*************************************************************************************

    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId

    xCon.CommitTrans
    
    MsgBox "El Cliente se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    Grabar = True
    Exit Function
    
LaCague:
    'Resume
    xCon.RollbackTrans
    Set RstCab = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function

Private Sub TxtTipPer_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusProv_Click
    End If
End Sub

Private Sub pExportar()
    TabOne1.CurrTab = 0
    

    Dim nSQL As String
    Dim oExport As New SGI2_funciones.formularios
    Dim RstTmp  As New ADODB.Recordset

    nSQL = "SELECT mae_tipoempresa.descripcion AS tipopersona, mae_dociden.abrev AS tipodoc, mae_cliente.numruc, mae_cliente.nombre, mae_cliente.nomcli1, mae_cliente.nomcli2, mae_cliente.apecli1, mae_cliente.apecli2, mae_cliente.dir, mae_cliente.tel, mae_cliente.fax, mae_cliente.nomcon AS contacto, mae_condpago.abrev AS CondPago, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS vendedor, IIf([mae_cliente].[activo]=-1,'Activo','De Baja') AS estado " _
        + vbCr + " FROM (((mae_tipoempresa RIGHT JOIN mae_cliente ON mae_tipoempresa.id = mae_cliente.tipper) LEFT JOIN mae_dociden ON mae_cliente.idtipdoc = mae_dociden.id) LEFT JOIN mae_condpago ON mae_cliente.idcondpag = mae_condpago.id) LEFT JOIN (pla_empleados RIGHT JOIN vta_vendedores ON pla_empleados.id = vta_vendedores.idper) ON mae_cliente.idven = vta_vendedores.id " _
        + vbCr + " ORDER BY mae_cliente.nombre; "
    
    RST_Busq RstTmp, nSQL, xCon
        
    Dim xCampos(14, 3) As String
    
    '0::Nombre a Mostrar;
    '1::nombre de Campo del Rst;
    '2::alineacion(0::derecha, 1::centro, 2::izquierda);
    '3::ancho de columna
    '--obs: el rst puede tener mas columnas solo se consideran los campos del array
    xCampos(0, 0) = "Tipo":         xCampos(0, 1) = "tipopersona":  xCampos(0, 2) = 0:   xCampos(0, 3) = "1500"
    xCampos(1, 0) = "Tipo Doc":     xCampos(1, 1) = "tipodoc":      xCampos(1, 2) = 0:   xCampos(1, 3) = "814"
    xCampos(2, 0) = "Num. Doc":     xCampos(2, 1) = "numruc":       xCampos(2, 2) = 0:   xCampos(2, 3) = "1200"
    xCampos(3, 0) = "Nombre":       xCampos(3, 1) = "nombre":       xCampos(3, 2) = 0:   xCampos(3, 3) = "4000"
    xCampos(4, 0) = "Nombre1":      xCampos(4, 1) = "nomcli1":      xCampos(4, 2) = 0:   xCampos(4, 3) = "1000"
    xCampos(5, 0) = "Nombre2":      xCampos(5, 1) = "nomcli2":      xCampos(5, 2) = 0:   xCampos(5, 3) = "1000"
    xCampos(6, 0) = "Apellido1":    xCampos(6, 1) = "apecli1":      xCampos(6, 2) = 0:   xCampos(6, 3) = "1000"
    xCampos(7, 0) = "Apellido2":    xCampos(7, 1) = "apecli2":      xCampos(7, 2) = 0:   xCampos(7, 3) = "1000"
    xCampos(8, 0) = "Dirección":    xCampos(8, 1) = "dir":          xCampos(8, 2) = 0:   xCampos(8, 3) = "4700"
    xCampos(9, 0) = "Teléfono":     xCampos(9, 1) = "tel":          xCampos(9, 2) = 0:   xCampos(9, 3) = "1600"
    xCampos(10, 0) = "Fax":         xCampos(10, 1) = "fax":         xCampos(10, 2) = 0:  xCampos(10, 3) = "800"
    
    xCampos(11, 0) = "Contacto":   xCampos(11, 1) = "contacto":     xCampos(11, 2) = 0:  xCampos(11, 3) = "1400"
    xCampos(12, 0) = "Cond. Pago": xCampos(12, 1) = "condpago":     xCampos(12, 2) = 0:  xCampos(12, 3) = "1057"
    xCampos(13, 0) = "Vendedor":   xCampos(13, 1) = "vendedor":     xCampos(13, 2) = 0:  xCampos(13, 3) = "1543"
    
    xCampos(14, 0) = "Estado":     xCampos(14, 1) = "estado":       xCampos(14, 2) = 1:  xCampos(14, 3) = "850"
        
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "Lista de Clientes", "", "", "Listado de Clientes", RstTmp, xCampos
    Set oExport = Nothing
    Set RstTmp = Nothing
    
End Sub


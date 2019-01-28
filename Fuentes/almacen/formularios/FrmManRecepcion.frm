VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmManRecepcion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Almacén - Recepción"
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
            Picture         =   "FrmManRecepcion.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRecepcion.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRecepcion.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRecepcion.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRecepcion.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRecepcion.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRecepcion.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRecepcion.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRecepcion.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRecepcion.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRecepcion.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRecepcion.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRecepcion.frx":277E
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
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6600
         Left            =   -12435
         TabIndex        =   23
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6090
            Left            =   0
            TabIndex        =   45
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
            Columns(1).Caption=   "Fch. Recep."
            Columns(1).DataField=   "fching"
            Columns(1).NumberFormat=   "Short Date"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Doc."
            Columns(2).DataField=   "fchdoc"
            Columns(2).NumberFormat=   "Short Date"
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
            Columns(5).Caption=   "Peso Bruto"
            Columns(5).DataField=   "pesbru"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Peso Neto"
            Columns(6).DataField=   "pesnet"
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
            Splits(0)._ColumnProps(14)=   "Column(2).Width=1773"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1693"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=2487"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2408"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=6244"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=6165"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=2619"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=2540"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=2408"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=2328"
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
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
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
            TabIndex        =   24
            Top             =   0
            Width           =   1860
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Consulta de Recepción"
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
            TabIndex        =   25
            Top             =   30
            Width           =   11610
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6600
         Left            =   45
         TabIndex        =   16
         Top             =   375
         Width           =   11790
         Begin VB.Frame Frame3 
            Caption         =   "[ Ítem ]"
            Height          =   555
            Left            =   120
            TabIndex        =   42
            Top             =   1740
            Width           =   11565
            Begin VB.CommandButton cmd 
               Height          =   240
               Index           =   1
               Left            =   1320
               Picture         =   "FrmManRecepcion.frx":2B10
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   210
               Width           =   240
            End
            Begin VB.TextBox txtiditem 
               Height          =   300
               Left            =   675
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   10
               Text            =   "txtiditem"
               Top             =   180
               Width           =   915
            End
            Begin VB.Label lbldesitem 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbldesitem"
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
               Left            =   1635
               TabIndex        =   43
               Top             =   195
               Width           =   9825
            End
         End
         Begin VB.CommandButton cmd 
            Height          =   240
            Index           =   4
            Left            =   7710
            Picture         =   "FrmManRecepcion.frx":2C42
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   1095
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.CommandButton cmd 
            Height          =   240
            Index           =   6
            Left            =   11430
            Picture         =   "FrmManRecepcion.frx":2D74
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   1440
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.CommandButton cmd 
            Height          =   240
            Index           =   0
            Left            =   7710
            Picture         =   "FrmManRecepcion.frx":2EA6
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   750
            Width           =   240
         End
         Begin VB.CommandButton CmdBusRes 
            Height          =   240
            Left            =   7710
            Picture         =   "FrmManRecepcion.frx":2FD8
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   405
            Width           =   240
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1080
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   3
            Text            =   "TxtNumSer"
            Top             =   1050
            Width           =   915
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2175
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   4
            Text            =   "TxtNumDoc"
            Top             =   1050
            Width           =   1440
         End
         Begin VB.CommandButton CmdBusProv 
            Height          =   240
            Left            =   5430
            Picture         =   "FrmManRecepcion.frx":310A
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   1440
            Width           =   240
         End
         Begin VB.Frame Frame4 
            Height          =   615
            Left            =   90
            TabIndex        =   18
            Top             =   5910
            Width           =   11610
            Begin VB.CommandButton cmd 
               Caption         =   "Eliminar Pesaje"
               Enabled         =   0   'False
               Height          =   330
               Index           =   2
               Left            =   1410
               TabIndex        =   14
               Top             =   180
               Width           =   1305
            End
            Begin VB.CommandButton cmd 
               Caption         =   "Agregar Pesaje"
               Enabled         =   0   'False
               Height          =   330
               Index           =   3
               Left            =   90
               TabIndex        =   13
               Top             =   180
               Width           =   1305
            End
            Begin VB.Label lblIdIngreso 
               AutoSize        =   -1  'True
               Caption         =   "lblIdIngreso"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   10290
               TabIndex        =   47
               Top             =   240
               Visible         =   0   'False
               Width           =   810
            End
         End
         Begin VB.CommandButton CmdBusTipDoc 
            Height          =   240
            Left            =   1725
            Picture         =   "FrmManRecepcion.frx":323C
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   750
            Width           =   240
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   3555
            Left            =   90
            TabIndex        =   11
            Top             =   2340
            Width           =   11595
            _cx             =   20452
            _cy             =   6271
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
            Cols            =   15
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmManRecepcion.frx":336E
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
            Left            =   1080
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
            Left            =   7065
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   6
            Text            =   "TxtIdRes"
            Top             =   360
            Width           =   915
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchDoc 
            Height          =   300
            Left            =   4470
            TabIndex        =   1
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
         Begin VB.TextBox TxtIdAlm 
            Height          =   300
            Left            =   7065
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   7
            Text            =   "TxtIdAlm"
            Top             =   705
            Width           =   915
         End
         Begin VB.TextBox TxtTipDoc 
            Height          =   300
            Left            =   1080
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   2
            Text            =   "TxtTipDoc"
            Top             =   705
            Width           =   915
         End
         Begin VB.TextBox TxtProv 
            Height          =   300
            Left            =   1080
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   5
            Text            =   "TxtProv"
            Top             =   1395
            Width           =   4620
         End
         Begin VB.TextBox TxtIdTipDocRef 
            Height          =   300
            Left            =   7065
            Locked          =   -1  'True
            TabIndex        =   8
            Text            =   "TxtIdTipDocRef"
            Top             =   1050
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.TextBox txtNumDocRef 
            Height          =   300
            Left            =   7050
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   9
            Text            =   "txtNumDocRef"
            Top             =   1395
            Visible         =   0   'False
            Width           =   4665
         End
         Begin VB.Label lbliddocref 
            AutoSize        =   -1  'True
            Caption         =   "lbliddocref"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   10770
            TabIndex        =   46
            Top             =   1110
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tip. Doc. Ref."
            Height          =   195
            Index           =   5
            Left            =   6000
            TabIndex        =   41
            Top             =   1095
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label LblTipDocRef 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipDocRef"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   8025
            TabIndex        =   40
            Top             =   1065
            Visible         =   0   'False
            Width           =   3690
         End
         Begin VB.Label LblIdProveedor 
            AutoSize        =   -1  'True
            Caption         =   "LblIdProveedor"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   4620
            TabIndex        =   38
            Top             =   1110
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Doc. Ref."
            Height          =   195
            Index           =   7
            Left            =   6030
            TabIndex        =   37
            Top             =   1440
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Almacen"
            Height          =   195
            Index           =   8
            Left            =   6000
            TabIndex        =   34
            Top             =   750
            Width           =   615
         End
         Begin VB.Label LblAlmacen 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblAlmacen"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   8025
            TabIndex        =   33
            Top             =   720
            Width           =   3690
         End
         Begin VB.Label LblTipDoc 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipDoc"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   2010
            TabIndex        =   31
            Top             =   720
            Width           =   3690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Doc."
            Height          =   195
            Index           =   4
            Left            =   3600
            TabIndex        =   30
            Top             =   405
            Width           =   705
         End
         Begin VB.Label LblResp 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblResp"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   8025
            TabIndex        =   29
            Top             =   375
            Width           =   3690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Responsable"
            Height          =   195
            Index           =   2
            Left            =   6000
            TabIndex        =   28
            Top             =   405
            Width           =   930
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000001&
            BackStyle       =   1  'Opaque
            Height          =   90
            Left            =   2025
            Top             =   1170
            Width           =   105
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Doc."
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   26
            Top             =   1095
            Width           =   795
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Doc."
            Height          =   195
            Index           =   1
            Left            =   105
            TabIndex        =   22
            Top             =   750
            Width           =   930
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   105
            TabIndex        =   21
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Recepción"
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
            Top             =   405
            Width           =   720
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
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
Attribute VB_Name = "FrmManRecepcion"
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

Dim RstRecep As New ADODB.Recordset                  ' RECORDSET PRINCIPAL QUE CARGARA TODAS LAS OPERACIONES REGISTRADAS
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
    COLUMNAHORA_ = 1
    COLUMNAENVASE_
    COLUMNAUNIMED_
    COLUMNAPESOENV_
    COLUMNAPESOPARIHUELA_
    COLUMNANUMEROENV_
    COLUMNAPBRUTOENV_
    COLUMNAPBRUTOTOTAL_
    COLUMNAPNETOTOTAL_
    COLUMNAESTADO_
    COLUMNAOBS_
    COLUMNAIDENV_
    COLUMNAIDUNIMED_
    COLUMNAIDESTADO_
End Enum

Const ESTADOPENDIENTE_ = 1
Const ESTADOPROCESADO_ = 2
Const ESTADOAPROBADO_ = 3
Const ESTADOANULADO_ = 4
Dim ESTADOANTERIOR_ As Double
Dim PROCESADO_ As Boolean

'*****************************************************************************************************
'* Nombre           : Buscar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : BUSCA UN REGISTRO EN EL RECORDSET RstRecep
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
        RstRecep.MoveFirst
        RstRecep.Find "id = " & xRs("id") & ""
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

Private Sub cmd_Click(Index As Integer)
    Dim xRs As New ADODB.Recordset
    Dim xRsAux As New ADODB.Recordset
    Dim nTitulo As String
    Dim xCampos() As String
    Dim nSQLId As String
        
    If QueHace = 3 Then Exit Sub
    
    Select Case Index
        Case 0
            ' PERMITE BUSCAR UN ALMACEN
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
            txtiditem.SetFocus
            Set xRs = Nothing
            
        Case 1
            ' BUSCA UN ITEM
            ReDim xCampos(3, 4) As String
            
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Producto":    xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":      xCampos(1, 1) = "codpro":        xCampos(1, 2) = "1500":         xCampos(1, 3) = "C"
            xCampos(2, 0) = "Uni. Med":    xCampos(2, 1) = "abrev":         xCampos(2, 2) = "1200":         xCampos(2, 3) = "C"
                    
            nTitulo = "Buscando Ítems"
            
            cSQL = "SELECT alm_almacenesdet.iditem, alm_inventario.descripcion, alm_inventario.codpro, mae_unidades.abrev " _
                + vbCr + "FROM ((alm_almacenes INNER JOIN alm_almacenesdet ON alm_almacenes.id = alm_almacenesdet.idalm) INNER JOIN alm_inventario ON alm_almacenesdet.iditem = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
                + vbCr + "WHERE (((alm_almacenes.id) = " & NulosN(TxtIdAlm.Text) & ") And ((alm_almacenes.idtippro) = 0)) " _
                + vbCr + "UNION " _
                + vbCr + "SELECT alm_inventario.id AS iditem, alm_inventario.descripcion, alm_inventario.codpro, mae_unidades.abrev " _
                + vbCr + "FROM (alm_almacenes INNER JOIN alm_inventario ON alm_almacenes.idtippro = alm_inventario.tippro) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
                + vbCr + "WHERE (((alm_almacenes.id) = " & NulosN(TxtIdAlm.Text) & "))"
            
'            cSQL = "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.stckact, alm_inventario.activo " _
'                + vbCr + "FROM mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed " _
'                + vbCr + "WHERE (((alm_inventario.activo)=-1)) " _
'                + vbCr + "ORDER BY alm_inventario.codpro"
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "descripcion", "descripcion", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            txtiditem.Text = NulosN(xRs("iditem"))
            lbldesitem.Caption = NulosC(xRs("descripcion"))
            cmd(3).SetFocus
            Set xRs = Nothing
        
        Case 2 ' Eliminar Pesaje
            If Fg1.Rows = 3 Then Fg1.Rows = 1: Exit Sub
            If Fg1.Rows <= 1 Then Exit Sub
            If Fg1.Row = Fg1.Rows - 1 Then Exit Sub
            If Fg1.TextMatrix(Fg1.Row, COLUMNAESTADO_) > ESTADOPENDIENTE_ Then
                MsgBox "No se puede eliminar un registro en este estado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
            
            Fg1.RemoveItem Fg1.Row
            Fg1.Select Fg1.Row - 1, 1
            Fg1.SetFocus
        
        Case 3 ' Agregar pesaje
            If Fg1.Rows > 2 Then Fg1.Rows = Fg1.Rows - 1
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAESTADO_) = ESTADOPENDIENTE_
            hallarTotales Fg1.Rows - 1, 1
            Fg1.Select Fg1.Rows - 2, 1
            Fg1_EnterCell
            Fg1.SetFocus
        
        Case 4 ' Tipo de documento de referencia' BUSCA EL TIPO DE DOCUMENTO
            ReDim xCampos(2, 4) As String
            
            If QueHace = 3 Then Exit Sub
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
                  
            nTitulo = "Buscando Tipo de Documento de Referencia"
            
            cSQL = "SELECT mae_documento.* FROM mae_documento WHERE (tipo = 1 OR tipo = 3)"
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "descripcion", "descripcion", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            TxtIdTipDocRef.Text = NulosN(xRs("id"))
            LblTipDocRef.Caption = NulosC(xRs("descripcion"))
            txtNumDocRef.SetFocus
            Set xRs = Nothing
            
        Case 6 ' Numero de Referencia
        
    End Select
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
        
    xform.SQLCad = "SELECT mae_prov.id, mae_prov.numruc, mae_prov.nombre, mae_prov.activo From mae_prov " _
        & " Where (((mae_prov.activo) = -1)) ORDER BY mae_prov.nombre"
    xform.Titulo = "Buscando Proveedores"
        
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        LblIdProveedor.Caption = NulosN(xRs("id"))
        TxtProv.Text = NulosC(xRs("nombre"))
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
        TxtIdAlm.SetFocus
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
        
'        TxtIdAlm.SetFocus
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
    TDB_FiltroGenerar Dg1, RstRecep
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
    
    If Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAENVASE_) = "TOTAL" Then
        NUEVO_ = False
    Else
        NUEVO_ = True
    End If
    
    If Not NUEVO_ Then
        Fg1.Rows = Fg1.Rows - 1
    End If
       
    For A = 1 To Fg1.Rows - 1
        TOTALPPARIH_ = TOTALPPARIH_ + NulosN(Fg1.TextMatrix(A, COLUMNAPESOPARIHUELA_))
        TOTALNUMENV_ = TOTALNUMENV_ + NulosN(Fg1.TextMatrix(A, COLUMNANUMEROENV_))
        TOTALPBRUTOENV_ = TOTALPBRUTOENV_ + NulosN(Fg1.TextMatrix(A, COLUMNAPBRUTOENV_))
        TOTALBRUTO_ = TOTALBRUTO_ + NulosN(Fg1.TextMatrix(A, COLUMNAPBRUTOTOTAL_))
        TOTALNETO_ = TOTALNETO_ + NulosN(Fg1.TextMatrix(A, COLUMNAPNETOTOTAL_))
    Next A
    
    Agregando = True
    Fg1.Rows = Fg1.Rows + 1
    Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAENVASE_) = "TOTAL"
    Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAPESOPARIHUELA_) = Format(TOTALPPARIH_, FORMAT_CANTIDAD)
    Fg1.TextMatrix(Fg1.Rows - 1, COLUMNANUMEROENV_) = Format(TOTALNUMENV_, FORMAT_CANTIDAD)
    Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAPBRUTOENV_) = Format(TOTALPBRUTOENV_, FORMAT_CANTIDAD)
    Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAPBRUTOTOTAL_) = Format(TOTALBRUTO_, FORMAT_CANTIDAD)
    Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAPNETOTOTAL_) = Format(TOTALNETO_, FORMAT_CANTIDAD)
    
    Fg1.Select Fg1.Rows - 1, COLUMNAENVASE_, Fg1.Rows - 1, COLUMNAPNETOTOTAL_
    Fg1.FillStyle = flexFillRepeat
    Fg1.CellBackColor = &H8000000F
    Fg1.CellFontBold = True
    Fg1.Select FILA_, COLUMNA_
    
    Agregando = False
    
    'Fg1.SetFocus
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA LA COLUMNA SELECCIONADA DEL CONTROL DataGrid Dg1
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstRecep.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
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
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then VerMovimientos1 IdMenuActivo, NulosN(RstRecep("id")), xCon
End Sub

Private Sub Fg1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Fg1.Rows > Fg1.FixedRows And Row = Fg1.Rows - 1 Then Cancel = True
    If Col = COLUMNAESTADO_ Then
        ' Se llena el estado anterior
        ESTADOANTERIOR_ = NulosN(Fg1.TextMatrix(Row, Col))
    End If
End Sub

Private Sub fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
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
        
        Fg1.TextMatrix(Fg1.Row, COLUMNAENVASE_) = NulosC(xRs("descripcion"))
        Fg1.TextMatrix(Fg1.Row, COLUMNAIDENV_) = NulosN(xRs("iditem"))
        Fg1.TextMatrix(Fg1.Row, COLUMNAUNIMED_) = NulosC(xRs("abrev"))
        Fg1.TextMatrix(Fg1.Row, COLUMNAIDUNIMED_) = NulosN(xRs("idunimed"))
        Fg1.TextMatrix(Fg1.Row, COLUMNAPESOENV_) = NulosN(xRs("peso"))
        Fg1.Select Fg1.Row, COLUMNAPESOPARIHUELA_
        
        Set xRs = Nothing
    End If
    
    If Col = COLUMNAUNIMED_ Then
'        ' BUSCA UN ITEM
'        ReDim xCampos(3, 4) As String
'
'        If NulosN(Fg1.TextMatrix(Fg1.Row, COLUMNAIDTIPO_)) = 0 Then
'            MsgBox "Seleccione el tipo de producto para el registro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'            Fg1.Col = COLUMNAENVASE_
'            Exit Sub
'        End If
'
'        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
'        xCampos(0, 0) = "Producto":    xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
'        xCampos(1, 0) = "Codigo":      xCampos(1, 1) = "codpro":        xCampos(1, 2) = "1500":         xCampos(1, 3) = "C"
'        xCampos(2, 0) = "Uni. Med":    xCampos(2, 1) = "abrev":         xCampos(2, 2) = "1200":         xCampos(2, 3) = "C"
'
'        nTitulo = "Buscando " & NulosC(Fg1.TextMatrix(Fg1.Row, COLUMNAENVASE_))
'
'        cSQL = "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.stckact, alm_inventario.activo " _
'            + vbCr + "FROM mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed " _
'            + vbCr + "WHERE (((alm_inventario.activo)=-1) AND ((alm_inventario.tippro)=" & NulosN(Fg1.TextMatrix(Fg1.Row, COLUMNAIDTIPO_)) & ")) " _
'            + vbCr + "ORDER BY alm_inventario.codpro"
'
'        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
'                                                        "descripcion", "descripcion", Principio, ""
'
'        If xRs.State = 0 Then Exit Sub
'        If xRs.RecordCount = 0 Then Exit Sub
'
'        Fg1.TextMatrix(Row, COLUMNAPESOPARIHUELA_) = NulosC(xRs("descripcion"))
'        Fg1.TextMatrix(Row, COLUMNANUMEROENV_) = NulosC(xRs("abrev"))
'        Fg1.TextMatrix(Row, COLUMNAIDPROD_) = NulosN(xRs("id"))
'
'        '****************************************************************************************
'        ' Se llena el lote Si es ingreso
'        If OptIng.Value = True Then
'            Fg1.TextMatrix(Row, COLUMNAPNETOTOTAL_) = crearLote(NulosN(xRs("id")), NulosC(TxtFchIng.Valor))
'            Fg1.TextMatrix(Row, COLUMNAESTADO_) = -1
'            Fg1.TextMatrix(Row, COLUMNAIDENV_) = -1
'        End If
'        '****************************************************************************************
'
'        Fg1.Col = COLUMNAPBRUTOTOTAL_
'        Fg1.SetFocus
'
'        Set xRs = Nothing
    End If
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim NUEVO_ As Boolean
    
    If Agregando = True Then Exit Sub
    If Row <= 0 Then Exit Sub
    
    Select Case Col
        Case COLUMNAHORA_
            If IsDate(Fg1.TextMatrix(Row, Col)) Then
                Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), FORMAT_HORA_SIN_SEGUNDO)
            Else
                MsgBox "Ingrese una hora correcta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Fg1.TextMatrix(Row, Col) = ""
                Fg1.Col = Col
            End If
        
        Case COLUMNAPESOENV_
            Fg1.TextMatrix(Row, Col) = Format(NulosN(Fg1.TextMatrix(Row, Col)), FORMAT_CANTIDAD)
            Fg1.TextMatrix(Row, COLUMNAPBRUTOENV_) = NulosN(Fg1.TextMatrix(Row, COLUMNANUMEROENV_)) * NulosN(Fg1.TextMatrix(Row, COLUMNAPESOENV_))
            Fg1.TextMatrix(Row, COLUMNAPBRUTOENV_) = Format(Fg1.TextMatrix(Row, COLUMNAPBRUTOENV_), FORMAT_CANTIDAD)
            Fg1.TextMatrix(Row, COLUMNAPNETOTOTAL_) = NulosN(Fg1.TextMatrix(Row, COLUMNAPBRUTOTOTAL_)) - NulosN(Fg1.TextMatrix(Row, COLUMNAPBRUTOENV_)) - NulosN(Fg1.TextMatrix(Row, COLUMNAPESOPARIHUELA_))
            Fg1.TextMatrix(Row, COLUMNAPNETOTOTAL_) = Format(Fg1.TextMatrix(Row, COLUMNAPNETOTOTAL_), FORMAT_CANTIDAD)
            hallarTotales Row, Col
            
        Case COLUMNAPESOPARIHUELA_
            Fg1.TextMatrix(Row, Col) = Format(NulosN(Fg1.TextMatrix(Row, Col)), FORMAT_CANTIDAD)
            Fg1.TextMatrix(Row, COLUMNAPNETOTOTAL_) = NulosN(Fg1.TextMatrix(Row, COLUMNAPBRUTOTOTAL_)) - NulosN(Fg1.TextMatrix(Row, COLUMNAPBRUTOENV_)) - NulosN(Fg1.TextMatrix(Row, COLUMNAPESOPARIHUELA_))
            Fg1.TextMatrix(Row, COLUMNAPNETOTOTAL_) = Format(Fg1.TextMatrix(Row, COLUMNAPNETOTOTAL_), FORMAT_CANTIDAD)
            
            hallarTotales Row, Col
            
        Case COLUMNANUMEROENV_
            Fg1.TextMatrix(Row, Col) = Format(NulosN(Fg1.TextMatrix(Row, Col)), "000")
            Fg1.TextMatrix(Row, COLUMNAPBRUTOENV_) = NulosN(Fg1.TextMatrix(Row, COLUMNANUMEROENV_)) * NulosN(Fg1.TextMatrix(Row, COLUMNAPESOENV_))
            Fg1.TextMatrix(Row, COLUMNAPBRUTOENV_) = Format(Fg1.TextMatrix(Row, COLUMNAPBRUTOENV_), FORMAT_CANTIDAD)
            Fg1.TextMatrix(Row, COLUMNAPNETOTOTAL_) = NulosN(Fg1.TextMatrix(Row, COLUMNAPBRUTOTOTAL_)) - NulosN(Fg1.TextMatrix(Row, COLUMNAPBRUTOENV_)) - NulosN(Fg1.TextMatrix(Row, COLUMNAPESOPARIHUELA_))
            Fg1.TextMatrix(Row, COLUMNAPNETOTOTAL_) = Format(Fg1.TextMatrix(Row, COLUMNAPNETOTOTAL_), FORMAT_CANTIDAD)
            hallarTotales Row, Col
        
        Case COLUMNAPBRUTOTOTAL_
            Fg1.TextMatrix(Row, Col) = Format(NulosN(Fg1.TextMatrix(Row, Col)), FORMAT_CANTIDAD)
            Fg1.TextMatrix(Row, COLUMNAPNETOTOTAL_) = NulosN(Fg1.TextMatrix(Row, COLUMNAPBRUTOTOTAL_)) - NulosN(Fg1.TextMatrix(Row, COLUMNAPBRUTOENV_)) - NulosN(Fg1.TextMatrix(Row, COLUMNAPESOPARIHUELA_))
            Fg1.TextMatrix(Row, COLUMNAPNETOTOTAL_) = Format(Fg1.TextMatrix(Row, COLUMNAPNETOTOTAL_), FORMAT_CANTIDAD)
            hallarTotales Row, Col
            
            
    End Select
End Sub

Private Sub Fg1_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
    Dim IDORD_ As Double
    Dim ESTADO_ As Double
    Dim Rpta As Integer
    Dim MENSAJE_ As String
    
    If Col = COLUMNAESTADO_ Then
        If QueHace = 1 Then
            MsgBox "Esta acción no esta permitida, haga clic en grabar para que se habilite esta opción", vbInformation, xTitulo
            Fg1.TextMatrix(Row, Col) = ESTADOANTERIOR_
            Exit Sub
        End If
        Rpta = MsgBox("¿Esta seguro de cambiar el estado actual?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)

        If Rpta = vbNo Then
            Fg1.TextMatrix(Row, Col) = ESTADOANTERIOR_
            Exit Sub
        End If
        
        ESTADO_ = NulosN(Fg1.TextMatrix(Row, Col))
            
        If ESTADOANTERIOR_ > ESTADO_ Then
            MsgBox "Este cambio de estado no esta permitido", vbInformation, xTitulo
        Else
            If ESTADO_ = ESTADOANULADO_ Then Exit Sub
            ' Se graba el ingreso
            If GrabarIngreso Then
                ESTADOANTERIOR_ = ESTADO_
            Else
                Fg1.TextMatrix(Row, Col) = ESTADOANTERIOR_
                Exit Sub
            End If
        End If
    End If
End Sub

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

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then Fg1.Editable = flexEDNone: Exit Sub
    If Agregando Then Exit Sub
    If Fg1.Rows - 1 <= Fg1.FixedRows Then Exit Sub
    If Fg1.Row = Fg1.Rows - 1 Then Exit Sub
    If Fg1.TextMatrix(Fg1.Row, COLUMNAESTADO_) > ESTADOPENDIENTE_ Then Fg1.Editable = flexEDNone: Exit Sub
    
    Select Case Fg1.Col
        Case COLUMNAPBRUTOTOTAL_, COLUMNANUMEROENV_, COLUMNAPESOPARIHUELA_, _
                                COLUMNAENVASE_, COLUMNAESTADO_, COLUMNAPESOENV_, COLUMNAOBS_, COLUMNAHORA_
            Fg1.Editable = flexEDKbdMouse

        Case Else
            Fg1.Editable = flexEDNone
    End Select
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case Col
        Case COLUMNAPESOPARIHUELA_, COLUMNANUMEROENV_
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
            
        Case COLUMNAENVASE_, COLUMNAUNIMED_
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
        
    GRID_COMBOLIST Fg1, COLUMNAENVASE_
    GRID_COMBOLIST Fg1, COLUMNAUNIMED_
    Fg1.ColEditMask(COLUMNAHORA_) = "##:##"
    
    Fg1.ColWidth(COLUMNAIDENV_) = 0
    Fg1.ColWidth(COLUMNAIDUNIMED_) = 0
    Fg1.ColWidth(COLUMNAIDESTADO_) = 0
    
    Fg1.AllowUserResizing = flexResizeColumns
    Fg1.ExplorerBar = flexExSortShow
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.ForeColorSel = &H80000005
    Fg1.BackColorSel = &H80&
    Fg1.WordWrap = True
    
    ESTADOANTERIOR_ = 1
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
    TxtFchDoc.Valor = Date
    TxtFchIng.SetFocus
    
    llenarEstados Fg1, COLUMNAESTADO_
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
        If RstRecep.State = 0 Then Exit Sub
        If RstRecep.RecordCount = 0 And QueHace <> 1 Then
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
            PROCESADO_ = False
            RstRecep.Requery
            Dg1.Refresh
            Cancelar
            
            If RstRecep.RecordCount <> 0 Then
                RstRecep.MoveFirst
                RstRecep.Find "id=" & mIdRegistro
                If RstRecep.EOF = True Then RstRecep.MoveFirst
            End If
            
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then
        RstRecep.Filter = ""
        TDB_FiltroLimpiar Dg1
    End If
    
    If Button.Index = 10 Then Buscar
    
    If Button.Index = 11 Then
        mMesActivo = SeleccionaMes(xCon)
        pCargarDatos
    End If
        
    'If Button.Index = 12 Then FrmConsIngAlmacen.Show
    
    If Button.Index = 16 Then
        Unload Me
        Set RstRecep = Nothing
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
    If PROCESADO_ Then
        MsgBox "No se puede Cancelar la operación, haga clic en el botón Grabar primero", vbExclamation, xTitulo
        Exit Sub
    End If
    
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

Private Function GrabarIngreso() As Boolean
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
    Dim IDALM_ As Integer
    Dim NUMDOC_ As String
    Dim xRs As New ADODB.Recordset
    Dim xRsAux As New ADODB.Recordset
            
    ' Se busca si esq hay datos que modificar de una ingreso anterior
    cSQL = "SELECT alm_ingreso.id, alm_ingreso.numdoc, alm_inventariolotedet.idlote, alm_ingresodet.idlotedet, alm_ingresodet.cantidad " _
        + vbCr + "FROM (alm_ingreso LEFT JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN alm_inventariolotedet ON alm_ingresodet.idlotedet = alm_inventariolotedet.id " _
        + vbCr + "WHERE (((alm_ingreso.idtipdocref)=71) AND ((alm_ingreso.iddocref)=" & NulosN(RstRecep("id")) & "));"
    
    RST_Busq xRsAux, cSQL, xCon
    If xRsAux.State = 0 Then Exit Function
    ' Se llenan los detalles
    FCHMOV_ = Format(TxtFchIng.Valor, "dd/mm/yyyy")
    TIPDOC_ = 71
    NUMSER_ = NulosC(TxtNumSer.Text)
    NUMDOC_ = ""
    IDRESP_ = NulosN(TxtIdRes.Text)
    IDPROV_ = NulosN(LblIdProveedor.Caption)
    DESPROV_ = NulosC(TxtProv.Text)
    IDESTADO_ = ESTADOPROCESADO_
    IDTIPMOV_ = -1
    IDTIPDOCREF_ = 71
    IDDOCREF_ = NulosN(RstRecep("id"))
    IDALM_ = NulosN(TxtIdAlm.Text)
    ' Se prepara el Recordset
    If xRs.State = 0 Then preparaRST xRs
    limpiarRST xRs
    ' Se llena el recordset
    IDING_ = 0
    xRs.AddNew
    ' ---CAMPOS PARA EL INGRESO
    xRs("iditem") = NulosN(txtiditem.Text)
    xRs("cantidad") = NulosN(Fg1.TextMatrix(Fg1.Row, COLUMNAPNETOTOTAL_))
    xRs("hora") = NulosC(Fg1.TextMatrix(Fg1.Row, COLUMNAHORA_))
    ' ---CAMPOS PARA EL LOTE
    If xRsAux.RecordCount = 0 Then
        xRs("idlote") = 0
        xRs("idlotedet") = 0
        xRs("canant") = 0
        xRs("idloteant") = 0
        xRs("idlotedetant") = 0
    Else
        xRs("idlote") = xRsAux("idlote")
        xRs("idlotedet") = xRsAux("idlotedet")
        xRs("canant") = 0
        xRs("idloteant") = xRsAux("idlote")
        xRs("idlotedetant") = xRsAux("idlotedet")
    End If
    xRs.Update
    
    ' Se graba el movimiento
    GrabarIngreso = grabarMovimiento(FCHMOV_, TIPDOC_, NUMSER_, "", IDRESP_, IDPROV_, DESPROV_, IDESTADO_, _
                                IDTIPMOV_, IDTIPDOCREF_, IDDOCREF_, IDALM_, xRs, IDING_, NUMDOC_, 6, mMesActivo)
    
    PROCESADO_ = True
End Function

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
    '**************************************
'    Dim RstLote As New ADODB.Recordset
'    Dim RstLoteDet As New ADODB.Recordset
'    Dim xIdLote As Double
'    Dim xIdLoteDet As Double
    '**************************************
    
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
    
    If Not IsDate(TxtFchDoc.Valor) Then
        MsgBox "No ha especificado la fecha del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchDoc.SetFocus
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
    
    If NulosC(TxtProv.Text) = "" Then
        MsgBox "No ha especificado el nombre del proveedor", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtProv.SetFocus
        Exit Function
    End If
        
    If NulosC(TxtIdRes.Text) = "" Then
        MsgBox "No ha especificado el responsable el movimiento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdRes.SetFocus
        Exit Function
    End If
       
    If Fg1.Rows = 2 Then
        If NulosN(Fg1.TextMatrix(1, COLUMNAIDENV_)) = 0 Then
            MsgBox "Falta ingresar el detalle...!", vbExclamation, "Mensaje...!"
            Exit Function
        End If
    ElseIf Fg1.Rows = 1 Then
        MsgBox "Falta ingresar el detalle...!", vbExclamation, "Mensaje...!"
        Exit Function
    End If
    
On Error GoTo LaCague
    xCon.BeginTrans
    
    If QueHace = 1 Then
        xId = HallaCodigoTabla("alm_recepcion", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM alm_recepcion", xCon
        RST_Busq RstDet, "SELECT TOP 1 * FROM alm_recepciondet", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = NulosN(RstRecep("id"))
        RST_Busq RstCab, "SELECT * FROM alm_recepcion WHERE id = " & RstRecep("id") & "", xCon
        xCon.Execute "DELETE * FROM alm_recepciondet WHERE idrecep = " & RstRecep("id") & " "
        RST_Busq RstDet, "SELECT * FROM alm_recepciondet", xCon
    End If
    
    mIdRegistro = xId
    xIdDet = HallaCodigoTabla("alm_recepciondet", xCon, "id")
    
    '******************************************************************
'    RST_Busq RstLote, "SELECT * FROM alm_inventariolote", xCon
'    RST_Busq RstLoteDet, "SELECT * FROM alm_inventariolotedet", xCon
'    xIdLote = HallaCodigoTabla("alm_inventariolote", xCon, "id")
'    xIdLoteDet = HallaCodigoTabla("alm_inventariolotedet", xCon, "id")
    '******************************************************************
        
    RstCab("fchdoc") = CDate(TxtFchDoc.Valor)
    RstCab("fching") = CDate(TxtFchIng.Valor)
    RstCab("tipdoc") = NulosN(TxtTipDoc.Text)
    RstCab("numser") = NulosC(TxtNumSer.Text)
    RstCab("numdoc") = NulosC(TxtNumDoc.Text)
    RstCab("ano") = AnoTra
    RstCab("idmes") = mMesActivo
    RstCab("idresp") = NulosN(TxtIdRes.Text)
    RstCab("idalm") = NulosN(TxtIdAlm.Text)
    RstCab("idtipdocref") = NulosN(TxtIdTipDocRef.Text)
    RstCab("iddocref") = NulosN(lbliddocref.Caption)
    RstCab("numdocref") = NulosN(txtNumDocRef.Text)
    RstCab("idprov") = NulosN(LblIdProveedor.Caption)
    RstCab("pesbru") = NulosN(Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAPBRUTOTOTAL_))
    RstCab("pesnet") = NulosN(Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAPNETOTOTAL_))
    RstCab("iditem") = NulosN(txtiditem.Text)
    RstCab.Update
    
    For A = 1 To Fg1.Rows - 2
        RstDet.AddNew
        RstDet("id") = xIdDet
        RstDet("idrecep") = xId
        RstDet("idenv") = NulosN(Fg1.TextMatrix(A, COLUMNAIDENV_))
        RstDet("idunimed") = NulosN(Fg1.TextMatrix(A, COLUMNAIDUNIMED_))
        RstDet("pesenv") = NulosN(Fg1.TextMatrix(A, COLUMNAPESOENV_))
        RstDet("pespar") = NulosN(Fg1.TextMatrix(A, COLUMNAPESOPARIHUELA_))
        RstDet("numenv") = NulosN(Fg1.TextMatrix(A, COLUMNANUMEROENV_))
        RstDet("pesbruenv") = NulosN(Fg1.TextMatrix(A, COLUMNAPBRUTOENV_))
        RstDet("pesbrutot") = NulosN(Fg1.TextMatrix(A, COLUMNAPBRUTOTOTAL_))
        RstDet("pesnettot") = NulosN(Fg1.TextMatrix(A, COLUMNAPNETOTOTAL_))
        RstDet("obs") = NulosC(Fg1.TextMatrix(A, COLUMNAOBS_))
        RstDet("hora") = NulosC(Fg1.TextMatrix(A, COLUMNAHORA_))
        
        RstDet("idestado") = NulosN(Fg1.TextMatrix(A, COLUMNAESTADO_))
        
        xIdDet = xIdDet + 1
        RstDet.Update
    Next A
    
    ' grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    xCon.CommitTrans
    
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
    
    TabOne1.CurrTab = 0
    If RstRecep.State = 0 Then Exit Sub
    If RstRecep.RecordCount = 0 Then
        MsgBox "No hay Registros de Recepcion de Almacén para eliminar", vbInformation, xTitulo
        Exit Sub
    End If
    
    Rpta = MsgBox("¿Esta seguro de eliminar la Recepcion Nº " + Trim(RstRecep("numdoc")) + "?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        ' Se elimina los registros relacionados
        cSQL = "SELECT alm_ingresodet.id, alm_ingresodet.idlotedet, alm_inventariolotedet.idlote, alm_inventariolote.cantidad " _
            + vbCr + "FROM alm_inventariolote RIGHT JOIN ((alm_ingresodet RIGHT JOIN alm_ingreso ON alm_ingresodet.id = alm_ingreso.id) LEFT JOIN alm_inventariolotedet ON alm_ingresodet.idlotedet = alm_inventariolotedet.id) ON alm_inventariolote.id = alm_inventariolotedet.idlote " _
            + vbCr + "WHERE (((alm_ingreso.idrecep)=" & NulosN(RstRecep("id")) & "));"
            
        Set xRs = Nothing
        RST_Busq xRs, cSQL, xCon
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then GoTo SIGUIENTE_
        xRs.MoveFirst
        While Not xRs.EOF
            ' se Reestablece los lotes
            xCon.Execute "DELETE * FROM alm_inventariolotedet WHERE id = " & NulosN(xRs("idlotedet"))
            xCon.Execute "UPDATE alm_inventariolote " _
                + vbCr + "SET alm_inventariolote.cantidad = alm_inventariolote.cantidad-" & NulosN(RstRecep("pesnet")) & " " _
                + vbCr + "WHERE (((alm_inventariolote.id)=" & NulosN(xRs("idlote")) & "));"
            ' Se elimina el detalle del ingreso
            xCon.Execute "DELETE * FROM alm_ingresodet WHERE id = " & NulosN(xRs("id"))
            xRs.MoveNext
        Wend
        ' Se elimina el Ingreso
        xCon.Execute "DELETE * FROM alm_ingreso WHERE idrecep = " & NulosN(RstRecep("id"))
        
SIGUIENTE_:
        ' Se elimina la recepcion
        xCon.Execute "DELETE * FROM alm_recepciondet WHERE idrecep = " & NulosN(RstRecep("id"))
        xCon.Execute "DELETE * FROM alm_recepcion WHERE id = " & NulosN(RstRecep("id"))
        
        ' Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & NulosN(RstRecep("id")) & " AND idform = " & IdMenuActivo

        RstRecep.Requery
        Dg1.Refresh
        MsgBox "La recepcion se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
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
    TxtFchDoc.Valor = ""
    TxtTipDoc.Text = ""
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    TxtProv.Text = ""
    TxtIdRes.Text = ""
    LblResp.Caption = ""
    LblTipDoc.Caption = ""
    TxtIdAlm.Text = ""
    LblAlmacen.Caption = ""
    TxtIdTipDocRef.Text = ""
    LblTipDocRef.Caption = ""
    lbliddocref.Caption = ""
    txtNumDocRef.Text = ""
    txtiditem.Text = ""
    lbldesitem.Caption = ""
    TxtTipDoc.Text = "71"
    TxtTipDoc_Validate True
    PROCESADO_ = False
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
    TxtTipDoc.Locked = Not TxtTipDoc.Locked
    TxtNumSer.Locked = Not TxtNumSer.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    TxtProv.Locked = Not TxtProv.Locked
    TxtIdRes.Locked = Not TxtIdRes.Locked
    TxtIdAlm.Locked = Not TxtIdAlm.Locked
    TxtIdTipDocRef.Locked = Not TxtIdTipDocRef.Locked
    txtNumDocRef.Locked = Not txtNumDocRef.Locked
    
    cmd(3).Enabled = Not cmd(3).Enabled
    cmd(2).Enabled = Not cmd(2).Enabled
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

Private Sub txtiditem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        lbldesitem.Caption = ""
    End If
End Sub

Private Sub txtiditem_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        cmd_Click 1
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
            TxtNumDoc.Text = hallarNumDoc("alm_recepcion", NulosC(TxtTipDoc.Text), "tipdoc", "'" & NulosC(TxtNumSer.Text) & "'", "numser")
        End If
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
        CmdBusProv_Click
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
End Sub

'Private Function crearLote(IDITEM_ As Double, FCHING_ As String) As String
'    Dim LOTE_ As String
'    Dim A As Integer
'    Dim MAXLIS_ As Integer
'    Dim MAXBD_ As Integer
'    Dim NUMERO_ As Integer
'    Dim xRs As New ADODB.Recordset
'
'    LOTE_ = Format(IDITEM_, "0000") & Format(CDate(FCHING_), "yy") & Format(Month(CDate(FCHING_)), "00") & Format(Day(CDate(FCHING_)), "00")
'    ' Se verifica el mayor de el listado
'    MAXLIS_ = 0
'    NUMERO_ = 0
'    For A = 1 To Fg1.Rows - 1
'        If A = Fg1.Row Then GoTo SIGUIENTE_
'        If NulosN(Fg1.TextMatrix(A, COLUMNAIDPROD_)) <> IDITEM_ Then GoTo SIGUIENTE_
'
'        NUMERO_ = NulosN(Mid(NulosC(Fg1.TextMatrix(A, COLUMNAPNETOTOTAL_)), 11, 2))
'        If NUMERO_ > MAXLIS_ Then
'            MAXLIS_ = NUMERO_
'        End If
'SIGUIENTE_:
'    Next A
'
'    ' Se verifica el mayor en la base de datos
'    cSQL = "SELECT Max(Mid([alm_inventariolote].[descripcion],10,2)) AS orden, alm_inventariolote.iditem, alm_inventariolote.fching " _
'        + vbCr + "FROM alm_inventariolote " _
'        + vbCr + "GROUP BY alm_inventariolote.iditem, alm_inventariolote.fching " _
'        + vbCr + "HAVING (((alm_inventariolote.iditem)=" & IDITEM_ & ") AND ((alm_inventariolote.fching)=CDate('" & FCHING_ & "')))"
'
'    Set xRs = Nothing
'    RST_Busq xRs, cSQL, xCon
'
'    MAXBD_ = 0
'    If xRs.State = 0 Then GoTo SALIR_
'    If xRs.RecordCount = 0 Then GoTo SALIR_
'
'    MAXBD_ = NulosN(xRs("orden"))
'SALIR_:
'
'    NUMERO_ = 0
'    If MAXLIS_ > MAXBD_ Then
'        NUMERO_ = MAXLIS_ + 1
'    Else
'        NUMERO_ = MAXBD_ + 1
'    End If
'
'    LOTE_ = LOTE_ & Format(NUMERO_, "00")
'
'    crearLote = LOTE_
'End Function

Private Sub llenarEstados(ByRef FGGRID As VSFlexGrid, columna As Integer)
    Dim CAMPOS As String
    Dim xRs As New ADODB.Recordset
    
    cSQL = "SELECT * FROM mae_estados ORDER BY id"
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then
        MsgBox "No se ha encontrado estados, Ingrese estados", vbInformation, xTitulo
        Exit Sub
    End If
    
    xRs.MoveFirst
    CAMPOS = "#" & NulosN(xRs("id")) & ";" & UCase(NulosC(xRs("descripcion")))
    xRs.MoveNext
    While Not xRs.EOF
        CAMPOS = CAMPOS & "|#" & NulosN(xRs("id")) & ";" & UCase(NulosC(xRs("descripcion")))
        xRs.MoveNext
    Wend
    FGGRID.ColComboList(columna) = CAMPOS
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
    Dim xRsDet As New ADODB.Recordset
    
    Blanquea
    llenarEstados Fg1, COLUMNAESTADO_
    
    If RstRecep.RecordCount = 0 Then Exit Sub
    If RstRecep.BOF = True Or RstRecep.EOF = True Then Exit Sub
    
    cSQL = "SELECT alm_recepcion.*, alm_inventario.descripcion AS desitem, alm_almacenes.descripcion AS desalm, mae_prov.nombre AS desprov, pla_empleados.nombre AS desresp, mae_documento.descripcion AS destipdoc, mae_documento_1.descripcion AS destipdocref, alm_recepcion.id " _
        + vbCr + "FROM ((((alm_almacenes RIGHT JOIN (alm_recepcion LEFT JOIN mae_prov ON alm_recepcion.idprov = mae_prov.id) ON alm_almacenes.id = alm_recepcion.idalm) LEFT JOIN alm_inventario ON alm_recepcion.iditem = alm_inventario.id) LEFT JOIN pla_empleados ON alm_recepcion.idresp = pla_empleados.id) LEFT JOIN mae_documento ON alm_recepcion.tipdoc = mae_documento.id) LEFT JOIN mae_documento AS mae_documento_1 ON alm_recepcion.idtipdocref = mae_documento_1.id " _
        + vbCr + "WHERE (((alm_recepcion.id)=" & NulosN(RstRecep("id")) & "));"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    
    TxtFchIng.Valor = xRs("fching")
    TxtFchDoc.Valor = xRs("fchdoc")
    TxtNumSer.Text = NulosC(xRs("numser"))
    TxtNumDoc.Text = NulosC(xRs("numdoc"))
    LblIdProveedor.Caption = NulosN(xRs("idprov"))
    TxtProv.Text = NulosC(xRs("desprov"))
    TxtIdRes.Text = NulosN(xRs("idresp"))
    LblResp.Caption = NulosC(xRs("desresp"))
    TxtIdAlm.Text = NulosN(xRs("idalm"))
    LblAlmacen.Caption = NulosC(xRs("desalm"))
    TxtTipDoc.Text = NulosN(xRs("tipdoc"))
    LblTipDoc.Caption = NulosC(xRs("destipdoc"))
    TxtIdTipDocRef.Text = NulosN(xRs("idtipdocref"))
    LblTipDocRef.Caption = NulosC(xRs("destipdocref"))
    txtNumDocRef.Text = NulosC(xRs("numdocref"))
    lbliddocref.Caption = NulosN(xRs("idtipdocref"))
    txtiditem.Text = NulosN(xRs("iditem"))
    lbldesitem.Caption = NulosC(xRs("desitem"))

    Fg1.Rows = 1
'    cSQL = "SELECT alm_recepciondet.*, alm_inventario.descripcion AS desenv, mae_unidades.abrev AS desunimed, mae_estados.descripcion AS desestado, mae_equivalencia.peso " _
'        + vbCr + "FROM (((alm_recepciondet LEFT JOIN alm_inventario ON alm_recepciondet.idenv = alm_inventario.id) LEFT JOIN mae_unidades ON alm_recepciondet.idunimed = mae_unidades.id) LEFT JOIN mae_estados ON alm_recepciondet.idestado = mae_estados.id) LEFT JOIN mae_equivalencia ON (alm_recepciondet.idunimed = mae_equivalencia.idunimed) AND (alm_recepciondet.idenv = mae_equivalencia.iditem) " _
'        + vbCr + "WHERE (((alm_recepciondet.idrecep) = " & NulosN(RstRecep("id")) & "));"
        
    cSQL = "SELECT alm_recepciondet.*, alm_inventario.descripcion AS desenv, mae_unidades.abrev AS desunimed, mae_estados.descripcion AS desestado " _
        + vbCr + "FROM ((alm_recepciondet LEFT JOIN alm_inventario ON alm_recepciondet.idenv = alm_inventario.id) LEFT JOIN mae_unidades ON alm_recepciondet.idunimed = mae_unidades.id) LEFT JOIN mae_estados ON alm_recepciondet.idestado = mae_estados.id " _
        + vbCr + "WHERE (((alm_recepciondet.idrecep)=" & NulosN(RstRecep("id")) & "));"

    
    RST_Busq xRsDet, cSQL, xCon
    If xRsDet.State = 0 Then Exit Sub
    If xRsDet.RecordCount = 0 Then Exit Sub
    
    Agregando = True
    xRsDet.MoveFirst
    For A = 1 To xRsDet.RecordCount
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(A, COLUMNAENVASE_) = NulosC(xRsDet("desenv"))
        Fg1.TextMatrix(A, COLUMNAUNIMED_) = NulosC(xRsDet("desunimed"))
        Fg1.TextMatrix(A, COLUMNAPESOPARIHUELA_) = Format(NulosN(xRsDet("pespar")), FORMAT_CANTIDAD)
        Fg1.TextMatrix(A, COLUMNANUMEROENV_) = Format(NulosN(xRsDet("numenv")), "000")
        Fg1.TextMatrix(A, COLUMNAPBRUTOENV_) = Format(NulosN(xRsDet("pesbruenv")), FORMAT_CANTIDAD)
        Fg1.TextMatrix(A, COLUMNAPBRUTOTOTAL_) = Format(NulosN(xRsDet("pesbrutot")), FORMAT_CANTIDAD)
        Fg1.TextMatrix(A, COLUMNAPNETOTOTAL_) = Format(NulosN(xRsDet("pesnettot")), FORMAT_CANTIDAD)
        Fg1.TextMatrix(A, COLUMNAESTADO_) = NulosN(xRsDet("idestado"))
        Fg1.TextMatrix(A, COLUMNAIDENV_) = NulosN(xRsDet("idenv"))
        Fg1.TextMatrix(A, COLUMNAIDUNIMED_) = NulosN(xRsDet("idunimed"))
        Fg1.TextMatrix(A, COLUMNAIDESTADO_) = NulosN(xRsDet("idestado"))
        Fg1.TextMatrix(A, COLUMNAPESOENV_) = Format(NulosN(xRsDet("pesenv")), FORMAT_CANTIDAD)
        Fg1.TextMatrix(A, COLUMNAOBS_) = NulosC(xRsDet("obs"))
        Fg1.TextMatrix(A, COLUMNAHORA_) = Format(NulosC(xRsDet("hora")), FORMAT_HORA_SIN_SEGUNDO)
        
        xRsDet.MoveNext
    Next A
    
    Agregando = False
    hallarTotales 1, 1
End Sub

Sub pCargarDatos()
     Dim NomMes As String
     Dim Cerrado As Boolean
     
    LblPeriodo.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
    
    TDB_FiltroLimpiar Dg1
    Set RstRecep = Nothing

    cSQL = "SELECT alm_recepcion.id AS id, [alm_recepcion].[fchdoc] & '' AS fchdoc, [alm_recepcion].[fching] & '' AS fching, [alm_recepcion].[numser] & '-' & [alm_recepcion].[numdoc] AS numdoc, alm_inventario.descripcion AS desitem, alm_recepcion.pesbru, alm_recepcion.pesnet " _
        + vbCr + "FROM (alm_almacenes RIGHT JOIN (alm_recepcion LEFT JOIN mae_prov ON alm_recepcion.idprov = mae_prov.id) ON alm_almacenes.id = alm_recepcion.idalm) LEFT JOIN alm_inventario ON alm_recepcion.iditem = alm_inventario.id " _
        + vbCr + "Where (((alm_recepcion.ano) = " & AnoTra & ") And ((alm_recepcion.idmes) = " & mMesActivo & ")) " _
        + vbCr + "ORDER BY [alm_recepcion].[fchdoc] & '' DESC;"
        
    RST_Busq RstRecep, cSQL, xCon
        
    Set Dg1.DataSource = RstRecep
End Sub


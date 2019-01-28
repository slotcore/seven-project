VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SIZERONE.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmManItems 
   Caption         =   "Punto de Venta - Items"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10710
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8595
      Top             =   0
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
            Picture         =   "FrmManItems.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManItems.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManItems.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManItems.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManItems.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManItems.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManItems.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManItems.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManItems.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManItems.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManItems.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManItems.frx":2236
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   5970
      Left            =   15
      TabIndex        =   0
      Top             =   360
      Width           =   10665
      _cx             =   18812
      _cy             =   10530
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
      Appearance      =   2
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
         Height          =   5550
         Left            =   45
         TabIndex        =   2
         Top             =   375
         Width           =   10575
         Begin TrueOleDBGrid70.TDBGrid Dg3 
            Height          =   5145
            Left            =   15
            TabIndex        =   8
            Top             =   360
            Width           =   10410
            _ExtentX        =   18362
            _ExtentY        =   9075
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Código"
            Columns(0).DataField=   "codpro"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Descripción"
            Columns(1).DataField=   "descripcion"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Unidad"
            Columns(2).DataField=   "abrev"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Moneda"
            Columns(3).DataField=   "simbolo"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Precio"
            Columns(4).DataField=   "precio"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Tipo Item"
            Columns(5).DataField=   "tipproddesc"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   20
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Activo"
            Columns(6).DataField=   "activo"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2805"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2725"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=6694"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=6615"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1032"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=953"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=1349"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1270"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=512"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=1773"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1693"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=514"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=2884"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2805"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=1032"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=953"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=513"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HDBFDFD&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H800000&"
            _StyleDefs(11)  =   ":id=2,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H80&"
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
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H80&"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80&"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=74,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=71,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=72,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=73,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=78,.parent=13,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=75,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=76,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=77,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=1"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
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
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Items"
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
            Left            =   15
            TabIndex        =   3
            Top             =   30
            Width           =   11550
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   5550
         Left            =   11310
         TabIndex        =   4
         Top             =   375
         Width           =   10575
         Begin VB.TextBox txt 
            BackColor       =   &H00E3DFE0&
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
            Index           =   3
            Left            =   5895
            MaxLength       =   50
            TabIndex        =   17
            Text            =   "txt(3)"
            Top             =   810
            Width           =   1200
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00E3DFE0&
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
            Index           =   2
            Left            =   3720
            MaxLength       =   50
            TabIndex        =   15
            Text            =   "txt(2)"
            Top             =   810
            Width           =   1200
         End
         Begin VB.CommandButton cb 
            Height          =   240
            Index           =   0
            Left            =   1920
            Picture         =   "FrmManItems.frx":277E
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   465
            Width           =   225
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   1050
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   1
            Text            =   "txt(1)"
            Top             =   810
            Width           =   1200
         End
         Begin VB.TextBox txt 
            BackColor       =   &H0000FFFF&
            Height          =   315
            Index           =   0
            Left            =   9435
            TabIndex        =   7
            Tag             =   "null"
            Text            =   "txt(0)"
            Top             =   420
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   0
            Left            =   1050
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   10
            Text            =   "txt_cb(0)"
            Top             =   435
            Width           =   1125
         End
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   4035
            Left            =   15
            TabIndex        =   19
            Top             =   1380
            Width           =   10530
            _cx             =   18574
            _cy             =   7117
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   2
            MousePointer    =   0
            _ConvInfo       =   1
            Version         =   700
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            FrontTabColor   =   -2147483633
            BackTabColor    =   -2147483633
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   8388608
            Caption         =   "    [Descuento en General]   |   [Descuento Corporativo]   "
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
            Begin VB.Frame fr 
               BorderStyle     =   0  'None
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
               Height          =   3660
               Index           =   1
               Left            =   45
               TabIndex        =   30
               Top             =   330
               Width           =   10440
               Begin VB.Frame fra 
                  Height          =   525
                  Index           =   0
                  Left            =   15
                  TabIndex        =   31
                  Top             =   2805
                  Width           =   4740
                  Begin VB.CommandButton cmd 
                     Caption         =   "Agregar"
                     Enabled         =   0   'False
                     Height          =   315
                     Index           =   4
                     Left            =   60
                     TabIndex        =   33
                     ToolTipText     =   "Agregar Documentos"
                     Top             =   165
                     Width           =   1200
                  End
                  Begin VB.CommandButton cmd 
                     Caption         =   "&Eliminar"
                     Enabled         =   0   'False
                     Height          =   315
                     Index           =   5
                     Left            =   1305
                     TabIndex        =   32
                     ToolTipText     =   "Eliminar Documentos Seleccionados"
                     Top             =   165
                     Width           =   1200
                  End
               End
               Begin VSFlex7Ctl.VSFlexGrid fg 
                  Height          =   2625
                  Index           =   2
                  Left            =   15
                  TabIndex        =   34
                  Top             =   195
                  Width           =   4740
                  _cx             =   8361
                  _cy             =   4630
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
                  ForeColorSel    =   16777215
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   2
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmManItems.frx":28B0
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
            Begin VB.Frame fr 
               BorderStyle     =   0  'None
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
               Height          =   3660
               Index           =   0
               Left            =   11175
               TabIndex        =   20
               Top             =   330
               Width           =   10440
               Begin VB.Frame fra 
                  Height          =   525
                  Index           =   3
                  Left            =   6495
                  TabIndex        =   26
                  Top             =   2805
                  Width           =   3900
                  Begin VB.CommandButton cmd 
                     Caption         =   "Agregar"
                     Enabled         =   0   'False
                     Height          =   315
                     Index           =   2
                     Left            =   60
                     TabIndex        =   28
                     ToolTipText     =   "Agregar Documentos"
                     Top             =   165
                     Width           =   1200
                  End
                  Begin VB.CommandButton cmd 
                     Caption         =   "&Eliminar"
                     Enabled         =   0   'False
                     Height          =   315
                     Index           =   3
                     Left            =   1320
                     TabIndex        =   27
                     ToolTipText     =   "Eliminar Documentos Seleccionados"
                     Top             =   165
                     Width           =   1200
                  End
               End
               Begin VB.Frame fra 
                  Height          =   525
                  Index           =   2
                  Left            =   15
                  TabIndex        =   23
                  Top             =   2805
                  Width           =   6305
                  Begin VB.CommandButton cmd 
                     Caption         =   "Agregar"
                     Enabled         =   0   'False
                     Height          =   315
                     Index           =   0
                     Left            =   60
                     TabIndex        =   25
                     ToolTipText     =   "Agregar Documentos"
                     Top             =   165
                     Width           =   1200
                  End
                  Begin VB.CommandButton cmd 
                     Caption         =   "&Eliminar"
                     Enabled         =   0   'False
                     Height          =   315
                     Index           =   1
                     Left            =   1305
                     TabIndex        =   24
                     ToolTipText     =   "Eliminar Documentos Seleccionados"
                     Top             =   165
                     Width           =   1200
                  End
               End
               Begin VSFlex7Ctl.VSFlexGrid fg 
                  Height          =   2625
                  Index           =   0
                  Left            =   15
                  TabIndex        =   21
                  Top             =   195
                  Width           =   6305
                  _cx             =   11121
                  _cy             =   4630
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
                  ForeColorSel    =   16777215
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   2
                  Cols            =   4
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmManItems.frx":295B
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
               Begin VSFlex7Ctl.VSFlexGrid fg 
                  Height          =   2625
                  Index           =   1
                  Left            =   6495
                  TabIndex        =   22
                  Top             =   195
                  Width           =   3900
                  _cx             =   6879
                  _cy             =   4630
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
                  ForeColorSel    =   16777215
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   2
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmManItems.frx":29D6
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
               Begin VB.Label lbl_cabecera 
                  Caption         =   "lbl_cabecera(1)"
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
                  Height          =   255
                  Index           =   1
                  Left            =   15
                  TabIndex        =   29
                  Top             =   3360
                  Width           =   10365
               End
            End
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Index           =   4
            Left            =   5175
            TabIndex        =   18
            Top             =   915
            Width           =   630
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Unidad:"
            Height          =   195
            Index           =   3
            Left            =   3090
            TabIndex        =   16
            Top             =   915
            Width           =   555
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "Item"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   14
            Top             =   540
            Width           =   915
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Precio Unit."
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   13
            Top             =   915
            Width           =   825
         End
         Begin VB.Label lbl_cb_cod 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb_cod(0)"
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
            Index           =   0
            Left            =   7320
            TabIndex        =   12
            Top             =   450
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label lbl_cb 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb(0)"
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
            Index           =   0
            Left            =   2190
            TabIndex        =   11
            Top             =   435
            Width           =   6255
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "codigo"
            Height          =   195
            Index           =   0
            Left            =   8850
            TabIndex        =   6
            Top             =   525
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Item"
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
            Height          =   210
            Left            =   15
            TabIndex        =   5
            Top             =   30
            Width           =   11550
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   10710
      _ExtentX        =   18891
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
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
                  Text            =   "Modificar Item"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Activar Item"
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
                  Text            =   "Eliminar un Item"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Retirar Item"
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
            Object.ToolTipText     =   "Imprimir Guia"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Eliminar"
      End
   End
   Begin VB.Menu Menu2 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu Menu2_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu2_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu2_3 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu Menu2_4 
         Caption         =   "Eliminar Todo"
      End
   End
   Begin VB.Menu Menu3 
      Caption         =   "Menu3"
      Visible         =   0   'False
      Begin VB.Menu Menu3_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu3_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu3_3 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu Menu3_4 
         Caption         =   "Eliminar Todo"
      End
   End
End
Attribute VB_Name = "FrmManItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim RstFrm As New ADODB.Recordset
Dim Agregando As Boolean
'---
Dim RstTmp As New ADODB.Recordset
Dim IdCorrelativo As Integer   '--INDICA EL INCREMENTO DE LOS DESCUENTOS SOLO CORPORATIVO

Private Sub REGISTRO_ADD(Index As Integer, Optional F_SELECCION_VARIOS As Boolean = False)
    '--
    If QueHace = 3 Then Exit Sub
    If txt_cb(0) = "" Then
        MsgBox "Seleccione el Item", vbExclamation, xTitulo
        txt_cb(0).SetFocus
        Exit Sub
    End If
    
    Agregando = True
    Select Case Index
        Case 0 '--
            If fg(Index).TextMatrix(fg(Index).Rows - 1, 2) = "" Or fg(Index).TextMatrix(fg(Index).Rows - 1, 3) = "" Then
                MsgBox "Primero Seleccione al Cliente" + vbCr + "Luego continue", vbExclamation, xTitulo
                fg(Index).Col = IIf(fg(Index).TextMatrix(fg(Index).Rows - 1, 3) = "", 3, 4)
                fg(Index).Row = fg(Index).Rows - 1
                GoTo Salir
            End If
        
            fg(Index).AddItem ""
            fg(Index).Row = fg(Index).Rows - 1
            fg(Index).Col = 2
            GoTo Salir
        Case 1, 2
            If fg(Index).Rows > 1 Then
                If fg(Index).TextMatrix(fg(Index).Rows - 1, 2) = "" Or fg(Index).TextMatrix(fg(Index).Rows - 1, 3) = "" Then
                    MsgBox "Ingrese la Cantidad " + IIf(fg(Index).TextMatrix(fg(Index).Rows - 1, 2) = "", "Inicial", "Final") + _
                    vbCr + "Luego continue", vbExclamation, xTitulo
                    
                    fg(Index).Row = fg(Index).Rows - 1
                    fg(Index).Col = IIf(fg(Index).TextMatrix(fg(Index).Rows - 1, 2) = "", 1, 2)
                    
                    GoTo Salir:
                End If
            End If
            fg(Index).AddItem ""
            fg(Index).TextMatrix(fg(Index).Rows - 1, 1) = IdCorrelativo
            fg(Index).Row = fg(Index).Rows - 1
            fg(Index).Col = 2
            IdCorrelativo = IdCorrelativo + 1
    End Select
    '-------------------------------

Salir:
    fg(Index).SetFocus
    Agregando = False
    Exit Sub
error:
    Agregando = False
    SHOW_ERROR Me.Name, "Registro_Add"
End Sub

Private Sub REGISTRO_DEL(Index As Integer, Optional del_todos As Boolean = False)
    If QueHace = 3 Then Exit Sub
    If fg(Index).Row <= 0 Then
        MsgBox "La fila no se puede eliminar" + vbCr + "Seleccione una fila correcta", vbExclamation, xTitulo
        Exit Sub
    End If
    If del_todos = False Then
        If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
    End If
    '--ELIMINAR EL REGISTRO
    If Index = 0 Then '--  ELIMINAR TODOS LOS REGISTROS
        Agregando = True
        LimpiarGrid fg(Index + 1), True, 1
        '--1:.ACTIVO  2::PASIVO
        '--ELIMINAR DATOS DEL TEMPORAL
        If fg(Index).TextMatrix(fg(Index).Row, 1) <> "" Then
            RstTmp.Filter = "idcli = " + fg(Index).TextMatrix(fg(Index).Row, 1)
            If RstTmp.RecordCount <> 0 Then
                RstTmp.MoveFirst
                Do While Not RstTmp.EOF
                    RstTmp.Delete
                    RstTmp.MoveNext
                Loop
            End If
        End If
        lbl_cabecera(1) = ""
    Else
        If fg(Index).TextMatrix(fg(Index).Row, 1) <> "" Then
            If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
            Do While Not RstTmp.EOF
                If NulosN(RstTmp.Fields("corr")) = NulosN(fg(Index).TextMatrix(fg(Index).Row, 1)) Then
                    RstTmp.Delete
                End If
                RstTmp.MoveNext
            Loop
        End If
    End If
    fg(Index).RemoveItem (fg(Index).Row)
    Agregando = False
End Sub



Private Sub cmd_Click(Index As Integer)
    Select Case Index
        '--DESCUENTO CORPORATIVO
        Case 0 '--ADD
            REGISTRO_ADD 0
        Case 1 '--SEL
            REGISTRO_DEL 0
        Case 2 '--DEL
            REGISTRO_ADD 1
        Case 3 '--DEL
            REGISTRO_DEL 1
        '--DESCUENTO GENERAL
        Case 4 '--ADD REG
            REGISTRO_ADD 2
        Case 5 '--DEL REG
            REGISTRO_DEL 2
    End Select
End Sub

Private Sub Dg3_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg3_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    RstFrm.Sort = CStr(Dg3.Columns(ColIndex).DataField)
    Err.Clear
End Sub

Private Sub fg_EnterCell(Index As Integer)
    If QueHace = 3 Or (Index = 0 And fg(0).Col <> 2) Then
        fg(Index).Editable = flexEDNone
        Exit Sub
    End If
    If fg(Index).Col = 1 Then
        fg(Index).Editable = flexEDNone
    Else
        fg(Index).Editable = flexEDKbdMouse
    End If
End Sub
Private Sub fg_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Row = 0 Or Index = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case Col
        Case Is <> 1
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub fg_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Or KeyCode = vbKeyInsert Then 'F3 = Agregar Item
        REGISTRO_ADD Index
    End If
    If KeyCode = 115 Or KeyCode = vbKeyDelete Then
        REGISTRO_DEL Index  'F4 = Eliminar Item
    End If
End Sub
Private Sub fg_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)
'    On Error GoTo error
    If Agregando = True Then Exit Sub
    If Row < 1 Or Index = 0 Then Exit Sub
    Select Case Col
        Case 2, 3, 4, 5 '--VALIDAR QUE EL NUMERO DE ORDEN SEA UNICO
            If NulosN(fg(Index).TextMatrix(Row, Col)) = 0 Then Exit Sub
            
            If IsNumeric(fg(Index).TextMatrix(Row, Col)) = False Then
                MsgBox "El valor ingresado no es numérico", vbExclamation, xTitulo
                fg(Index).TextMatrix(Row, Col) = "":    Exit Sub
            End If
            If Col = 4 Then '--DEL PORCENTAJE
                Dim VN_Porcentaje As Double
                If NulosN(fg(Index).TextMatrix(Row, Col)) >= 0 Then
                    VN_Porcentaje = IIf(NulosN(fg(Index).TextMatrix(Row, Col)) > 1, (NulosN(fg(Index).TextMatrix(Row, Col))) / 100, NulosN(fg(Index).TextMatrix(Row, Col)))
                Else
                    VN_Porcentaje = NulosN(fg(Index).TextMatrix(Row, Col))
                End If
                If VN_Porcentaje > 1 Then VN_Porcentaje = 0
                fg(Index).TextMatrix(Row, Col) = FormatPercent(VN_Porcentaje, 2)
            Else
                fg(Index).TextMatrix(Row, Col) = NulosN(fg(Index).TextMatrix(Row, Col))
            End If
            
    End Select
    Exit Sub
error:
    SHOW_ERROR Me.Name, "fg_CellChanged (" + CStr(Index) + ")"
End Sub



Private Sub Fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Index <> 0 And Col <> 2 Then Exit Sub
    
    Agregando = True
        
    '----------
    '----------
    Dim xRs As New ADODB.Recordset
    Dim xCampos() As String
    Dim N_SQL As String
    Dim ID_CLIENTE As String
    On Error GoTo error
    ID_CLIENTE = GRID_GENERAR_SQL_ID(fg(0), 1, "mae_cliente.id", " NOT IN ", True)
    If ID_CLIENTE <> "" Then ID_CLIENTE = " WHERE " + ID_CLIENTE
    
    N_SQL = " SELECT mae_cliente.id, mae_cliente.numruc, mae_cliente.nombre " _
        + vbCr + " FROM mae_cliente " + ID_CLIENTE _
        + vbCr + " ORDER BY mae_cliente.nombre;"


    ReDim xCampos(2, 3) As String
    xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":  xCampos(0, 2) = "4500":     xCampos(0, 3) = "C"
    xCampos(1, 0) = "RUC":      xCampos(1, 1) = "numruc":  xCampos(1, 2) = "1500":     xCampos(1, 3) = "C"

    CARGAR_DLL_EPSBUSCAR xCon, xRs, N_SQL, xCampos(), "Buscando Cliente", "nombre", "nombre", Principio

    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir
    
    fg(0).TextMatrix(fg(0).Rows - 1, 1) = xRs.Fields("id") & ""
    fg(0).TextMatrix(fg(0).Rows - 1, 2) = xRs.Fields("numruc") & ""
    fg(0).TextMatrix(fg(0).Rows - 1, 3) = xRs.Fields("nombre") & ""
    
Salir:
    Set xRs = Nothing
    Agregando = False
    Exit Sub
error:
    SHOW_ERROR Me.Name, "Fg_CellButtonClick (" + CStr(Index) + ")"
End Sub

Private Sub Fg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace <> 3 Then
            Select Case Index
                Case 0: PopupMenu Menu1
                Case 1: PopupMenu Menu2
                Case 2: PopupMenu Menu3
            End Select
        End If
    End If
End Sub

Private Sub Fg_RowColChange(Index As Integer)
    
    If Agregando = True Then Exit Sub
    If Index = 2 Then Exit Sub
    If fg(Index).Rows = 1 Then Exit Sub
    If Index = 0 Then
        fg(Index + 1).Rows = 1
        lbl_cabecera(1) = fg(Index).TextMatrix(fg(Index).Row, 3)
        If fg(Index).Row <= 0 Then Exit Sub
        If fg(Index).TextMatrix(fg(Index).Row, 1) = "" Then Exit Sub
        '--FILTRANDO POR CLIENTE
        RstTmp.Filter = "idcli=" + fg(Index).TextMatrix(fg(Index).Row, 1)
        If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
        Agregando = True
        Dim VN_Porcentaje As Double
        Do While Not RstTmp.EOF
            fg(Index + 1).Rows = fg(Index + 1).Rows + 1
            fg(Index + 1).TextMatrix(fg(Index + 1).Rows - 1, 1) = RstTmp.Fields("corr") & ""
            fg(Index + 1).TextMatrix(fg(Index + 1).Rows - 1, 2) = RstTmp.Fields("inicio") & ""
            fg(Index + 1).TextMatrix(fg(Index + 1).Rows - 1, 3) = RstTmp.Fields("fin") & ""
            '------------------
            If NulosN(RstTmp.Fields("porcentaje") & "") <> 0 Then
                VN_Porcentaje = IIf(NulosN(RstTmp.Fields("porcentaje") & "") > 1, (NulosN(RstTmp.Fields("porcentaje") & "") / 100), NulosN(RstTmp.Fields("porcentaje") & ""))
            Else
                VN_Porcentaje = NulosN(RstTmp.Fields("porcentaje") & "")
            End If
            fg(Index + 1).TextMatrix(fg(Index + 1).Rows - 1, 4) = FormatPercent(VN_Porcentaje, 2)
            'fg(Index + 1).TextMatrix(fg(Index + 1).Rows - 1, 4) = RstTmp.Fields("porcentaje") & ""
            '------------------
            fg(Index + 1).TextMatrix(fg(Index + 1).Rows - 1, 5) = RstTmp.Fields("valor") & ""
            RstTmp.MoveNext
        Loop
    Else
        If fg(0).Row < 0 Then Exit Sub
        RstTmp.Filter = "corr = " + CStr(NulosN(fg(1).TextMatrix(fg(1).Row, 1)))
        If RstTmp.RecordCount = 0 Then RstTmp.AddNew
        RstTmp.Fields("idcli") = NulosN(fg(0).TextMatrix(fg(0).Row, 1))
        RstTmp.Fields("corr") = NulosN(fg(1).TextMatrix(fg(1).Row, 1))
        RstTmp.Fields("inicio") = NulosN(fg(1).TextMatrix(fg(1).Row, 2))
        RstTmp.Fields("fin") = NulosN(fg(1).TextMatrix(fg(1).Row, 3))
        RstTmp.Fields("porcentaje") = NulosN(Replace(fg(1).TextMatrix(fg(1).Row, 4), "%", "")) / 100
        RstTmp.Fields("valor") = NulosN(fg(1).TextMatrix(fg(1).Row, 5))
        
        RstTmp.Filter = "idcli=" + fg(Index).TextMatrix(fg(Index).Row, 1)
    End If
    
    Agregando = False
    
End Sub


Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
    Dim Rpta As Integer

    SeEjecuto = False
    CARGAR_GRID
    SeEjecuto = True
    If RstFrm.RecordCount = 0 Then
        If MsgBox("No se ha registrado ninguna cuenta por rendir, ¿Desea agergar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
            Nuevo
        End If
    End If
End Sub

Private Sub CARGAR_GRID()

    Dim xSQL  As String
        
    xSQL = "SELECT pvt_items.iditem as id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, mae_moneda.simbolo, pvt_items.precio, mae_tipoproducto.descripcion AS tipproddesc, pvt_items.activo " _
        + vbCr + " FROM (mae_unidades RIGHT JOIN (mae_tipoproducto RIGHT JOIN (mae_moneda RIGHT JOIN alm_inventario ON mae_moneda.id = alm_inventario.idmon) ON mae_tipoproducto.id = alm_inventario.tippro) ON mae_unidades.id = alm_inventario.idunimed) RIGHT JOIN pvt_items ON alm_inventario.id = pvt_items.iditem " _
        + vbCr + " ORDER BY alm_inventario.descripcion;"
    
    '--CARGANDO_DATOS
    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, xSQL, xCon
    
    Set Dg3.DataSource = RstFrm
    Me.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    SeEjecuto = False
    Agregando = False
    QueHace = 3
    
    
    Dg3.BatchUpdates = False
    
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    '--
    Habilitar_Obj False
    '----
    fg(0).Tag = fg(0).FormatString
    fg(1).Tag = fg(1).FormatString
    fg(2).Tag = fg(2).FormatString
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation, xTitulo
        Cancel = 1
        Exit Sub
    End If
    Set RstFrm = Nothing
    Set Dg3.DataSource = Nothing
End Sub


Private Sub Menu1_1_Click()
    cmd_Click 0
End Sub

Private Sub Menu1_3_Click()
    cmd_Click 1
End Sub


Private Sub Menu2_1_Click()
    cmd_Click 2
End Sub

Private Sub Menu2_3_Click()
    cmd_Click 3
End Sub

Private Sub Menu2_4_Click()
    Dim Q_ROW As Long
    If fg(1).Rows <= 1 Then Exit Sub
    If MsgBox("Seguro desea eliminar todos los registros", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
    Agregando = True
    Do While fg(1).Rows > 1
        fg(1).Row = 1
        REGISTRO_DEL 1, True
    Loop
    Agregando = False
End Sub

Private Sub Menu3_1_Click()
    cmd_Click 4
End Sub

Private Sub Menu3_2_Click()
    cmd_Click 5
End Sub

Private Sub Menu3_3_Click()
    cmd_Click 5
End Sub



Private Sub Menu3_4_Click()
    Dim Q_ROW As Long
    If fg(2).Rows <= 1 Then Exit Sub
    If MsgBox("Seguro desea eliminar todos los registros", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
    Agregando = True
    Do While fg(2).Rows > 1
        fg(2).Row = 1
        REGISTRO_DEL 2, True
    Loop
    Agregando = False
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 3 Then MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            RstFrm.Requery
            Dg3.Refresh
            Cancelar
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 8 Then
        TabOne1.CurrTab = 0
        Filtrar
    End If
    
    If Button.Index = 9 Then
        TabOne1.CurrTab = 0
        RstFrm.Filter = adFilterNone
    End If
    
    If Button.Index = 10 Then
        TabOne1.CurrTab = 0
        Buscar
    End If
    
    If Button.Index = 12 Then
'        frmImprimirItem.Show vbModal
    End If
    
    If Button.Index = 14 Then
        Unload Me
    End If
End Sub

Sub Eliminar()
    On Error GoTo error
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
        MsgBox "No hay registros", vbExclamation, xTitulo
        Exit Sub
    End If
    
    If MsgBox("¿Esta seguro de eliminar el registro?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
        xCon.Execute "DELETE * FROM pvt_desccorporativo   WHERE iditem = " & RstFrm("id") & ""
        xCon.Execute "DELETE * FROM pvt_descgeneral WHERE iditem = " & RstFrm("id") & ""
        xCon.Execute "DELETe * FROM pvt_items WHERE iditem = " & RstFrm("id") & ""
        
        MsgBox "El registro fue eliminado con éxito", vbInformation, xTitulo
        RstFrm.Requery
        Dg3.Refresh
        If RstFrm.RecordCount = 0 Then
            If MsgBox("No hay registrado ningún Item, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
                Nuevo
            Else
                TabOne1.CurrTab = 0
            End If
        End If
    End If
    
Exit Sub
error:
    SHOW_ERROR Me.Name, "Eliminar"
End Sub

Sub Cancelar()
    QueHace = 3
    TabOne1.TabEnabled(0) = True
    ActivaTool
    Habilitar_Obj False
    Label1.Caption = "Detalle del Estado"
    TabOne1.CurrTab = 0
    
    Dg3.SetFocus
End Sub

Private Sub Modificar()
   '------
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
        MsgBox "No hay Registros", vbExclamation, xTitulo
        Exit Sub
    End If
    
    QueHace = 2
    ActivaTool
    Habilitar_Obj True
    MuestraSegundoTab
    
    IdCorrelativo = 999
    GRID_COMBOLIST fg(0), 2
    
    Label1.Caption = "Modificando Item"
    
    txt_cb(0).SetFocus
    

End Sub

Sub MuestraSegundoTab()
'    On Error GoTo error
    With RstFrm
        Blanquea
        If .State = 0 Then Exit Sub
        If .EOF = True Or .BOF = True Then Exit Sub
        
        txt(0).Text = .Fields("id") & "" '--CODIGO
        txt(1).Text = .Fields("precio") & ""
        '--DEL PRODUCTO
        txt_cb(0).Text = .Fields("id") & ""
        lbl_cb(0).Caption = .Fields("descripcion") & ""
        lbl_cb_cod(0).Caption = .Fields("id") & ""
        '----
        txt(2).Text = .Fields("abrev") & ""
        txt(3).Text = .Fields("simbolo") & ""
        '----
        MuestraDetalle
        Me.TabOne2.CurrTab = 0
    End With
    
    Exit Sub
error:
    SHOW_ERROR Me.Name, "MuestraSegundoTab"
End Sub

Private Sub MuestraDetalle()
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    Dim xCol, xFil As Integer
    Dim xSQL As String
    Dim xFch As Date
    Dim xFila  As Integer
'    On Error GoTo error
    
    '--CARGANDO EL LISTADO DE CLIENTES QUE TIENEN DESCUENTO
    xSQL = "SELECT pvt_desccorporativo.idcli, mae_cliente.numruc, mae_cliente.nombre " _
        + vbCr + " FROM mae_cliente INNER JOIN pvt_desccorporativo ON mae_cliente.id = pvt_desccorporativo.idcli " _
        + vbCr + " GROUP BY pvt_desccorporativo.iditem, pvt_desccorporativo.idcli, mae_cliente.numruc, mae_cliente.nombre " _
        + vbCr + " HAVING (((pvt_desccorporativo.iditem) = " + CStr(RstFrm.Fields("id")) + ")) " _
        + vbCr + " ORDER BY mae_cliente.nombre;"
    RST_Busq xRs, xSQL, xCon
    fg(0).Rows = 1
    Agregando = True
    If xRs.BOF = False Or xRs.EOF = False Or xRs.RecordCount <> 0 Then xRs.MoveFirst
    Do While Not xRs.EOF
        fg(0).Rows = fg(0).Rows + 1
        fg(0).TextMatrix(fg(0).Rows - 1, 1) = xRs.Fields("idcli") & ""
        fg(0).TextMatrix(fg(0).Rows - 1, 2) = xRs.Fields("numruc") & ""
        fg(0).TextMatrix(fg(0).Rows - 1, 3) = xRs.Fields("nombre") & ""
        xRs.MoveNext
    Loop
    Set xRs = Nothing
    '----------------------------------------------------
    '--CARGANDO EL LISTADO DE CLIENTES QUE TIENEN DESCUENTO
    xSQL = "SELECT pvt_descgeneral.corr, pvt_descgeneral.inicio, pvt_descgeneral.fin, pvt_descgeneral.porcentaje, pvt_descgeneral.valor " _
        + vbCr + " FROM pvt_descgeneral " _
        + vbCr + " WHERE (((pvt_descgeneral.iditem) = " + CStr(RstFrm.Fields("id")) + ")) " _
        + vbCr + " ORDER BY pvt_descgeneral.inicio;"
    RST_Busq xRs, xSQL, xCon
    fg(2).Rows = 1
    If xRs.BOF = False Or xRs.EOF = False Or xRs.RecordCount <> 0 Then xRs.MoveFirst
    Dim VN_Porcentaje As Double
    Do While Not xRs.EOF
        fg(2).Rows = fg(2).Rows + 1
        fg(2).TextMatrix(fg(2).Rows - 1, 1) = xRs.Fields("corr") & ""
        fg(2).TextMatrix(fg(2).Rows - 1, 2) = xRs.Fields("inicio") & ""
        fg(2).TextMatrix(fg(2).Rows - 1, 3) = xRs.Fields("fin") & ""
        '------------------
        If NulosN(xRs.Fields("porcentaje") & "") <> 0 Then
            VN_Porcentaje = IIf(NulosN(xRs.Fields("porcentaje") & "") > 1, (NulosN(xRs.Fields("porcentaje") & "") / 100), NulosN(xRs.Fields("porcentaje") & ""))
        Else
            VN_Porcentaje = NulosN(xRs.Fields("porcentaje") & "")
        End If
        fg(2).TextMatrix(fg(2).Rows - 1, 4) = FormatPercent(VN_Porcentaje, 2)
        '------------------
        fg(2).TextMatrix(fg(2).Rows - 1, 5) = xRs.Fields("valor") & ""
        xRs.MoveNext
    Loop
    Set xRs = Nothing
    '--CARGANDO DATOS DE LAS CUENTAS
    Dim N_SQL As String
    Dim RST_TMP As New ADODB.Recordset
    N_SQL = GENERAR_CONSULTA(RstFrm.Fields("id") & "")
    If N_SQL <> "" Then
        RST_Busq RST_TMP, N_SQL, xCon
        CARGAR_RST_TMP RstTmp, RST_TMP
    End If
    Set RST_TMP = Nothing
    '----------------------------------------------------
    Set xRs = Nothing
    Agregando = False
    '--CARGANDO LOS DATOS DE LAS CUENTAS AL ACTIVO Y PASIVO
    Fg_RowColChange 0
    Exit Sub
error:
    Set xRs = Nothing:  Set RST_TMP = Nothing
    Agregando = False
    SHOW_ERROR Me.Name, "MuestraDetalle"
End Sub


Private Sub Habilitar_Obj(band As Boolean)
    habilitar_Locked txt, Not band
    habilitar cmd, band
    
    TabOne1.CurrTab = IIf(band = False, 0, 1)
    TabOne1.TabEnabled(0) = Not band
    
    If band = False Then
        fg(0).SelectionMode = flexSelectionByRow
        fg(1).SelectionMode = flexSelectionByRow
        fg(2).SelectionMode = flexSelectionByRow
    Else
        fg(0).SelectionMode = flexSelectionFree
        fg(1).SelectionMode = flexSelectionFree
        fg(2).SelectionMode = flexSelectionFree
    End If
    
End Sub

Private Sub Blanquea()
    LimpiaText txt
    LimpiaText lbl_cabecera
    
    LimpiaText txt_cb
    LimpiaText lbl_cb_cod
    LimpiaText lbl_cb
    Agregando = True
    LimpiarGrid fg(0), True, 1
    LimpiarGrid fg(1), True, 1
    LimpiarGrid fg(2), True, 1
    Agregando = False
    OCULTAR_COL fg(0), 1, 1
    OCULTAR_COL fg(1), 1, 1
    OCULTAR_COL fg(2), 1, 1

    DEFINIR_RST
    
End Sub

Private Sub ActivaTool()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    Toolbar1.Buttons(2).Enabled = Not Toolbar1.Buttons(2).Enabled
    Toolbar1.Buttons(3).Enabled = Not Toolbar1.Buttons(3).Enabled
    
    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
    Toolbar1.Buttons(6).Enabled = Not Toolbar1.Buttons(6).Enabled
    
    Toolbar1.Buttons(8).Enabled = Not Toolbar1.Buttons(8).Enabled
    Toolbar1.Buttons(9).Enabled = Not Toolbar1.Buttons(9).Enabled
    Toolbar1.Buttons(10).Enabled = Not Toolbar1.Buttons(10).Enabled
    
    Toolbar1.Buttons(11).Enabled = Not Toolbar1.Buttons(11).Enabled
    Toolbar1.Buttons(12).Enabled = Not Toolbar1.Buttons(12).Enabled
    Toolbar1.Buttons(13).Enabled = Not Toolbar1.Buttons(13).Enabled
    
End Sub

Private Sub Nuevo()
    QueHace = 1
    ActivaTool
    Blanquea
    Habilitar_Obj True
    Label1.Caption = "Agregando Items"
    '------------
    Dim XGrid As Integer
    For XGrid = 0 To fg.Count - 1
        fg(XGrid).Editable = flexEDKbdMouse
        fg(XGrid).SelectionMode = flexSelectionFree
    Next XGrid
    '------------
    
    GRID_COMBOLIST fg(0), 2

    IdCorrelativo = 1
    
    For XGrid = 1 To 2
        fg(XGrid).ColFormat(2) = "###,###.00"
        fg(XGrid).ColFormat(3) = "###,###.00"
        fg(XGrid).ColFormat(5) = FORMAT_MONTO
    Next XGrid
    txt_cb(0).SetFocus
End Sub


Function Grabar() As Boolean
    If VALIDAR_DATOS() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modficar") + " el Registro", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo Salir
    
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstDet_1 As New ADODB.Recordset
    Dim TmpRst As New ADODB.Recordset '--PARA BUSCAR SI EL NUMERO DE CABECERA YA ESTA REGISTRADO
    
    Dim xCod As Integer
    Dim xCodDet As Integer '--al detalle
    Dim xCol, xFil As Integer
    Dim xCorr As Integer
    Dim vCorr As Integer '--correlativo
    
'On Error GoTo LaCague
    Me.MousePointer = vbHourglass
    xCon.BeginTrans
    
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT top 1 * FROM pvt_items", xCon
        RST_Busq RstDet, "SELECT top 1 * FROM pvt_descgeneral", xCon
        RST_Busq RstDet_1, "SELECT top 1 * FROM pvt_desccorporativo", xCon
        xCod = NulosN(txt_cb(0).Text)  '--CODIGO DE ITEM
        RstCab.AddNew
        RstCab("iditem") = xCod
    Else
        RST_Busq RstCab, "SELECT * FROM pvt_items WHERE iditem =" & RstFrm("id") & "", xCon
        xCon.Execute "DELETE * FROM pvt_desccorporativo WHERE iditem = " & RstFrm("id") & ""
        xCon.Execute "DELETE * FROM pvt_descgeneral WHERE iditem = " & RstFrm("id") & ""
        
        RST_Busq RstDet, "SELECT top 1 * FROM pvt_descgeneral", xCon
        RST_Busq RstDet_1, "SELECT top 1 * FROM pvt_desccorporativo", xCon
        xCod = RstFrm("id")
    End If
    
    RstCab("precio") = NulosN(Format(Trim(txt(1).Text) & "", "##.000"))
    RstCab.Update
    '--DESCUENTO EN GENERAL
    With fg(2)
        For xFil = 1 To .Rows - 1
            If NulosN(.TextMatrix(xFil, 2)) > 0 And .TextMatrix(xFil, 3) <> "" Then
                RstDet.AddNew
                '--LLAVE
                RstDet("iditem") = xCod
                RstDet.Fields("corr") = xFil
                '--FIN LLAVE
                RstDet.Fields("inicio") = NulosN(.TextMatrix(.Row, 2))
                RstDet.Fields("fin") = NulosN(.TextMatrix(.Row, 3))
                RstDet.Fields("porcentaje") = NulosN(Replace(.TextMatrix(.Row, 4), "%", "")) / 100
                RstDet.Fields("valor") = NulosN(.TextMatrix(.Row, 5))
                
                RstDet.Update
                
                
            End If
        Next xFil
    End With
    '--DESCUENTO CORPORATIVO
    With fg(0)
        For xFil = 1 To .Rows - 1
            If NulosN(.TextMatrix(xFil, 1)) > 0 And .TextMatrix(xFil, 3) <> "" Then
                '----
                RstTmp.Filter = "idcli=" + .TextMatrix(xFil, 1)
                RstTmp.Sort = "inicio"
                vCorr = 1
                If RstTmp.RecordCount > 0 Then RstTmp.MoveFirst
                Do While Not RstTmp.EOF
                    RstDet_1.AddNew
                    RstDet_1.Fields("iditem") = xCod
                    RstDet_1.Fields("idcli") = NulosN(fg(0).TextMatrix(fg(0).Row, 1))
                    RstDet_1.Fields("corr") = vCorr
                    
                    RstDet_1.Fields("inicio") = RstTmp.Fields("inicio")
                    RstDet_1.Fields("fin") = RstTmp.Fields("fin")
                    
                    RstDet_1.Fields("porcentaje") = RstTmp.Fields("porcentaje")
                    
                    RstDet_1.Fields("valor") = RstTmp.Fields("valor")
                    RstDet_1.Update
                    vCorr = vCorr + 1
                    RstTmp.MoveNext
                Loop
            End If
        Next xFil
    End With

    MsgBox "El registro se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo
    
    xCon.CommitTrans
    Grabar = True
Salir:
    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstDet_1 = Nothing:    Set TmpRst = Nothing
    Me.MousePointer = vbDefault
    Exit Function
LaCague:
    xCon.RollbackTrans
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstDet_1 = Nothing:    Set TmpRst = Nothing
    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo :"
    Grabar = False
End Function


Private Function VALIDAR_DATOS() As Boolean
    '--VALIDAR QUE LA GRILLA DE ACTIVO Y PASIVO TENGAN VALORES TANTO DE ORDEN Y DESCRIPCION
    Dim Q_ROW  As Long
    Dim QGrid As Integer
    
    If Trim(txt_cb(0).Text) = "" Then
        MsgBox "Seleccione el Item", vbExclamation, xTitulo
        txt_cb(0).SetFocus
        Exit Function
    End If
    If IsNumeric(txt(1).Text) = False Then
        MsgBox "Ingrese el Precio Unitario", vbExclamation, xTitulo
        txt(1).Text = ""
        txt(1).SetFocus
        Exit Function
    End If
    
    '--------------------------------
    '--VALIDAR QUE EL REGISTRO NO ESTE REGISTRADO
    Dim RstTmp As New ADODB.Recordset
    Set RstTmp = Nothing
    '--------------------------------
    With fg(QGrid)
        For Q_ROW = 1 To .Rows - 1
'            If IsNumeric(.TextMatrix(Q_ROW, 3)) = False Or .TextMatrix(Q_ROW, 3) = "0" Then
'                MsgBox "Ingrese El N° de Orden:", vbExclamation, xTitulo
'                Agregando = True:  .Row = Q_ROW: .Col = 3: Agregando = False
'
'                Exit Function
'            ElseIf .TextMatrix(Q_ROW, 4) = "" Then
'                MsgBox "Ingrese la Descripción:", vbExclamation, xTitulo
'                Agregando = True:  .Row = Q_ROW: .Col = 4: Agregando = False
'
'                Exit Function
'            End If
        Next Q_ROW
    End With
    '-----
    VALIDAR_DATOS = True
End Function
 

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.Index = 2 Then
        If ButtonMenu.Index = 1 Then
            Modificar
        End If
        If ButtonMenu.Index = 2 Then
            ActivarItem
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


Sub Buscar()
    On Error GoTo error
    TabOne1.CurrTab = 0
     
    Dim xRs As New ADODB.Recordset
    Dim N_SQL As String
   
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "descripcion":  xCampos(0, 2) = "3500":     xCampos(0, 3) = "C"
    xCampos(1, 0) = "Tipo Producto":    xCampos(1, 1) = "tipproddesc":  xCampos(1, 2) = "1500":     xCampos(1, 3) = "C"
    xCampos(2, 0) = "Unid.":            xCampos(2, 1) = "abrev":        xCampos(2, 2) = "450":      xCampos(2, 3) = "C"
    xCampos(3, 0) = "M":                xCampos(3, 1) = "simbolo":      xCampos(3, 2) = "450":      xCampos(3, 3) = "C"
        
    N_SQL = "SELECT pvt_items.iditem as id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, mae_moneda.simbolo, pvt_items.precio, mae_tipoproducto.descripcion AS tipproddesc, pvt_items.activo " _
        + vbCr + " FROM (mae_unidades RIGHT JOIN (mae_tipoproducto RIGHT JOIN (mae_moneda RIGHT JOIN alm_inventario ON mae_moneda.id = alm_inventario.idmon) ON mae_tipoproducto.id = alm_inventario.tippro) ON mae_unidades.id = alm_inventario.idunimed) RIGHT JOIN pvt_items ON alm_inventario.id = pvt_items.iditem " _
        + vbCr + " ORDER BY alm_inventario.descripcion;"
        
    CARGAR_DLL_EPSBUSCAR xCon, xRs, N_SQL, xCampos(), "Buscando Items", "descripcion", "descripcion", Principio
    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir
    
    RstFrm.MoveFirst
    RstFrm.Find "id = " + CStr(xRs("id"))
Salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub




Private Sub Filtrar()
    
    Dim xCampos(2, 4) As String
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha

    xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "descripcion":  xCampos(0, 2) = "C":     xCampos(0, 3) = "3500"
    xCampos(1, 0) = "Código Item":      xCampos(1, 1) = "tipproddesc":  xCampos(1, 2) = "C":     xCampos(1, 3) = "1500"
    xCampos(2, 0) = "Tipo Producto":    xCampos(2, 1) = "tipproddesc":  xCampos(2, 2) = "C":     xCampos(2, 3) = "1500"
    xCampos(3, 0) = "Unidad":           xCampos(3, 1) = "abrev":        xCampos(3, 2) = "C":      xCampos(3, 3) = "450"
    xCampos(4, 0) = "Moneda":           xCampos(4, 1) = "simbolo":      xCampos(4, 2) = "C":      xCampos(4, 3) = "450"
    
    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg3

    TabOne1.CurrTab = 0
End Sub


Private Sub Imprimir(Optional IMP_LISTADO As Boolean = False)

    On Error GoTo error

    Me.MousePointer = vbHourglass
    If IMP_LISTADO = False Then
        If Me.TabOne1.CurrTab = 0 Then
        
        Else
''            MsgBox "Primero muestre el detalle del Registro" + vbCr + _
''                   "Luego inténtelo otra vez", vbExclamation, xTitulo
        End If
    Else
    
        TDB_IMPRIMIR Dg3, "IMPRESIÓN DE ESTADOS", " "
   
    End If

    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "IMPRIMIR"

End Sub



Private Function GENERAR_CONSULTA(X_ID As String) As String
    
    Dim N_SQL As String


    N_SQL = "SELECT pvt_desccorporativo.idcli, pvt_desccorporativo.corr, pvt_desccorporativo.inicio, pvt_desccorporativo.fin, pvt_desccorporativo.porcentaje, pvt_desccorporativo.valor " _
        + vbCr + " FROM pvt_desccorporativo " _
        + vbCr + " WHERE (((pvt_desccorporativo.iditem) = " + X_ID + ")) " _
        + vbCr + " ORDER BY pvt_desccorporativo.idcli, pvt_desccorporativo.inicio;"


    GENERAR_CONSULTA = N_SQL
End Function


Private Sub DEFINIR_RST()
    '--DEFINIR EL RECORSET TEMPORAL PARA INSUMO Y TAREA
    Dim RST_ORIGEN As New ADODB.Recordset
    Dim N_SQL As String
    N_SQL = GENERAR_CONSULTA("-1")
    RST_Busq RST_ORIGEN, N_SQL, xCon
    DEFINIR_RST_TMP RstTmp, RST_ORIGEN
    Set RST_ORIGEN = Nothing
    
End Sub
Sub Retirar()
    Dim Rpta As Integer
    Rpta = MsgBox("Esta seguro de retirar el item " + StrConv(Trim(RstFrm("descripcion") & ""), 3), vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "UPDATE pvt_items SET pvt_items.activo = 0 WHERE (((pvt_items.iditem)=" & RstFrm("id") & "))"
        MsgBox "El item " + Trim(RstFrm("descripcion")) + " se Retiró con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstFrm.Requery
        Dg3.Refresh
    End If
End Sub


Sub ActivarItem()
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    Dim N_SQL As String
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "6000":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "codpro":         xCampos(1, 2) = "2000":    xCampos(1, 3) = "C"
    
    N_SQL = "SELECT alm_inventario.id, alm_inventario.descripcion, alm_inventario.codpro, pvt_items.activo " _
        + vbCr + " FROM alm_inventario INNER JOIN pvt_items ON alm_inventario.id = pvt_items.iditem " _
        + vbCr + " WHERE (((pvt_items.activo) = 0)) " _
        + vbCr + " ORDER BY alm_inventario.descripcion;"

    CARGAR_DLL_EPSBUSCAR xCon, xRs, N_SQL, xCampos(), "Buscando Items Retirados", "descripcion", "descripcion", Principio

    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir
    Dim Rpta As Integer
    Rpta = MsgBox("¿Esta seguro de activar el item seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "UPDATE pvt_items SET pvt_items.activo = -1 WHERE (((pvt_items.iditem)=" & xRs("id") & "))"
        MsgBox "El item " + Trim(xRs("descripcion")) + " se activo con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstFrm.Requery
        Dg3.Refresh
    End If
Salir:
    Set xRs = Nothing
End Sub




'----------------------------
'----------------------------

Private Sub cb_Click(Index As Integer)
    Dim xRs As New ADODB.Recordset
    Dim xCampos() As String
    Dim N_SQL As String
    On Error GoTo error
    If QueHace = 3 Then Exit Sub
    
    N_SQL = "SELECT alm_inventario.id, alm_inventario.descripcion, alm_inventario.codpro, mae_tipoproducto.descripcion AS tipproddesc, mae_unidades.abrev, mae_moneda.simbolo,alm_inventario.preuni " _
        + vbCr + " FROM mae_unidades RIGHT JOIN (mae_moneda RIGHT JOIN (mae_tipoproducto RIGHT JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) ON mae_moneda.id = alm_inventario.idmon) ON mae_unidades.id = alm_inventario.idunimed " _
        + vbCr + " WHERE (((alm_inventario.id) Not In (select  [pvt_items]![iditem]  from pvt_items )) AND ((alm_inventario.activo)=-1)) " _
        + vbCr + " ORDER BY alm_inventario.descripcion;"

    ReDim xCampos(4, 3) As String
    xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "descripcion":  xCampos(0, 2) = "5500":     xCampos(0, 3) = "C"
    xCampos(1, 0) = "Tipo Producto":    xCampos(1, 1) = "tipproddesc":  xCampos(1, 2) = "1500":     xCampos(1, 3) = "C"
    xCampos(2, 0) = "Unid.":            xCampos(2, 1) = "abrev":        xCampos(2, 2) = "450":      xCampos(2, 3) = "C"
    xCampos(3, 0) = "M":                xCampos(3, 1) = "simbolo":      xCampos(3, 2) = "450":      xCampos(3, 3) = "C"

    CARGAR_DLL_EPSBUSCAR xCon, xRs, N_SQL, xCampos(), "Buscando Items", "descripcion", "descripcion", Principio

    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir
    txt_cb(Index).Text = xRs.Fields("id") & ""  '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields("descripcion") & "" '--NOMBRE
    lbl_cb_cod(Index).Caption = xRs.Fields("codpro") & "" '--CODIGO
    txt(1).Text = xRs.Fields("preuni") & ""
    txt(2).Text = xRs.Fields("abrev") & ""
    txt(3).Text = xRs.Fields("simbolo") & ""
    '----------------
    txt(1).SetFocus
    
Salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo error
    If QueHace = 3 Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb_Click Index
        Exit Sub
    End If
    Exit Sub
error:
    SHOW_ERROR Me.Name, "txt_cb_KeyDown(" + CStr(Index) + ")"
End Sub
'----------------------------
'----------------------------
Private Sub txt_Change(Index As Integer)

End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmManRegProd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Produccion - Parte de Producción"
   ClientHeight    =   7410
   ClientLeft      =   165
   ClientTop       =   1590
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   11835
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
            Picture         =   "FrmManRegProd.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegProd.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegProd.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegProd.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegProd.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegProd.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegProd.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegProd.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegProd.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegProd.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegProd.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegProd.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegProd.frx":277E
            Key             =   "IMG13"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7020
      Left            =   0
      TabIndex        =   10
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
         TabIndex        =   14
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6090
            Left            =   0
            TabIndex        =   15
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
            Columns(1).Caption=   "Fch. Prod."
            Columns(1).DataField=   "fchdoc"
            Columns(1).NumberFormat=   "Short Date"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nº Documento"
            Columns(2).DataField=   "numdoc"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "N° OP"
            Columns(3).DataField=   "numop"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Responsable"
            Columns(4).DataField=   "responsable"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Glosa"
            Columns(5).DataField=   "glosa"
            Columns(5).NumberFormat=   "Short Date"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2223"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2143"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=2963"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2884"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=3122"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=3043"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=5847"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=5768"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=5186"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=5106"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=131588"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
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
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=110,.parent=75"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=107,.parent=76"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=108,.parent=77"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=109,.parent=79"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=16,.parent=75"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=13,.parent=76"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=14,.parent=77"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=15,.parent=79"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=20,.parent=75"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=17,.parent=76"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=18,.parent=77"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=19,.parent=79"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=118,.parent=75,.alignment=3"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=115,.parent=76,.alignment=2"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=116,.parent=77,.alignment=2"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=117,.parent=79"
            _StyleDefs(60)  =   "Named:id=33:Normal"
            _StyleDefs(61)  =   ":id=33,.parent=0"
            _StyleDefs(62)  =   "Named:id=34:Heading"
            _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(64)  =   ":id=34,.wraptext=-1"
            _StyleDefs(65)  =   "Named:id=35:Footing"
            _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(67)  =   "Named:id=36:Selected"
            _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(69)  =   "Named:id=37:Caption"
            _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(71)  =   "Named:id=38:HighlightRow"
            _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=39:EvenRow"
            _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(75)  =   "Named:id=40:OddRow"
            _StyleDefs(76)  =   ":id=40,.parent=33"
            _StyleDefs(77)  =   "Named:id=41:RecordSelector"
            _StyleDefs(78)  =   ":id=41,.parent=34"
            _StyleDefs(79)  =   "Named:id=42:FilterBar"
            _StyleDefs(80)  =   ":id=42,.parent=33"
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label lblperiodo 
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
            Index           =   0
            Left            =   9810
            TabIndex        =   16
            Top             =   0
            Width           =   1860
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Consulta de Producción"
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
         TabIndex        =   11
         Top             =   375
         Width           =   11790
         Begin VB.CommandButton Cmd 
            Enabled         =   0   'False
            Height          =   240
            Index           =   3
            Left            =   5775
            Picture         =   "FrmManRegProd.frx":2B10
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   780
            Width           =   240
         End
         Begin VB.Frame Frame4 
            Height          =   600
            Left            =   80
            TabIndex        =   25
            Top             =   5880
            Width           =   11655
            Begin VB.CommandButton Cmd 
               Caption         =   "Eliminar Item"
               Enabled         =   0   'False
               Height          =   330
               Index           =   2
               Left            =   1440
               TabIndex        =   9
               Top             =   180
               Width           =   1305
            End
            Begin VB.CommandButton Cmd 
               Caption         =   "Agregar Item"
               Enabled         =   0   'False
               Height          =   330
               Index           =   1
               Left            =   90
               TabIndex        =   8
               Top             =   180
               Width           =   1305
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "( Periodo )"
            Height          =   1100
            Left            =   9960
            TabIndex        =   23
            Top             =   360
            Width           =   1770
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
               Left            =   240
               TabIndex        =   24
               Top             =   480
               Width           =   1245
            End
         End
         Begin VB.TextBox GlosaText 
            Height          =   300
            Left            =   1170
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   7
            Text            =   "GlosaText"
            Top             =   1110
            Width           =   8565
         End
         Begin VB.CommandButton Cmd 
            Enabled         =   0   'False
            Height          =   240
            Index           =   0
            Left            =   5775
            Picture         =   "FrmManRegProd.frx":2C42
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   480
            Width           =   240
         End
         Begin VB.TextBox NumSerText 
            Height          =   300
            Left            =   1170
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   1
            Text            =   "NumS"
            Top             =   750
            Width           =   915
         End
         Begin VB.TextBox NumDocText 
            Height          =   300
            Left            =   2265
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   2
            Text            =   "NumDocText"
            Top             =   750
            Width           =   1440
         End
         Begin AspaTextBoxFecha.TextBoxFecha FechaText 
            Height          =   300
            Left            =   1170
            TabIndex        =   0
            Top             =   450
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
         Begin VB.TextBox IdResponsableText 
            Height          =   300
            Left            =   5130
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   3
            Text            =   "IdResponsableText"
            Top             =   450
            Width           =   915
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   4335
            Left            =   75
            TabIndex        =   26
            Top             =   1560
            Width           =   11655
            _cx             =   20558
            _cy             =   7646
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
            Cols            =   17
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmManRegProd.frx":2D74
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
         Begin VB.TextBox IdAlmOrigenText 
            Height          =   300
            Left            =   5130
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   5
            Text            =   "IdAlmOrigen"
            Top             =   750
            Width           =   915
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Almacén"
            Height          =   195
            Index           =   1
            Left            =   4080
            TabIndex        =   28
            Top             =   795
            Width           =   615
         End
         Begin VB.Label AlmOrigenLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "AlmOrigenLabel"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   6075
            TabIndex        =   27
            Top             =   750
            Width           =   3660
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Glosa"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   22
            Top             =   1155
            Width           =   405
         End
         Begin VB.Label ResponsableLabel 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ResponsableLabel"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   6075
            TabIndex        =   20
            Top             =   450
            Width           =   3660
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Responsable"
            Height          =   195
            Index           =   2
            Left            =   4080
            TabIndex        =   19
            Top             =   480
            Width           =   930
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000001&
            BackStyle       =   1  'Opaque
            Height          =   90
            Left            =   2115
            Top             =   870
            Width           =   105
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Num. Doc."
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   18
            Top             =   795
            Width           =   765
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Producción"
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
            Caption         =   "Fch. Prod."
            Height          =   195
            Index           =   3
            Left            =   105
            TabIndex        =   12
            Top             =   480
            Width           =   735
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   11835
      _ExtentX        =   20876
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
            Object.Visible         =   0   'False
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
Attribute VB_Name = "FrmManRegProd"
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
        Case 0 ' Responsable
            ReDim xCampos(2, 4) As String
            
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Apellidos y Nombres":  xCampos(0, 1) = "apenom":     xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":               xCampos(1, 1) = "id":         xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
                      
            
            nTitulo = "Buscando Responsables"
            
            cSQL = "SELECT pla_empleados.nombre AS apenom, pla_empleados.id " _
                + vbCr + "FROM pla_empleados " _
                + vbCr + "ORDER BY pla_empleados.nombre;"
                        
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "apenom", "apenom", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            IdResponsableText.Text = NulosN(xRs("id"))
            ResponsableLabel.Caption = NulosC(xRs("apenom"))
            IdAlmOrigenText.SetFocus
            
            Set xRs = Nothing
        
        Case 1 ' Agregar Item
            AgregarItem
            
        Case 2 ' Eliminar Item
            EliminarItem
            
        Case 3 ' Agregar Almacen
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
            
            IdAlmOrigenText.Text = NulosN(xRs("id"))
            AlmOrigenLabel.Caption = UCase(NulosC(xRs("descripcion")))
            GlosaText.SetFocus
            Set xRs = Nothing
        
    End Select
End Sub

Sub AgregarItem()
    Dim nTitulo As String
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion                     'campo                       'tamaño                         'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "desitem":       xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "codpro":        xCampos(1, 2) = "1500":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Uni. Med.":     xCampos(2, 1) = "desunimed":     xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
      
    cSQL = "SELECT pro_receta.iditem, alm_inventario.descripcion AS desitem, alm_inventario.codpro, pro_receta.id AS idrec, pro_receta.codrec, pro_receta.idunimed, mae_unidades.abrev AS desunimed " _
        + vbCr + "FROM (((pro_receta LEFT JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_receta.idunimed = mae_unidades.id) LEFT JOIN pro_tiptrab ON pro_receta.idtiptrab = pro_tiptrab.id) LEFT JOIN pro_formapag ON pro_receta.idformapag = pro_formapag.id " _
        + vbCr + "WHERE (((pro_receta.prirec)=1) AND ((alm_inventario.activo)=-1));"
        
    nTitulo = "Buscando Ítems"
    
    Set xRs = Nothing
    CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                    "codpro", "codpro", Principio, ""
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    
    Fg1.Rows = Fg1.Rows + 1
'    ITEM
    Fg1.TextMatrix(Fg1.Rows - 1, Fg1.ColIndex("ITEM")) = NulosC(xRs("desitem"))
    Fg1.TextMatrix(Fg1.Rows - 1, Fg1.ColIndex("CODIGO")) = NulosC(xRs("codpro"))
    Fg1.TextMatrix(Fg1.Rows - 1, Fg1.ColIndex("IDITEM")) = NulosN(xRs("iditem"))
'    RECETA
    Fg1.TextMatrix(Fg1.Rows - 1, Fg1.ColIndex("IDRECETA")) = NulosN(xRs("idrec"))
    Fg1.TextMatrix(Fg1.Rows - 1, Fg1.ColIndex("RECETA")) = NulosC(xRs("codrec"))
'    UNIDADES DE MEDIDA
    Fg1.TextMatrix(Fg1.Rows - 1, Fg1.ColIndex("IDUNIMED")) = NulosN(xRs("idunimed"))
    Fg1.TextMatrix(Fg1.Rows - 1, Fg1.ColIndex("UM")) = NulosC(xRs("desunimed"))
    
    Fg1.Select Fg1.Rows - 1, Fg1.ColIndex("ORDPROD")
    Fg1.SetFocus
End Sub

Sub agregarOrden()
    Dim nTitulo As String
    Dim xRs As New ADODB.Recordset
    Dim xCampos(5, 4) As String
    Dim IdOrdenProd As Integer
    Dim IdReceta As Integer
    Dim F As New SistemaLogica.Funciones
            
    xCampos(0, 0) = "Fecha.":           xCampos(0, 1) = "fchpro":           xCampos(0, 2) = "1000":          xCampos(0, 3) = "C"
    xCampos(1, 0) = "Num. Ord.":        xCampos(1, 1) = "numdoc":           xCampos(1, 2) = "1400":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Lote":             xCampos(2, 1) = "lote":             xCampos(2, 2) = "1400":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Responsable":      xCampos(3, 1) = "desresp":          xCampos(3, 2) = "1900":         xCampos(3, 3) = "C"
    xCampos(4, 0) = "Cantidad":         xCampos(4, 1) = "cantidad":         xCampos(4, 2) = "900":          xCampos(4, 3) = "N"
      
    IdOrdenProd = F.NuloNumeric(Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("IDORDPROD")))
    IdReceta = F.NuloNumeric(Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("IDRECETA")))
    
    cSQL = "SELECT pro_ordenprod.id, pro_ordenprod.fchpro, [pro_ordenprod].[numser] & '-' & [pro_ordenprod].[numdoc] AS numdoc, pro_ordenprod.lote, pro_ordenprod.cantidad, pla_empleados.nombre AS desresp " _
        + vbCr + "FROM pro_ordenprod LEFT JOIN pla_empleados ON pro_ordenprod.idresp = pla_empleados.id " _
        + vbCr + "WHERE (((pro_ordenprod.idrec)=" & IdReceta & ") AND ((pro_ordenprod.estado)=2) AND ((pro_ordenprod.id) Not In (SELECT pro_ordenprod.id FROM pro_producciondet INNER JOIN pro_ordenprod ON pro_producciondet.idord = pro_ordenprod.id WHERE pro_ordenprod.id <> " & IdOrdenProd & ")))"

    nTitulo = "Buscando Ordenes"
    
    Set xRs = Nothing
    CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                    "numdoc", "numdoc", Principio, ""
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    
'    ITEM
    Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("ORDPROD")) = NulosC(xRs("numdoc"))
    Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("LOTE")) = NulosC(xRs("lote"))
    Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("CANPROG")) = Format(NulosN(xRs("cantidad")), FORMAT_CANTIDAD)
    Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("IDORDPROD")) = NulosN(xRs("id"))
    
    Fg1.Select Fg1.Row, Fg1.ColIndex("CANPROD")
    Fg1.SetFocus
End Sub
 
Sub EliminarItem()
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
    
    Fg1.RemoveItem Fg1.Row
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

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If QueHace = 3 Then Exit Sub
    
    Select Case Col
        Case Fg1.ColIndex("ORDPROD")
            agregarOrden
    End Select
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If Row <= 0 Then Exit Sub
    
    Select Case Col
        Case Fg1.ColIndex("CANPROD")
            Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), FORMAT_CANTIDAD)
            
        Case Fg1.ColIndex("HORINI")
            If (Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("HORINI")) = "") Then Exit Sub
            If (Not IsDate(Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("HORINI")))) Then
                MsgBox "Ingrese correctamente la hora de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("HORINI")) = ""
                Exit Sub
            End If
            If (Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("HORFIN")) <> "") Then
                If (CDate(Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("HORINI"))) > CDate(Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("HORFIN")))) Then
                    MsgBox "La hora de inicio no puede ser mayor a la hora de fin", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("HORINI")) = ""
                End If
            End If
        Case Fg1.ColIndex("HORFIN")
            If (Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("HORFIN")) = "") Then Exit Sub
            If (Not IsDate(Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("HORFIN")))) Then
                MsgBox "Ingrese correctamente la hora de fin", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("HORFIN")) = ""
                Exit Sub
            End If
            If (Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("HORINI")) <> "") Then
                If (CDate(Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("HORINI"))) > CDate(Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("HORFIN")))) Then
                    MsgBox "La hora de fin no puede ser menor a la hora de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    Fg1.TextMatrix(Fg1.Row, Fg1.ColIndex("HORFIN")) = ""
                End If
            End If
        
    End Select
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then Fg1.Editable = flexEDNone: Exit Sub
    
    Select Case Fg1.Col
        Case Fg1.ColIndex("ORDPROD"), Fg1.ColIndex("CANPROD"), _
                        Fg1.ColIndex("HORINI"), Fg1.ColIndex("HORFIN"), Fg1.ColIndex("OBS")
            Fg1.Editable = flexEDKbdMouse
            
        Case Else
            Fg1.Editable = flexEDNone
    End Select
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case Col
        Case Fg1.ColIndex("CANPROD"), _
                        Fg1.ColIndex("HORINI"), Fg1.ColIndex("HORFIN")
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
            
        Case Fg1.ColIndex("OBS")
        
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    
    If KeyCode = 45 Then
        AgregarItem
    End If
    If KeyCode = 46 Then
        EliminarItem
    End If
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If QueHace = 3 Then Exit Sub
    If Button = 2 Then PopupMenu menu1
End Sub

Private Sub Form_Activate()
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
    Mostrando = False
    
    Fg1.AllowUserResizing = flexResizeColumns
    Fg1.ExplorerBar = flexExSortShow
    Fg1.SelectionMode = flexSelectionFree
    Fg1.ForeColorSel = &H80000005
    Fg1.BackColorSel = &H80&
    Fg1.ColEditMask(Fg1.ColIndex("HORINI")) = "##:##"
    Fg1.ColEditMask(Fg1.ColIndex("HORFIN")) = "##:##"
        
    GRID_COMBOLIST Fg1, Fg1.ColIndex("ORDPROD")
End Sub

Sub Nuevo()
    QueHace = 1
    TabOne1.CurrTab = 1
    ActivaTool
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Agregando Parte Produccion"
    Bloquea
    Blanquea
    
    Fg1.Editable = flexEDKbdMouse
    Fg1.Rows = Fg1.FixedRows
    Fg1.SelectionMode = flexSelectionFree
    
    xHorIni = Time
    FechaText.valor = Date
    FechaText.SetFocus
End Sub

Private Sub Form_Load()
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

Private Sub Menu1_1_Click()
    AgregarItem
End Sub

Private Sub Menu1_3_Click()
    EliminarItem
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
    Dim xCampos(3, 3) As String
    
    ' 0 Nombre a Mostrar;
    ' 1 nombre de Campo del Rst;
    ' 2 alineacion(0::derecha, 1::centro, 2::izquierda);
    ' 3 ancho de columna
    ' obs: el rst puede tener mas columnas solo se consideran los campos del array
    xCampos(0, 0) = "Fch. Prod.":           xCampos(0, 1) = "fchdoc":           xCampos(0, 2) = 0:  xCampos(0, 3) = "1000"
    xCampos(1, 0) = "Nº Documento":         xCampos(1, 1) = "numdoc":           xCampos(1, 2) = 0:  xCampos(1, 3) = "1200"
    xCampos(2, 0) = "Responsable":          xCampos(2, 1) = "responsable":      xCampos(2, 2) = 0:  xCampos(2, 3) = "2500"
    xCampos(3, 0) = "Glosa":                xCampos(3, 1) = "glosa":            xCampos(3, 2) = 0:  xCampos(3, 3) = "2500"
    '**********************************************************************************************************************************
        
    Set RstTmp = RstIng
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "Lista de Parte de Produccion", "Periodo: " & lblperiodo(0).Caption & "  -  " & AnoTra, "", "Listado de Parte de Produccion", RstTmp, xCampos
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

Function validarDatos() As Boolean
    Dim A As Integer
    
    If FechaText.valor = "" Then
        MsgBox "No ha especificado la fecha de movimiento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        FechaText.SetFocus
        validarDatos = False
        Exit Function
    End If
    If Year(FechaText.valor) <> AnoTra Then
        MsgBox "El año ingresado en la " & Label3(3).Caption & " no coincide con el Ejercicio" & vbCr & "Corrija la fecha o registre en su año que corresponde", vbInformation, xTitulo
        FechaText.valor = ""
        FechaText.SetFocus
        validarDatos = False
        Exit Function
    End If
    If NulosC(NumSerText.Text) = "" Then
        MsgBox "No ha especificado el numero de serie del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        NumSerText.SetFocus
        validarDatos = False
        Exit Function
    End If
    If NulosC(NumDocText.Text) = "" Then
        MsgBox "No ha especificado el numero de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        NumDocText.SetFocus
        validarDatos = False
        Exit Function
    End If
    If NulosN(IdResponsableText.Text) = 0 Then
        MsgBox "No ha especificado el responsable el movimiento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        IdResponsableText.SetFocus
        validarDatos = False
        Exit Function
    End If
    If NulosN(IdAlmOrigenText.Text) = 0 Then
        MsgBox "No ha especificado el almacen para el registro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        IdAlmOrigenText.SetFocus
        validarDatos = False
        Exit Function
    End If
    If Fg1.Rows = Fg1.FixedRows Then
        MsgBox "Falta ingresar el detalle...!", vbExclamation, "Mensaje...!"
        validarDatos = False
        Exit Function
    End If
    
    For A = 1 To Fg1.Rows - 1
        If (Not IsDate(Fg1.TextMatrix(A, Fg1.ColIndex("HORINI")))) Then
            MsgBox "La hora de inicio ingresada no es correcta", vbExclamation, "Mensaje...!"
            Fg1.TextMatrix(A, Fg1.ColIndex("HORINI")) = ""
            Fg1.Select A, Fg1.ColIndex("HORINI")
            Fg1.SetFocus
            validarDatos = False
            Exit Function
        End If
        If (Not IsDate(Fg1.TextMatrix(A, Fg1.ColIndex("HORFIN")))) Then
            MsgBox "La hora de fin ingresada no es correcta", vbExclamation, "Mensaje...!"
            Fg1.TextMatrix(A, Fg1.ColIndex("HORFIN")) = ""
            Fg1.Select A, Fg1.ColIndex("HORFIN")
            Fg1.SetFocus
            validarDatos = False
            Exit Function
        End If
    Next
    
    
    validarDatos = True
End Function
'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : Function
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA Alm_ingreso, DEVUELVE VERDADERO SI TIENE EXITO
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Function Grabar() As Boolean
    Dim F As New SistemaLogica.Funciones
    Dim ParteProduccion As New ProduccionEntidad.EParteProd
    Dim A As Integer

On Error GoTo BloqueError
    If Not validarDatos Then Grabar = False: Exit Function
    
    ' Se llenan cabecera
    If QueHace = 1 Then ParteProduccion.IdParteProduccion = 0 Else ParteProduccion.IdParteProduccion = NulosN(RstIng("id"))
    ParteProduccion.FechaParteProduccion = Format(FechaText.valor, "dd/mm/yyyy")
    ParteProduccion.NumeroSerie = NulosC(NumSerText.Text)
    ParteProduccion.NumeroDocumento = NulosC(NumDocText.Text)
    ParteProduccion.IdResponsable = NulosN(IdResponsableText.Text)
    ParteProduccion.IdEstado = 1
    ParteProduccion.Glosa = NulosC(GlosaText.Text)
    ParteProduccion.IdAlmacen = NulosN(IdAlmOrigenText.Text)
    ParteProduccion.MesTrabajo = mMesActivo
    ParteProduccion.AnhoTrabajo = AnoTra
    ' Se llena detalle
    For A = 1 To Fg1.Rows - 1
        Dim ParteProduccionDet As New ProduccionEntidad.EParteProdDet
        
        ParteProduccionDet.IDITEM = NulosN(Fg1.TextMatrix(A, Fg1.ColIndex("IDITEM")))
        ParteProduccionDet.CantidadProgramada = NulosN(Fg1.TextMatrix(A, Fg1.ColIndex("CANPROG")))
        ParteProduccionDet.CantidadProducida = NulosN(Fg1.TextMatrix(A, Fg1.ColIndex("CANPROD")))
        ParteProduccionDet.OrdenProduccion = NulosC(Fg1.TextMatrix(A, Fg1.ColIndex("ORDPROD")))
        ParteProduccionDet.CodigoItem = NulosC(Fg1.TextMatrix(A, Fg1.ColIndex("CODIGO")))
        ParteProduccionDet.Item = NulosC(Fg1.TextMatrix(A, Fg1.ColIndex("ITEM")))
        ParteProduccionDet.IdOrdenProduccion = NulosN(Fg1.TextMatrix(A, Fg1.ColIndex("IDORDPROD")))
        ParteProduccionDet.IdUnidadMedida = NulosN(Fg1.TextMatrix(A, Fg1.ColIndex("IDUNIMED")))
        ParteProduccionDet.IdReceta = NulosN(Fg1.TextMatrix(A, Fg1.ColIndex("IDRECETA")))
        ParteProduccionDet.Glosa = NulosC(Fg1.TextMatrix(A, Fg1.ColIndex("OBS")))
        ParteProduccionDet.HoraInicio = CDate(Fg1.TextMatrix(A, Fg1.ColIndex("HORINI")))
        ParteProduccionDet.HoraFin = CDate(Fg1.TextMatrix(A, Fg1.ColIndex("HORFIN")))
        
        ParteProduccion.LParteProdDet.Add ParteProduccionDet
        Set ParteProduccionDet = Nothing
    Next A
    
    xCon.BeginTrans
    Set ParteProduccion.Conexion = xCon
    ParteProduccion.Called = True
    
    ' Creamos el movimiento de ingreso automatico en almacen
    If F.NuloNumeric(F.KeyValue("CreacionMovimientoAutoProduccion", xCon)) = -1 Then
        ' Se valida la fecha de cierre de mes
        If F.MesCerradoOpcion(mMesActivo, CLng(F.KeyValue("IdOpcionSistemaMovimientoAlmacen", xCon)), xCon) Then
            Err.Raise &HFFFFFF01, , "No se puede crear el movimiento de ingreso automatico. El presente mes para la opcion: [Ingresos y Salidas de Almacén] se encuentra cerrado, modifique la fecha o aperture el mes cerrado"
        End If
        GrabarMovimiento ParteProduccion
    End If
    ' Creamos el movimiento de salida automatico en almacen
    If F.NuloNumeric(F.KeyValue("GeneraDespachoAutomaticoParte", xCon)) = -1 Then
        ' Se valida la fecha de cierre de mes
        If F.MesCerradoOpcion(mMesActivo, CLng(F.KeyValue("IdOpcionSistemaMovimientoAlmacen", xCon)), xCon) Then
            Err.Raise &HFFFFFF01, , "No se puede crear el movimiento de salida automatico. El presente mes para la opcion: [Ingresos y Salidas de Almacén] se encuentra cerrado, modifique la fecha o aperture el mes cerrado"
        End If
        GrabarMovimientoDespacho ParteProduccion
    End If
    ' Grabamos el parte
    If Not ParteProduccion.Save(0, "") Then Err.Raise &HFFFFFF01, , F.ErrorDescriptionDLL(Err.LastDllError)
    
    xCon.CommitTrans

    MsgBox "El Parte de Produccion se registró con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    mIdRegistro = ParteProduccion.IdParteProduccion
    Set ParteProduccion = Nothing
    Grabar = True
    Exit Function

BloqueError:
    xCon.RollbackTrans
    MsgBox "No se pudo registrar el Parte de Produccion por el siguiente motivo :" + Trim(Err.Description)
    Set ParteProduccion = Nothing
    Grabar = False
End Function

Sub GrabarMovimiento(mParteProd As ProduccionEntidad.EParteProd)
    Dim F As New SistemaLogica.Funciones
    ' Verificamos si ya tiene registro en movimientos
    Dim database As New SistemaData.EDataBase
    Dim record As New ADODB.Recordset
    Dim Movimiento As New AlmacenEntidad.EMovimiento
    
    Set database.Connection = xCon
    database.CommandText = "SELECT alm_ingreso.id AS idmov " _
                + vbCr + "FROM alm_ingreso " _
                + vbCr + "WHERE (((alm_ingreso.idtipdocref)=" & F.NuloNumeric(F.KeyValue("IdDocumentoParteProduccion", xCon)) & ") AND ((alm_ingreso.iddocref)=" & mParteProd.IdParteProduccion & "))"
    Set record = database.GetRecordset
    ' Se eliminan todos los movimientos
    If record.RecordCount > 0 Then
        record.MoveFirst
        While Not record.EOF
            Dim mMovAux As New AlmacenEntidad.EMovimiento
            mMovAux.IdMovimiento = F.NuloNumeric(record("idmov"))
            Set mMovAux.Conexion = xCon
            mMovAux.Delete CLng(xIdUsuario), F.MachineName
            record.MoveNext
        Wend
    End If
    ' Se crea el movimiento en almacen
    ' Cabecera
    Movimiento.IdTipoMovimiento = -1
    Movimiento.FechaMovimiento = CDate(FechaText.valor)
    Movimiento.NumeroSerie = F.NuloString(NumSerText.Text)
    Movimiento.NumeroDocumento = F.HallaNumeroDocumento("alm_ingreso", "'" & Movimiento.NumeroSerie & "'", "numser", xCon)
    Movimiento.IdEstado = F.NuloNumeric(F.KeyValue("EstadoAprobadoMovimiento", xCon))
    Movimiento.IdAlmacen = F.NuloNumeric(IdAlmOrigenText.Text)
    Movimiento.Glosa = F.NuloString(GlosaText.Text)
    Movimiento.IdTipoDocumentoReferencia = F.NuloNumeric(F.KeyValue("IdDocumentoParteProduccion", xCon))
    Movimiento.IdDocumentoReferencia = mParteProd.IdParteProduccion
    Movimiento.DocumentoReferencia = F.NuloString(NumSerText.Text & " - " & NumDocText.Text)
    Movimiento.MesTrabajo = mMesActivo
    Movimiento.AnhoTrabajo = AnoTra
    ' Detalle
    Dim ParteProdDet As New ProduccionEntidad.EParteProdDet
    For Each ParteProdDet In mParteProd.LParteProdDet
        Dim MovimientoDet As New AlmacenEntidad.EMovimientoDet
        MovimientoDet.IDITEM = ParteProdDet.IDITEM
        MovimientoDet.cantidad = ParteProdDet.CantidadProducida
        MovimientoDet.CantidadTeorica = ParteProdDet.CantidadProgramada
        ' Se agrega una referencia al parte detalle
        MovimientoDet.IdDocumentoReferencia = ParteProdDet.IdParteProduccionDet
        ' Se agrega al padre
        Movimiento.LMovimientoDet.Add MovimientoDet
        Set MovimientoDet = Nothing
    Next
    ' Se graba el movimiento
    Set Movimiento.Conexion = xCon
    Movimiento.Called = True
    If Not Movimiento.Save(CLng(xIdUsuario), F.MachineName) Then Err.Raise &HFFFFFF01, , F.ErrorDescriptionDLL(Err.LastDllError)
End Sub

Sub GrabarMovimientoDespacho(mParteProd As ProduccionEntidad.EParteProd)
    Dim F As New SistemaLogica.Funciones
    Dim mIdAlmacenItem As Integer
    Dim mDataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
                
    ' Se eliminan todos los movimientos generados por el parte
    Set mDataBase.Connection = xCon
    mDataBase.CommandText = "SELECT alm_ingreso.id " _
        + vbCr + "FROM alm_ingreso " _
        + vbCr + "WHERE (((alm_ingreso.idtipdocref2)=" & F.NuloNumeric(F.KeyValue("IdDocumentoParteProduccion", xCon)) & ") AND ((alm_ingreso.iddocref2)=" & F.NuloNumeric(mParteProd.IdParteProduccion) & "))"
    Set mRecord = mDataBase.GetRecordset
    If mRecord.RecordCount > 0 Then
        mRecord.MoveFirst
        While Not mRecord.EOF
            Dim mMovAux As New AlmacenEntidad.EMovimiento
            mMovAux.IdMovimiento = F.NuloNumeric(mRecord("id"))
            Set mMovAux.Conexion = xCon
            mMovAux.Delete CLng(xIdUsuario), F.MachineName
            mRecord.MoveNext
        Wend
    End If
    Set mRecord = Nothing
                
    ' Detalle
    Dim ParteProdDet As New ProduccionEntidad.EParteProdDet
    For Each ParteProdDet In mParteProd.LParteProdDet
        mIdAlmacenItem = F.DespachaEn(ParteProdDet.IDITEM, xCon)
        If mIdAlmacenItem > 0 Then
            Dim mMovimiento As New AlmacenEntidad.EMovimiento
            Dim mIdTipoDocumentoReferencia As Integer
            Dim mIdDocumentoReferencia As Integer
            Dim mNumeroSerieReferencia As String
            Dim mNumeroDocumentoReferencia As String
                        
            ' Se busca la primera solicitud de materiales
            mDataBase.ClearParameter
            mDataBase.CommandText = "SELECT pro_solicitudmat.id, pro_solicitudmat.numser, pro_solicitudmat.numdoc " _
                + vbCr + "FROM pro_ordenprod INNER JOIN pro_solicitudmat ON pro_ordenprod.id = pro_solicitudmat.iddocref " _
                + vbCr + "WHERE (((pro_ordenprod.idtipdocref)=" & F.NuloNumeric(F.KeyValue("IdDocumentoOrdenProduccion", xCon)) & ") " _
                            & "AND ((pro_ordenprod.iddocref)=" & ParteProdDet.IdOrdenProduccion & "))"
                                      
            mIdTipoDocumentoReferencia = 0
            Set mRecord = mDataBase.GetRecordset
            If mRecord.RecordCount > 0 Then
                mRecord.MoveFirst
                mIdTipoDocumentoReferencia = F.NuloNumeric(F.KeyValue("IdDocumentoSolictudMateriales", xCon))
                mIdDocumentoReferencia = F.NuloNumeric(mRecord("id"))
                mNumeroSerieReferencia = F.NuloString(mRecord("numser"))
                mNumeroDocumentoReferencia = F.NuloString(mRecord("numdoc"))
                Set mRecord = Nothing
            Else
                Err.Raise &HFFFFFF01, , "El sistema no puede encontrar la Solicitud de Materiales que hace referencia a la orden actual: " _
                                + vbCr + "Orden de Produccion: " & ParteProdDet.OrdenProduccion _
                                + vbCr + "Codigo de Item: " & ParteProdDet.CodigoItem _
                                + vbCr + "Item: " & ParteProdDet.Item
            End If
            Set mRecord = Nothing
            
            ' Se crea el movimiento en almacen
            ' Cabecera
            mMovimiento.IdTipoMovimiento = 0
            mMovimiento.FechaMovimiento = mParteProd.FechaParteProduccion
            mMovimiento.NumeroSerie = mNumeroSerieReferencia
            mMovimiento.NumeroDocumento = F.HallaNumeroDocumento("alm_ingreso", "'" & mMovimiento.NumeroSerie & "'", "numser", xCon)
            mMovimiento.IdEstado = F.NuloNumeric(F.KeyValue("EstadoAprobadoMovimiento", xCon))
            mMovimiento.IdAlmacen = mIdAlmacenItem
            mMovimiento.Glosa = mParteProd.Glosa
            mMovimiento.IdTipoDocumentoReferencia = mIdTipoDocumentoReferencia
            mMovimiento.IdDocumentoReferencia = mIdDocumentoReferencia
            mMovimiento.DocumentoReferencia = F.NuloString(mNumeroSerieReferencia & " - " & mNumeroDocumentoReferencia)
            mMovimiento.MesTrabajo = mParteProd.MesTrabajo
            mMovimiento.AnhoTrabajo = mParteProd.AnhoTrabajo
            mMovimiento.IdTipoDocumentoReferencia2 = F.NuloNumeric(F.KeyValue("IdDocumentoParteProduccion", xCon))
            mMovimiento.IdDocumentoReferencia2 = mParteProd.IdParteProduccion
        
            Dim MovimientoDet As New AlmacenEntidad.EMovimientoDet
            MovimientoDet.IDITEM = ParteProdDet.IDITEM
            MovimientoDet.cantidad = ParteProdDet.CantidadProducida
            MovimientoDet.CantidadTeorica = ParteProdDet.CantidadProgramada
            ' Se agrega al padre
            mMovimiento.LMovimientoDet.Add MovimientoDet
            ' Se graba el movimiento
            Set mMovimiento.Conexion = xCon
            mMovimiento.Called = True
            mMovimiento.IsRecursive = True
            If Not mMovimiento.Save(CLng(xIdUsuario), "") Then Err.Raise &HFFFFFF01, , F.ErrorDescriptionDLL(Err.LastDllError)
            
            Set mMovimiento = Nothing
            Set MovimientoDet = Nothing
        End If
    Next
End Sub

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO DE LA TABLA alm_ingreso
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    Dim Rpta As Integer
    Dim ParteProd As New ProduccionEntidad.EParteProd
    Dim F As New SistemaLogica.Funciones
    
On Error GoTo BloqueError
    TabOne1.CurrTab = 0
    If RstIng.State = 0 Then Exit Sub
    If RstIng.RecordCount = 0 Then
        MsgBox "No hay Registros de Ingreso/Salida de Almacén para eliminar", vbInformation, xTitulo
        Exit Sub
    End If
    
    Rpta = MsgBox("¿Esta seguro de eliminar el Parte de Produccion Nº " + Trim(RstIng("numdoc")) + "?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.BeginTrans
        ParteProd.IdParteProduccion = NulosN(RstIng("id"))
        ' Se eliminan los movimientos
        If F.NuloNumeric(F.KeyValue("CreacionMovimientoAutoProduccion", xCon)) = -1 Then
            ' Verificamos si ya tiene registro en movimientos
            Dim database As New SistemaData.EDataBase
            Dim record As New ADODB.Recordset
            Dim Movimiento As New AlmacenEntidad.EMovimiento
            
            Set database.Connection = xCon
            database.CommandText = "SELECT alm_ingreso.id AS idmov " _
                        + vbCr + "FROM alm_ingreso " _
                        + vbCr + "WHERE (((alm_ingreso.idtipdocref)=" & F.NuloNumeric(F.KeyValue("IdDocumentoParteProduccion", xCon)) & ") AND ((alm_ingreso.iddocref)=" & ParteProd.IdParteProduccion & "))"
            Set record = database.GetRecordset
            If record.RecordCount > 0 Then
                Movimiento.IdMovimiento = F.NuloNumeric(record("idmov"))
                Set Movimiento.Conexion = xCon
                Movimiento.Called = True
                Movimiento.Delete 0, ""
            End If
        End If
        ' Se elimina el parte de Produccion
        Set ParteProd.Conexion = xCon
        ParteProd.Called = True
        If Not ParteProd.Delete(0, "") Then Err.Raise &HFFFFFF01, , F.ErrorDescriptionDLL(Err.LastDllError)
        xCon.CommitTrans
        RstIng.Requery
        Dg1.Refresh
        MsgBox "El Parte Produccion se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
    Set ParteProd = Nothing
    Exit Sub
    
BloqueError:
    xCon.RollbackTrans
    MsgBox "No se pudo eliminar la ParteProduccion por el siguiente motivo :" + Trim(Err.Description)
    Set ParteProd = Nothing
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
    Label5.Caption = "Modificando Parte Produccion"
    QueHace = 2
    Bloquea
    Blanquea
    Fg1.Editable = flexEDKbdMouse
    Fg1.Rows = Fg1.FixedRows
    Fg1.Rows = Fg1.Rows + 1
    Fg1.SelectionMode = flexSelectionFree
    MuestraSegundoTab
    xHorIni = Time
    FechaText.SetFocus
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
    
    FechaText.valor = ""
    NumSerText.Text = ""
    NumDocText.Text = ""
    IdResponsableText.Text = ""
    ResponsableLabel.Caption = ""
    IdAlmOrigenText.Text = ""
    AlmOrigenLabel.Caption = ""
    GlosaText.Text = ""
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
    FechaText.Locked = Not FechaText.Locked
    NumSerText.Locked = Not NumSerText.Locked
    NumDocText.Locked = Not NumDocText.Locked
    IdResponsableText.Locked = Not IdResponsableText.Locked
    GlosaText.Locked = Not GlosaText.Locked
    
    habilitar Cmd, Not Cmd(0).Enabled
End Sub

Private Sub IdResponsableText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub IdResponsableText_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then  'TECLA F5
        Cmd(0).Value = True
    End If
End Sub

Private Sub IdResponsableText_Validate(Cancel As Boolean)
    Dim xRs As New ADODB.Recordset
    
    If NulosC(IdResponsableText.Text) = "" Then Exit Sub
    xRs.CursorLocation = adUseClient
    
    cSQL = "SELECT pla_empleados.id, pla_empleados.nombre AS descripcion " _
        + vbCr + "FROM pla_empleados " _
        + vbCr + "WHERE id = " & NulosN(IdResponsableText.Text) & ""
    
    RST_Busq xRs, cSQL, xCon
    
    If xRs.RecordCount = 0 Then
        IdResponsableText.Text = ""
        ResponsableLabel.Caption = ""
    Else
        ResponsableLabel.Caption = xRs("descripcion")
    End If
    
    Set xRs = Nothing
End Sub

Private Sub IdAlmOrigenText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub IdAlmOrigenText_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        Cmd(3).Value = True
    End If
End Sub

Private Sub IdAlmOrigenText_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(IdAlmOrigenText.Text) = "" Then Exit Sub
    Dim xRs As New ADODB.Recordset
    xRs.CursorLocation = adUseClient
    
    cSQL = "SELECT alm_almacenes.id, alm_almacenes.descripcion " _
        + vbCr + "FROM alm_almacenes " _
        + vbCr + "WHERE id = " & NulosN(IdAlmOrigenText.Text) & ""
    
    RST_Busq xRs, cSQL, xCon
    
    If xRs.RecordCount = 0 Then
        IdAlmOrigenText.Text = ""
        AlmOrigenLabel.Caption = ""
    Else
        AlmOrigenLabel.Caption = NulosC(xRs("descripcion"))
    End If
    
    Set xRs = Nothing
End Sub

'*************************
' NUMEROS DE DOCUMENTO
'*************************
Private Sub NumSerText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub NumSerText_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    
    If NulosC(NumSerText.Text) <> "" Then
        NumSerText.Text = Format(NumSerText.Text, "0000")
        NumDocText.Text = hallarNumDoc("pro_produccion", "'" & NulosC(NumSerText.Text) & "'", "numser")
        If NulosC(NumDocText.Text) = "" Then NumSerText.Text = ""
    End If
End Sub

Private Sub NumDocText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub NumDocText_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(NumDocText.Text) = "" Then Exit Sub
    
    NumDocText.Text = Format(NumDocText.Text, "0000000000")
    
    If existeNumeroDoc("pro_produccion", "'" & NumDocText.Text & "'", "numdoc", "'" & NumSerText.Text & "'", "numser") Then
        MsgBox "El documento ingresado ya existe" & vbCr & "Corrija el numero de documento", vbInformation, xTitulo
        NumDocText.Text = ""
        NumDocText.SetFocus
        Exit Sub
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
    Dim xRs As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim A As Integer
    
    If RstIng.RecordCount = 0 Then Exit Sub
    If RstIng.BOF = True Or RstIng.EOF = True Then Exit Sub
    
    cSQL = "SELECT pro_produccion.* " _
        + vbCr + "FROM pro_produccion " _
        + vbCr + "WHERE pro_produccion.id=" & NulosN(RstIng("id"))
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    
    lblperiodo(1).Caption = lblperiodo(0).Caption
    FechaText.valor = xRs("fchdoc")
    NumSerText.Text = NulosC(xRs("numser"))
    NumDocText.Text = NulosC(xRs("numdoc"))
    IdResponsableText.Text = NulosN(xRs("idresponsable"))
    ResponsableLabel.Caption = UCase(Busca_Codigo(NulosN(xRs("idresponsable")), "id", "nombre", "pla_empleados", "N", xCon))
    
    IdAlmOrigenText.Text = NulosN(xRs("idalm"))
    AlmOrigenLabel.Caption = UCase(Busca_Codigo(NulosN(xRs("idalm")), "id", "descripcion", "alm_almacenes", "N", xCon))
    
    GlosaText.Text = NulosC(xRs("glosa"))
    
    cSQL = "SELECT pro_producciondet.idproddet, pro_producciondet.idpro, pro_producciondet.idrec, pro_producciondet.iditem, pro_producciondet.idunimed, pro_producciondet.canprog, pro_producciondet.cantidad, pro_producciondet.horini, pro_producciondet.horfin, pro_producciondet.obs, pro_producciondet.idord, pro_producciondet.obs, mae_unidades.abrev, pro_receta.codrec, [pro_ordenprod].[numser] & '-' & [pro_ordenprod].[numdoc] AS numordprod, pro_ordenprod.lote, alm_inventario.codpro, alm_inventario.descripcion " _
        + vbCr + "FROM (((pro_producciondet LEFT JOIN mae_unidades ON pro_producciondet.idunimed = mae_unidades.id) LEFT JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN pro_ordenprod ON pro_producciondet.idord = pro_ordenprod.id) LEFT JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id " _
        + vbCr + "WHERE (((pro_producciondet.idpro) = " & NulosN(RstIng("id")) & "));"
            
    Set RstDet = Nothing
    RST_Busq RstDet, cSQL, xCon

    Fg1.Rows = Fg1.FixedRows
    If RstDet.RecordCount <> 0 Then
        RstDet.MoveFirst
        For A = 1 To RstDet.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, Fg1.ColIndex("CODIGO")) = NulosC(RstDet("codpro"))
            Fg1.TextMatrix(A, Fg1.ColIndex("ITEM")) = NulosC(RstDet("descripcion"))
            Fg1.TextMatrix(A, Fg1.ColIndex("RECETA")) = NulosC(RstDet("codrec"))
            Fg1.TextMatrix(A, Fg1.ColIndex("ORDPROD")) = NulosC(RstDet("numordprod"))
            Fg1.TextMatrix(A, Fg1.ColIndex("LOTE")) = NulosC(RstDet("lote"))
            Fg1.TextMatrix(A, Fg1.ColIndex("UM")) = NulosC(RstDet("abrev"))
            Fg1.TextMatrix(A, Fg1.ColIndex("CANPROG")) = Format(NulosN(RstDet("canprog")), FORMAT_CANTIDAD)
            Fg1.TextMatrix(A, Fg1.ColIndex("CANPROD")) = Format(NulosN(RstDet("cantidad")), FORMAT_CANTIDAD)
            Fg1.TextMatrix(A, Fg1.ColIndex("HORINI")) = Format(NulosC(RstDet("horini")), "HH:mm")
            Fg1.TextMatrix(A, Fg1.ColIndex("HORFIN")) = Format(NulosC(RstDet("horfin")), "HH:mm")
            Fg1.TextMatrix(A, Fg1.ColIndex("OBS")) = NulosC(RstDet("obs"))
            Fg1.TextMatrix(A, Fg1.ColIndex("IDPRODDET")) = NulosN(RstDet("idproddet"))
            Fg1.TextMatrix(A, Fg1.ColIndex("IDORDPROD")) = NulosN(RstDet("idord"))
            Fg1.TextMatrix(A, Fg1.ColIndex("IDITEM")) = NulosN(RstDet("iditem"))
            Fg1.TextMatrix(A, Fg1.ColIndex("IDRECETA")) = NulosN(RstDet("idrec"))
            Fg1.TextMatrix(A, Fg1.ColIndex("IDUNIMED")) = NulosN(RstDet("idunimed"))
            
            RstDet.MoveNext
            
            If RstDet.EOF = True Then
                Exit For
            End If
        Next A
    End If
End Sub

Sub pCargarDatos()
    TDB_FiltroLimpiar Dg1
    Set RstIng = Nothing
    
    cSQL = "SELECT pro_produccion.id, pro_produccion.fchdoc, [pro_produccion].[numser] & ' - ' & [pro_produccion].[numdoc] AS numdoc, [pro_ordenprod].[numser] & '-' & [pro_ordenprod].[numdoc] AS numop, pro_produccion.idresponsable, pla_empleados.nombre AS responsable, pro_produccion.glosa " _
        + vbCr + "FROM (pro_produccion LEFT JOIN pla_empleados ON pro_produccion.idresponsable = pla_empleados.id) LEFT JOIN (pro_producciondet LEFT JOIN pro_ordenprod ON pro_producciondet.idord = pro_ordenprod.id) ON pro_produccion.id = pro_producciondet.idpro " _
        + vbCr + "WHERE (((pro_produccion.ano) = " & AnoTra & ") And ((pro_produccion.idmes) = " & mMesActivo & ")) " _
        + vbCr + "ORDER BY pro_produccion.fchdoc DESC;"

    RST_Busq RstIng, cSQL, xCon
    Set Dg1.DataSource = RstIng
    
    '********************************************************************************************
    lblperiodo(0).Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    '********************************************************************************************
    
    '------------------------------------------------------------------------------------------
    ' bloqueamos los botones del toolbar
    CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
    '------------------------------------------------------------------------------------------
End Sub


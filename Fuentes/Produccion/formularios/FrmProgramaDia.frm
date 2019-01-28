VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmProgramaDia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Producción - Programación Diaria"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11940
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7245
      Left            =   15
      TabIndex        =   4
      Top             =   360
      Width           =   11910
      _cx             =   21008
      _cy             =   12779
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
         Height          =   6825
         Left            =   -12465
         TabIndex        =   5
         Top             =   375
         Width           =   11820
         Begin TrueOleDBGrid70.TDBGrid Dg3 
            Height          =   6465
            Left            =   45
            TabIndex        =   18
            Top             =   345
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11404
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Fch Prod."
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Fch Término"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Producto"
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Cant. Producc"
            Columns(3).DataField=   ""
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Origen"
            Columns(4).DataField=   ""
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Cant. Inicial"
            Columns(5).DataField=   ""
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Receta"
            Columns(6).DataField=   ""
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Responsable"
            Columns(7).DataField=   ""
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Nº Lote"
            Columns(8).DataField=   ""
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1349"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1270"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=1693"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1614"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=4524"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=4445"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=1931"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1852"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=1588"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1508"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=2064"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1984"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(43)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=516"
            Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(49)=   "Column(8).Width=1296"
            Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=1217"
            Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=516"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
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
         Begin VB.Label lblperiodo 
            AutoSize        =   -1  'True
            Caption         =   "lblperiodo"
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
            Left            =   9705
            TabIndex        =   17
            Top             =   30
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Programación Diaria"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
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
            TabIndex        =   7
            Top             =   30
            Width           =   11610
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
            TabIndex        =   6
            Top             =   30
            Width           =   1275
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Detalle de la Cuenta"
         Height          =   6825
         Left            =   45
         TabIndex        =   8
         Top             =   375
         Width           =   11820
         Begin VB.CommandButton Cmd 
            Caption         =   "Agregar"
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   23
            ToolTipText     =   "Agregar Documentos"
            Top             =   3375
            Width           =   1275
         End
         Begin VB.CommandButton Cmd 
            Caption         =   "&Eliminar"
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   1365
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "Eliminar Documentos Seleccionados"
            Top             =   3375
            Width           =   1275
         End
         Begin VB.TextBox txt 
            BackColor       =   &H0080FF80&
            Height          =   315
            Index           =   0
            Left            =   10575
            TabIndex        =   15
            Tag             =   "null"
            Text            =   "txt(0)"
            Top             =   0
            Visible         =   0   'False
            Width           =   1170
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
            Height          =   300
            Index           =   0
            Left            =   1440
            TabIndex        =   0
            Top             =   360
            Width           =   1230
            _ExtentX        =   2170
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
            Valor           =   "21/11/2007"
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   2610
            Left            =   90
            TabIndex        =   2
            Top             =   690
            Width           =   11700
            _cx             =   20637
            _cy             =   4604
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
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   11
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmProgramaDia.frx":0000
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
         Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
            Height          =   300
            Index           =   1
            Left            =   3930
            TabIndex        =   19
            Top             =   360
            Width           =   1230
            _ExtentX        =   2170
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
            Valor           =   "21/11/2007"
         End
         Begin VB.CommandButton cb 
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   7815
            Picture         =   "FrmProgramaDia.frx":0146
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   390
            Width           =   225
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   1
            Left            =   6855
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   1
            Text            =   "txt_cb(1)"
            ToolTipText     =   "Ingrese DNI del Supervisor"
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton CmdMover 
            Height          =   240
            Index           =   0
            Left            =   5070
            Picture         =   "FrmProgramaDia.frx":0278
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Ampliar"
            Top             =   6585
            Width           =   6720
         End
         Begin VB.CommandButton CmdMover 
            Height          =   240
            Index           =   1
            Left            =   5070
            Picture         =   "FrmProgramaDia.frx":03B6
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Reducir"
            Top             =   6585
            Visible         =   0   'False
            Width           =   6720
         End
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   2940
            Left            =   30
            TabIndex        =   24
            Top             =   3885
            Width           =   11835
            _cx             =   20876
            _cy             =   5186
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9
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
            FrontTabForeColor=   -2147483630
            Caption         =   " Lista de Materiales |  Hoja de Ruta  "
            Align           =   0
            CurrTab         =   1
            FirstTab        =   0
            Style           =   0
            Position        =   1
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
            Begin VB.Frame Frame3 
               BorderStyle     =   0  'None
               Caption         =   "Frame6"
               Height          =   2520
               Left            =   45
               TabIndex        =   27
               Top             =   45
               Width           =   11745
               Begin VSFlex7Ctl.VSFlexGrid Fg 
                  Height          =   2535
                  Index           =   1
                  Left            =   30
                  TabIndex        =   28
                  Top             =   30
                  Width           =   11640
                  _cx             =   20532
                  _cy             =   4471
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
                  AllowUserResizing=   1
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   1
                  Cols            =   10
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmProgramaDia.frx":04F4
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
            Begin VB.Frame Frame4 
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   2520
               Left            =   -12390
               TabIndex        =   25
               Top             =   45
               Width           =   11745
               Begin VSFlex7Ctl.VSFlexGrid Fg 
                  Height          =   2505
                  Index           =   0
                  Left            =   15
                  TabIndex        =   26
                  Top             =   30
                  Width           =   11685
                  _cx             =   20611
                  _cy             =   4419
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
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   1
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmProgramaDia.frx":0628
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
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Resumen de Requerimientos"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   75
            TabIndex        =   21
            Top             =   3690
            Width           =   2625
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Fch Finalización"
            Height          =   195
            Index           =   1
            Left            =   2715
            TabIndex        =   20
            Top             =   450
            Width           =   1140
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Index           =   0
            Left            =   10020
            TabIndex        =   16
            Top             =   120
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lbl_cb_cod 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb_cod(1)"
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
            Index           =   1
            Left            =   10200
            TabIndex        =   12
            Top             =   360
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Fch Producción"
            Height          =   195
            Index           =   8
            Left            =   105
            TabIndex        =   10
            Top             =   450
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle Producción"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   75
            TabIndex        =   9
            Top             =   15
            Width           =   11610
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "Programado Por"
            Height          =   195
            Index           =   1
            Left            =   5655
            TabIndex        =   14
            Top             =   450
            Width           =   1140
         End
         Begin VB.Label lbl_cb 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb(1)"
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
            Index           =   1
            Left            =   8070
            TabIndex        =   13
            Top             =   360
            Width           =   3675
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11940
      _ExtentX        =   21061
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
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar Mes"
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
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Listado"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Registro"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6600
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
               Picture         =   "FrmProgramaDia.frx":06E2
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramaDia.frx":0C26
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramaDia.frx":0FB8
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramaDia.frx":113C
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramaDia.frx":1590
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramaDia.frx":16A8
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramaDia.frx":1BEC
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramaDia.frx":2130
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramaDia.frx":2244
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramaDia.frx":2358
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramaDia.frx":27AC
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramaDia.frx":2918
               Key             =   "IMG11"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmProgramaDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QueHace As Integer
Dim Agregando As Boolean
Dim SeEjecuto As Boolean
Dim RstFrm As New ADODB.Recordset
Dim RstTmp As New ADODB.Recordset
'----
Dim mMesActivo  As Integer '--INDICA EL MES ACTIVO
Dim RST_INSUMO As New ADODB.Recordset '--PARA LOS INSUMOS
Dim RST_TAREA As New ADODB.Recordset '--PARA LAS TAREAS
'----
Dim M_NUM_PARTE As Long      '--INDICA EL NUMERO PARTE DE PRODUCCION

Private Const FORMAT_NUM_PRODUCCION As String = "000000" '--INDICA EL FORMATO DE LA COLUMNA CON EL NUMERO DE PRODUCCION
Private Const FORMAT_NUM_PARTE As String = "00000000" '--INDICA EL FORMATO DE LA COLUMNA CON EL NUMERO DE PRODUCCION
Private Const FORMAT_CANT As String = "#0.000000"
Dim F_GRUPO  As Boolean '--CAMBIA DE COLOR EL GRUPO
Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta


Private Sub cb_Click(Index As Integer)
    Dim xRs As New ADODB.Recordset
    Dim xCampos() As String
    Dim nTitulo As String
    Dim nSQL As String
    On Error GoTo error
    
    nSQL = "SELECT pla_empleados.numdoc, [pla_empleados].[ape] & ' ' & [pla_empleados].[nom] AS nombre, pro_emp.id " _
        + vbCr + " FROM pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp " _
        + vbCr + " Where (((pro_emp.prog) = -1)) " _
        + vbCr + " ORDER BY pla_empleados.ape;"
        
    nTitulo = "Buscando Programadores"
    
    If Index = 1 Then
        nTitulo = "Buscando Supervisores"
        nSQL = Replace(nSQL, "(pro_emp.prog)", "(pro_emp.sup)")
    End If
    
    ReDim xCampos(2, 3) As String
    xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "DNI":      xCampos(1, 1) = "numdoc":    xCampos(1, 2) = "1500":   xCampos(1, 3) = "C"
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
    
    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_cb_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO
SALIR:
    Set xRs = Nothing
Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub


Private Sub pRegistroAdd(Optional F_SELECCION_VARIOS As Boolean = True, _
                         Optional fAddRegistroSinPrograma As Boolean = False)
                         
    If IsDate(TxtFecha(0).Valor) = False Then
        MsgBox "Ingrese la fecha de Producción", vbExclamation, xTitulo
        Exit Sub
    End If
    Dim SQL_IDREC As String
    '--GENERAR EL WHERE DE LOS ID'S RECETA PARA QUE NO SE REPITAN
    If fAddRegistroSinPrograma = False Then SQL_IDREC = GENERAR_SQL_ID(Fg1, 12, "pro_receta.id", "NOT IN")
    If SQL_IDREC <> "" Then SQL_IDREC = " AND " + SQL_IDREC
    '----
    On Error GoTo error
    Dim xRs  As New ADODB.Recordset
    
    Dim nSQL As String
    
    
    If fAddRegistroSinPrograma = False Then
        ReDim xCampos(4, 5) As String
        xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "descripcion":  xCampos(0, 2) = "5000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "N"
        xCampos(1, 0) = "Receta":           xCampos(1, 1) = "codrec":       xCampos(1, 2) = "1000":     xCampos(1, 3) = "C":    xCampos(1, 4) = "S"
        xCampos(2, 0) = "U.M.":             xCampos(2, 1) = "abrev":        xCampos(2, 2) = "600":      xCampos(2, 3) = "C":    xCampos(2, 4) = "N"
        xCampos(3, 0) = "Cant.Prog":        xCampos(3, 1) = "canprog":      xCampos(3, 2) = "1000":      xCampos(3, 3) = "N":    xCampos(3, 4) = "N"

        nSQL = "SELECT pro_programadet.idprod, pro_receta.id AS idrec, pro_receta.iditem,pro_receta.idunimed, alm_inventario.descripcion, pro_receta.codrec, mae_unidades.abrev, pro_programadet.canpro as canprog " _
            + vbCr + " FROM (alm_inventario INNER JOIN (pro_receta LEFT JOIN mae_unidades ON pro_receta.idunimed = mae_unidades.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN pro_programadet ON pro_receta.id = pro_programadet.idrec " _
            + vbCr + " WHERE (((pro_programadet.idpro) = 0) And ((pro_programadet.dia) = CDATE('" + TxtFecha(0).Valor + "'))) " + SQL_IDREC _
            + vbCr + " ORDER BY alm_inventario.descripcion;"
    Else
    
        ReDim xCampos(3, 5) As String
        xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "descripcion":  xCampos(0, 2) = "5000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
        xCampos(1, 0) = "Familia":          xCampos(1, 1) = "famdesc":      xCampos(1, 2) = "1500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
        xCampos(2, 0) = "U.M.":             xCampos(2, 1) = "abrev":        xCampos(2, 2) = "600":      xCampos(2, 3) = "C":    xCampos(2, 4) = "N"

        nSQL = "SELECT alm_inventario.id as iditem, alm_inventario.descripcion , mae_familia.descripcion AS famdesc, mae_unidades.abrev,0 AS canprog " _
            + vbCr + " FROM mae_unidades RIGHT JOIN (alm_inventario LEFT JOIN mae_familia ON alm_inventario.idfam = mae_familia.id) ON mae_unidades.id = alm_inventario.idunimed " _
            + vbCr + " WHERE alm_inventario.tippro = 3  AND alm_inventario.activo = -1 " _
            + vbCr + " ORDER BY alm_inventario.descripcion, mae_familia.descripcion;"

    End If
    
    If F_SELECCION_VARIOS = True Then
        CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), "Buscando Programas de Producción para el dia " + TxtFecha(0).Valor
    Else
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Programas de Producción para el dia " + TxtFecha(0).Valor, "descripcion", "descripcion", Principio
    End If
    
    Agregando = True
    Dim a As Integer
    Dim xFila As Integer
    Dim RstReceta As New ADODB.Recordset   '--BUSCAR LA RECETA PREDETERMINADA
    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    If F_SELECCION_VARIOS = True Then xRs.MoveFirst
   
    '--SI NO HAY REGISTROS OBTENER EL NUMERO DE PRODUCCION
    If Fg1.Rows = 1 Then M_NUM_PARTE = HallaValor(xCon, "pro_producciondet", "numparte")
   
    Do While Not xRs.EOF
        ADD_REG Fg1, Fila_Ninguno
        With Fg1
            '--DEL NUMERO DE REGISTRO
            .TextMatrix(.Rows - 1, 1) = M_NUM_PARTE
            .TextMatrix(.Rows - 1, 2) = xRs.Fields("descripcion") & ""
            .TextMatrix(.Rows - 1, 5) = Format(NulosN(xRs.Fields("canprog")), FORMAT_CANTIDAD)
            .TextMatrix(.Rows - 1, 11) = M_NUM_PARTE
            .TextMatrix(.Rows - 1, 13) = xRs.Fields("iditem") & ""
            If fAddRegistroSinPrograma = False Then
                '--YA FUE PROGRAMADO
                .TextMatrix(.Rows - 1, 3) = xRs.Fields("codrec") & ""
                .TextMatrix(.Rows - 1, 4) = xRs.Fields("abrev") & ""
                .TextMatrix(.Rows - 1, 12) = xRs.Fields("idrec") & ""
                .TextMatrix(.Rows - 1, 14) = xRs.Fields("idunimed") & ""
                '---
                DATOS_TMP_ADD CStr(M_NUM_PARTE), xRs.Fields("idrec"), E_INSUMO, True, fAddRegistroSinPrograma
                DATOS_TMP_ADD CStr(M_NUM_PARTE), xRs.Fields("idrec"), e_TAREA, True, fAddRegistroSinPrograma
            Else
                '-----CARGAR RECETA PREDETERMINADA
                RST_Busq RstReceta, "SELECT TOP 1 pro_receta.id AS idrec, pro_receta.descripcion, pro_receta.codrec, mae_unidades.abrev , pro_receta.idunimed" _
                                    + vbCr + " FROM pro_receta INNER JOIN mae_unidades ON pro_receta.idunimed = mae_unidades.id " _
                                    + vbCr + " Where (((pro_receta.iditem) = " + CStr(xRs.Fields("iditem")) + "))" _
                                    + vbCr + " ORDER BY pro_receta.prirec;", xCon
                If RstReceta.EOF = False Or RstReceta.BOF = False Or RstReceta.RecordCount <> 0 Then
                    If VERIFICAR_LISTA(Fg1, 3, RstReceta.Fields("codrec") & "", False) = True Then
                        .TextMatrix(.Rows - 1, 3) = RstReceta.Fields("codrec") & ""
                        .TextMatrix(.Rows - 1, 4) = RstReceta.Fields("abrev") & ""
                        .TextMatrix(.Rows - 1, 12) = RstReceta.Fields("idrec") & ""
                        .TextMatrix(.Rows - 1, 14) = RstReceta.Fields("idunimed") & ""
                        
                        DATOS_TMP_ADD CStr(M_NUM_PARTE), RstReceta.Fields("idrec"), E_INSUMO, True, fAddRegistroSinPrograma
                        DATOS_TMP_ADD CStr(M_NUM_PARTE), RstReceta.Fields("idrec"), e_TAREA, True, fAddRegistroSinPrograma
                    End If
                End If
                
            End If
            '-----
            
            '---

            'M_NUM_PARTE = M_NUM_PARTE + 1
            If F_SELECCION_VARIOS = False Then Exit Do
            
            '---
        End With
        If F_SELECCION_VARIOS = False Then Exit Do
        xRs.MoveNext
    Loop
SALIR:
    Agregando = False
    If Fg1.Rows >= 2 Then Fg1.Row = Fg1.Rows - 1: Fg1.Col = 6:   'Fg1_RowColChange
    Set xRs = Nothing
    Fg1.SetFocus
    Exit Sub
error:
    Agregando = False
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "pRegistroAdd"
End Sub

Private Function pGenerarConsulta(mIdParte As String, mIdReceta As String, fTipo As e_PROGRAMA, dFecha As Date, Optional fAddRegistro As Boolean = False, Optional fAddRegistroSinPrograma As Boolean = False) As String
    '--ESTA FUNCION CONSTRUYE LAS CONSULTAS DE INSUMOS Y TAREAS EN FUNCION DE LA RECETA(AGREGAR NUEVA RECETA, AGREGAR DE PROGRAMACION DE PRODUCCION)
    Dim nSQL As String
    If fTipo = E_INSUMO Then
        If fAddRegistro = True Then '--NUEVO
            If fAddRegistroSinPrograma = False Then
                nSQL = "SELECT " + mIdParte + " AS idparte, pro_programadet.idrec, pro_recetains.iditem, pro_recetains.idunimed, mae_tipoproducto.descripcion AS tipprodesc, alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro AS unid, pro_recetains!canpro*pro_programadet!canpro AS canprog, '' AS canteo, '' AS canreal, '' AS dif " _
                    + vbCr + " FROM (pro_receta INNER JOIN pro_programadet ON pro_receta.id = pro_programadet.idrec) INNER JOIN (mae_unidades RIGHT JOIN ((mae_tipoproducto RIGHT JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) RIGHT JOIN pro_recetains ON alm_inventario.id = pro_recetains.iditem) ON mae_unidades.id = pro_recetains.idunimed) ON pro_receta.id = pro_recetains.idrec " _
                    + vbCr + " WHERE (((pro_programadet.idrec) = " + mIdReceta + ") And ((pro_programadet.dia) = CDATE('" + CStr(dFecha) + "') )) " _
                    + vbCr + " ORDER BY mae_tipoproducto.descripcion ASC, alm_inventario.descripcion;"
            Else '--DIRECTAMENTE DE RECETAS(SIN PROGRAMA)
                nSQL = "SELECT " + mIdParte + " AS idparte,pro_receta.id AS idrec,pro_recetains.iditem, pro_recetains.idunimed, mae_tipoproducto.descripcion AS tipprodesc, alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro AS unid, 0 AS canprog, '' AS canteo, '' AS canreal, '' AS dif " _
                    + vbCr + " FROM pro_receta INNER JOIN (mae_unidades RIGHT JOIN ((mae_tipoproducto RIGHT JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) RIGHT JOIN pro_recetains ON alm_inventario.id = pro_recetains.iditem) ON mae_unidades.id = pro_recetains.idunimed) ON pro_receta.id = pro_recetains.idrec " _
                    + vbCr + " WHERE (((pro_receta.ID) = " + mIdReceta + ")) " _
                    + vbCr + " ORDER BY mae_tipoproducto.descripcion, alm_inventario.descripcion;"
            End If
        Else '--CONSULTA O MODIFICAR
            nSQL = "SELECT pro_producciondet.numparte AS idparte, pro_producciondetins.idrec, pro_producciondetins.iditem, pro_producciondetins.idunimed, mae_tipoproducto.descripcion AS tipprodesc, alm_inventario.descripcion, mae_unidades.abrev, pro_producciondetins.canpro AS unid, IIf(pro_programadet.canpro Is Null,0,pro_programadet.canpro*pro_producciondetins.canpro) AS canprog, IIf(pro_producciondet.cantidad Is Null,0,(pro_producciondet.cantidad*pro_producciondetins.canpro)) AS canteo, pro_producciondetins.canutil AS canreal, [canteo]-pro_producciondetins!canutil AS dif " _
                + vbCr + " FROM (pro_producciondet LEFT JOIN pro_programadet ON (pro_producciondet.idpro = pro_programadet.idpro) AND (pro_producciondet.idrec = pro_programadet.idrec)) INNER JOIN (mae_tipoproducto RIGHT JOIN (((pro_producciondetins LEFT JOIN mae_unidades ON pro_producciondetins.idunimed = mae_unidades.id) LEFT JOIN pro_recetains ON (pro_producciondetins.idrec = pro_recetains.idrec) AND (pro_producciondetins.iditem = pro_recetains.iditem)) LEFT JOIN alm_inventario ON pro_producciondetins.iditem = alm_inventario.id) ON mae_tipoproducto.id = alm_inventario.tippro) ON (pro_producciondet.idrec = pro_producciondetins.idrec) AND (pro_producciondet.numparte = pro_producciondetins.numparte) AND (pro_producciondet.idpro = pro_producciondetins.idpro) " _
                + vbCr + " GROUP BY pro_producciondet.numparte, pro_producciondetins.idrec, pro_producciondetins.iditem, pro_producciondetins.idunimed, mae_tipoproducto.descripcion, alm_inventario.descripcion, mae_unidades.abrev, pro_producciondetins.canpro, IIf(pro_programadet.canpro Is Null,0,pro_programadet.canpro*pro_producciondetins.canpro), IIf(pro_producciondet.cantidad Is Null,0,(pro_producciondet.cantidad*pro_producciondetins.canpro)), pro_producciondetins.canutil " _
                + vbCr + " HAVING (((pro_producciondet.numparte) = '" + mIdParte + "') And ((pro_producciondetins.idrec) = " + mIdReceta + ")) " _
                + vbCr + " ORDER BY mae_tipoproducto.descripcion, alm_inventario.descripcion;"

            
        End If
    ElseIf fTipo = e_TAREA Then
        If fAddRegistro = True Then
            If fAddRegistroSinPrograma = False Then
                nSQL = "SELECT " + mIdParte + " AS idparte, pro_programadet.idrec,pro_recetatar.orden as corr, pro_recetatar.idtar, pro_recetatar.idunimed, pro_tareas.descripcion, mae_unidades.abrev, pro_recetatar.cantidad AS unid, pro_programadet!canpro*pro_recetatar!cantidad AS canprog,  '' AS horini, '' AS horfin, 0 as canper " _
                    + vbCr + " FROM  pro_tareas INNER JOIN ((pro_receta INNER JOIN pro_programadet ON pro_receta.id = pro_programadet.idrec) INNER JOIN (mae_unidades INNER JOIN pro_recetatar ON mae_unidades.id = pro_recetatar.idunimed) ON pro_receta.id = pro_recetatar.idrec) ON pro_tareas.id = pro_recetatar.idtar " _
                    + vbCr + " WHERE (((pro_programadet.idrec) = " + mIdReceta + ") And ((pro_programadet.dia) = CDATE('" + CStr(dFecha) + "') )) " _
                    + vbCr + " ORDER BY pro_recetatar.orden;"
            Else '--DIRECTAMENTE DE RECETAS
                nSQL = "SELECT " + mIdParte + " AS idparte,pro_receta.id AS idrec, pro_recetatar.idtar, pro_recetatar.orden as corr, pro_recetatar.idunimed, pro_tareas.descripcion, mae_unidades.abrev, pro_recetatar.cantidad AS unid, 0 AS canprog,'' AS horini, '' AS horfin, 0 AS canper " _
                    + vbCr + " FROM pro_tareas INNER JOIN (pro_receta INNER JOIN (mae_unidades INNER JOIN pro_recetatar ON mae_unidades.id = pro_recetatar.idunimed) ON pro_receta.id = pro_recetatar.idrec) ON pro_tareas.id = pro_recetatar.idtar " _
                    + vbCr + " WHERE (((pro_receta.ID) = " + mIdReceta + ")) " _
                    + vbCr + " ORDER BY pro_recetatar.orden;"
            End If
        Else
            
           nSQL = "SELECT pro_producciondet.numparte as idparte ,pro_producciondettar.idrec, pro_recetatar.idtar,pro_producciondettar.corr, pro_recetatar.idunimed, pro_tareas.descripcion, mae_unidades.abrev, pro_recetatar.cantidad AS unid, IIF(pro_programadet.canpro IS NULL,0,pro_programadet.canpro*pro_recetatar.cantidad) AS canprog, Format([pro_producciondettar].[horini],'hh:nn AM/PM') AS horini, Format([pro_producciondettar].[horfin],'hh:nn AM/PM') AS horfin, pro_producciondettar.canper " _
                + vbCr + " FROM (pro_producciondet LEFT JOIN pro_programadet ON (pro_producciondet.idpro = pro_programadet.idpro) AND (pro_producciondet.idrec = pro_programadet.idrec)) INNER JOIN (pro_tareas RIGHT JOIN ((pro_producciondettar INNER JOIN pro_recetatar ON (pro_producciondettar.idtar = pro_recetatar.idtar) AND (pro_producciondettar.idrec = pro_recetatar.idrec)) LEFT JOIN mae_unidades ON pro_producciondettar.idunimed = mae_unidades.id) ON pro_tareas.id = pro_producciondettar.idtar) ON (pro_producciondet.idpro = pro_producciondettar.idpro) AND (pro_producciondet.numparte = pro_producciondettar.numparte) AND (pro_producciondet.idrec = pro_producciondettar.idrec) " _
                + vbCr + " GROUP BY pro_producciondet.numparte, pro_producciondettar.idrec, pro_recetatar.idtar,pro_producciondettar.corr, pro_recetatar.idunimed, pro_tareas.descripcion, mae_unidades.abrev, pro_recetatar.cantidad, IIf(pro_programadet.canpro Is Null,0,pro_programadet.canpro*pro_recetatar.cantidad), pro_producciondettar.canper, pro_producciondettar.horini, pro_producciondettar.horfin, pro_recetatar.orden, pro_programadet.dia, pro_recetatar.orden " _
                + vbCr + " HAVING (((pro_producciondet.numparte) = '" + mIdParte + "') And ((pro_producciondettar.idrec) = " + mIdReceta + ")) " _
                + vbCr + " ORDER BY pro_tareas.descripcion, pro_recetatar.orden, pro_recetatar.orden;"
        End If
    ElseIf fTipo = e_EQUIPO Then
        '--FALTA DEFINIR
        '-------------
        '----------------
    End If
    pGenerarConsulta = nSQL
End Function

Private Sub DATOS_TMP_ADD(mIdParte As String, mIdReceta As String, fTipo As e_PROGRAMA, _
                            Optional fAddRegistro As Boolean = False, _
                            Optional fAddRegistroSinPrograma As Boolean = False)
                            
    '--ESTA FUNCION CARGA LOS DATOS RELACIONADO A LOS INSUMOS, TAREAS, PARA LUEGO SER MOSTRADO EN EL GRID DE INSUMOS Y TAREAS
    On Error GoTo error
    Dim RST_ORIGEN As New ADODB.Recordset
    Dim nSQL As String
    Me.MousePointer = vbHourglass
    nSQL = pGenerarConsulta(mIdParte, mIdReceta, fTipo, TxtFecha(0).Valor, fAddRegistro, fAddRegistroSinPrograma)
    If nSQL = "" Then GoTo SALIR
    RST_Busq RST_ORIGEN, nSQL, xCon
    If RST_ORIGEN.State = 0 Then GoTo SALIR
    If RST_ORIGEN.EOF = True Or RST_ORIGEN.BOF = True Or RST_ORIGEN.RecordCount = 0 Then GoTo SALIR:
    If fTipo = E_INSUMO Then
        CARGAR_RST_TMP RST_INSUMO, RST_ORIGEN
    ElseIf fTipo = e_TAREA Then
        CARGAR_RST_TMP RST_TAREA, RST_ORIGEN
    ElseIf fTipo = e_EQUIPO Then
    
    End If
SALIR:
    Me.MousePointer = vbDefault
    Set RST_ORIGEN = Nothing
    Exit Sub
error:
    Me.MousePointer = vbDefault
    Set RST_ORIGEN = Nothing
    SHOW_ERROR Me.Name, "DATOS_TMP_ADD"
End Sub

Private Sub DATOS_TMP_DEL(mIdParte As String, mIdReceta As String, RST_TMP As ADODB.Recordset)
    If mIdReceta = "" Then Exit Sub
    '--ELIMINAR DATOS DEL TEMPORAL
    RST_TMP.Filter = "idparte= " + mIdParte + " AND idrec=" + mIdReceta
    If RST_TMP.RecordCount = 0 Then Exit Sub
    RST_TMP.MoveFirst
    Do While Not RST_TMP.EOF
        RST_TMP.Delete
        RST_TMP.MoveNext
    Loop
    
End Sub


Private Sub pRegistroDel()
    If Fg1.Row < 0 Then Exit Sub
    If Fg1.Row = 0 Then
        MsgBox "La fila no se puede eliminar" + vbCr + "Seleccione una correcta", vbExclamation
        Exit Sub
    End If
    If MsgBox("Seguro desea eliminar el Producto", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
    
    If Fg1.TextMatrix(Fg1.Row, 12) <> "" Then
        '--ELIMINANDO LOS REGISTROS DE RECORSET TEMPORAL
        DATOS_TMP_DEL Fg1.TextMatrix(Fg1.Row, 11), Fg1.TextMatrix(Fg1.Row, 12), RST_INSUMO
        DATOS_TMP_DEL Fg1.TextMatrix(Fg1.Row, 11), Fg1.TextMatrix(Fg1.Row, 12), RST_TAREA
    End If
    Label2.Caption = ""
    '--ELIMINAR EL PRODUCTO
    Fg1.RemoveItem (Fg1.Row)
    If Fg1.Rows > 1 Then Fg1.Row = 1
End Sub

Private Sub Cmd_Click(Index As Integer)
    Select Case Index
        Case 0 '--AGREGAR PROGRAMA
            pRegistroAdd False, True
        Case 1 '--ELIMINAR REGISTROS AGREGADOS
            pRegistroDel
        Case 2
            pRegistroAdd True, True
        Case 3
            pRegistroAdd True, False
    End Select
End Sub

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace <> 3 Then
            If Index = 0 Then PopupMenu Menu3
            If Index = 1 Then PopupMenu menu2
        End If
    End If

End Sub

Private Sub Dg3_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg3_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
  
  'If F_GRUPO = False Then Exit Sub
'''''
'''''
'''''    RstFrm.Bookmark = Bookmark
''''''    If Bookmark <> 1 Then
''''''        RstTmp.MovePrevious
''''''    End If
'''''
'''''    'If RstFrm.Fields("id") <> RstTmp.Fields("id") Then
'''''    If InStr(Val(RstFrm.Fields("id")) / 2, ".") = 0 Then
'''''        RowStyle = "ItemSelected"
'''''        RowStyle = "ItemSelected"
'''''    End If
'''''
''''''    Set RST_TEPM_1 = Nothing

    
    
End Sub

Private Sub Dg3_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstFrm.Sort = CStr(Dg3.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    If Agregando = True Then Exit Sub
    If Row = 0 Then Exit Sub
    If NulosN(Fg1.TextMatrix(Row, 12)) = 0 Then
        MsgBox "Ingrese La Receta del Producto:" + vbCr + _
        "Producto:        " + Fg1.TextMatrix(Row, 2) & "", vbExclamation, xTitulo
        Fg1.TextMatrix(Row, Col) = ""
        Fg1.Col = 3: Exit Sub
    End If
    Dim RST_TMP As New ADODB.Recordset
    
    Select Case Col
        Case 1
            If Fg1.TextMatrix(Row, 12) = "" Then Exit Sub
            RST_Busq RST_TMP, "select numparte from pro_producciondet where numparte='" + Format(NulosN(Fg1.TextMatrix(Row, 1)), FORMAT_NUM_PARTE) + "' ;", xCon
            If RST_TMP.EOF = False And RST_TMP.BOF = False And RST_TMP.RecordCount <> 0 Then
                Set RST_TMP = Nothing
                If MsgBox("El número de producción ya existe, desea continuar?", vbQuestion + vbYesNo, xTitulo) = vbNo Then
                    Fg1.TextMatrix(Row, Col) = ""
                    Fg1.Col = 1:    Exit Sub
                End If
            End If
            Set RST_TMP = Nothing
        Case 6
            If IsNumeric(Fg1.TextMatrix(Row, Col)) = False Then
                MsgBox "El valor ingresado no es numérico", vbCritical, xTitulo
                Fg1.TextMatrix(Row, Col) = ""
            Else
                Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), FORMAT_CANTIDAD)
                '---ACTUALIZAR LA CANTIDAD TEORICA
                If RST_INSUMO.EOF = False Or RST_INSUMO.BOF = False Or RST_INSUMO.RecordCount <> 0 Then RST_INSUMO.MoveFirst
                Do While Not RST_INSUMO.EOF
                    RST_INSUMO.Fields("canteo") = NulosN(RST_INSUMO.Fields("unid")) * NulosN(Fg1.TextMatrix(Row, Col))
                    RST_INSUMO.MoveNext
                Loop
                '------------------------
                '------------------------
            End If
        Case 7, 8
            If IsDate(Fg1.TextMatrix(Row, Col)) = False Then
                MsgBox "El valor ingresado no es una Hora correcta", vbCritical, xTitulo
                Fg1.TextMatrix(Row, Col) = ""
            Else
                If IsDate(Fg1.TextMatrix(Row, 7)) = True And IsDate(Fg1.TextMatrix(Row, 8)) = True Then '--HORA INICIO
                    If CDate(Fg1.TextMatrix(Row, 7)) >= CDate(Fg1.TextMatrix(Row, 8)) Then
                        MsgBox "La hora " + IIf(Col = 7, "Inicial debe ser menor ", "Final debe ser mayor") + " a la hora " + IIf(Col = 7, "Final", "Inicial"), vbExclamation, xTitulo
                        Fg1.TextMatrix(Row, Col) = "":  Exit Sub
                   End If
                End If
                Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), FORMAT_HORA_SIN_SEGUNDO)
            End If
    End Select
    Exit Sub
error:
    Set RST_TMP = Nothing
    SHOW_ERROR Me.Name, "Fg1_CellChanged"
End Sub


Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col <> 3 And Col <> 9 And Col <> 10 Then Exit Sub
    If QueHace = 3 Then Exit Sub
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim nTitulo As String
    
    If NulosN(Fg1.TextMatrix(Row, 12)) = 0 And Col <> 3 Then
        MsgBox "Ingrese La Receta del Producto:" + vbCr + _
        "Producto:        " + Fg1.TextMatrix(Row, 2) & "", vbExclamation, xTitulo
        Fg1.TextMatrix(Row, Col) = ""
        Fg1.Col = 3: Exit Sub
    End If
    Select Case Col
        Case 3 '--DE LAS RECETAS
            ReDim xCampos(3, 4) As String
            xCampos(0, 0) = "Descripción":     xCampos(0, 1) = "nombre":        xCampos(0, 2) = "5000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
            xCampos(1, 0) = "Código":          xCampos(1, 1) = "codrec":        xCampos(1, 2) = "1500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
            xCampos(2, 0) = "U.M.":            xCampos(2, 1) = "abrev":         xCampos(2, 2) = "500":      xCampos(2, 3) = "C":    xCampos(2, 4) = "N"

            nSQL = "SELECT pro_receta.id AS idrec, pro_receta.descripcion as nombre, pro_receta.codrec, mae_unidades.abrev, pro_receta.idunimed " _
                    + vbCr + " FROM alm_inventario INNER JOIN (pro_receta INNER JOIN mae_unidades ON pro_receta.idunimed = mae_unidades.id) ON alm_inventario.id = pro_receta.iditem " _
                    + vbCr + " WHERE (((alm_inventario.id) = " + CStr(Fg1.TextMatrix(Row, 13)) + ")) " _
                    + vbCr + " ORDER BY pro_receta.descripcion;"
    
            nTitulo = "Buscando Recetas"
        
        Case 9 '--DEL RESPONSABLE
            ReDim xCampos(2, 4) As String
            xCampos(0, 0) = "Nombre":       xCampos(0, 1) = "nombre":       xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
            xCampos(1, 0) = "DNI":          xCampos(1, 1) = "numdoc":       xCampos(1, 2) = "1500":    xCampos(1, 3) = "C"
                
            nSQL = "SELECT pro_emp.id, [pla_empleados].[ape] & ' ' & [pla_empleados].[nom] AS nombre, pla_empleados.numdoc " _
                + vbCr + " FROM pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp " _
                + vbCr + " Where (((pro_emp.res) = -1)) " _
                + vbCr + " ORDER BY pla_empleados.ape;"
    
            nTitulo = "Buscando Responsables de Producción"
    
        Case 10 '--DEL TURNO
            ReDim xCampos(1, 4) As String
            xCampos(0, 0) = "Turno":       xCampos(0, 1) = "nombre":       xCampos(0, 2) = "3500":    xCampos(0, 3) = "C"
                
            nSQL = "SELECT mae_turnos.id, mae_turnos.descripcion as nombre FROM mae_turnos;"
            
            nTitulo = "Buscando Turnos"
            
    End Select

    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio, ""

    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    Agregando = True
    If Col = 3 Then '--DE LA RECETA
        If GRID_BUSCAR_VALOR(Fg1, 3, xRs.Fields("codrec") & "", False, , Row) <> "-1" Then
            MsgBox "La receta ya existe" + vbCr + "Seleccione otra Receta o Elimine el Producto", vbExclamation, xTitulo
            GoTo SALIR:
        Else
            Fg1.TextMatrix(Row, 3) = xRs.Fields("codrec") & ""
            Fg1.TextMatrix(Row, 4) = xRs.Fields("abrev") & ""
            Fg1.TextMatrix(Row, 12) = xRs.Fields("idrec") & ""
            Fg1.TextMatrix(Row, 14) = xRs.Fields("idunimed") & ""
            '--ELIMINAR RECETA ANTERIOR
            DATOS_TMP_DEL Fg1.TextMatrix(Row, 11), Fg1.TextMatrix(Row, 12), RST_INSUMO
            DATOS_TMP_DEL Fg1.TextMatrix(Row, 11), Fg1.TextMatrix(Row, 12), RST_TAREA
            '--AGREGANDO NUEVA RECETA
            DATOS_TMP_ADD Fg1.TextMatrix(Row, 11), xRs.Fields("idrec"), E_INSUMO, True, True
            DATOS_TMP_ADD Fg1.TextMatrix(Row, 11), xRs.Fields("idrec"), e_TAREA, True, True
            
            If RST_INSUMO.EOF = False Or RST_INSUMO.BOF = False Or RST_INSUMO.RecordCount <> 0 Then RST_INSUMO.MoveFirst
            Do While Not RST_INSUMO.EOF
                RST_INSUMO.Fields("canteo") = NulosN(RST_INSUMO.Fields("unid")) * NulosN(Fg1.TextMatrix(Row, 6))
                RST_INSUMO.MoveNext
            Loop
        End If
        
    ElseIf Col = 9 Then '--DEL RESPONSABLE DE PRODUCCION
        Fg1.TextMatrix(Row, 9) = xRs.Fields("nombre") & ""
        Fg1.TextMatrix(Row, 15) = xRs.Fields("id") & ""
    ElseIf Col = 10 Then '--DE LA UNIDAD DE MEDIDA
        Fg1.TextMatrix(Row, 10) = xRs.Fields("nombre") & ""
        Fg1.TextMatrix(Row, 16) = xRs.Fields("id") & ""
        
    End If
    
    Agregando = False
    Set xRs = Nothing
    Exit Sub
SALIR:
    Set xRs = Nothing
    Agregando = False
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
        Exit Sub
    End If
    If Fg1.Col = 4 Then
        Fg1.Editable = flexEDNone
    Else
        If Fg1.Col = Fg1.Cols - 1 Then
            Fg1.Editable = flexEDNone
        Else
            Fg1.Editable = flexEDKbdMouse
        End If
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Row = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case Col
        Case 1, 6, 7, 8
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo error
    If KeyCode = 117 Then
        Dg(TabOne2.CurrTab).SetFocus
    End If
    If KeyCode = 114 Or KeyCode = vbKeyInsert Then 'F3 = Agregar Item
        Cmd_Click 0
    End If
    If KeyCode = 115 Or KeyCode = vbKeyDelete Then
        Cmd_Click 1  'F4 = Eliminar Item
    End If
    Exit Sub
error:
    SHOW_ERROR Me.Name, "Fg1_KeyUp"
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace = 3 Then
            PopupMenu Menu4
        Else
            PopupMenu Menu1
        End If
    End If
End Sub

Private Sub Fg1_RowColChange()
    Dim N_FILTER As String
'    If Agregando = True Then Exit Sub
'    If Fg1.Rows = 1 Then
'        Exit Sub
'    End If
'    If Fg1.TextMatrix(Fg1.Row, 12) = "" Then '--NO HAY RECETA
'        N_FILTER = "-999"
'    Else
'        N_FILTER = Fg1.TextMatrix(Fg1.Row, 12)
'    End If
'    If Fg1.Row <= 0 Then Exit Sub
'    Label2.Caption = Fg1.TextMatrix(Fg1.Row, 2)
    
    '--FILTRANDO LOS INSUMOS Y TAREAS
'    RST_INSUMO.Filter = "idparte ='" + Fg1.TextMatrix(Fg1.Row, 11) + "' AND idrec='" + N_FILTER + "'"
'    RST_TAREA.Filter = "idparte ='" + Fg1.TextMatrix(Fg1.Row, 11) + "' AND idrec='" + N_FILTER + "'"
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
   
    SeEjecuto = False
    
    mMesActivo = Month(Date)
    pCargarGrid
    
    SeEjecuto = True
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.RecordCount = 0 Then
        If MsgBox("No se ha registrado ningúna producción, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
            Nuevo
        End If
    End If
    
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    Agregando = False
    QueHace = 3
    
    F_GRUPO = True
    
    Fg1.SelectionMode = flexSelectionByRow
    
    LimpiarGrid Fg1
    
    OCULTAR_COL Fg1, 5, 5
    OCULTAR_COL Fg1, 11, 16
    
'    Dg3.Columns("dia").NumberFormat = FORMAT_DATE
        
'    Dg(0).BatchUpdates = False
'    Dg(1).BatchUpdates = False
'
'    Dg(0).Columns("unid").NumberFormat = FORMAT_PU:
'    Dg(0).Columns("canteo").NumberFormat = FORMAT_CANT:
'    Dg(0).Columns("canprog").NumberFormat = FORMAT_CANT:
'    Dg(0).Columns("canreal").NumberFormat = FORMAT_CANT:
'    Dg(0).Columns("dif").NumberFormat = FORMAT_CANT:
'
'    Dg(0).Columns("descripcion").Button = False
'    Dg(0).Columns("descripcion").ButtonAlways = False
'
'
'    Dg(1).Columns("horini").EditMask = "##:##"
'    Dg(1).Columns("horfin").EditMask = "##:##"
'    Dg(1).Columns("canper").NumberFormat = FORMAT_CANT
    
    TabOne1.CurrTab = 0
    TabOne2.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Fg1.FrozenCols = 4
    Fg1.Tag = Fg1.FormatString
    
    
    Dg3.HeadLines = 2
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 1
        Exit Sub
    Else
        Set RstFrm = Nothing
        Set RstTmp = Nothing
    End If
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
            Cancelar
            RstFrm.Requery
            Dg3.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    If Button.Index = 8 Then Filtrar
    If Button.Index = 9 Then RstFrm.Filter = ""
    
    If Button.Index = 10 Then CambiarMes
    
    If Button.Index = 11 Then Buscar
    
    If Button.Index = 15 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub

Sub Eliminar()
    On Error GoTo error
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
        MsgBox "No hay registros", vbExclamation, xTitulo
        Exit Sub
    End If
    If MsgBox("¿Esta seguro de eliminar la Producción?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
        xCon.Execute "DELETe * FROM pro_produccion WHERE id = " & RstFrm("id") & ""
        xCon.Execute "UPDATE pro_programadet SET idpro = 0 WHERE idpro = " & RstFrm("id") & ""
        
        MsgBox "La Producción del dia " + Format(RstFrm("dia"), "dd/mm/yy") + " fue eliminado con éxito", vbInformation + vbOKOnly, xTitulo
        RstFrm.Requery
        Dg3.Refresh
        If RstFrm.RecordCount = 0 Then
            If MsgBox("No hay registrado ningúna producción, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
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
    Label1.Caption = "Detalle Programa de Producción"
    Fg1.SelectionMode = flexSelectionByRow
    TabOne1.CurrTab = 0
    Dg3.SetFocus
End Sub

Sub Modificar()
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
        MsgBox "No hay Registros", vbExclamation, xTitulo
        Exit Sub
    End If

    QueHace = 2
    
    TabOne1.TabEnabled(0) = False
    ActivaTool
    Habilitar_Obj True
    If TabOne1.CurrTab = 0 Then
        TabOne1.CurrTab = 1
        MuestraSegundoTab
    End If
    Label1.Caption = "Modificando la Producción"
    GRID_COMBOLIST Fg1, 3
    GRID_COMBOLIST Fg1, 9
    GRID_COMBOLIST Fg1, 10
   
    Fg1.ColEditMask(7) = "##:##"
    Fg1.ColEditMask(8) = "##:##"
    Fg1.ColFormat(1) = FORMAT_NUM_PARTE

    '--SI DESEA AGREGAR PRODUCTOS AL GRID OBTENER EL ULTIMO NUMERO DE PRODUCCION
    M_NUM_PARTE = HallaValor(xCon, "pro_producciondet", "numparte")
    
    TxtFecha(0).Enabled = False
    txt_cb(1).SetFocus
End Sub

Private Sub MuestraSegundoTab()
    With RstFrm
        TabOne2.CurrTab = 0
        Blanquea
        If .State = 0 Then Exit Sub
        If .EOF = True Or .BOF = True Or .RecordCount = 0 Then Exit Sub
        TxtFecha(0).Valor = .Fields("dia") & ""
        
       
        txt_cb(1).Text = .Fields("supnum") & ""
        lbl_cb(1).Caption = .Fields("sup") & ""
        lbl_cb_cod(1).Caption = .Fields("idsup") & ""
        
        txt(0).Text = .Fields("id") & ""
        Fg1.ColFormat(1) = FORMAT_NUM_PARTE
        MuestraDetalle
        If Fg1.Rows >= 2 Then Fg1.Row = 1:      Fg1.Col = 1:         Fg1_RowColChange

    End With
End Sub

Private Sub MuestraDetalle()
    Dim xRs As New ADODB.Recordset
    Dim a As Integer
    Dim xCol, xFil As Integer
    Dim xSQL As String
    Dim xFch As Date
    Dim xFila  As Integer
    On Error GoTo error

    xSQL = "SELECT pro_producciondet.numparte AS idparte, pro_producciondet.idrec, pro_producciondet.iditem, pro_producciondet.idunimed, mae_turnos.id AS idturno, pro_producciondet.numparte, alm_inventario.descripcion AS proddesc, pro_receta.codrec, mae_unidades.abrev, pro_programadet.canpro AS canprog, pro_producciondet.cantidad AS canreal, pro_producciondet.horini, pro_producciondet.horfin, pro_producciondet.idres, pla_empleados.ape & ' ' & pla_empleados.nom AS resnom, mae_turnos.descripcion AS turdesc " _
        + vbCr + " FROM pla_empleados RIGHT JOIN (alm_inventario INNER JOIN ((((mae_unidades RIGHT JOIN (pro_producciondet LEFT JOIN pro_emp ON pro_producciondet.idres = pro_emp.id) ON mae_unidades.id = pro_producciondet.idunimed) INNER JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN pro_programadet ON (pro_producciondet.idpro = pro_programadet.idpro) AND (pro_producciondet.idrec = pro_programadet.idrec)) LEFT JOIN mae_turnos ON pro_producciondet.idturno = mae_turnos.id) ON alm_inventario.id = pro_receta.iditem) ON pla_empleados.id = pro_emp.idemp " _
        + vbCr + " WHERE (((pro_producciondet.idpro) = " + CStr(RstFrm.Fields("id")) + ")) "


    RST_Busq xRs, xSQL, xCon
    If xRs.RecordCount <> 0 Then
        Agregando = True
        With Fg1
            .Rows = 1
            xRs.MoveFirst
            For a = 1 To xRs.RecordCount
                xFila = .Rows
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = xRs.Fields("numparte") & ""
                .TextMatrix(.Rows - 1, 2) = xRs.Fields("proddesc") & ""
                .TextMatrix(.Rows - 1, 3) = xRs.Fields("codrec") & ""
                .TextMatrix(.Rows - 1, 4) = xRs.Fields("abrev") & ""
                .TextMatrix(.Rows - 1, 5) = Format(NulosN(xRs.Fields("canprog")), FORMAT_CANTIDAD) '--CANTIDAD PROGRAMADA
                .TextMatrix(.Rows - 1, 6) = Format(NulosN(xRs.Fields("canreal")), FORMAT_CANTIDAD) '--CANTIDAD REAL
                '--DE LA HORAS
                If IsDate(xRs.Fields("horini")) = True Then .TextMatrix(.Rows - 1, 7) = Format(xRs.Fields("horini"), FORMAT_HORA_SIN_SEGUNDO)
                If IsDate(xRs.Fields("horfin")) = True Then .TextMatrix(.Rows - 1, 8) = Format(xRs.Fields("horfin"), FORMAT_HORA_SIN_SEGUNDO)
                .TextMatrix(.Rows - 1, 9) = xRs.Fields("resnom") & "" '--ID DEL RESPONSABLE
                .TextMatrix(.Rows - 1, 10) = xRs.Fields("turdesc") & ""
                .TextMatrix(.Rows - 1, 11) = xRs.Fields("idparte") & ""
                .TextMatrix(.Rows - 1, 12) = xRs.Fields("idrec") & ""
                .TextMatrix(.Rows - 1, 13) = xRs.Fields("iditem") & ""
                .TextMatrix(.Rows - 1, 14) = xRs.Fields("idunimed") & ""
                .TextMatrix(.Rows - 1, 15) = xRs.Fields("idres") & "" '--NOMBRE DEL RESPONSABLE
                
                .TextMatrix(.Rows - 1, 16) = xRs.Fields("idturno") & ""
                '---
                DATOS_TMP_ADD xRs.Fields("idparte") & "", xRs.Fields("idrec"), E_INSUMO
                DATOS_TMP_ADD xRs.Fields("idparte") & "", xRs.Fields("idrec"), e_TAREA
                
                xRs.MoveNext
                If xRs.EOF = True Then Exit For
            Next a
        End With
        
    End If
        
    Set xRs = Nothing
    Agregando = False
    Me.MousePointer = vbDefault
    Exit Sub
error:
    SHOW_ERROR Me.Name, "MuestraDetalle"
    Me.MousePointer = vbDefault
    Set xRs = Nothing
    Agregando = False
End Sub

Private Sub Habilitar_Obj(band As Boolean)
    habilitar_Locked TxtFecha, Not band
    habilitar_Locked txt_cb, Not band
    habilitar Me.cb, band
    habilitar Cmd, band
    Dg(0).Splits(0).Locked = Not band
    Dg(1).Splits(0).Locked = Not band
    
    If band = True Then
        Dg(0).MarqueeStyle = dbgHighlightCell
        Dg(1).MarqueeStyle = dbgHighlightCell
        Dg(0).Columns("descripcion").Button = True
'        Dg(0).Columns("descripcion").ButtonAlways = False'--BOTON VISIBLE EN TODAS LAS FILAS
    Else
        Dg(0).MarqueeStyle = dbgHighlightRow
        Dg(1).MarqueeStyle = dbgHighlightRow
        Dg(0).Columns("descripcion").Button = False
'        Dg(0).Columns("descripcion").ButtonAlways = False
    End If

End Sub

Sub Blanquea()
'    LimpiaText TxtFecha
'    LimpiaText txt
'    LimpiaText txt_cb
'    LimpiarGrid Fg1, True, 1
'    OCULTAR_COL Fg1, 5, 5
'    OCULTAR_COL Fg1, 11, 16
'
'    Set Dg(0).DataSource = Nothing
'    Set Dg(1).DataSource = Nothing
'
'    pDefinirRst RST_INSUMO, E_INSUMO
'    pDefinirRst RST_TAREA, e_TAREA
'    Set Dg(0).DataSource = RST_INSUMO
'    Set Dg(1).DataSource = RST_TAREA
End Sub

Sub ActivaTool()
    Dim a&
    For a = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(a).Enabled = Not Toolbar1.Buttons(a).Enabled
    Next a
End Sub

Private Sub Nuevo()
    QueHace = 1
    TabOne1.CurrTab = 1
    TabOne2.CurrTab = 0
    TabOne1.TabEnabled(0) = False
    ActivaTool
    TxtFecha(0).Valor = Date
    Blanquea
    Habilitar_Obj True
    Label1.Caption = "Agregando Producción"
    GRID_COMBOLIST Fg1, 3
    GRID_COMBOLIST Fg1, 9
    GRID_COMBOLIST Fg1, 10
    
    M_NUM_PARTE = HallaValor(xCon, "pro_producciondet", "numparte")
    
    
    Fg1.ColEditMask(7) = "##:##"
    Fg1.ColEditMask(8) = "##:##"
    Fg1.ColFormat(1) = FORMAT_NUM_PARTE
    
    TxtFecha(0).Enabled = True
    TxtFecha(0).SetFocus
End Sub


Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index = 1 Then pImprimir True

    If ButtonMenu.Index = 2 Then pImprimir

End Sub

Function Grabar() As Boolean
    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modificar") + " la Producción", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo SALIR
    
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstIns As New ADODB.Recordset
    Dim RstTar As New ADODB.Recordset
    Dim xCod As Integer
    Dim xCodDet As Integer '--al detalle
    Dim xCol, xFil As Integer
    Dim xCorr As Integer
    
On Error GoTo LaCague

    xCon.BeginTrans
    Me.MousePointer = vbHourglass
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT top 1 * FROM pro_produccion ", xCon
        RST_Busq RstDet, "SELECT top 1 * FROM pro_producciondet", xCon
        RST_Busq RstIns, "SELECT top 1 * FROM pro_producciondetins", xCon
        RST_Busq RstTar, "SELECT top 1 * FROM pro_producciondettar", xCon
        xCod = HallaCodigoTabla("pro_produccion", xCon, "id")
        RstCab.AddNew
        RstCab("id") = xCod
    Else
        RST_Busq RstCab, "SELECT * FROM pro_produccion WHERE id =" & RstFrm("id") & "", xCon
        xCon.Execute "DELETE * FROM pro_producciondet WHERE idpro = " & RstFrm("id") & ""
        xCon.Execute "DELETE * FROM pro_producciondetins WHERE idpro = " & RstFrm("id") & ""
        xCon.Execute "DELETE * FROM pro_producciondettar WHERE idpro = " & RstFrm("id") & ""
        '--ACTUALIZANDO EL PROGRAMA A 0
        xCon.Execute "UPDATE pro_programadet SET idpro =0 WHERE idpro = " & RstFrm("id") & ""
        
        RST_Busq RstDet, "SELECT top 1 * FROM pro_producciondet", xCon
        RST_Busq RstIns, "SELECT top 1 * FROM pro_producciondetins", xCon
        RST_Busq RstTar, "SELECT top 1 * FROM pro_producciondettar", xCon
        xCod = RstFrm("id")
    End If
    
    RstCab("dia") = CDate(TxtFecha(0).Valor)
    RstCab("idsup") = NulosN(lbl_cb_cod(1).Caption)
    RstCab("num") = Format(xCod, FORMAT_NUM_PRODUCCION) '  Trim(txt(1).Text)
    RstCab("obs") = ""
    
    RstCab.Update
    
    Dim RstTmp As New ADODB.Recordset
    Dim F_CAMBIO_PRODUCCION As Boolean
    Dim M_PRODUCCION As String
    
    F_CAMBIO_PRODUCCION = False
    For xFil = 1 To Fg1.Rows - 1
        '--1=NUM PRODUCCION    12=RECETA
        If NulosN(Fg1.TextMatrix(xFil, 1)) > 0 And NulosN(Fg1.TextMatrix(xFil, 12)) > 0 Then
            RstDet.AddNew
            
            xCodDet = NulosN(Fg1.TextMatrix(xFil, 1))
            
            RstDet("idpro") = xCod
            '--VALIDAR QUE EL NUMERO DE PRODUCCION SEA DIFERENTE
            
            M_PRODUCCION = Format(xCodDet, FORMAT_NUM_PARTE)
            
            If QueHace = 1 Then
                RST_Busq RstTmp, "SELECT pro_producciondet.numparte FROM pro_producciondet WHERE (((pro_producciondet.numparte)='" + M_PRODUCCION + "') AND ((pro_producciondet.idpro)<>" + CStr(xCod) + "));", xCon
                If RstTmp.EOF = False Or RstTmp.BOF = False Or RstTmp.RecordCount <> 0 And xFil = 1 Then
                    M_PRODUCCION = Format(HallaValor(xCon, "pro_producciondet", "numparte"), FORMAT_NUM_PARTE)
                    F_CAMBIO_PRODUCCION = True
                End If
                Set RstTmp = Nothing
            End If
            RstDet("numparte") = M_PRODUCCION
            RstDet("idrec") = Fg1.TextMatrix(xFil, 12)
            '--FIN
            
            
            RstDet("iditem") = Fg1.TextMatrix(xFil, 13)
            RstDet("idunimed") = NulosN(Fg1.TextMatrix(xFil, 14))
            RstDet("cantidad") = NulosN(Fg1.TextMatrix(xFil, 6))
            RstDet("horini") = Fg1.TextMatrix(xFil, 7)
            RstDet("horfin") = Fg1.TextMatrix(xFil, 8)
            RstDet("idres") = NulosN(Fg1.TextMatrix(xFil, 15))
            RstDet("idturno") = NulosN(Fg1.TextMatrix(xFil, 16))
            
            RstDet("obs") = ""
            RstDet.Update
            '-- ACTUALIZAR EL PROGRAMA DE ACUERDO A LA FECHA Y RECETA
            xCon.Execute "UPDATE pro_programadet SET idpro =" + CStr(xCod) + " WHERE dia = CDATE('" + TxtFecha(0).Valor + "') AND idrec = " + CStr(Fg1.TextMatrix(xFil, 12)) + ";"
            '----ADD INSUMOS
            RST_INSUMO.Filter = "idparte = " + Fg1.TextMatrix(xFil, 11) + " AND idrec=" + Fg1.TextMatrix(xFil, 12)
            
            If RST_INSUMO.RecordCount > 0 Then
                RST_INSUMO.MoveFirst
                Do While Not RST_INSUMO.EOF
                    If NulosN(RST_INSUMO.Fields("iditem")) <> 0 Then
                        RstIns.AddNew
                        '--CLAVE
                        RstIns("idpro") = xCod
                        RstIns("numparte") = M_PRODUCCION
                        RstIns("idrec") = Fg1.TextMatrix(xFil, 12)
                        RstIns("iditem") = NulosN(RST_INSUMO.Fields("iditem"))
                        '-FIN CLAVE
                        
                        RstIns("idunimed") = NulosN(RST_INSUMO.Fields("idunimed"))
                        RstIns("canutil") = NulosN(RST_INSUMO.Fields("canreal"))
                        RstIns("canpro") = NulosN(RST_INSUMO.Fields("unid"))
                        RstIns.Update
                    End If
                    RST_INSUMO.MoveNext
                Loop
            End If
            '----ADD TAREAS
            RST_TAREA.Filter = "idparte = " + Fg1.TextMatrix(xFil, 11) + " AND idrec=" + Fg1.TextMatrix(xFil, 12)
            xCorr = 1
            If RST_TAREA.RecordCount > 0 Then
                RST_TAREA.MoveFirst
                Do While Not RST_TAREA.EOF
                    RstTar.AddNew
                    '--CLAVE
                    RstTar("idpro") = xCod
                    RstTar("numparte") = M_PRODUCCION
                    RstTar("idrec") = Fg1.TextMatrix(xFil, 12)
                    RstTar("idtar") = NulosN(RST_TAREA.Fields("idtar"))
                    
                    RstTar("corr") = xCorr
                    '-FIN CLAVE
                    
                    RstTar("idunimed") = NulosN(RST_TAREA.Fields("idunimed"))

                    If IsDate(RST_TAREA.Fields("horini")) = True Then RstTar("horini") = CDate(RST_TAREA.Fields("horini"))
                    If IsDate(RST_TAREA.Fields("horfin")) = True Then RstTar("horfin") = CDate(RST_TAREA.Fields("horfin"))
                    RstTar("canper") = NulosN(RST_TAREA.Fields("canper"))


                    RstTar.Update
                    RST_TAREA.MoveNext
                    xCorr = xCorr + 1
                Loop
            End If
        End If
    Next xFil
    
    MsgBox "La Producción se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito" + _
    IIf(F_CAMBIO_PRODUCCION = True, vbCr + "El número de Producción se cambió", ""), vbInformation, xTitulo
    
    xCon.CommitTrans
    Grabar = True
SALIR:
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstIns = Nothing:    Set RstTar = Nothing:    Set RstTmp = Nothing
    Exit Function
LaCague:
    xCon.RollbackTrans
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstIns = Nothing:    Set RstTar = Nothing:    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo :"
    Grabar = False
End Function

Private Function fValidarDatos() As Boolean
    If TxtFecha(0).Valor = "" Or IsDate(TxtFecha(0).Valor) = False Then
        MsgBox "No ha especificado la fecha de Producción ", vbInformation, xTitulo
        TxtFecha(0).SetFocus
        Exit Function
    End If
    
    Dim band As Integer
    band = Validar(txt_cb)
    If band <> -1 Then
       MsgBox "Llene el Campo de " & lbl_cb_capt(band).Caption, vbInformation, xTitulo
       txt_cb(band).SetFocus
       Exit Function
    End If

    
    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado los productos para la producción", vbInformation, xTitulo
        Fg1.SetFocus
        Exit Function
    End If
    '---------------------------------------------------------------------------
    '--VALIDAR EL INGRESO DE LOS DATOS
    Dim Q_ROW  As Long
    Dim Q_COL As Long '--COLUMNA A POSICIONAR SI FALTAN DATOS
    Q_COL = -1
    For Q_ROW = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(Q_ROW, 1)) = 0 Then
            MsgBox "Ingrese El número de Producción del Producto:" + vbCr + _
            "Producto:        " + Fg1.TextMatrix(Q_ROW, 2) & "", vbExclamation, xTitulo
            Q_COL = 1:          Exit For
        ElseIf NulosN(Fg1.TextMatrix(Q_ROW, 12)) = 0 Then
            MsgBox "Ingrese La Receta del Producto:" + vbCr + _
            "Producto:        " + Fg1.TextMatrix(Q_ROW, 2) & "", vbExclamation, xTitulo
            Q_COL = 12:          Exit For
        ElseIf IsNumeric(Fg1.TextMatrix(Q_ROW, 6)) = False Or Fg1.TextMatrix(Q_ROW, 6) = "0" Then
            MsgBox "Ingrese el total de Producción:" + vbCr + _
            "Producto:        " + Fg1.TextMatrix(Q_ROW, 2) & "" + vbCr + _
            "Receta:         " + Fg1.TextMatrix(Q_ROW, 3) & "" + vbCr, vbExclamation, xTitulo
            
            Q_COL = 6:          Exit For
        ElseIf IsDate(Fg1.TextMatrix(Q_ROW, 7)) = False Or IsDate(Fg1.TextMatrix(Q_ROW, 8)) = False Then
            MsgBox "Ingrese el la Hora " + IIf(IsDate(Fg1.TextMatrix(Q_ROW, 7)) = False, "Inicial", "Final") + " de la Producción" + vbCr + _
            "Producto:  " + Fg1.TextMatrix(Q_ROW, 2) & "" + vbCr + _
            "Receta:         " + Fg1.TextMatrix(Q_ROW, 3) & "" + vbCr, vbExclamation, xTitulo
            
            Q_COL = IIf(IsDate(Fg1.TextMatrix(Q_ROW, 7)) = False, 7, 8):        Exit For
            
        ElseIf IsNumeric(Fg1.TextMatrix(Q_ROW, 15)) = False Or Fg1.TextMatrix(Q_ROW, 15) = "0" Then
            MsgBox "Ingrese el Responsable de la Producción:" + vbCr + _
            "Producto:  " + Fg1.TextMatrix(Q_ROW, 2) & "" + vbCr + _
            "Receta:         " + Fg1.TextMatrix(Q_ROW, 3) & "" + vbCr, vbExclamation, xTitulo
            
            Q_COL = 9:       Exit For
        ElseIf IsNumeric(Fg1.TextMatrix(Q_ROW, 16)) = False Or Fg1.TextMatrix(Q_ROW, 16) = "0" Then
            MsgBox "Ingrese el Turno de la Producción:" + vbCr + _
            "Producto:  " + Fg1.TextMatrix(Q_ROW, 2) & "" + vbCr + _
            "Receta:         " + Fg1.TextMatrix(Q_ROW, 3) & "" + vbCr, vbExclamation, xTitulo
            
            Q_COL = 10:       Exit For
        End If
    Next Q_ROW
    If Q_COL <> -1 Then
        Agregando = True:  Fg1.Row = Q_ROW: Fg1.Col = Q_COL: Agregando = False
        Exit Function
    End If
    '---------------------------------------------------------------------------

    fValidarDatos = True
End Function

Private Sub pCargarGrid()
    On Error GoTo error
    Dim xSQL  As String
    lblperiodo.Caption = MonthName(mMesActivo)
    F_GRUPO = True
    
    xSQL = " SELECT pro_produccion.id, pro_produccion.num, pro_produccion.dia, pro_produccion.idsup, pla_empleados_1.numdoc AS supnum, pla_empleados_1.ape & ' ' & pla_empleados_1.nom AS sup, alm_inventario.descripcion AS proddesc, pro_emp_2.id AS idres, pla_empleados_2.ape & ' ' & pla_empleados_2.nom AS resnom, pro_producciondet.numparte, pro_producciondet.cantidad, pro_receta.codrec " _
    + vbCr + " FROM ((pro_produccion LEFT JOIN pro_emp AS pro_emp_1 ON pro_produccion.idsup = pro_emp_1.id) LEFT JOIN pla_empleados AS pla_empleados_1 ON pro_emp_1.idemp = pla_empleados_1.id) LEFT JOIN ((((pro_producciondet LEFT JOIN pro_emp AS pro_emp_2 ON pro_producciondet.idres = pro_emp_2.id) LEFT JOIN pla_empleados AS pla_empleados_2 ON pro_emp_2.idemp = pla_empleados_2.id) LEFT JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id) LEFT JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) ON pro_produccion.id = pro_producciondet.idpro " _
    + vbCr + " WHERE YEAR(pro_produccion.dia)= " + AnoTra + " AND MONTH(pro_produccion.dia)= " + CStr(mMesActivo) + " " _
    + vbCr + " ORDER BY pro_produccion.dia,pro_produccion.num;"
    
    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, xSQL, xCon
    
    Set Dg3.DataSource = RstFrm
    Me.MousePointer = vbDefault
    F_GRUPO = False
Exit Sub
error:
    F_GRUPO = False
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pCargarGrid"
End Sub

Private Sub CambiarMes()
    Dim xMes  As Integer
    xMes = SeleccionaMes(xCon)
    If xMes = 0 Or xMes = 13 Then Exit Sub
    mMesActivo = xMes
    pCargarGrid
    TabOne1.CurrTab = 0
End Sub
Private Sub txt_cb_Change(Index As Integer)
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cb_cod(Index).Caption = ""
    End If
End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo error
    If txt_cb(Index).Locked = True Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb_Click Index
        Exit Sub
    End If

End Sub

Private Sub txt_cb_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txt_cb(1).Text) <> "" Then
            Cmd(0).SetFocus
        Else
            SendKeys vbTab
        End If
        Exit Sub
    End If
    Select Case Index
        Case 3: If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case Else: If validar_numero(KeyAscii) = True And KeyAscii = 46 Then KeyAscii = 0
    End Select
    
End Sub

Private Sub pDefinirRst(rst As ADODB.Recordset, fTipo As e_PROGRAMA)
    '--DEFINIR EL RECORSET TEMPORAL PARA INSUMO Y TAREA
    Dim RST_ORIGEN As New ADODB.Recordset
    Dim nSQL As String
    nSQL = pGenerarConsulta("-1", "-1", fTipo, CDate("01/01/07"), True)
    RST_Busq RST_ORIGEN, nSQL, xCon
    DEFINIR_RST_TMP rst, RST_ORIGEN
    Set RST_ORIGEN = Nothing
End Sub


Private Sub pImprimir(Optional IMP_LISTADO As Boolean = False)

    On Error GoTo error

    Me.MousePointer = vbHourglass
    If IMP_LISTADO = False Then
        If Me.TabOne1.CurrTab = 0 Then
        
        Else
''            MsgBox "Primero muestre el detalle del Registro" + vbCr + _
''                   "Luego inténtelo otra vez", vbExclamation, xTitulo
        End If
    Else
    
        TDB_IMPRIMIR Dg3, "IMPRESIÓN DE PRODUCCIÓN", "LISTADO DE PRODUCCIÓN  -  Periodo: " + MonthName(mMesActivo, False)
   
    End If

    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimir"

End Sub

Sub Buscar()
    On Error GoTo error
    TabOne1.CurrTab = 0
     
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
   
    Dim xCampos(6, 4) As String
    
    xCampos(0, 0) = "Fch.Prod":         xCampos(0, 1) = "dia":       xCampos(0, 2) = "850":    xCampos(0, 3) = "F"
    xCampos(1, 0) = "N°.Prod":          xCampos(1, 1) = "num":       xCampos(1, 2) = "900":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Receta":           xCampos(2, 1) = "codrec":    xCampos(2, 2) = "1000":   xCampos(2, 3) = "C"
    xCampos(3, 0) = "Producto":         xCampos(3, 1) = "proddesc":  xCampos(3, 2) = "3200":   xCampos(3, 3) = "C"
    xCampos(4, 0) = "Cantidad":         xCampos(4, 1) = "cantidad":  xCampos(4, 2) = "800":    xCampos(4, 3) = "N"
    xCampos(5, 0) = "Responsable":      xCampos(5, 1) = "resnom":    xCampos(5, 2) = "1500":   xCampos(5, 3) = "C"
        
        
    nSQL = " SELECT pro_produccion.id, pro_produccion.num, format(pro_produccion.dia,'dd/mm/yy') as dia, pro_produccion.idsup, pla_empleados_1.numdoc AS supnum, pla_empleados_1.ape & ' ' & pla_empleados_1.nom AS sup, alm_inventario.descripcion AS proddesc, pro_emp_2.id AS idres, pla_empleados_2.ape & ' ' & pla_empleados_2.nom AS resnom, pro_producciondet.numparte, pro_producciondet.cantidad, pro_receta.codrec " _
    + vbCr + " FROM ((pro_produccion LEFT JOIN pro_emp AS pro_emp_1 ON pro_produccion.idsup = pro_emp_1.id) LEFT JOIN pla_empleados AS pla_empleados_1 ON pro_emp_1.idemp = pla_empleados_1.id) LEFT JOIN ((((pro_producciondet LEFT JOIN pro_emp AS pro_emp_2 ON pro_producciondet.idres = pro_emp_2.id) LEFT JOIN pla_empleados AS pla_empleados_2 ON pro_emp_2.idemp = pla_empleados_2.id) LEFT JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id) LEFT JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) ON pro_produccion.id = pro_producciondet.idpro " _
    + vbCr + " WHERE YEAR(pro_produccion.dia)= " + AnoTra + " AND MONTH(pro_produccion.dia)= " + CStr(mMesActivo) + ""

    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Producción", "dia", "proddesc", Principio
    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    
    RstFrm.MoveFirst
    RstFrm.Find "id = " + CStr(xRs("id"))
SALIR:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub




Private Sub Filtrar()
    
    Dim xCampos(5, 4) As String
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Producto":     xCampos(0, 1) = "proddesc": xCampos(0, 2) = "C":         xCampos(0, 3) = "1500"
    xCampos(1, 0) = "Fch. Pro":     xCampos(1, 1) = "dia":      xCampos(1, 2) = "F":         xCampos(1, 3) = "1000"
    xCampos(2, 0) = "N° Prod.":     xCampos(2, 1) = "num":      xCampos(2, 2) = "C":         xCampos(2, 3) = "800"
    xCampos(3, 0) = "Receta":       xCampos(3, 1) = "codrec":   xCampos(3, 2) = "C":         xCampos(3, 3) = "1500"
    xCampos(4, 0) = "Cantidad":     xCampos(4, 1) = "cantidad": xCampos(4, 2) = "N":         xCampos(4, 3) = "1500"
    xCampos(5, 0) = "Responsable":  xCampos(5, 1) = "resnom":   xCampos(5, 2) = "C":         xCampos(5, 3) = "1500"
    
    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg3

    TabOne1.CurrTab = 0
End Sub


Private Function HallaValor(conn As ADODB.Connection, tabla As String, campo As String) As Long
Dim xRs As New ADODB.Recordset
On Error GoTo error
RST_Busq xRs, "SELECT top 1 CLng([" + campo + "]) AS num FROM " + tabla + " ORDER BY CLng([" + campo + "]) DESC;", conn
If xRs.State = 1 Then
    If xRs.EOF = False And xRs.BOF = False And xRs.RecordCount <> 0 Then
        HallaValor = NulosN(xRs.Fields(0)) + 1
    End If
Else
    HallaValor = -1
End If
Set xRs = Nothing
Exit Function
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "HallarValor"
End Function

Private Sub txt_cb_Validate(Index As Integer, Cancel As Boolean)
    If txt_cb(Index).Text = "" Then Exit Sub
    Dim RST_TMP As New ADODB.Recordset
    Dim nSQL As String

    nSQL = "SELECT pla_empleados.numdoc, [pla_empleados].[ape] & ' ' & [pla_empleados].[nom] AS nombre, pro_emp.id " _
        + vbCr + " FROM pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp " _
        + vbCr + " Where (((pro_emp.prog) = -1)) and pla_empleados.numdoc ='" + Trim(txt_cb(Index).Text) + "'" _
        + vbCr + " ORDER BY pla_empleados.ape;"
    If Index = 1 Then
        nSQL = Replace(nSQL, "(pro_emp.prog)", "(pro_emp.sup)")
    End If

    If xCon.State = 0 Then Exit Sub
    RST_Busq RST_TMP, nSQL, xCon
    
    If RST_TMP.State = 0 Then Exit Sub
    If RST_TMP.RecordCount > 0 Then
        txt_cb(Index).Text = RST_TMP.Fields(0) & ""  '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = RST_TMP.Fields(1) & "" '--NOMBRE
        lbl_cb_cod(Index).Caption = RST_TMP.Fields(2) & "" '--CODIGO
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cb_cod(Index).Caption = ""
    End If
    Set RST_TMP = Nothing
    Exit Sub
error:
    Set RST_TMP = Nothing
    SHOW_ERROR Me.Name, "txt_cb_KeyDown(" + CStr(Index) + ")"
End Sub

'*********************************************************************
Private Sub CmdMover_Click(Index As Integer)
    Ocultar CmdMover, False
    
    If Index = 0 Then '--arriba
        CmdMover(1).Visible = True
        Ocultar Cmd, False
        Fg1.Visible = False
        TabOne2.Height = 6120
        TabOne2.Top = 765
        Fg(0).Height = 5535
        Fg(1).Height = 5535
        
    Else '--abajo
        CmdMover(0).Visible = True
        Ocultar Cmd, True
        Fg1.Visible = True
        TabOne2.Height = 2910
        TabOne2.Top = 3885
        
        Fg(0).Height = 2535
        Fg(1).Height = 2535
    End If
End Sub


Private Sub pConfigurarGrilla()
    Dim RstTmp As New ADODB.Recordset
    Dim tFormat$
   
    With Fg1 '--de los ingredientes
        .Rows = 2
        .Cols = 9
        .FixedRows = 2
        .RowHeight(0) = 250
                
        GRID_COMBINAR Fg1, 0, 1, 0, 4, "Origen", flexAlignCenterCenter, True, flexMergeFree
        GRID_COMBINAR Fg1, 0, 5, 0, 8, "Resultado", flexAlignCenterCenter, True, flexMergeFree
        .TextMatrix(1, 1) = "Tipo":              .ColWidth(1) = 650:   .ColAlignment(1) = flexAlignLeftCenter:    .Row = 1: .Col = 1: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 2) = "Descripción":       .ColWidth(2) = 1800:  .ColAlignment(2) = flexAlignLeftCenter:    .Row = 1: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 3) = "Cant Inicial":      .ColWidth(3) = 1000:  .ColAlignment(3) = flexAlignRightCenter:   .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 4) = "U.M.":              .ColWidth(4) = 650:   .ColAlignment(4) = flexAlignCenterCenter:  .Row = 1: .Col = 4: .CellAlignment = flexAlignLeftCenter
        
        .TextMatrix(1, 5) = "Producto":          .ColWidth(5) = 3500:  .ColAlignment(5) = flexAlignRightCenter:    .Row = 1: .Col = 5: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 6) = "Receta":            .ColWidth(6) = 1200:  .ColAlignment(6) = flexAlignLeftCenter:     .Row = 1: .Col = 6: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 7) = "Cant Final":        .ColWidth(7) = 1000:  .ColAlignment(7) = flexAlignRightCenter:    .Row = 1: .Col = 7: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 8) = "U.M.":              .ColWidth(8) = 650:   .ColAlignment(8) = flexAlignCenterCenter:   .Row = 1: .Col = 8: .CellAlignment = flexAlignCenterCenter
        
        .TextMatrix(1, 9) = "Responsable":       .ColWidth(9) = 2500:  .ColAlignment(9) = flexAlignLeftCenter:     .Row = 1: .Col = 9: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 10) = "Nº Lote":          .ColWidth(10) = 800:  .ColAlignment(10) = flexAlignLeftCenter:    .Row = 1: .Col = 10: .CellAlignment = flexAlignLeftCenter
        
        .TextMatrix(1, 11) = "IdTipo":           .ColWidth(11) = 0:
        .TextMatrix(1, 12) = "IdOrigen":         .ColWidth(12) = 0:
        .TextMatrix(1, 13) = "IdProducto":       .ColWidth(13) = 0:
        .TextMatrix(1, 14) = "IdRec":            .ColWidth(14) = 0:
        .TextMatrix(1, 15) = "IdResponsable":    .ColWidth(15) = 0:
        
        
        .ColFormat(4) = FORMAT_MONTO
        .ColFormat(7) = FORMAT_MONTO
                
        .SelectionMode = flexSelectionByRow
                
        GRID_COMBOLIST Fg1, 2  '--origen
        GRID_COMBOLIST Fg1, 5 '--producto
        GRID_COMBOLIST Fg1, 6 '--receta
        GRID_COMBOLIST Fg1, 9 '--responsable
        
        '--Tipo de Origen (Materia Prima; Producto)
        RST_Busq RstTmp, "SELECT mae_tipoproducto.id, mae_tipoproducto.descripcion FROM mae_tipoproducto WHERE (((mae_tipoproducto.id) In (1,3))) ORDER BY mae_tipoproducto.descripcion; ", xCon
        tFormat = Fg1.BuildComboList(RstTmp, "descripcion", "id", vbYellow)
        Fg1.ColComboList(1) = tFormat
        Set RstTmp = Nothing
        DoEvents
                
    End With
    
    With Fg(0) '--Requerimiento de materiales
        .Rows = 1
        .Cols = 13
        .FixedRows = 1
        .RowHeight(0) = 300
        
        .TextMatrix(0, 1) = "Tipo":             .ColWidth(1) = 1000:  .ColAlignment(1) = flexAlignLeftCenter:  .Row = 0: .Col = 1: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 2) = "Descripción":      .ColWidth(2) = 3500:  .ColAlignment(2) = flexAlignLeftCenter:  .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 3) = "Stock":            .ColWidth(3) = 1200:  .ColAlignment(3) = flexAlignRightCenter: .Row = 0: .Col = 3: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 4) = "Cant Requerida":   .ColWidth(4) = 1200:  .ColAlignment(4) = flexAlignRightCenter: .Row = 0: .Col = 4: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 5) = "Saldo":            .ColWidth(5) = 1200:  .ColAlignment(5) = flexAlignRightCenter: .Row = 0: .Col = 5: .CellAlignment = flexAlignRightCenter
        
        
        .ColFormat(3) = FORMAT_MONTO
        .ColFormat(4) = FORMAT_MONTO
        .ColFormat(5) = FORMAT_MONTO
        
        .ColFormat(10) = "0.0000000"
        .ColFormat(11) = "0.0000000"
            
        .SelectionMode = flexSelectionByRow
        
    End With
    
    '*****************************************
    '--Ingredientes
    GRID_COMBOLIST Fg1, 2  '--descripcion
    GRID_COMBOLIST Fg1, 4 '--unidad
    '--Tareas
    GRID_COMBOLIST Fg3, 1 '--tarea
    '*****************************************

End Sub



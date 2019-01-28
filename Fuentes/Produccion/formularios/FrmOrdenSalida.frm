VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SIZERONE.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmOrdenSalida 
   Caption         =   "Produccion - Solicitud de Materiales"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   11205
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6930
      Left            =   15
      TabIndex        =   6
      Top             =   375
      Width           =   11190
      _cx             =   19738
      _cy             =   12224
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
         Height          =   6510
         Left            =   -11745
         TabIndex        =   12
         Top             =   375
         Width           =   11100
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6195
            Left            =   45
            TabIndex        =   13
            Top             =   315
            Width           =   11040
            _ExtentX        =   19473
            _ExtentY        =   10927
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Fch. Sol."
            Columns(0).DataField=   "fchped"
            Columns(0).NumberFormat=   "dd/mm/yy"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Tip. Doc."
            Columns(1).DataField=   "abredoc"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nº Documento"
            Columns(2).DataField=   "numdoc"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Solicitante"
            Columns(3).DataField=   "apenomsol"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Area"
            Columns(4).DataField=   "descripcion"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Producto"
            Columns(5).DataField=   "descpro"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Receta"
            Columns(6).DataField=   "codrec"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Cantidad"
            Columns(7).DataField=   "can"
            Columns(7).NumberFormat=   "0.00"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=8"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1561"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1482"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=513"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=1561"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1482"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2328"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2249"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=4445"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=4366"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=512"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=2302"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2223"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=1773"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1693"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(43)=   "Column(7).Width=1720"
            Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=1640"
            Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=514"
            Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HE0FEFE&,.bold=0,.fontsize=825"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13,.alignment=2"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=1"
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
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Solicitud de Materiales"
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
            TabIndex        =   14
            Top             =   30
            Width           =   11070
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   6510
         Left            =   45
         TabIndex        =   7
         Top             =   375
         Width           =   11100
         Begin VB.Frame Frame3 
            Height          =   1800
            Left            =   9195
            TabIndex        =   25
            Top             =   2325
            Width           =   1830
            Begin VB.CommandButton CmdDelOrden 
               Caption         =   "Eliminar Orden de Produccion"
               Height          =   480
               Left            =   180
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   975
               Width           =   1500
            End
            Begin VB.CommandButton CmdAddOrden 
               Caption         =   "Agregar Orden de Produccion"
               Height          =   480
               Left            =   180
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   465
               Width           =   1500
            End
         End
         Begin VB.CommandButton CmdBusSol 
            Height          =   240
            Left            =   6960
            Picture         =   "FrmOrdenSalida.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   1830
            Width           =   240
         End
         Begin VB.CommandButton CmdBusAre 
            Height          =   240
            Left            =   2100
            Picture         =   "FrmOrdenSalida.frx":0132
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   1515
            Width           =   240
         End
         Begin VB.CommandButton CmdBusDoc 
            Height          =   240
            Left            =   2100
            Picture         =   "FrmOrdenSalida.frx":0264
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   885
            Width           =   240
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchPed 
            Height          =   300
            Left            =   1515
            TabIndex        =   0
            Top             =   540
            Width           =   1200
            _ExtentX        =   2117
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
         Begin VB.TextBox TxtSolicitante 
            Height          =   300
            Left            =   1515
            Locked          =   -1  'True
            TabIndex        =   5
            Text            =   "TxtSolicitante"
            Top             =   1800
            Width           =   5715
         End
         Begin VB.TextBox TxtIdArea 
            Height          =   300
            Left            =   1515
            Locked          =   -1  'True
            TabIndex        =   4
            Text            =   "TxtIdArea"
            Top             =   1485
            Width           =   855
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2535
            Locked          =   -1  'True
            TabIndex        =   3
            Text            =   "TxtNumDoc"
            Top             =   1170
            Width           =   1515
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1515
            Locked          =   -1  'True
            TabIndex        =   2
            Text            =   "TxtNumSer"
            Top             =   1170
            Width           =   855
         End
         Begin VB.TextBox TxtTipDoc 
            Height          =   300
            Left            =   1515
            Locked          =   -1  'True
            TabIndex        =   1
            Text            =   "TxtTipDoc"
            Top             =   855
            Width           =   855
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   1710
            Left            =   105
            TabIndex        =   16
            Top             =   2415
            Width           =   9045
            _cx             =   15954
            _cy             =   3016
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
            Cols            =   9
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmOrdenSalida.frx":0396
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
            Height          =   1950
            Left            =   105
            TabIndex        =   29
            Top             =   4455
            Width           =   10905
            _cx             =   19235
            _cy             =   3440
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
            Cols            =   9
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmOrdenSalida.frx":04AD
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
         Begin VB.Label LblIdSol 
            AutoSize        =   -1  'True
            Caption         =   "LblIdSol"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   7380
            TabIndex        =   30
            Top             =   1875
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Insumos a Utilizar"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   4215
            Width           =   1230
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Top             =   585
            Width           =   450
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Solicitante"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   1845
            Width           =   735
         End
         Begin VB.Label LblTipDocumento 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipDocumento"
            Height          =   300
            Left            =   2430
            TabIndex        =   19
            Top             =   855
            Width           =   4800
         End
         Begin VB.Label LblArea 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblArea"
            Height          =   300
            Left            =   2430
            TabIndex        =   18
            Top             =   1485
            Width           =   4800
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Area"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   1530
            Width           =   330
         End
         Begin VB.Line Line1 
            BorderWidth     =   5
            X1              =   2415
            X2              =   2460
            Y1              =   1305
            Y2              =   1305
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Solicitud"
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
            TabIndex        =   11
            Top             =   30
            Width           =   11070
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   900
            Width           =   1185
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Productos a Generar"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   2175
            Width           =   1470
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Codigo"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   1215
            Width           =   495
         End
      End
   End
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
            Picture         =   "FrmOrdenSalida.frx":05CF
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenSalida.frx":0B13
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenSalida.frx":0EA5
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenSalida.frx":1029
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenSalida.frx":147D
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenSalida.frx":1595
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenSalida.frx":1AD9
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenSalida.frx":201D
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenSalida.frx":2131
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenSalida.frx":2245
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenSalida.frx":2699
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenSalida.frx":2805
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11205
      _ExtentX        =   19764
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   8
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Opciones de Impresion"
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
End
Attribute VB_Name = "FrmOrdenSalida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstDetIns As New ADODB.Recordset
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim Agregando As Boolean
Dim CaracteresNumericos As String
Dim RstOrd As New ADODB.Recordset

Function Grabar() As Boolean
    Grabar = False
    
    If NulosC(TxtFchPed.Valor) = "" Then
        MsgBox "No ha especificado la fecha del pedido", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchPed.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtTipDoc.Text) = "" Then
        MsgBox "No ha especificado el tipo de documento para la solicitud de materiales", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtNumSer.Text) = "" Then
        MsgBox "No ha especificado el numero de serie para el documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumSer.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtNumDoc.Text) = "" Then
        MsgBox "No ha especificado el numero de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDoc.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtIdArea.Text) = "" Then
        MsgBox "No ha especificado el areas solicitante", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdArea.SetFocus
        Exit Function
    End If
    
    If TxtSolicitante.Text = "" Then
        MsgBox "No ha especificado el nombre del solicitante", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtSolicitante.SetFocus
        Exit Function
    End If
    
    
    Dim A, xId As Integer
    Dim Rst As New ADODB.Recordset
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT * FROM pro_ordensalida WHERE serord = '" & TxtNumSer.Text & "' and numord = '" & TxtNumDoc.Text & "'", xCon
    If Rst.RecordCount <> 0 Then
        MsgBox "El numero de documento " + Trim(TxtNumSer.Text) + "-" + Trim(TxtNumDoc.Text) + " ya existe, " & Chr(16) _
            & " se ha asignado el numero " + Trim(TxtNumSer.Text) + "-" + HallaNumDoc(TxtNumSer.Text), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDoc.Text = HallaNumDoc(TxtNumSer.Text)
    End If
    
On Error GoTo LaCague
    
    xCon.BeginTrans
    If QueHace = 1 Then
        xId = HallaCodigoTabla("pro_ordensalida", xCon, "id")
        RST_Busq RstCab, "SELECT * FROM pro_ordensalida", xCon
        RST_Busq RstDet, "SELECT * FROM pro_ordensalidadet", xCon
        RstCab.AddNew
        RstCab("id") = xId
    End If
    
    RstCab("serord") = NulosC(TxtNumSer.Text)
    RstCab("numord") = NulosC(TxtNumDoc.Text)
    RstCab("idtipdoc") = NulosN(TxtTipDoc.Text)
    RstCab("idare") = NulosN(TxtIdArea.Text)
    RstCab("fchped") = TxtFchPed.Valor
    RstCab("idres") = Val(LblIdSol.Caption)
    RstCab.Update
    
    If RstDetIns.RecordCount <> 0 Then
        RstDetIns.Filter = adFilterNone
        RstDetIns.MoveFirst
        For A = 1 To RstDetIns.RecordCount
            RstDet.AddNew
            RstDet("id") = xId
            RstDet("idpro") = RstDetIns("idorden")
            RstDet("idrec") = RstDetIns("idrec")
            RstDet("iditem") = RstDetIns("idprod")
            RstDet("idins") = RstDetIns("iditem")
            RstDet("idunimed") = RstDetIns("idunimed")
            RstDet("canteo") = RstDetIns("cantidad")
            RstDet.Update
            
            RstDetIns.MoveNext
            If RstDetIns.EOF = True Then
                Exit For
            End If
        Next A
    End If
    
    xCon.Execute "UPDATE pro_produccion SET pro_produccion.idsolmat = " & xId & " WHERE (((pro_produccion.id)=" & Val(Fg1.TextMatrix(1, 8)) & "))"

    xCon.CommitTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    
    MsgBox "La Solicitud de materiales se grabo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Exit Function
LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
End Function

Sub Nuevo()
    QueHace = 1
    ActivaTool
    Blanquea
    Bloquea
    Label1.Caption = "Agregando Solicitud de Materiales"
    TabOne1.CurrTab = 1
    Fg1.Rows = 1
    Fg2.Rows = 1
    TabOne1.TabEnabled(0) = False
    PreparaRST
    TxtNumSer.Text = "0001"
    TxtNumDoc.Text = HallaNumDoc(TxtNumSer.Text)
    TxtFchPed.SetFocus
End Sub

Function HallaNumDoc(NumeroSerie As String)
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT * FROM pro_ordensalida WHERE serord = '" & NumeroSerie & "' ORDER BY numord ", xCon
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveLast
        HallaNumDoc = Format(Val(Rst("numord")) + 1, "00000000")
    Else
        HallaNumDoc = Format("1", "00000000")
    End If
End Function

Private Sub CmdAddOrden_Click()
    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(5, 4) As String
    Dim xCampos2(4, 4) As String
    Dim xCampos3(2, 4) As String
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
        xCampos(0, 0) = "Producto":           xCampos(0, 1) = "descripcion":     xCampos(0, 2) = "4100":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Uni. Med.":          xCampos(1, 1) = "abrev":           xCampos(1, 2) = "950":         xCampos(1, 3) = "C"
        xCampos(2, 0) = "Receta":             xCampos(2, 1) = "codrec":          xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
        xCampos(3, 0) = "Cantidad":           xCampos(3, 1) = "can":             xCampos(3, 2) = "900":         xCampos(3, 3) = "N"
        xCampos(4, 0) = "Nº Ord. Prod.":      xCampos(4, 1) = "id":              xCampos(4, 2) = "1250":         xCampos(4, 3) = "N"
        
        xForm.SQLCad = "SELECT pro_produccion.id, pro_produccion.fchini, pro_produccion.fchfin, alm_inventario.codpro, alm_inventario.descripcion, " _
            & " pro_producciondet.iditem, pro_receta.codrec, mae_unidades.abrev, pro_producciondet.can, UCase([pla_empleados]![ape])+', '+[pla_empleados]![nom] AS apenomsup, " _
            & " pro_producciondet.idrec, pro_produccion.idres, pro_produccion.idsup, UCase([pla_empleados_1]![ape])+', '+[pla_empleados_1]![nom] AS apenomresp, " _
            & " pro_producciondet.idunimed FROM ((pro_produccion LEFT JOIN pla_empleados ON pro_produccion.idres = pla_empleados.id) LEFT JOIN " _
            & " pla_empleados AS pla_empleados_1 ON pro_produccion.idsup = pla_empleados_1.id) INNER JOIN (((pro_producciondet LEFT JOIN pro_receta " _
            & " ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id) LEFT JOIN mae_unidades " _
            & " ON pro_producciondet.idunimed = mae_unidades.id) ON pro_produccion.id = pro_producciondet.id Where (((pro_produccion.idsolmat) = 0)) " _
            & " ORDER BY pro_produccion.id DESC "
    
        xForm.Titulo = "Buscando Ordenes de Produccion"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        xForm.Ordenado = "descripcion"
        xForm.CampoBusca = "descripcion"
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Dim Rst As New ADODB.Recordset
                Dim A As Integer
                
                Agregando = True
                RST_Busq Rst, "SELECT pro_produccion.id, pro_produccion.fchini, pro_produccion.fchfin, alm_inventario.codpro, pro_producciondet.iditem, " _
                    & " pro_producciondet.idrec, alm_inventario.descripcion, pro_receta.codrec, mae_unidades.abrev, pro_producciondet.can, " _
                    & " UCase([pla_empleados]![ape])+', '+[pla_empleados]![nom] AS apenomsup, pro_produccion.idres, pro_produccion.idsup, " _
                    & " UCase([pla_empleados_1]![ape])+', '+[pla_empleados_1]![nom] AS apenomresp, pro_producciondet.idunimed FROM ((pro_produccion LEFT JOIN " _
                    & " pla_empleados ON pro_produccion.idres = pla_empleados.id) LEFT JOIN pla_empleados AS pla_empleados_1 ON pro_produccion.idsup = pla_empleados_1.id) " _
                    & " INNER JOIN (((pro_producciondet LEFT JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN alm_inventario " _
                    & " ON pro_producciondet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_producciondet.idunimed = mae_unidades.id) " _
                    & " ON pro_produccion.id = pro_producciondet.id Where (((pro_produccion.id) = " & xRs("id") & ")) ORDER BY pro_produccion.id DESC", xCon

                If Rst.RecordCount <> 0 Then
                    Rst.MoveFirst
                    For A = 1 To Rst.RecordCount
                        Fg1.Rows = Fg1.Rows + 1
                        Fg1.TextMatrix(A, 1) = Rst("descripcion")
                        Fg1.TextMatrix(A, 2) = Rst("codrec")
                        Fg1.TextMatrix(A, 3) = Rst("abrev")
                        Fg1.TextMatrix(A, 4) = Format(Rst("can"), "0.00")
                        
                        Fg1.TextMatrix(A, 5) = Rst("iditem")
                        Fg1.TextMatrix(A, 6) = Rst("idrec")
                        Fg1.TextMatrix(A, 7) = Rst("idunimed")
                        Fg1.TextMatrix(A, 8) = Rst("id")
                        AgregarItems Rst("idrec"), Rst("id"), Rst("iditem")
                        Rst.MoveNext
                        If Rst.EOF = True Then
                            Exit For
                        End If
                        
                    Next A
                    Rst.MoveFirst
                    
                    MuestraItems Rst("idrec"), Rst("id"), Rst("can")
                End If
                Agregando = False
            End If
        End If
        
        Set xForm = Nothing
        Set xRs = Nothing
End Sub

Sub MuestraItems(IdReceta As Integer, IdOrdenProduccion As Integer, CantidadProducir As Double)
    Dim A As Integer
    
    RstDetIns.Filter = adFilterNone
    RstDetIns.Filter = "idrec = " & IdReceta & " AND idorden = " & IdOrdenProduccion & ""
    
    Fg2.Rows = 1
    If RstDetIns.RecordCount <> 0 Then
        RstDetIns.MoveFirst
        For A = 1 To RstDetIns.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(A, 1) = RstDetIns("descripcion")
            Fg2.TextMatrix(A, 2) = RstDetIns("tipoinsumo")
            Fg2.TextMatrix(A, 3) = RstDetIns("abrev")
            Fg2.TextMatrix(A, 4) = Format(RstDetIns("cantidad"), "0.0000")
            Fg2.TextMatrix(A, 5) = RstDetIns("cantidad") * CantidadProducir
            Fg2.TextMatrix(A, 5) = Format(Fg2.TextMatrix(A, 5), "0.0000")
            
            Fg2.TextMatrix(A, 7) = RstDetIns("iditem")
            Fg2.TextMatrix(A, 8) = RstDetIns("idunimed")
            
            RstDetIns.MoveNext
            If RstDetIns.EOF = True Then
                Exit For
            End If
        Next A
    End If
End Sub

Sub AgregarItems(IdReceta As Integer, IdOrden As Integer, IdProducto As Integer)
    Dim A As Integer
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT pro_receta.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro, pro_recetains.iditem, " _
        & " pro_recetains.idunimed, mae_tipoproducto.descripcion AS destippro FROM mae_tipoproducto INNER JOIN (mae_unidades RIGHT JOIN (alm_inventario " _
        & " INNER JOIN (pro_receta INNER JOIN pro_recetains ON pro_receta.id = pro_recetains.idrec) ON alm_inventario.id = pro_recetains.iditem) " _
        & " ON mae_unidades.id = pro_recetains.idunimed) ON mae_tipoproducto.id = alm_inventario.tippro Where (pro_receta.id = " & IdReceta & ") " _
        & " ORDER BY alm_inventario.descripcion", xCon

    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            RstDetIns.AddNew
            RstDetIns("descripcion") = Rst("descripcion")
            RstDetIns("tipoinsumo") = Rst("destippro")
            RstDetIns("abrev") = Rst("abrev")
            RstDetIns("cantidad") = Rst("canpro")
            RstDetIns("iditem") = Rst("iditem")
            RstDetIns("idunimed") = Rst("idunimed")
            RstDetIns("idrec") = IdReceta
            RstDetIns("idorden") = IdOrden
            RstDetIns("idprod") = IdProducto
            Rst.MoveNext
            
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
    End If
End Sub

Sub PreparaRST()
    Dim xFun As New eps_librerias.FuncionesData
    
    Dim xCampos(9, 3) As String

    xCampos(0, 0) = "iditem":       xCampos(0, 1) = "N":      xCampos(0, 2) = "8"    '* codigo del insumo
    xCampos(1, 0) = "idunimed":     xCampos(1, 1) = "N":      xCampos(1, 2) = "8"    '* codigo de la unidad de medida
    xCampos(2, 0) = "descripcion":  xCampos(2, 1) = "C":      xCampos(2, 2) = "100"  '* descripcion
    xCampos(3, 0) = "tipoinsumo":   xCampos(3, 1) = "C":      xCampos(3, 2) = "20"   '* tipo de insumo
    xCampos(4, 0) = "abrev":        xCampos(4, 1) = "C":      xCampos(4, 2) = "10"   '* abreviatura
    xCampos(5, 0) = "cantidad":     xCampos(5, 1) = "D":      xCampos(5, 2) = "8"    '* cantidad
    xCampos(6, 0) = "idrec":        xCampos(6, 1) = "N":      xCampos(6, 2) = "8"    ' codigo de la receta                      (tabla pro_producciondet)
    xCampos(7, 0) = "idorden":      xCampos(7, 1) = "N":      xCampos(7, 2) = "8"    ' codigo de la orden de produccion         (tabla pro_producciondet)
    xCampos(8, 0) = "idprod":       xCampos(8, 1) = "N":      xCampos(8, 2) = "8"    ' codigo del producto que se esta hacieno  (tabla pro_producciondet)
    
    Set RstDetIns = xFun.CrearRstTMP(xCampos)
    RstDetIns.Open
End Sub

Sub Modificar()
    QueHace = 2
    ActivaTool
    Blanquea
    Bloquea
    Label1.Caption = "Modificando Solicitud de Materiales"
    TabOne1.CurrTab = 1
    Fg1.Rows = 1
    Fg2.Rows = 1
    TabOne1.TabEnabled(0) = False
    PreparaRST
    TxtFchPed.SetFocus
End Sub

Sub MuestraSegundoTab()
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    
    TxtFchPed.Valor = RstOrd("fchped")
    TxtTipDoc.Text = RstOrd("idtipdoc")
    LblTipDocumento.Caption = RstOrd("descdoc")
    TxtNumSer.Text = RstOrd("serord")
    TxtNumDoc.Text = RstOrd("numord")
    TxtIdArea.Text = RstOrd("idare")
    LblArea.Caption = RstOrd("descripcion")
    TxtSolicitante.Text = RstOrd("apenomsol")
    LblIdSol.Caption = RstOrd("idres")
    Fg1.Rows = 1
    Fg2.Rows = 1
    
    RST_Busq Rst, "SELECT DISTINCT pro_producciondet.id, pro_producciondet.idrec, pro_producciondet.iditem, alm_inventario.descripcion AS despro, pro_receta.codrec, " _
        & " mae_unidades.abrev, pro_producciondet.can, pro_producciondet.idunimed FROM (((pro_producciondet RIGHT JOIN pro_ordensalidadet " _
        & " ON (pro_producciondet.iditem = pro_ordensalidadet.iditem) AND (pro_producciondet.idrec = pro_ordensalidadet.idrec) AND (pro_producciondet.id = pro_ordensalidadet.idpro)) " _
        & " LEFT JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id) " _
        & " LEFT JOIN mae_unidades ON pro_producciondet.idunimed = mae_unidades.id WHERE (((pro_ordensalidadet.id)=" & RstOrd("id") & "))", xCon

    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        Agregando = True
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = Rst("despro")
            Fg1.TextMatrix(A, 2) = Rst("codrec")
            Fg1.TextMatrix(A, 3) = Rst("abrev")
            Fg1.TextMatrix(A, 4) = Rst("can")
            Fg1.TextMatrix(A, 5) = Rst("iditem")
            Fg1.TextMatrix(A, 6) = Rst("idrec")
            Fg1.TextMatrix(A, 7) = Rst("idunimed")
            Fg1.TextMatrix(A, 8) = Rst("id")
            
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
        Agregando = False
    End If
    
    Set Rst = Nothing
    
    PreparaRST
    
    RST_Busq Rst, "SELECT pro_ordensalidadet.id, alm_inventario.descripcion, mae_unidades.abrev, mae_tipoproducto.descripcion AS destippro, pro_ordensalidadet.canteo, " _
        & " pro_ordensalidadet.idins, pro_ordensalidadet.idunimed, pro_ordensalidadet.idpro, pro_ordensalidadet.idrec, pro_ordensalidadet.iditem" _
        & " FROM mae_tipoproducto RIGHT JOIN ((pro_ordensalidadet LEFT JOIN alm_inventario ON pro_ordensalidadet.idins = alm_inventario.id) LEFT JOIN mae_unidades " _
        & " ON pro_ordensalidadet.idunimed = mae_unidades.id) ON mae_tipoproducto.id = alm_inventario.tippro WHERE (((pro_ordensalidadet.id)=" & RstOrd("id") & "))", xCon

    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            RstDetIns.AddNew
            RstDetIns("iditem") = Rst("idins")
            RstDetIns("idunimed") = Rst("idunimed")
            RstDetIns("descripcion") = Rst("descripcion")
            RstDetIns("tipoinsumo") = Rst("destippro")
            RstDetIns("abrev") = Rst("abrev")
            RstDetIns("cantidad") = Rst("canteo")
            RstDetIns("idrec") = Rst("idrec")
            RstDetIns("idorden") = Rst("idpro")
            RstDetIns("idprod") = Rst("iditem")
            
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
        
        Rst.MoveFirst
        MuestraItems Val(Fg1.TextMatrix(1, 6)), Val(Fg1.TextMatrix(1, 8)), Val(Fg1.TextMatrix(1, 4))
        
    End If
    
'    xCampos(0, 0) = "iditem":       xCampos(0, 1) = "N":      xCampos(0, 2) = "8"    '* codigo del insumo
'    xCampos(1, 0) = "idunimed":     xCampos(1, 1) = "N":      xCampos(1, 2) = "8"    '* codigo de la unidad de medida
'    xCampos(2, 0) = "descripcion":  xCampos(2, 1) = "C":      xCampos(2, 2) = "100"  '* descripcion
'    xCampos(3, 0) = "tipoinsumo":   xCampos(3, 1) = "C":      xCampos(3, 2) = "20"   '* tipo de insumo
'    xCampos(4, 0) = "abrev":        xCampos(4, 1) = "C":      xCampos(4, 2) = "10"   '* abreviatura
'    xCampos(5, 0) = "cantidad":     xCampos(5, 1) = "D":      xCampos(5, 2) = "8"    '* cantidad
'    xCampos(6, 0) = "idrec":        xCampos(6, 1) = "N":      xCampos(6, 2) = "8"    ' codigo de la receta                      (tabla pro_producciondet)
'    xCampos(7, 0) = "idorden":      xCampos(7, 1) = "N":      xCampos(7, 2) = "8"    ' codigo de la orden de produccion         (tabla pro_producciondet)
'    xCampos(8, 0) = "idprod":       xCampos(8, 1) = "N":      xCampos(8, 2) = "8"    ' codigo del producto que se esta hacieno  (tabla pro_producciondet)
    
End Sub

Sub Bloquea()
    TxtFchPed.Locked = Not TxtFchPed.Locked
    TxtTipDoc.Locked = Not TxtTipDoc.Locked
    TxtNumSer.Locked = Not TxtNumSer.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    TxtIdArea.Locked = Not TxtIdArea.Locked
    
    CmdBusSol.Enabled = Not CmdBusSol.Enabled
    CmdBusDoc.Enabled = Not CmdBusDoc.Enabled
    CmdBusAre.Enabled = Not CmdBusAre.Enabled
End Sub

Sub Blanquea()
    TxtFchPed.Valor = ""
    TxtTipDoc.Text = ""
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    TxtIdArea.Text = ""
    TxtSolicitante.Text = ""
    LblIdSol.Caption = ""
    LblTipDocumento.Caption = ""
    LblArea.Caption = ""
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
    
    Toolbar1.Buttons(12).Enabled = Not Toolbar1.Buttons(12).Enabled
    Toolbar1.Buttons(14).Enabled = Not Toolbar1.Buttons(14).Enabled
End Sub

Private Sub CmdBusAre_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":     xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":          xCampos(1, 1) = "id":               xCampos(1, 2) = "1500":         xCampos(1, 3) = "C"
    
    xForm.SQLCad = "SELECT pla_area.id, pla_area.descripcion From pla_area ORDER BY pla_area.descripcion"

    xForm.Titulo = "Buscando Areas"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "descripcion"
    xForm.CampoBusca = "descripcion"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdArea.Text = xRs("id")
            LblArea.Caption = xRs("descripcion")
            TxtSolicitante.SetFocus
        End If
    End If
    
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusDoc_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(3, 4) As String
    
    xCampos(0, 0) = "Descripcion":     xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Abrev":           xCampos(1, 1) = "abrev":            xCampos(1, 2) = "1200":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Codigo":          xCampos(2, 1) = "id":               xCampos(2, 2) = "1200":         xCampos(2, 3) = "C"
    
    xForm.SQLCad = "SELECT mae_documento.abrev, mae_documento.descripcion, mae_documento.id, mae_documento.tipo From mae_documento WHERE (((mae_documento.tipo)=0))"


    xForm.Titulo = "Buscando Areas"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "descripcion"
    xForm.CampoBusca = "descripcion"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtTipDoc.Text = xRs("id")
            LblTipDocumento.Caption = xRs("descripcion")
            TxtTipDoc.SetFocus
        End If
    End If
    
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusSol_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(3, 4) As String
    
    xCampos(0, 0) = "Apellido Nombre":    xCampos(0, 1) = "apenom":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Cargo":              xCampos(1, 1) = "descar":      xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Codigo":             xCampos(2, 1) = "id":          xCampos(2, 2) = "1000":         xCampos(2, 3) = "N"
    
    xForm.SQLCad = "SELECT pla_empleados.id, UCase([pla_empleados]![ape])+', '+[pla_empleados]![nom] AS apenom, pla_cargos.descripcion AS descar " _
        & " FROM pla_cargos INNER JOIN pla_empleados ON pla_cargos.id = pla_empleados.idcargo " _
        & " ORDER BY UCase([pla_empleados]![ape])+', '+[pla_empleados]![nom]"

    xForm.Titulo = "Buscando Personal"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "apenom"
    xForm.CampoBusca = "apenom"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtSolicitante.Text = xRs("apenom")
            LblIdSol.Caption = xRs("id")
            Fg1.SetFocus
        End If
    End If
    
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub Fg1_RowColChange()
    If Agregando = True Then Exit Sub
    MuestraItems Val(Fg1.TextMatrix(Fg1.Row, 6)), Val(Fg1.TextMatrix(Fg1.Row, 8)), Val(Fg1.TextMatrix(Fg1.Row, 4))
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        Dim Rpta As Integer
        RST_Busq RstOrd, "SELECT DISTINCT pro_ordensalida.id, [pro_ordensalida]![serord]+'-'+[pro_ordensalida]![numord] AS numdoc, mae_documento.abrev AS abredoc, " _
            & " pla_area.descripcion, pro_ordensalida.fchped, UCase([pla_empleados]![ape])+', '+[pla_empleados]![nom] AS apenomsol, alm_inventario.descripcion AS descpro, " _
            & " pro_receta.codrec, mae_unidades.abrev, pro_producciondet.can, pro_ordensalida.idtipdoc, mae_documento.descripcion AS descdoc, pro_ordensalida.serord, " _
            & " pro_ordensalida.numord, pro_ordensalida.idare, pro_ordensalida.idres " _
            & " FROM (pro_producciondet LEFT JOIN mae_unidades ON pro_producciondet.idunimed = mae_unidades.id) RIGHT JOIN ((((pro_ordensalida LEFT JOIN mae_documento " _
            & " ON pro_ordensalida.idtipdoc = mae_documento.id) LEFT JOIN pla_area ON pro_ordensalida.idare = pla_area.id) LEFT JOIN pla_empleados " _
            & " ON pro_ordensalida.idres = pla_empleados.id) LEFT JOIN ((pro_ordensalidadet LEFT JOIN alm_inventario ON pro_ordensalidadet.iditem = alm_inventario.id) " _
            & " LEFT JOIN pro_receta ON pro_ordensalidadet.idrec = pro_receta.id) ON pro_ordensalida.id = pro_ordensalidadet.id) ON (pro_producciondet.iditem = pro_ordensalidadet.iditem) " _
            & " AND (pro_producciondet.idrec = pro_ordensalidadet.idrec) AND (pro_producciondet.id = pro_ordensalidadet.idpro) " _
            & " ORDER BY [pro_ordensalida]![serord]+'-'+[pro_ordensalida]![numord] DESC", xCon

        Set Dg1.DataSource = RstOrd
        Dg1.Refresh
        
        If RstOrd.RecordCount = 0 Then
            Rpta = MsgBox("No se han registrado solicitud de materiales ¿ Desea agregar uno ahora ?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                Nuevo
            Else
                Set RstOrd = Nothing
                Unload Me
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    QueHace = 3
    SeEjecuto = False
    Fg1.Editable = flexEDNone
    Fg1.SelectionMode = flexSelectionByRow
    
    Fg2.Editable = flexEDNone
    Fg2.SelectionMode = flexSelectionByRow
    TabOne1.CurrTab = 0
    
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    
    CaracteresNumericos = "0123456789." & Chr(8)
    
    Fg1.ColWidth(5) = 0
    Fg1.ColWidth(6) = 0
    Fg1.ColWidth(7) = 0
    Fg1.ColWidth(8) = 0
    
    Fg2.ColWidth(6) = 0
    Fg2.ColWidth(7) = 0
    Fg2.ColWidth(8) = 0
End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then
            MuestraSegundoTab
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
            RstOrd.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then
        Cancelar
    End If
    
    If Button.Index = 11 Then
        Set RstOrd = Nothing
        Unload Me
    End If
End Sub

Sub Cancelar()
    QueHace = 3
    ActivaTool
    Bloquea
    Label1.Caption = "Detalle de la Solicitud"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
End Sub

Sub Eliminar()

End Sub

Private Sub TxtIdArea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdArea_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusSol_Click
    End If
End Sub

Private Sub TxtIdArea_Validate(Cancel As Boolean)
    If NulosN(TxtIdArea.Text) <> 0 Then
        LblArea.Caption = Busca_Codigo(Val(TxtIdArea.Text), "id", "descripcion", "pla_area", "N", xCon)
        If LblArea.Caption = "" Then
            TxtIdArea.Text = ""
        End If
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
    If TxtNumDoc.Text <> "" Then
        TxtNumDoc.Text = Format(TxtNumDoc.Text, "00000000")
    End If
End Sub

Private Sub TxtNumSer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumSer_Validate(Cancel As Boolean)
    If TxtNumSer.Text <> "" Then
        TxtNumSer.Text = Format(TxtNumSer.Text, "0000")
    End If
End Sub

Private Sub TxtSolicitante_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtSolicitante_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusSol_Click
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
        CmdBusDoc_Click
    End If
End Sub

Private Sub TxtTipDoc_Validate(Cancel As Boolean)
    If NulosN(TxtTipDoc.Text) <> 0 Then
        If Val(Busca_Codigo(Val(TxtTipDoc.Text), "id", "tipo", " mae_documento", "N", xCon)) = 1 Then
            MsgBox "Documento no valido para esta operacion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtTipDoc.Text = ""
            LblTipDocumento.Caption = ""
            Exit Sub
        End If
        LblTipDocumento.Caption = Busca_Codigo(Val(TxtTipDoc.Text), "id", "descripcion", " mae_documento", "N", xCon)
        If NulosC(LblTipDocumento.Caption) = "" Then
            TxtTipDoc.Text = ""
        End If
    End If
End Sub

VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmComision 
   Caption         =   "Planillas - Comisión a Vendedores"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7215
      Left            =   0
      TabIndex        =   7
      Top             =   360
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
         Height          =   6795
         Left            =   45
         TabIndex        =   12
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6420
            Left            =   30
            TabIndex        =   21
            Top             =   375
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
            Columns(1).Caption=   "Fch. Com."
            Columns(1).DataField=   "fchcom"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Vendedor"
            Columns(2).DataField=   "apenom"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Nº Documentos"
            Columns(3).DataField=   "numdoc"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Imp. Fact."
            Columns(4).DataField=   "imptot"
            Columns(4).NumberFormat=   "0.00"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Imp. Cob."
            Columns(5).DataField=   "impabo"
            Columns(5).NumberFormat=   "0.00"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "% Comisión"
            Columns(6).DataField=   "marcom"
            Columns(6).NumberFormat=   "0.00"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Comisión"
            Columns(7).DataField=   "comision"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2037"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1958"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=7038"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=6959"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=2514"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2434"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=2090"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2011"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=514"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=2117"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=2037"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=514"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=1905"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1826"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=513"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=2011"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=1931"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=514"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=66,.parent=13,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=1"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=46,.parent=13,.alignment=1"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=43,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=44,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=45,.parent=17"
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
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Comisión de Vendedores"
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
            TabIndex        =   14
            Top             =   45
            Width           =   11595
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label lblperiodo 
            Caption         =   "lblperiodo(0)"
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
            Height          =   300
            Index           =   0
            Left            =   9450
            TabIndex        =   13
            Top             =   75
            Width           =   1980
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   6795
         Left            =   12525
         TabIndex        =   8
         Top             =   375
         Width           =   11790
         Begin VB.Frame Frame3 
            Height          =   840
            Left            =   9345
            TabIndex        =   20
            Top             =   945
            Width           =   2430
            Begin VB.CommandButton CmdImprimir 
               Height          =   570
               Left            =   900
               Picture         =   "FrmComision.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   4
               ToolTipText     =   "Imprimir"
               Top             =   180
               Width           =   630
            End
            Begin VB.CommandButton CmdExp 
               Height          =   570
               Left            =   1560
               Picture         =   "FrmComision.frx":030A
               Style           =   1  'Graphical
               TabIndex        =   5
               ToolTipText     =   "Exportar a Excel"
               Top             =   180
               Width           =   630
            End
            Begin VB.CommandButton CmdCargar 
               Enabled         =   0   'False
               Height          =   570
               Left            =   225
               Picture         =   "FrmComision.frx":0E14
               Style           =   1  'Graphical
               TabIndex        =   3
               ToolTipText     =   "Cargar Documentos"
               Top             =   180
               Width           =   630
            End
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
            Height          =   300
            Left            =   1305
            TabIndex        =   0
            Top             =   555
            Width           =   1290
            _ExtentX        =   2275
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
         Begin VB.TextBox TxtComision 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1305
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   2
            Text            =   "TxtComisio"
            Top             =   1200
            Width           =   915
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   4830
            Left            =   60
            TabIndex        =   6
            Top             =   1860
            Width           =   11700
            _cx             =   20637
            _cy             =   8520
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
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmComision.frx":1256
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
         Begin VB.CommandButton CmdBusProv 
            Height          =   240
            Left            =   1950
            Picture         =   "FrmComision.frx":13B5
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   900
            Width           =   240
         End
         Begin VB.TextBox TxtidVen 
            Height          =   300
            Left            =   1305
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   1
            Text            =   "TxtidVen"
            Top             =   870
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Documentos Emitidos y Cobrados"
            Height          =   195
            Left            =   90
            TabIndex        =   19
            Top             =   1590
            Width           =   2370
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha "
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   18
            Top             =   585
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Comisión"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   17
            Top             =   1230
            Width           =   630
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
            Left            =   2280
            TabIndex        =   16
            Top             =   870
            Width           =   5055
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Comisión"
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
            Left            =   90
            TabIndex        =   11
            Top             =   45
            Width           =   11610
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   10
            Top             =   900
            Width           =   690
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8790
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
               Picture         =   "FrmComision.frx":14E7
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmComision.frx":1A2B
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmComision.frx":1DBD
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmComision.frx":1F41
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmComision.frx":2395
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmComision.frx":24AD
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmComision.frx":29F1
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmComision.frx":2F35
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmComision.frx":3049
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmComision.frx":315D
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmComision.frx":35B1
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmComision.frx":371D
               Key             =   "IMG11"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmComision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim RstCom As New ADODB.Recordset

Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO


Function Grabar() As Boolean
    If TxtFecha.Valor = "" Then
        MsgBox "No ha especificado la fecha de vencimiento del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFecha.SetFocus
        Grabar = False
        Exit Function
    End If
    
    If NulosN(TxtidVen.Text) = 0 Then
        MsgBox "No ha especificado el vendedor", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtidVen.SetFocus
        Grabar = False
        Exit Function
    End If

    If Fg1.Rows = 1 Then
        MsgBox "No se ha especificado que documentos de venta se estan comisionando", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        CmdCargar.SetFocus
        Grabar = False
        Exit Function
    End If
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xId As Double
    Dim A As Integer
    
    xCon.BeginTrans
    
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT TOP 1 * FROM pla_comision", xCon
        RST_Busq RstDet, "SELECT TOP 1 * FROM pla_comisiondet", xCon
        xId = HallaCodigoTabla("pla_comision", xCon, "id")
        RstCab.AddNew
        RstCab("id") = xId
    Else
    
    End If
    
    RstCab("idven") = NulosN(TxtidVen.Text)
    RstCab("fchcom") = NulosC(TxtFecha.Valor)
    RstCab("comision") = NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 10))
    RstCab("numdoc") = Fg1.Rows - 2
    RstCab("imptot") = NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 6))
    RstCab("marcom") = NulosN(TxtComision.Text)
    RstCab.Update
    
    For A = 1 To Fg1.Rows - 2
        RstDet.AddNew
        RstDet("id") = xId
        RstDet("iddoc") = NulosN(Fg1.TextMatrix(A, 11))
        RstDet("impcom") = NulosN(Fg1.TextMatrix(A, 10))
        
        xCon.Execute "UPDATE vta_ventas SET vta_ventas.comision = -1 WHERE (((vta_ventas.id)=" & NulosN(Fg1.TextMatrix(A, 11)) & "));"

        RstDet.Update
    Next A
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId

    
    MsgBox "La Comisión del vendedor " + LblVendedor.Caption + " se guardó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    xCon.CommitTrans
    
    Set RstCab = Nothing
    Set RstDet = Nothing
    Grabar = True
    
    Exit Function
    
LaCague:
    Me.MousePointer = vbDefault
    xCon.RollbackTrans
    MsgBox "No se pudo guardar la Comisión por el siguiente motivo: " + vbCr + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set RstCab = Nothing
    Set RstDet = Nothing
End Function

Sub Blanquea()
    TxtFecha.Valor = ""
    TxtidVen.Text = ""
    TxtComision.Text = ""
    LblVendedor.Caption = ""
    Fg1.Rows = 1
End Sub

Sub Bloquea()
    TxtFecha.Locked = Not TxtFecha.Locked
    TxtidVen.Locked = Not TxtidVen.Locked
    CmdCargar.Enabled = Not CmdCargar.Enabled
    'TxtComision.Locked = Not TxtComision.Locked
    LblVendedor.Caption = ""
End Sub

Private Sub CmdBusProv_Click()
    If QueHace = 3 Then Exit Sub
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Nombre Vendedor":   xCampos(0, 1) = "apenom":     xCampos(0, 2) = "4500":      xCampos(0, 3) = "C"
    xCampos(1, 0) = "Comisión":          xCampos(1, 1) = "comision":   xCampos(1, 2) = "500":      xCampos(1, 3) = "N"

    CARGAR_DLL_EPSBUSCAR xCon, xRs, "SELECT vta_vendedores.id, vta_vendedores.basico, vta_vendedores.comision, UCase([pla_empleados]![apepat] & ' ' & [pla_empleados]![apemat] & ' ' &  [pla_empleados]![nom]) AS apenom" _
        & " FROM vta_vendedores LEFT JOIN pla_empleados ON vta_vendedores.idper = pla_empleados.id", xCampos(), "Buscando Venededor", "apenom", "apenom", Principio

    If xRs.State = 0 Then GoTo salir
    If xRs.RecordCount = 0 Then GoTo salir
    TxtidVen.Text = xRs("id") & ""
    LblVendedor.Caption = xRs("apenom") & ""
    TxtComision.Text = Format(xRs("comision"), "0.00")
    
salir:
    Set xRs = Nothing
End Sub

Private Sub CmdCargar_Click()
    If IsDate(TxtFecha.Valor) = False Then
        MsgBox "Falta especificar la fecha de Liquidación", vbExclamation, xTitulo
        TxtFecha.SetFocus
        Exit Sub
    End If
    
    If NulosN(TxtidVen.Text) = 0 Then
        MsgBox "No ha especificado el vendedor", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    CargarCobrados NulosN(TxtidVen.Text)
End Sub

Private Sub CmdExp_Click()
    If Fg1.Rows = 1 Then
        MsgBox "No se han especificado registros para exportar, haga clic en el botón [Cargar Documentos]", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Dim xMiFun As New SGI2_funciones.formularios
    xMiFun.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "Comisión de Vendedores", "Cobranzas hasta el : " + NulosC(TxtFecha.Valor), "Vendedor : " + Trim(LblVendedor.Caption)
    Set xMiFun = Nothing
    
End Sub

Private Sub CmdImprimir_Click()
    If Fg1.Rows = 1 Then
        MsgBox "No se han especificado registros para exportar, haga clic en el botón [Cargar Documentos]", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If

    Dim xFun As New SGI2_funciones.formularios
    xFun.Imprimir_x_VSFlexGrid Fg1, "Comisión de Vendedores", "Vendedor : " + NulosC(LblVendedor.Caption), "Hasta el : " + NulosC(TxtFecha.Valor), True, True
    Set xFun = Nothing
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstCom.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstCom("id")), xCon
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
    
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
        'RST_Busq RstCom, "SELECT pla_comision.*, UCase([pla_empleados]![apepat] & ' ' & [pla_empleados]![apemat] & ' ' & [pla_empleados]![nom]) AS apenom, " _
            & " (SELECT Sum([impabo]) AS total FROM pla_comisiondet LEFT JOIN (con_cajabanco RIGHT JOIN con_cajabancodet ON con_cajabanco.id = con_cajabancodet.id) " _
            & " ON pla_comisiondet.iddoc = con_cajabancodet.iddoc WHERE (((pla_comisiondet.id)=pla_comision.id) AND ((con_cajabanco.tipmov)=1))) AS impabo " _
            & " FROM (pla_comision LEFT JOIN vta_vendedores ON pla_comision.idven = vta_vendedores.id) LEFT JOIN pla_empleados ON vta_vendedores.idper = pla_empleados.id " _
            & " ORDER BY pla_comision.fchcom DESC", xCon
        
        RST_Busq RstCom, "SELECT pla_comision.*, UCase([pla_empleados]![apepat] & ' ' & [pla_empleados]![apemat] & ' ' & [pla_empleados]![nom]) AS apenom, " _
            & " 0 AS impabo FROM (pla_comision LEFT JOIN vta_vendedores ON pla_comision.idven = vta_vendedores.id) LEFT JOIN pla_empleados " _
            & " ON vta_vendedores.idper = pla_empleados.id ORDER BY pla_comision.fchcom DESC", xCon

        Set Dg1.DataSource = RstCom

    End If
End Sub

Sub MuestraSegundoTab()
    Dim A As Integer
    
    Fg1.Rows = 1
    TxtFecha.Valor = CDate(RstCom("fchcom"))
    TxtidVen.Text = NulosN(RstCom("idven"))
    TxtComision.Text = Format(NulosN(RstCom("marcom")), "0.00")
    LblVendedor.Caption = NulosC(RstCom("apenom"))
    
    Dim Rst As New ADODB.Recordset
    Dim xTotal, xTotalDoc, xTotalCob As Double

    RST_Busq Rst, "SELECT DISTINCT mae_cliente.id, mae_cliente.nombre, mae_documento.abrev, Left([vta_ventas].[numreg],2) & IIf([mae_libros].[codsun] Is " _
        & " Null,'--',[mae_libros].[codsun]) & Right([vta_ventas].[numreg],4) AS registro, vta_ventas.numreg, mae_libros.codsun, " _
        & " vta_ventas!numser+'-'+vta_ventas!numdoc AS numdoc, vta_ventas.fchdoc, vta_ventas.imptotdoc, pla_comisiondet.iddoc, tes_caja.fchope, " _
        & " tes_cajadestinodet.acuenta AS impabo FROM (tes_caja RIGHT JOIN tes_cajadestino ON tes_caja.id = tes_cajadestino.idtes) RIGHT JOIN (((pla_comisiondet " _
        & " LEFT JOIN ((mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) " _
        & " ON pla_comisiondet.iddoc = vta_ventas.id) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN tes_cajadestinodet " _
        & " ON vta_ventas.id = tes_cajadestinodet.iddoc) ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes) " _
        & " Where (((tes_caja.tipmov) = 1) And ((pla_comisiondet.id) = " & RstCom("id") & ")) ORDER BY mae_cliente.nombre, vta_ventas.fchdoc", xCon

    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        Do While Not Rst.EOF
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(Rst("registro"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(Rst("nombre"))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(Rst("abrev"))
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(Rst("numdoc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(NulosC(Rst("fchdoc")), FORMAT_DATE)
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(NulosN(Rst("imptotdoc")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(NulosC(Rst("fchope")), FORMAT_DATE)
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(NulosN(Rst("impabo")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = "0.00"
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = (NulosN(Rst("impabo")) * (NulosN(TxtComision.Text) / 100))
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 10)), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosN(Rst("iddoc"))
            
            xTotal = xTotal + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 10))
            xTotalDoc = xTotalDoc + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 6))
            xTotalCob = xTotalCob + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 8))
                    
            Rst.MoveNext
        Loop
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(xTotalDoc, FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(xTotalCob, FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(xTotal, FORMAT_MONTO)
        FormatearCeldas
    End If
End Sub

Private Sub Form_Load()
    Fg1.Rows = 1
    SeEjecuto = False
    TabOne1.CurrTab = 0
    Fg1.ColWidth(11) = 0
    Fg1.SelectionMode = flexSelectionByRow
    QueHace = 3
End Sub

Sub CargarCobrados(IdVendedor As Integer)
    Dim xRst As New ADODB.Recordset
    Dim A&
    Dim xTotal, xTotalDoc, xTotalCob As Double
    
    Fg1.Rows = 1
    DoEvents
    RST_Busq xRst, "SELECT Left([vta_ventas].[numreg],2) & IIf([mae_libros].[codsun] Is Null,'--',[mae_libros].[codsun]) & Right([vta_ventas].[numreg],4) AS registro, " _
        & " mae_cliente.nombre, vta_ventas!numser+'-'+vta_ventas!numdoc AS NUMDOC, vta_ventas.imptotdoc, vta_ventas.fchdoc, vta_ventas.impsal, mae_documento.abrev, " _
        & " tes_caja.fchope, vta_ventas.id, vta_ventas.idven, tes_cajadestinodet.acuenta AS impabo, tes_caja.tipmov " _
        & " FROM (tes_caja RIGHT JOIN tes_cajadestino ON tes_caja.id = tes_cajadestino.idtes) RIGHT JOIN ((((mae_cliente RIGHT JOIN vta_ventas " _
        & " ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_libros " _
        & " ON vta_ventas.idlib = mae_libros.id) LEFT JOIN tes_cajadestinodet ON vta_ventas.id = tes_cajadestinodet.iddoc) ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) " _
        & " AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes) Where (((vta_ventas.impsal) = 0) And ((vta_ventas.idven) = 1) And ((tes_caja.tipmov) = 1) " _
        & " And ((vta_ventas.anulado) = False) And ((vta_ventas.Comision) = 0)) and tes_caja.fchope <= cdate('" & TxtFecha.Valor & "') ORDER BY mae_cliente.nombre, vta_ventas.fchdoc", xCon

    If xRst.RecordCount <> 0 Then
        xRst.MoveFirst
        Do While Not xRst.EOF
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(xRst("registro"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(xRst("nombre"))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(xRst("abrev"))
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(xRst("numdoc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(xRst("fchdoc"), FORMAT_DATE)
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(xRst("imptotdoc"), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(xRst("fchope"), FORMAT_DATE)
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(xRst("impabo"), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(xRst("impsal"), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = NulosN(xRst("impabo")) * ((NulosN(TxtComision.Text)) / 100)
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(Fg1.TextMatrix(Fg1.Rows - 1, 10), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = xRst("id")
            
            xTotal = xTotal + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 10))
            xTotalDoc = xTotalDoc + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 6))
            xTotalCob = xTotalCob + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 8))

            xRst.MoveNext
        Loop
        Fg1.Rows = Fg1.Rows + 1
        
        Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(xTotalDoc, FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(xTotalCob, FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(xTotal, FORMAT_MONTO)
        FormatearCeldas
    Else
        MsgBox "No se han encontrado ventas cobradas para el vendedor : " + LblVendedor.Caption, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Set xRst = Nothing
End Sub

Sub FormatearCeldas()
    UNIR_CELDAS Fg1, Fg1.Rows - 1, 3, Fg1.Rows - 1, 4, "TOTALES ==>", flexAlignCenterTop, True
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 3, , True
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 5, , True
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 7, , True
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 9, , True
End Sub

Sub ActivaTool()
    Dim A As Integer
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub


Sub Nuevo()
    QueHace = 1
    xHorIni = Time
    TabOne1.CurrTab = 1
    ActivaTool
    TxtFecha.Valor = Date
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Calculando Comisión para el Vendedor"
    Blanquea
    Bloquea
    
    TxtFecha.SetFocus
End Sub

Sub Modificar()
xHorIni = Time

End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstCom.Requery
            Dg1.Refresh
        End If
    End If
    If Button.Index = 6 Then Cancelar
    If Button.Index = 12 Then
        Set RstCom = Nothing
        Unload Me
    End If
End Sub

Sub Cancelar()
    Bloquea
    ActivaTool
    QueHace = 3
    Label5.Caption = "Detalle de Comisión"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
End Sub

Private Sub TxtidVen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtidVen_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusProv_Click
    End If
End Sub

Private Sub TxtidVen_Validate(Cancel As Boolean)
    If NulosN(TxtidVen.Text) <> 0 Then
        Dim xRs As New ADODB.Recordset
        
        Set xRs = BuscaConCriterio("SELECT vta_vendedores.id, UCase([pla_empleados]![apepat] & ' ' & [pla_empleados]![apemat] & ' ' & [pla_empleados]![nom]) AS apenom, vta_vendedores.comision " _
            & " FROM pla_empleados RIGHT JOIN vta_vendedores ON pla_empleados.id = vta_vendedores.idper", xCon)

        If xRs.RecordCount <> 0 Then
            LblVendedor.Caption = NulosC(xRs("apenom"))
            TxtComision.Text = Format(NulosN(xRs("comision")), FORMAT_MONTO)
        Else
            TxtidVen.Text = ""
        End If
        Set xRs = Nothing
    End If
End Sub

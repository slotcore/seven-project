VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmManOrdCotiza 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compras - Orden de Cotizacion"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   2355
      Left            =   5850
      TabIndex        =   30
      Top             =   2970
      Visible         =   0   'False
      Width           =   5070
      Begin VB.CommandButton CmdBusOrdreq 
         Height          =   240
         Left            =   4185
         Picture         =   "FrmManOrdCotiza.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   570
         Width           =   240
      End
      Begin VB.CommandButton CmdAcepta 
         Caption         =   "&Aceptar"
         Height          =   435
         Left            =   1500
         TabIndex        =   34
         Top             =   1695
         Width           =   1020
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   435
         Left            =   2550
         TabIndex        =   35
         Top             =   1695
         Width           =   1020
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmi2 
         Height          =   300
         Left            =   1725
         TabIndex        =   32
         Top             =   840
         Width           =   1380
         _ExtentX        =   2434
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
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchVen2 
         Height          =   300
         Left            =   1725
         TabIndex        =   33
         Top             =   1140
         Width           =   1380
         _ExtentX        =   2434
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
      End
      Begin VB.TextBox TxtNumOrdReq 
         Height          =   300
         Left            =   1725
         TabIndex        =   31
         Text            =   "TxtNumOrdReq"
         Top             =   540
         Width           =   2730
      End
      Begin VB.Label LblIdReq 
         Caption         =   "LblIdreq"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   3765
         TabIndex        =   41
         Top             =   1020
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   15
         X2              =   5055
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   15
         X2              =   5055
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   5055
         X2              =   5040
         Y1              =   15
         Y2              =   2355
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   15
         Y1              =   -15
         Y2              =   2310
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cotizando Requerimientos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   210
         TabIndex        =   40
         Top             =   105
         Width           =   2220
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nº Requerimiento"
         Height          =   195
         Left            =   165
         TabIndex        =   39
         Top             =   585
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Emision"
         Height          =   195
         Left            =   165
         TabIndex        =   38
         Top             =   885
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Entrega"
         Height          =   195
         Left            =   165
         TabIndex        =   37
         Top             =   1185
         Width           =   915
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   330
         Left            =   60
         Top             =   45
         Width           =   4950
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7230
      Left            =   0
      TabIndex        =   12
      Top             =   360
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
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6810
         Left            =   45
         TabIndex        =   16
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6480
            Left            =   30
            TabIndex        =   17
            Top             =   315
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11430
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
            Columns(1).Caption=   "Documento"
            Columns(1).DataField=   "descdoc"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nº Documento"
            Columns(2).DataField=   "numdoccot"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Fch. Emi."
            Columns(3).DataField=   "fchemi"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Proveedor"
            Columns(4).DataField=   "nombre"
            Columns(4).NumberFormat=   "0.00"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Nº Requerimiento"
            Columns(5).DataField=   "numdocreq"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Estado"
            Columns(6).DataField=   "desest"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   4
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Enviado"
            Columns(7).DataField=   "envmail"
            Columns(7).NumberFormat=   "General Number"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   397
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
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2752"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2672"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=2408"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2328"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=1720"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1640"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=6138"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=6059"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=3598"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=3519"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=1614"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1535"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=1482"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=1402"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=0"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=86,.parent=13,.alignment=0"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=83,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=84,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=85,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=32,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=29,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=30,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=31,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=25,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=26,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=27,.parent=17"
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
            TabIndex        =   20
            Top             =   0
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta Ordenes de Cotizacion"
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
            TabIndex        =   19
            Top             =   30
            Width           =   11610
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
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
            TabIndex        =   18
            Top             =   30
            Visible         =   0   'False
            Width           =   765
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6810
         Left            =   12525
         TabIndex        =   13
         Top             =   375
         Width           =   11790
         Begin VB.CommandButton CmdBusArea 
            Height          =   240
            Left            =   2160
            Picture         =   "FrmManOrdCotiza.frx":0132
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   2400
            Width           =   240
         End
         Begin VB.Frame Frame6 
            Caption         =   "[ Estado ]"
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
            Height          =   1020
            Left            =   8820
            TabIndex        =   64
            Top             =   285
            Width           =   2970
            Begin VB.Label LblEstado 
               Alignment       =   2  'Center
               Caption         =   "LblEstado"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   315
               Left            =   120
               TabIndex        =   65
               Top             =   450
               Width           =   2715
            End
         End
         Begin VB.CommandButton CmdBusMon 
            Height          =   240
            Left            =   6360
            Picture         =   "FrmManOrdCotiza.frx":0264
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   1770
            Width           =   240
         End
         Begin VB.TextBox TxtObs 
            Height          =   540
            Left            =   1635
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Text            =   "FrmManOrdCotiza.frx":0396
            Top             =   2685
            Width           =   10050
         End
         Begin VB.CommandButton Command1 
            Height          =   240
            Left            =   2145
            Picture         =   "FrmManOrdCotiza.frx":039D
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   2085
            Width           =   240
         End
         Begin VB.Frame Frame4 
            Caption         =   "[ Datos de la Cotizacion ]"
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
            Height          =   1110
            Left            =   3165
            TabIndex        =   42
            Top             =   4215
            Visible         =   0   'False
            Width           =   6525
            Begin VB.CommandButton CmdBusCondPag 
               Height          =   240
               Left            =   2025
               Picture         =   "FrmManOrdCotiza.frx":04CF
               Style           =   1  'Graphical
               TabIndex        =   45
               Top             =   660
               Width           =   240
            End
            Begin VB.TextBox TxtNumCot 
               Height          =   300
               Left            =   1515
               Locked          =   -1  'True
               TabIndex        =   43
               Text            =   "TxtNumCot"
               Top             =   315
               Width           =   2295
            End
            Begin VB.TextBox TxtIdCondPag 
               Height          =   300
               Left            =   1515
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   46
               Text            =   "TxtIdCondPag"
               Top             =   630
               Width           =   780
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Condicion Pago"
               Height          =   195
               Left            =   195
               TabIndex        =   48
               Top             =   675
               Width           =   1125
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
               Left            =   2355
               TabIndex        =   47
               Top             =   630
               Width           =   3945
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Nº Cotizacion"
               Height          =   195
               Left            =   195
               TabIndex        =   44
               Top             =   360
               Width           =   960
            End
         End
         Begin VB.CommandButton Command2 
            Height          =   240
            Left            =   3270
            Picture         =   "FrmManOrdCotiza.frx":0601
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   1455
            Width           =   240
         End
         Begin VB.CommandButton Command3 
            Height          =   240
            Left            =   2145
            Picture         =   "FrmManOrdCotiza.frx":0733
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   720
            Width           =   240
         End
         Begin VB.TextBox TxtNumDocReq 
            Height          =   300
            Left            =   2490
            Locked          =   -1  'True
            TabIndex        =   1
            Text            =   "TxtNumDocReq"
            Top             =   375
            Width           =   1365
         End
         Begin VB.TextBox TxtIdDoc 
            Height          =   300
            Left            =   1635
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   2
            Text            =   "TxtIdDoc"
            Top             =   690
            Width           =   780
         End
         Begin VB.TextBox TxtIdpro 
            Height          =   300
            Left            =   1635
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   5
            Text            =   "TxtIdpro"
            Top             =   1425
            Width           =   1905
         End
         Begin VB.TextBox TxtNumSerReq 
            Height          =   300
            Left            =   1635
            Locked          =   -1  'True
            TabIndex        =   0
            Text            =   "TxtNumSerReq"
            Top             =   375
            Width           =   780
         End
         Begin VB.TextBox TxtNumDocCot 
            Height          =   300
            Left            =   2490
            Locked          =   -1  'True
            TabIndex        =   4
            Text            =   "TxtNumDocCot"
            Top             =   1005
            Width           =   1365
         End
         Begin VB.TextBox TxtNumSerCot 
            Height          =   300
            Left            =   1635
            Locked          =   -1  'True
            TabIndex        =   3
            Text            =   "TxtNumSerCot"
            Top             =   1005
            Width           =   780
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   2670
            Left            =   105
            TabIndex        =   11
            Top             =   3480
            Width           =   11535
            _cx             =   20346
            _cy             =   4710
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
            Cols            =   13
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmManOrdCotiza.frx":0865
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
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmi 
            Height          =   300
            Left            =   1635
            TabIndex        =   6
            Top             =   1740
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
         Begin VB.TextBox TxtIdSol 
            Height          =   300
            Left            =   1635
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   8
            Text            =   "TxtIdSol"
            Top             =   2055
            Width           =   780
         End
         Begin VB.TextBox TxtIdMon 
            Height          =   300
            Left            =   5850
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   7
            Text            =   "TxtIdMon"
            Top             =   1740
            Width           =   780
         End
         Begin VB.Frame Frame5 
            Height          =   690
            Left            =   105
            TabIndex        =   58
            Top             =   6105
            Width           =   8520
            Begin VB.CommandButton CmdNewItem 
               Caption         =   "Nuevo Item"
               Height          =   405
               Left            =   3180
               TabIndex        =   61
               Top             =   195
               Width           =   1380
            End
            Begin VB.CommandButton CMdDel 
               Caption         =   "Eliminar Item"
               Height          =   405
               Left            =   1770
               TabIndex        =   60
               Top             =   195
               Width           =   1380
            End
            Begin VB.CommandButton CmdAdd 
               Caption         =   "Agregar Item"
               Height          =   405
               Left            =   360
               TabIndex        =   59
               Top             =   195
               Width           =   1380
            End
         End
         Begin VB.TextBox TxtIdArea 
            Height          =   300
            Left            =   1635
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   9
            Text            =   "TxtIdArea"
            Top             =   2370
            Width           =   780
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Area"
            Height          =   195
            Left            =   135
            TabIndex        =   68
            Top             =   2430
            Width           =   330
         End
         Begin VB.Label LblArea 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblArea"
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
            Left            =   2430
            TabIndex        =   67
            Top             =   2370
            Width           =   3435
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Total =>"
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
            Left            =   9405
            TabIndex        =   63
            Top             =   6210
            Width           =   720
         End
         Begin VB.Label LblTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTotal"
            Height          =   300
            Left            =   10410
            TabIndex        =   62
            Top             =   6165
            Width           =   990
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Left            =   4920
            TabIndex        =   57
            Top             =   1800
            Width           =   585
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
            Left            =   6645
            TabIndex        =   56
            Top             =   1740
            Width           =   2400
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
            Height          =   195
            Left            =   135
            TabIndex        =   54
            Top             =   2745
            Width           =   1065
         End
         Begin VB.Label LblSolicitante 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblSolicitante"
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
            Left            =   2430
            TabIndex        =   53
            Top             =   2055
            Width           =   6615
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Solicitante"
            Height          =   195
            Left            =   135
            TabIndex        =   51
            Top             =   2115
            Width           =   735
         End
         Begin VB.Label LblIdProv 
            AutoSize        =   -1  'True
            Caption         =   "LblIdProv"
            ForeColor       =   &H00800080&
            Height          =   195
            Left            =   9210
            TabIndex        =   50
            Top             =   1470
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label LblProveedor 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblProveedor"
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
            Left            =   3555
            TabIndex        =   29
            Top             =   1425
            Width           =   5490
         End
         Begin VB.Label LblDocumento 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDocumento"
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
            Left            =   2430
            TabIndex        =   28
            Top             =   690
            Width           =   3945
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Documento"
            Height          =   195
            Left            =   135
            TabIndex        =   27
            Top             =   735
            Width           =   825
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Emision"
            Height          =   195
            Left            =   135
            TabIndex        =   26
            Top             =   1800
            Width           =   900
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   135
            TabIndex        =   25
            Top             =   1470
            Width           =   735
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Nº Requerimiento"
            Height          =   195
            Left            =   135
            TabIndex        =   24
            Top             =   435
            Width           =   1245
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Nº Orden Cotizacion"
            Height          =   195
            Left            =   135
            TabIndex        =   23
            Top             =   1050
            Width           =   1440
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Orden de Cotizacion"
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
            TabIndex        =   15
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "[  Lista de Items  ]"
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
            Height          =   195
            Left            =   135
            TabIndex        =   14
            Top             =   3240
            Width           =   1560
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
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdCotiza.frx":09E5
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdCotiza.frx":0F29
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdCotiza.frx":12BB
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdCotiza.frx":143F
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdCotiza.frx":1893
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdCotiza.frx":19AB
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdCotiza.frx":1EEF
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdCotiza.frx":2433
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdCotiza.frx":2547
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdCotiza.frx":265B
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdCotiza.frx":2AAF
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdCotiza.frx":2C1B
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdCotiza.frx":3163
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManOrdCotiza.frx":34F5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   609
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Con Orden de Requerimiento"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Directo"
               EndProperty
            EndProperty
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
            Object.ToolTipText     =   "Enviar por Correo Electronico"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir "
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
Attribute VB_Name = "FrmManOrdCotiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstLista As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim QueHace As Integer
Dim Pagina As Integer
Dim oPDF As cPDF
Dim CaracteresNumericos As String

Public DeDonde As Integer                ' ESPECIFICA DESDE DONDE SE ESTA LLAMANDO AL FORMULARIO
                                         ' 1 = MENU DEL SISTEMAS
                                         ' 2 = OTRO FORMULARIO
Public xIdOR As Integer                  ' ESPECIFICA EL ID DE LA ORDEN QUE SE MOSTRARA CUANDO LA VARIABLE DEDONDE SEA 2
Dim fOrdenLista As Boolean                 ' --especfica el orden de la lista de la consulta
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO

Sub Bloquea()
    TxtNumSerCot.Locked = Not TxtNumSerCot.Locked
    TxtNumDocCot.Locked = Not TxtNumDocCot.Locked
    TxtFchEmi.Locked = Not TxtFchEmi.Locked
    TxtidSol.Locked = Not TxtidSol.Locked
    TxtIdpro.Locked = Not TxtIdpro.Locked
    TxtObs.Locked = Not TxtObs.Locked
    TxtIdMon.Locked = Not TxtIdMon.Locked
    TxtIdArea.Locked = Not TxtIdArea.Locked
End Sub

Sub ActivaTool()
    Dim A As Integer
    For A = 1 To 15
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Sub Cancelar()
    QueHace = 3
    Bloquea
    ActivaTool
    TabOne1.TabEnabled(0) = True
    Label5.Caption = "Detalle de la Orden de Cotizacion"
    Fg1.Editable = flexEDNone
    Fg1.SelectionMode = flexSelectionByRow
    TabOne1.CurrTab = 0
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    
    If RstLista("idest") = 3 Then
        MsgBox "No se puede eliminar una orden de cotizacion procesada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Rpta = MsgBox("Esta seguro de eliminar el registro seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    
    If Rpta = vbYes Then
        If RstLista("tipcot") = 2 Then
            xCon.Execute "UPDATE com_ordenreq SET com_ordenreq.idest = 1 WHERE (((com_ordenreq.id)=" & RstLista("idor") & "))"
        End If
        
        xCon.Execute "DELETE * FROM com_ordencotdet WHERE idoc = " & RstLista("id") & ""
        xCon.Execute "DELETE * FROM com_ordencot WHERE id = " & RstLista("id") & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstLista("id") & " AND idform = " & IdMenuActivo

        
        
        MsgBox "El registro se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstLista.Requery
        Dg1.Refresh

    End If
End Sub

Private Sub CmdAcepta_Click()
    If NulosC(TxtNumOrdReq.Text) = "" Then
        MsgBox "No ha especificado el requerimiento que se va a cotizar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumOrdReq.SetFocus
        Exit Sub
    End If
    
    GenerarCotizacion NulosN(LblIdReq.Caption), "0001"
    CmdCancelar_Click
    RstLista.Requery
    RstLista.Requery
    Dg1.Refresh
    Me.Refresh
End Sub

Private Sub CmdAdd_Click()
    If NulosC(Fg1.TextMatrix(Fg1.Rows - 1, 4)) = "" Then
        Exit Sub
    End If
    Fg1.Rows = Fg1.Rows + 1
End Sub

Private Sub CmdBusArea_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xForm.SQLCad = "SELECT * FROM mae_area"

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
            TxtObs.SetFocus
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing

End Sub

Private Sub CmdBusCondPag_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xForm.SQLCad = "SELECT * FROM mae_condpago ORDER BY descripcion"
    
    xForm.Titulo = "Buscando Condicion de Pago"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "descripcion"
    xForm.CampoBusca = "descripcion"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdCondPag.Text = xRs("id")
            LblCondPag.Caption = xRs("descripcion")
            Fg1.SetFocus
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMon_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xForm.SQLCad = "SELECT * FROM mae_moneda"

    xForm.Titulo = "Buscando Monedas"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "descripcion"
    xForm.CampoBusca = "descripcion"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdMon.Text = xRs("id")
            LblMoneda.Caption = xRs("descripcion")
            TxtObs.SetFocus
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusOrdreq_Click()
    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(3, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Documento":       xCampos(0, 1) = "descripcion":        xCampos(0, 2) = "1500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº Documento":    xCampos(1, 1) = "numdoc":             xCampos(1, 2) = "1500":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Solicitante":     xCampos(2, 1) = "nomsol":             xCampos(2, 2) = "3000":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Fch. Emi":        xCampos(3, 1) = "fchemi":             xCampos(3, 2) = "1200":         xCampos(3, 3) = "F"
    
    xForm.SQLCad = "SELECT com_ordenreq.id, mae_documento.descripcion, [com_ordenreq]![numser] & '-' & [com_ordenreq]![numdoc] AS numdoc, com_ordenreq.fchent, " _
        & " UCase([pla_empleados]![apepat]) & ' ' & UCase([pla_empleados]![apemat]) & ', ' & [pla_empleados]![nom] AS nomsol, com_ordenreq.fchemi " _
        & " FROM ((com_ordenreq LEFT JOIN mae_documento ON com_ordenreq.idtipdoc = mae_documento.id) LEFT JOIN com_usuario ON com_ordenreq.idsol = com_usuario.id) " _
        & " LEFT JOIN pla_empleados ON com_usuario.idper = pla_empleados.id WHERE (((com_ordenreq.idsit)=2) AND ((com_ordenreq.idest)=1))"

    'SELECT com_ordenreq.id, mae_documento.descripcion, [com_ordenreq]![numser] & '-' & [com_ordenreq]![numdoc] AS numdoc, " _
        & " com_ordenreq.fchent, UCase([pla_empleados]![apepat]) & ' ' & UCase([pla_empleados]![apemat]) & ', ' & [pla_empleados]![nom] AS nomsol, " _
        & " com_ordenreq.fchemi, com_ordenreq.idsit FROM ((com_ordenreq LEFT JOIN mae_documento ON com_ordenreq.idtipdoc = mae_documento.id) " _
        & " LEFT JOIN com_usuario ON com_ordenreq.idsol = com_usuario.id) LEFT JOIN pla_empleados ON com_usuario.idper = pla_empleados.id " _
        & " WHERE (((com_ordenreq.idsit)=2))"
    
    xForm.Titulo = "Buscando Requerimientos Pendientes"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "numdoc"
    xForm.CampoBusca = "numdoc"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtFchEmi2.Valor = xRs("fchemi")
            TxtFchVen2.Valor = xRs("fchent")
            TxtNumOrdReq.Text = xRs("numdoc")
            LblIdReq.Caption = xRs("id")
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing

End Sub

Private Sub CmdCancelar_Click()
    ActivEntorno True
    Frame3.Visible = False
End Sub

Sub ActivEntorno(Valor As Boolean)
    Toolbar1.Enabled = Valor
    TabOne1.Enabled = Valor
End Sub

Function GeneraNumDoc(NumSerie As String) As String
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT * FROM com_ordencot WHERE numser = '" & NumSerie & "' ORDER BY numdoc", xCon
    If Rst.RecordCount = 0 Then
        GeneraNumDoc = "00000001"
    Else
        Rst.MoveLast
        GeneraNumDoc = Format(Val(Rst("numdoc")) + 1, "00000000")
    End If
    Set Rst = Nothing
End Function

Sub GenerarCotizacion(idOrdenRequerimiento As Integer, NumSerie As String)
    Dim A As Integer
    Dim RstCotDet As New ADODB.Recordset
    Dim RstCotRes As New ADODB.Recordset
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xNumDoc As String
    Dim xId As Double
    Dim B As Integer
    
    ' OBTENEMOS EL RESUMEN PARA GENERAR LAS ORDENES DE COTIZACION
    RST_Busq RstCotRes, "SELECT DISTINCT com_ordenreqprocot.idor, com_ordenreqprocot.idpro, mae_prov.nombre FROM com_ordenreqprocot LEFT JOIN mae_prov " _
        & " ON com_ordenreqprocot.idpro = mae_prov.id Where (((com_ordenreqprocot.idor) = " & idOrdenRequerimiento & ")) ORDER BY mae_prov.nombre", xCon

    
    'SELECT com_ordenreqprocot.idor, com_ordenreqprocot.idpro, mae_prov.nombre FROM com_ordenreqprocot LEFT JOIN mae_prov " _
        & " ON com_ordenreqprocot.idpro = mae_prov.id Where (((com_ordenreqprocot.idor) = " & idOrdenRequerimiento & ")) ORDER BY mae_prov.nombre", xCon

    ' OBTENEMOS EL DETALLE
    RST_Busq RstCotDet, "SELECT com_ordenreqprocot.idor, com_ordenreqprocot.idpro, com_ordenreqprocot.idite, mae_prov.nombre, com_ordenreqdet.idunimed, " _
        & " com_ordenreqdet.cantidad FROM (com_ordenreqprocot LEFT JOIN mae_prov ON com_ordenreqprocot.idpro = mae_prov.id) LEFT JOIN com_ordenreqdet " _
        & " ON com_ordenreqprocot.idite = com_ordenreqdet.iditem Where (((com_ordenreqprocot.idor) = " & idOrdenRequerimiento & ") And ((com_ordenreqdet.idor) = " & idOrdenRequerimiento & ")) " _
        & " ORDER BY mae_prov.nombre", xCon

    RST_Busq RstCab, "SELECT * FROM com_ordencot", xCon
    RST_Busq RstDet, "SELECT * FROM com_ordencotdet", xCon
    
    RstCotRes.MoveFirst
    
    For A = 1 To RstCotRes.RecordCount
        xId = HallaCodigoTabla("com_ordencot", xCon, "id")
        xNumDoc = GeneraNumDoc(NumSerie)
        RstCab.AddNew
        RstCab("id") = xId
        RstCab("idtipdoc") = 107
        RstCab("numser") = NumSerie
        RstCab("numdoc") = xNumDoc
        RstCab("idpro") = RstCotRes("idpro")
        RstCab("fchemi") = Date
        RstCab("idor") = idOrdenRequerimiento
        RstCab.Update
        
        RstCotDet.Filter = adFilterNone
        RstCotDet.Filter = "idpro = " & RstCotRes("idpro") & ""
        
        If RstCotDet.RecordCount <> 0 Then
            RstCotDet.MoveFirst
            For B = 1 To RstCotDet.RecordCount
                RstDet.AddNew
                RstDet("idoc") = xId
                RstDet("iditem") = RstCotDet("idite")
                RstDet("idunimed") = RstCotDet("idunimed")
                RstDet("cantidad") = RstCotDet("cantidad")
                RstDet.Update
                
                RstCotDet.MoveNext
                If RstCotDet.EOF = True Then
                    Exit For
                End If
            Next B
        End If
        
        RstCotRes.MoveNext
        If RstCotRes.EOF = True Then Exit For
    Next A
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, 1, Time, Time, Date, xCon, xId
    QueHace = 3
    
    xCon.Execute "UPDATE com_ordenreq SET com_ordenreq.idest = 2 WHERE (((com_ordenreq.id)=" & idOrdenRequerimiento & "))"

End Sub

Sub Nuevo()
    QueHace = 1
    xHorIni = Time
    ActivEntorno False
    TxtNumOrdReq.Text = ""
    TxtFchEmi2.Valor = ""
    TxtFchVen2.Valor = ""
    
    Frame3.Left = 3480
    Frame3.Top = 1965
    Frame3.Visible = True
    Bloquea
End Sub

Sub Modificar()
    If RstLista("idest") = 2 Then
        MsgBox "No se puede modificar una orden de cotizacion aprobada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    If RstLista("idest") = 3 Then
        MsgBox "No se puede modificar una orden de cotizacion procesada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    QueHace = 2
    xHorIni = Time
    ActivaTool
    Bloquea
    TabOne1.CurrTab = 1
    Label5.Caption = "Modificando Orden de Cotizacion"
    
    'Fg1.ColWidth(7) = 1200
    Fg1.Editable = flexEDKbdMouse
    Fg1.SelectionMode = flexSelectionFree
    
    MuestraSegundoTab
    Fg1.ColComboList(2) = "|..."
    Fg1.ColComboList(3) = "|..."
    Fg1.ColComboList(4) = "|..."
    
    Fg1.SetFocus
End Sub

Private Sub CMdDel_Click()
    If Fg1.Rows = 1 Then Exit Sub
    
    Fg1.RemoveItem Fg1.Row
End Sub

Private Sub CmdNewItem_Click()
    If QueHace = 3 Then Exit Sub
    
    Dim xFun As New Sgi2_Procesos.Procesos
    Dim xIdProducto As Integer
    Dim xRs As New ADODB.Recordset
    
    xIdProducto = xFun.IngRapidoItems(xCon)
    If xIdProducto <> 0 Then
        RST_Busq xRs, "SELECT alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, mae_tipoproducto.descripcion AS desctippro, alm_inventario.id, " _
            & " alm_inventario.idunimed FROM mae_tipoproducto RIGHT JOIN (mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed) " _
            & " ON mae_tipoproducto.id = alm_inventario.tippro Where (((alm_inventario.activo) = -1) And ((alm_inventario.id) = " & xIdProducto & ")) ORDER BY alm_inventario.descripcion", xCon
    
        If xRs.RecordCount <> 0 Then
            Fg1.TextMatrix(Fg1.Row, 4) = xRs("descripcion")
            Fg1.TextMatrix(Fg1.Row, 5) = xRs("abrev")
            Fg1.TextMatrix(Fg1.Row, 9) = xRs("id")
            Fg1.TextMatrix(Fg1.Row, 10) = xRs("idunimed")
        End If
        Set xRs = Nothing
        Fg1.SetFocus
    End If
End Sub

Private Sub Command1_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Empleado":    xCampos(0, 1) = "apenom":           xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":      xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xForm.SQLCad = "SELECT com_usuario.id, com_usuario.idper, UCase([pla_empleados]![apepat]) & ' ' & UCase([pla_empleados]![apemat]) & ', ' & [pla_empleados]![nom] AS apenom " _
        & " FROM com_usuario LEFT JOIN pla_empleados ON com_usuario.idper = pla_empleados.id"

    xForm.Titulo = "Buscando Usuarios"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "apenom"
    xForm.CampoBusca = "apenom"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtidSol.Text = xRs("id")
            LblSolicitante.Caption = xRs("apenom")
            TxtIdArea.SetFocus
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub Command2_Click()
    If QueHace = 3 Then Exit Sub
    'If Option1.Value = True Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
        
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Nombre":       xCampos(0, 1) = "nombre":      xCampos(0, 2) = "6000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":    xCampos(1, 1) = "numruc":      xCampos(1, 2) = "1200":         xCampos(1, 3) = "C"
    
    xForm.SQLCad = "SELECT * FROM mae_prov ORDER BY nombre"
    
    xForm.Titulo = "Buscando Proveedor"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "nombre"
    xForm.CampoBusca = "nombre"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            LblIdProv.Caption = xRs("id")
            TxtIdpro.Text = xRs("numruc")
            LblProveedor.Caption = xRs("nombre")
            TxtFchEmi.SetFocus
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstLista
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA LAS COLUMNAS DEL CONTROL Dg1
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstLista.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstLista("id")), xCon
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim xCampos2(2, 4) As String
    Dim xCampos3(2, 4) As String
    Dim xCampos1(3, 4) As String
    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
       
    If Col = 2 Then
        xCampos2(0, 0) = "Descripcion":    xCampos2(0, 1) = "descripcion":      xCampos2(0, 2) = "4000":         xCampos2(0, 3) = "C"
        xCampos2(1, 0) = "Codigo":         xCampos2(1, 1) = "id":               xCampos2(1, 2) = "1400":         xCampos2(1, 3) = "N"
        
        xForm.SQLCad = "SELECT * FROM man_equipotipo"
        xForm.Titulo = "Buscando Tipos"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        xForm.Ordenado = "descripcion"
        xForm.CampoBusca = "descripcion"
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos2)
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 2) = xRs("descripcion")
                Fg1.TextMatrix(Fg1.Row, 11) = xRs("id")
            End If
        End If
    End If
    
    If Col = 3 Then
        If NulosN(Fg1.TextMatrix(Fg1.Row, 11)) = 0 Then Exit Sub
        xCampos1(0, 0) = "Descripcion":     xCampos1(0, 1) = "nombre":          xCampos1(0, 2) = "3500":   xCampos1(0, 3) = "C"
        xCampos1(1, 0) = "Caracteristicas": xCampos1(1, 1) = "caracteristicas": xCampos1(1, 2) = "5000":   xCampos1(1, 3) = "C"
        xCampos1(2, 0) = "Codigo":          xCampos1(2, 1) = "id":              xCampos1(2, 2) = "1000":   xCampos1(2, 3) = "N"
        
        xForm.SQLCad = "SELECT man_equipos.nombre, man_equipos.id, man_equipos.caracteristicas From man_equipos WHERE (((man_equipos.idtip)=" & NulosN(Fg1.TextMatrix(Fg1.Row, 11)) & "))"

        xForm.Titulo = "Buscando Equipos e Instalaciones"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        xForm.Ordenado = "nombre"
        xForm.CampoBusca = "nombre"
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos1)
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 3) = xRs("nombre")
                Fg1.TextMatrix(Fg1.Row, 12) = xRs("id")
            End If
        End If
    End If
    
    If Col = 4 Then
        Dim xCampos(3, 4) As String
        
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "codpro":           xCampos(1, 2) = "1400":         xCampos(1, 3) = "c"
        xCampos(2, 0) = "Abreviatura":    xCampos(2, 1) = "abrev":            xCampos(2, 2) = "1000":         xCampos(2, 3) = "c"
        xCampos(3, 0) = "Tipo Producto":  xCampos(3, 1) = "desctippro":       xCampos(3, 2) = "1200":         xCampos(3, 3) = "c"
        
        xForm.SQLCad = "SELECT alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, mae_tipoproducto.descripcion AS desctippro, " _
            & " alm_inventario.id, alm_inventario.idunimed FROM mae_tipoproducto RIGHT JOIN (mae_unidades RIGHT JOIN alm_inventario " _
            & " ON mae_unidades.id = alm_inventario.idunimed) ON mae_tipoproducto.id = alm_inventario.tippro Where (((alm_inventario.activo) = -1)) " _
            & " ORDER BY alm_inventario.descripcion"
        
        xForm.Titulo = "Buscando Items"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        xForm.Ordenado = "descripcion"
        xForm.CampoBusca = "descripcion"
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 4) = xRs("descripcion")
                Fg1.TextMatrix(Fg1.Row, 5) = xRs("abrev")
                Fg1.TextMatrix(Fg1.Row, 9) = xRs("id")
                Fg1.TextMatrix(Fg1.Row, 10) = xRs("idunimed")
                'Fg1.TextMatrix(Fg1.Row, 1) = (NulosN(Fg1.TextMatrix(Fg1.Row - 1, 1)) + 1)
                
                If NulosC(Fg1.TextMatrix(Fg1.Rows - 1, 2)) <> "" Then
                    Fg1.Rows = Fg1.Rows + 1
                    Fg1.TextMatrix(Fg1.Rows - 1, 1) = (NulosN(Fg1.TextMatrix(Fg1.Rows - 2, 1)) + 1)
                End If
            End If
        End If
    
    End If
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 6 Or Col = 7 Then
        Fg1.TextMatrix(Row, 6) = Format(Fg1.TextMatrix(Row, 6), "0.00")
        Fg1.TextMatrix(Row, 7) = Format(Fg1.TextMatrix(Row, 7), "0.00")
        
        Fg1.TextMatrix(Row, 8) = Format(NulosN(Fg1.TextMatrix(Row, 6)) * NulosN(Fg1.TextMatrix(Row, 7)), "0.00")
    End If
    
    LblTotal.Caption = Format(GRID_SUMAR_COL(Fg1, 8, 1, Fg1.Rows - 1), "0.00")
End Sub

Private Sub Fg1_EnterCell()
    If Fg1.Col = 2 Or Fg1.Col = 3 Or Fg1.Col = 4 Or Fg1.Col = 6 Or Fg1.Col = 7 Then
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = 6 Or Col = 7 Then
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Else
        KeyAscii = 0
    End If

'    If Col = 4 Or Col = 5 Then
'        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
'    Else
'        KeyAscii = 0
'    End If

End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
    
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
        If DeDonde <> 2 Then
            RST_Busq RstLista, "SELECT com_ordencot.*, mae_prov.nombre, [com_ordencot]![numser]+'-'+[com_ordencot]![numdoc] AS numdoccot, " _
                & " [com_ordenreq]![numser] & '-' & [com_ordenreq]![numdoc] AS numdocreq, mae_documento.descripcion AS descdoc, mae_estados.descripcion AS desest " _
                & " FROM (((com_ordencot LEFT JOIN mae_prov ON com_ordencot.idpro = mae_prov.id) LEFT JOIN com_ordenreq ON com_ordencot.idor = com_ordenreq.id) " _
                & " LEFT JOIN mae_documento ON com_ordencot.idtipdoc = mae_documento.id) LEFT JOIN mae_estados ON com_ordencot.idest = mae_estados.id", xCon
        Else
            RST_Busq RstLista, "SELECT com_ordencot.*, mae_prov.nombre, [com_ordencot]![numser]+'-'+[com_ordencot]![numdoc] AS numdoccot, " _
                & " [com_ordenreq]![numser] & '-' & [com_ordenreq]![numdoc] AS numdocreq, mae_documento.descripcion AS descdoc, mae_estados.descripcion AS desest " _
                & " FROM (((com_ordencot LEFT JOIN mae_prov ON com_ordencot.idpro = mae_prov.id) LEFT JOIN com_ordenreq ON com_ordencot.idor = com_ordenreq.id) " _
                & " LEFT JOIN mae_documento ON com_ordencot.idtipdoc = mae_documento.id) LEFT JOIN mae_estados ON com_ordencot.idest = mae_estados.id " _
                & " WHERE (((com_ordencot.id)=" & xIdOR & "))", xCon
        End If
        
        Set Dg1.DataSource = RstLista
        
        
        If DeDonde = 2 Then
            Toolbar1.Buttons(1).Visible = False
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = False
            Toolbar1.Buttons(7).Visible = False
            TabOne1.CurrTab = 1
        End If
    End If
End Sub

Private Sub Form_Load()
    CaracteresNumericos = "0123456789." & Chr(8) & Chr(13)
    SeEjecuto = False
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    QueHace = 3

    Fg1.ColWidth(4) = Fg1.ColWidth(4) + 465
    Fg1.ColWidth(1) = 0
    Fg1.ColWidth(9) = 0
    Fg1.ColWidth(10) = 0
    Fg1.ColWidth(11) = 0
    Fg1.ColWidth(12) = 0
    TabOne1.CurrTab = 0
    Fg1.SelectionMode = flexSelectionByRow
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
        If xOrigen = 0 Then
            If RstLista.State = 0 Then Exit Sub
            If RstLista.RecordCount = 0 And QueHace <> 1 Then
                Cancel = 1
                Exit Sub
            End If
            If QueHace = 3 Then MuestraSegundoTab
        End If
    End If
End Sub

Sub MuestraSegundoTab()
    Dim A As Integer
    
    If RstLista("tipcot") = 1 Then
        TxtNumSerReq.BackColor = &H8000000F
        TxtNumDocReq.BackColor = &H8000000F
        TxtNumSerReq.Text = ""
        TxtNumDocReq.Text = ""
    Else
    End If
    
    LblEstado.Caption = Busca_Codigo(RstLista("idest"), "id", "descripcion", "mae_estados", "N", xCon)
    If RstLista("idest") = 1 Then LblEstado.ForeColor = &H8000&   ' verde    pediente
    If RstLista("idest") = 2 Then LblEstado.ForeColor = &HC00000  ' azul     aprobada
    If RstLista("idest") = 3 Then LblEstado.ForeColor = &H800080     ' amarillo procesada
    If RstLista("idest") = 4 Then LblEstado.ForeColor = &HFF&     ' rojo     rechasada
    
    TxtIdDoc.Text = RstLista("idtipdoc")
    TxtIdDoc_Validate True
    TxtNumSerCot.Text = NulosC(RstLista("numser"))
    TxtNumDocCot.Text = NulosC(RstLista("numdoc"))
    LblIdProv.Caption = RstLista("idpro")
    TxtIdpro.Text = Busca_Codigo(RstLista("idpro"), "id", "numruc", "mae_prov", "N", xCon)
    LblProveedor.Caption = NulosC(RstLista("nombre"))
    If IsNull(RstLista("fchemi")) = False Then TxtFchEmi.Valor = RstLista("fchemi")
    TxtIdMon.Text = RstLista("idmon")
    TxtIdMon_Validate True
    TxtidSol.Text = RstLista("idsol")
    TxtIdSol_Validate True
    TxtObs.Text = NulosC(RstLista("obs"))
    TxtIdArea.Text = NulosN(RstLista("idarea"))
    TxtIdArea_Validate True
    
    Dim RstDet As New ADODB.Recordset
    
    RST_Busq RstDet, "SELECT com_ordencotdet.*, alm_inventario.descripcion AS descitem, mae_unidades.abrev AS descuni, man_equipotipo.descripcion AS desctip, man_equipos.nombre AS nomequi " _
        & " FROM (((com_ordencotdet LEFT JOIN alm_inventario ON com_ordencotdet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON com_ordencotdet.idunimed = mae_unidades.id) " _
        & " LEFT JOIN man_equipotipo ON com_ordencotdet.idtip = man_equipotipo.id) LEFT JOIN man_equipos ON com_ordencotdet.idequi = man_equipos.id " _
        & " WHERE (((com_ordencotdet.idoc)=" & NulosN(RstLista("id")) & "))", xCon
    Fg1.Rows = 1
    
    If RstDet.RecordCount <> 0 Then
        RstDet.MoveFirst
        For A = 1 To RstDet.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = A
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstDet("desctip"))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(RstDet("nomequi"))
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(RstDet("descitem"))
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(RstDet("descuni"))
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(RstDet("cantidad"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(RstDet("precio"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = RstDet("iditem")
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = RstDet("idunimed")
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosN(RstDet("idtip"))
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = NulosN(RstDet("idequi"))
                        
            RstDet.MoveNext
            If RstDet.EOF = True Then Exit For
        Next A
    End If
    Set RstDet = Nothing
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then NuevoDirecto
        
    If Button.Index = 2 Then Modificar
    
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            RstLista.Requery
            Dg1.Refresh
            Cancelar
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstLista.Filter = ""
    End If
    
    If Button.Index = 13 Then Imprimir RstLista("id"), 1
    
    If Button.Index = 12 Then
        Imprimir RstLista("id"), 2
    End If

    If Button.Index = 15 Then
        Set RstLista = Nothing
        Unload Me
    End If
End Sub

Sub Imprimir(IdCotizacion As Integer, Opcion As Integer)
    ' OPCION = 1 SE ABRE EL DOCUMENTO PDF
    ' OPCION = 2 SE ENVIA POR CORREO EL ARCHIVO PDF
    Dim Li As Integer
    Dim strSource As String
    Dim xArea, xEmp, xDir, xCuerpo, xCad  As String
    Dim xEmpleado As String
    Dim Pagina As Integer
    Dim Lineas As Integer
    
    'On Error Resume Next
    Set oPDF = New cPDF
    
    If oPDF.PDFCreate(App.Path & "\OCT" & RstLista("numdoccot") & ".pdf") = True Then
        oPDF.Fonts.Add "Tit", Times_BoldItalic, WinAnsiEncoding
        oPDF.Fonts.Add "Head", Times_Italic, WinAnsiEncoding
        oPDF.Fonts.Add "Cont", Courier, WinAnsiEncoding
        oPDF.Fonts.Add "Time", Times_Roman, WinAnsiEncoding
        
        CrearCabecera RstLista("numdoccot")
        xCad = xDisEmp & " " & Format(RstLista("fchemi"), "dd") & " de " & Format(RstLista("fchemi"), "mmmm") & " del " & Format(RstLista("fchemi"), "yyyy")
        
        oPDF.WTextBox 100, 55, 10, 420, xCad, "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        oPDF.WTextBox 120, 55, 10, 420, "Para :", "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        
        xArea = Busca_Codigo(RstLista("idpro"), "id", "areacon", "mae_prov", "N", xCon)
        xEmp = RstLista("nombre")
        xDir = Busca_Codigo(RstLista("idpro"), "id", "dir", "mae_prov", "N", xCon)
        
        oPDF.WTextBox 130, 80, 10, 420, NulosC(xArea), "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        oPDF.WTextBox 140, 80, 10, 420, xEmp, "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        oPDF.WTextBox 150, 80, 10, 420, NulosC(xDir), "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        
        ' ESCRIBIMOS EL CONTENIDO DEL CUERPO
        xCuerpo = "Por medio de la presente le saludamos y solicitamos nos envié en el mas breve plazo la cotización de los siguientes ítems"
        oPDF.WTextBox 170, 55, 10, 420, xCuerpo, "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        
        oPDF.WTextBox 195, 55, 15, 33, "Item", "Time", 9, hCenter, vMiddle, vbWhite, 1, vbBlack, True
        oPDF.WTextBox 195, 90, 15, 278, "Descripcion", "Time", 9, hCenter, vMiddle, vbWhite, 1, vbBlack, True
        oPDF.WTextBox 195, 370, 15, 48, "Uni. Med.", "Time", 9, hCenter, vMiddle, vbWhite, 1, vbBlack, True
        oPDF.WTextBox 195, 420, 15, 40, "Cantidad", "Time", 9, hCenter, vMiddle, vbWhite, 1, vbBlack, True
        
        Dim Rst As New ADODB.Recordset
        Dim A, Fila As Integer
        RST_Busq Rst, "SELECT com_ordencotdet.*, alm_inventario.descripcion AS descitem, mae_unidades.abrev AS descuni " _
            & " FROM (com_ordencotdet LEFT JOIN alm_inventario ON com_ordencotdet.iditem = alm_inventario.id) LEFT JOIN mae_unidades " _
            & " ON com_ordencotdet.idunimed = mae_unidades.id WHERE (((com_ordencotdet.idoc)=" & RstLista("id") & "))", xCon

        If Rst.RecordCount <> 0 Then
            Fila = 215
            Rst.MoveFirst
            For A = 1 To Rst.RecordCount
                oPDF.WTextBox Fila, 55, 10, 33, A, "Time", 9, hCenter, vMiddle, vbBlack, , vbBlack
                oPDF.WTextBox Fila, 90, 10, 278, Rst("descitem"), "Time", 9, hLeft, vMiddle, vbBlack, , vbBlack
                oPDF.WTextBox Fila, 370, 10, 48, Rst("descuni"), "Time", 9, hCenter, vMiddle, vbBlack, , vbBlack
                oPDF.WTextBox Fila, 420, 10, 40, Format(Rst("cantidad"), "0.00"), "Time", 9, hRight, vMiddle, vbBlack, , vbBlack
                
                Fila = Fila + 10
                Rst.MoveNext
                If Rst.EOF = True Then Exit For
            Next A
        End If
        'opdf.WRectangle 190, 55, 20, 400, 1.5, vbBlack
        
        ' ESCRIBIMOS EL FINAL DEL DOCUMENTO
        xCuerpo = "Especificar:"
        oPDF.WTextBox 470, 55, 10, 420, xCuerpo, "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        xCuerpo = "* Validez de la oferta "
        oPDF.WTextBox 480, 75, 10, 420, xCuerpo, "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        xCuerpo = "* Lugar de entrega"
        oPDF.WTextBox 490, 75, 10, 420, xCuerpo, "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        xCuerpo = "* Formato de pago"
        oPDF.WTextBox 500, 75, 10, 420, xCuerpo, "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        xCuerpo = "* Plazo de entrega"
        oPDF.WTextBox 510, 75, 10, 420, xCuerpo, "Time", 9, hJustify, vMiddle, vbBlack, , vbBlack
        
        xCuerpo = "Atentamente"
        oPDF.WTextBox 530, 55, 10, 420, xCuerpo, "Time", 9, hCenter, vMiddle, vbBlack, , vbBlack

        ' ESCRIBIMOS LA FIRMA DEL ENCARGADO
        xCuerpo = "--------------------------------"
        oPDF.WTextBox 560, 55, 10, 420, xCuerpo, "Time", 9, hCenter, vMiddle, vbBlack, , vbBlack
        xEmpleado = "Juan Perez Martinez"
        oPDF.WTextBox 570, 55, 10, 420, xEmpleado, "Time", 9, hCenter, vMiddle, vbBlack, , vbBlack
        xCuerpo = "Jefe de Compras"
        oPDF.WTextBox 580, 55, 10, 420, xCuerpo, "Time", 9, hCenter, vMiddle, vbBlack, , vbBlack

        
'        opdf.WLineTo 493, Li, 32, Li
'        opdf.LineStroke
        
        oPDF.PDFClose
        Set oPDF = Nothing
        
        If Opcion = 1 Then
            Shell ("rundll32.exe url.dll,FileProtocolHandler " & Trim(App.Path) & ("\OC" & RstLista("numdoccot") & ".pdf")), vbMaximizedFocus
        End If
        
        If Opcion = 2 Then
            Dim xIdPro As Integer
            Dim eMail As String
            xIdPro = Busca_Codigo(IdCotizacion, "id", "idpro", "com_ordencot", "N", xCon)
            eMail = NulosC(Busca_Codigo(xIdPro, "id", "email", "mae_prov", "N", xCon))
            If NulosC(eMail) = "" Then
                MsgBox "EL proveedor " & Trim(xEmp) & " no tiene correo electronico, agregue la direccion de correo electronico del proveedor para efectuar esta operacion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
            
            Dim xFun As New eps_librerias.Correo
            Dim xAdjunto(2) As String
            xFun.ServidorSMTP = "mail.agro-vado.com"
            'xFun.ServidorSMTP = "smtpseguro.speedy.com.pe"
            xFun.NomRemitente = "Sistema de Compras"
            xFun.MailRemitente = "seven@seven.com"
            xFun.MailDestino = eMail
            xFun.Asunto = "Orden de Cotizacion Nº " & RstLista("numdoccot")
            xFun.Cuerpo = "Buenos dias remito orden de cotizacion favor de contestar a la brevedad prosible"
            
            xAdjunto(0) = Trim(App.Path) & "\OC" & RstLista("numdoccot") & ".pdf"
            If xFun.EnviarCorreo(xAdjunto) = True Then
                xCon.Execute "UPDATE com_ordencot SET com_ordencot.envmail = -1 WHERE (((com_ordencot.id)=" & RstLista("id") & "))"
                RstLista.Requery
                Dg1.Refresh
            End If
        End If
    Else
        MsgBox "No se Puede Mostrar Documento", vbCritical, "Error"
    End If
End Sub

Sub CrearCabecera(NumDoc As String)
    Dim xTelEmp, xNumDoc As String
    
    xTelEmp = "Telf: 493-0808   Tele Fax: 295-6868"
    xNumDoc = NumDoc

    oPDF.NewPage UsarAnchoAlto, 525, 675
    Pagina = Pagina + 1
    oPDF.WTextBox 32, 55, 20, 250, xNomEmp, "Tit", 12, hCenter, vMiddle, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 55, 55, 10, 250, xDirEmp, "Head", 9, hCenter, vMiddle, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 65, 55, 10, 250, xTelEmp, "Head", 9, hCenter, vMiddle, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 75, 55, 10, 250, xPagEmp, "Head", 9, hCenter, vMiddle, RGB(0, 0, 128), , vbRed
    
    oPDF.WTextBox 46, 330, 10, 150, "ORDEN DE COTIZACION", "Head", 10, hCenter, vMiddle, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 60, 330, 10, 150, "Nº " & xNumDoc, "Head", 10, hCenter, vMiddle, RGB(0, 0, 128), , vbRed
    
    oPDF.WRectangle 32, 330, 53, 150, 1.5, vbBlack
End Sub

Function Grabar()
    Dim xCampos(11, 5) As String
    Dim xCampos2(6, 5) As String
    Dim xId As Double
    Dim A, B As Integer
    
    ' ELIMINAMOS LAS FILAS EN BLANCO DEL CONTROL Fg1
    For A = 1 To Fg1.Rows - 1
        If Fg1.TextMatrix(A, 4) = "" Then
            Fg1.RemoveItem A
        End If
    Next A
    
    Dim xTipCot As Integer
    xTipCot = 1
        
On Error GoTo LaCague
    xCon.BeginTrans
    If QueHace = 1 Then
        xId = HallaCodigoTabla("com_ordencot", xCon, "id")
    Else
        xId = RstLista("id")
    End If
    
    'Columna    | Descripcion
    '------------------------
    '0          | campo
    '1          | Valor
    '2          | requerido
    '3          | tipo
    '4          | msj error
    '5          | INDICA QUE EL CAMPO ES INDICE Y NO SE ESCRIBIRA CUANDO SE MODIFIQUE EL REGISTRO
    '--------------------------------
    'GRABAMOS LA CABECERA DE LA ORDEN DE REQUERIMIENTO
    xCampos(0, 0) = "id":           xCampos(0, 1) = Str(xId):                   xCampos(0, 2) = "S":    xCampos(0, 3) = "N":    xCampos(0, 4) = "":                                                                     xCampos(0, 5) = "S"
    xCampos(1, 0) = "idtipdoc":     xCampos(1, 1) = "108":                      xCampos(1, 2) = "S":    xCampos(1, 3) = "N":    xCampos(1, 4) = "":                                                                     xCampos(1, 5) = ""
    xCampos(2, 0) = "numser":       xCampos(2, 1) = NulosC(TxtNumSerCot.Text):  xCampos(2, 2) = "S":    xCampos(2, 3) = "C":    xCampos(2, 4) = "":                                                                     xCampos(2, 5) = ""
    xCampos(3, 0) = "numdoc":       xCampos(3, 1) = NulosC(TxtNumDocCot.Text):  xCampos(3, 2) = "S":    xCampos(3, 3) = "C":    xCampos(3, 4) = "":                                                                     xCampos(3, 5) = ""
    xCampos(4, 0) = "idpro":        xCampos(4, 1) = NulosC(LblIdProv.Caption):  xCampos(4, 2) = "S":    xCampos(4, 3) = "N":    xCampos(4, 4) = "No ha especificado el proveedor":                                      xCampos(4, 5) = ""
    xCampos(5, 0) = "fchemi":       xCampos(5, 1) = TxtFchEmi.Valor:            xCampos(5, 2) = "S":    xCampos(5, 3) = "F":    xCampos(5, 4) = "No ha especificado la fecha de emision de la orden de requerimiento":  xCampos(5, 5) = ""
    xCampos(6, 0) = "idmon":        xCampos(6, 1) = NulosC(TxtIdMon.Text):      xCampos(6, 2) = "S":    xCampos(6, 3) = "N":    xCampos(6, 4) = "No ha especificado la moneda de la orden de requerimiento":            xCampos(6, 5) = ""
    xCampos(7, 0) = "idsol":        xCampos(7, 1) = TxtidSol.Text:              xCampos(7, 2) = "S":    xCampos(7, 3) = "N":    xCampos(7, 4) = "No ha especificado el solicitante del requerimiento":                  xCampos(7, 5) = ""
    xCampos(8, 0) = "obs":          xCampos(8, 1) = TxtObs.Text:                xCampos(8, 2) = "N":    xCampos(8, 3) = "C":    xCampos(8, 4) = "":                                                                     xCampos(8, 5) = ""
    xCampos(9, 0) = "idest":        xCampos(9, 1) = 1:                          xCampos(9, 2) = "S":    xCampos(9, 3) = "N":    xCampos(9, 4) = "":                                                                     xCampos(9, 5) = ""
    xCampos(10, 0) = "tipcot":      xCampos(10, 1) = xTipCot:                   xCampos(10, 2) = "S":   xCampos(10, 3) = "N":   xCampos(10, 4) = "":                                                                    xCampos(10, 5) = ""
    xCampos(11, 0) = "idarea":      xCampos(11, 1) = TxtIdArea.Text:            xCampos(11, 2) = "S":   xCampos(11, 3) = "N":   xCampos(11, 4) = "":                                                                    xCampos(11, 5) = ""
    
    If QueHace = 1 Then
        If EscribirNuevoRegistro(xCampos, "com_ordencot", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
    Else
        ' ELIMINAMOS LOS DETALLES DE LA ORDEN DE COTIZACION
        xCon.Execute "DELETE * FROM com_ordencotdet WHERE idoc = " & RstLista("id") & ""
        ' MODIFICAMOS EL REGISTRO
        If ModificarRegistro(xCampos, "com_ordencot", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
    End If
    
    ' GRABAMOS EL DETALLE DE LA ORDEN DE REQUERIMIENTO
    For A = 1 To Fg1.Rows - 1
        xCampos2(0, 0) = "idoc":       xCampos2(0, 1) = Str(xId):               xCampos2(0, 2) = "S":   xCampos2(0, 3) = "N":  xCampos2(0, 4) = "":   xCampos2(0, 5) = ""
        xCampos2(1, 0) = "iditem":     xCampos2(1, 1) = Fg1.TextMatrix(A, 9):   xCampos2(1, 2) = "S":   xCampos2(1, 3) = "N":  xCampos2(1, 4) = "":   xCampos2(1, 5) = ""
        xCampos2(2, 0) = "idunimed":   xCampos2(2, 1) = Fg1.TextMatrix(A, 10):  xCampos2(2, 2) = "S":   xCampos2(2, 3) = "N":  xCampos2(2, 4) = "":   xCampos2(2, 5) = ""
        xCampos2(3, 0) = "cantidad":   xCampos2(3, 1) = Fg1.TextMatrix(A, 6):   xCampos2(3, 2) = "S":   xCampos2(3, 3) = "N":  xCampos2(3, 4) = "":   xCampos2(3, 5) = ""
        xCampos2(4, 0) = "precio":     xCampos2(4, 1) = Fg1.TextMatrix(A, 7):   xCampos2(4, 2) = "S":   xCampos2(4, 3) = "N":  xCampos2(4, 4) = "":   xCampos2(4, 5) = ""
        xCampos2(5, 0) = "idtip":      xCampos2(5, 1) = Fg1.TextMatrix(A, 11):  xCampos2(5, 2) = "N":   xCampos2(5, 3) = "N":  xCampos2(5, 4) = "":   xCampos2(5, 5) = ""
        xCampos2(6, 0) = "idequi":     xCampos2(6, 1) = Fg1.TextMatrix(A, 12):  xCampos2(6, 2) = "N":   xCampos2(6, 3) = "N":  xCampos2(6, 4) = "":   xCampos2(6, 5) = ""
        
        If EscribirNuevoRegistro(xCampos2, "com_ordencotdet", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
    Next A
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    xCon.CommitTrans
    MsgBox "El registro se grabo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Exit Function

LaCague:
    Resume
    xCon.RollbackTrans
    MsgBox "No se pudo guardar el registro por el siguiente motivo: " & Err.Description, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = False
End Function

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index = 1 Then
        Nuevo
    End If
    If ButtonMenu.Index = 2 Then
        NuevoDirecto
    End If
End Sub

Sub Blanquea()

    TxtNumSerReq.Text = ""
    TxtNumDocReq.Text = ""
    TxtIdDoc.Text = ""
    LblDocumento.Caption = ""
    TxtIdpro.Text = ""
    LblProveedor.Caption = ""
    
    TxtNumSerCot.Text = ""
    TxtNumDocCot.Text = ""
    TxtFchEmi.Valor = ""
    TxtIdMon.Text = ""
    LblMoneda.Caption = ""
    
    TxtidSol.Text = ""
    LblSolicitante.Caption = ""
    TxtObs.Text = ""
    TxtIdArea.Text = ""
    LblArea.Caption = ""
    
End Sub

Sub NuevoDirecto()
    TxtNumSerReq.BackColor = &H8000000F
    TxtNumDocReq.BackColor = &H8000000F
    
    Blanquea
    Bloquea
    ActivaTool
    TxtIdDoc.Text = "108"
    TxtIdDoc_Validate False
    
    QueHace = 1
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Agregando Orden de Cotizacion"
    Fg1.Rows = 1
    Fg1.Rows = Fg1.Rows + 1
    Fg1.ColComboList(2) = "|..."
    Fg1.ColComboList(3) = "|..."
    Fg1.ColComboList(4) = "|..."
    LblEstado.Caption = ""
    
    Fg1.SelectionMode = flexSelectionFree
    Fg1.Editable = flexEDKbdMouse
    TxtIdpro.SetFocus
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
        CmdBusArea_Click
    End If
End Sub

Private Sub TxtIdArea_Validate(Cancel As Boolean)
    If NulosN(TxtIdArea.Text) = 0 Then
        TxtIdArea.Text = ""
        LblArea.Caption = ""
        Exit Sub
    End If

    LblArea.Caption = Busca_Codigo(TxtIdArea.Text, "id", "descripcion", "mae_area", "N", xCon)
    If NulosC(LblArea.Caption) = "" Then
        TxtIdArea.Text = ""
        LblArea.Caption = ""
    End If
End Sub

Private Sub TxtIdCondPag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdCondPag_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCondPag_Click
    End If
End Sub

Private Sub TxtIdCondPag_Validate(Cancel As Boolean)
    If NulosN(TxtIdCondPag.Text) = 0 Then
        LblCondPag.Caption = ""
        Exit Sub
    End If
        
    LblCondPag.Caption = Busca_Codigo(TxtIdCondPag.Text, "id", "descripcion", "mae_condpago", "N", xCon)
    If NulosC(LblCondPag.Caption) = "" Then
        TxtIdCondPag.Text = ""
    End If
End Sub

Private Sub TxtIdDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdDoc_Validate(Cancel As Boolean)
    If TxtIdDoc.Text = "" Then
        Exit Sub
    End If
    LblDocumento.Caption = Busca_Codigo(TxtIdDoc.Text, "id", "descripcion", "mae_documento", "N", xCon)
    If LblDocumento.Caption = "" Then
        TxtIdDoc.Text = ""
    Else
        TxtNumSerCot.Text = "0001"
        TxtNumDocCot.Text = GeneraNumDoc(TxtNumSerCot.Text)
    End If
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdMon_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMon_Click
    End If
End Sub

Private Sub TxtIdMon_Validate(Cancel As Boolean)
    If NulosN(TxtIdMon.Text) = 0 Then
        LblSolicitante.Caption = ""
        Exit Sub
    End If
   
    LblMoneda.Caption = Busca_Codigo(TxtIdMon.Text, "id", "descripcion", "mae_moneda", "N", xCon)
    If NulosC(LblMoneda.Caption) = "" Then
        TxtIdMon.Text = ""
        LblMoneda.Caption = ""
    End If
End Sub

Private Sub TxtIdpro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdpro_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        Command2_Click
    End If
End Sub

Private Sub TxtIdpro_Validate(Cancel As Boolean)
    If TxtIdpro.Text = "" Then
        Exit Sub
    End If
    LblProveedor.Caption = Busca_Codigo(TxtIdpro.Text, "numruc", "nombre", "mae_prov", "C", xCon)
    If LblProveedor.Caption = "" Then
        TxtIdpro.Text = ""
    End If
End Sub

Private Sub TxtIdSol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdSol_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        Command1_Click
    End If
End Sub

Private Sub TxtIdSol_Validate(Cancel As Boolean)
    If NulosN(TxtidSol.Text) = 0 Then
        LblSolicitante.Caption = ""
        Exit Sub
    End If
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT com_usuario.id, com_usuario.idper, UCase([pla_empleados]![apepat]) & ' ' & UCase([pla_empleados]![apemat]) & ', ' & [pla_empleados]![nom] AS apenom " _
        & " FROM com_usuario LEFT JOIN pla_empleados ON com_usuario.idper = pla_empleados.id WHERE (((com_usuario.id)=" & NulosN(TxtidSol.Text) & "))", xCon
    If Rst.RecordCount <> 0 Then
        LblSolicitante.Caption = Rst("apenom")
    End If
    If NulosC(LblSolicitante.Caption) = "" Then
        TxtidSol.Text = ""
        LblSolicitante.Caption = ""
    End If
End Sub

Private Sub TxtNumDocCot_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumDocReq_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumOrdReq_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusOrdreq_Click
    End If
End Sub

Private Sub TxtNumSerCot_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumSerReq_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

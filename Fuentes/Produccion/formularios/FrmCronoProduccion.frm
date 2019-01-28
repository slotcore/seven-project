VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCronoProduccion 
   Caption         =   "Produccion - Cronograma de Produccion"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3330
      Left            =   5565
      TabIndex        =   28
      Top             =   450
      Visible         =   0   'False
      Width           =   6075
      Begin VB.TextBox TxtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4005
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "TxtTotal"
         Top             =   2460
         Width           =   945
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   350
         Left            =   3045
         TabIndex        =   36
         Top             =   2865
         Width           =   1155
      End
      Begin VB.CommandButton CmAcepta 
         Caption         =   "&Aceptar"
         Height          =   350
         Left            =   1860
         TabIndex        =   35
         Top             =   2865
         Width           =   1155
      End
      Begin VB.TextBox TxtCan 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1170
         TabIndex        =   33
         Text            =   "TxtCan"
         Top             =   660
         Width           =   1185
      End
      Begin VB.TextBox TxtMP 
         Height          =   300
         Left            =   1170
         TabIndex        =   29
         Text            =   "TxtMP"
         Top             =   360
         Width           =   4845
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg2 
         Height          =   1470
         Left            =   75
         TabIndex        =   30
         Top             =   990
         Width           =   5940
         _cx             =   10477
         _cy             =   2593
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
         BackColorSel    =   64
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmCronoProduccion.frx":0000
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
      Begin VB.Label Label11 
         Caption         =   "Total ==>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   3090
         TabIndex        =   38
         Top             =   2505
         Width           =   825
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   15
         X2              =   6045
         Y1              =   3315
         Y2              =   3315
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   75
         TabIndex        =   34
         Top             =   690
         Width           =   630
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   6060
         X2              =   6060
         Y1              =   15
         Y2              =   3330
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   15
         X2              =   0
         Y1              =   0
         Y2              =   3315
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   15
         X2              =   6045
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccion de Productos"
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
         Left            =   120
         TabIndex        =   32
         Top             =   60
         Width           =   2040
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Materia Prima"
         Height          =   195
         Left            =   75
         TabIndex        =   31
         Top             =   390
         Width           =   960
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   255
         Left            =   30
         Top             =   45
         Width           =   6015
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   5010
      Left            =   15
      TabIndex        =   7
      Top             =   360
      Width           =   11805
      _cx             =   20823
      _cy             =   8837
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
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
      FrontTabForeColor=   -2147483630
      Caption         =   "  &Consulta  |   &Detalle   "
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
      Begin SizerOneLibCtl.ElasticOne ElasticOne3 
         Height          =   4575
         Left            =   -12360
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   390
         Width           =   11715
         _cx             =   20664
         _cy             =   8070
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
         Appearance      =   4
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   700
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   8
         BorderWidth     =   2
         ChildSpacing    =   2
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   2
         GridCols        =   1
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmCronoProduccion.frx":009B
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   4170
            Left            =   30
            TabIndex        =   15
            Top             =   375
            Width           =   11655
            _ExtentX        =   20558
            _ExtentY        =   7355
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Fch. Ini."
            Columns(1).DataField=   "fchini"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Fin."
            Columns(2).DataField=   "fchfin"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Tipo Produccion"
            Columns(3).DataField=   "destippro"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Programador"
            Columns(4).DataField=   "apenom"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1296"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1217"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=1746"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1667"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1667"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1588"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=3757"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=3678"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=512"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=8731"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=8652"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=70,.parent=13,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Named:id=33:Normal"
            _StyleDefs(57)  =   ":id=33,.parent=0"
            _StyleDefs(58)  =   "Named:id=34:Heading"
            _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(60)  =   ":id=34,.wraptext=-1"
            _StyleDefs(61)  =   "Named:id=35:Footing"
            _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(63)  =   "Named:id=36:Selected"
            _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(65)  =   "Named:id=37:Caption"
            _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(67)  =   "Named:id=38:HighlightRow"
            _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(69)  =   "Named:id=39:EvenRow"
            _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(71)  =   "Named:id=40:OddRow"
            _StyleDefs(72)  =   ":id=40,.parent=33"
            _StyleDefs(73)  =   "Named:id=41:RecordSelector"
            _StyleDefs(74)  =   ":id=41,.parent=34"
            _StyleDefs(75)  =   "Named:id=42:FilterBar"
            _StyleDefs(76)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Cronogramas"
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
            Height          =   315
            Left            =   30
            TabIndex        =   14
            Top             =   30
            Width           =   11655
         End
      End
      Begin SizerOneLibCtl.ElasticOne ElasticOne1 
         Height          =   4575
         Left            =   45
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   390
         Width           =   11715
         _cx             =   20664
         _cy             =   8070
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
         Appearance      =   4
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   700
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   8
         BorderWidth     =   2
         ChildSpacing    =   2
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   4
         GridCols        =   1
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmCronoProduccion.frx":00DD
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   375
            Left            =   30
            TabIndex        =   25
            Top             =   4170
            Width           =   11655
            Begin VB.Label LblProducto 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblProducto"
               Height          =   300
               Left            =   1755
               TabIndex        =   27
               Top             =   30
               Width           =   9900
            End
            Begin VB.Label Label7 
               Caption         =   "MP / Producto =>"
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
               Left            =   120
               TabIndex        =   26
               Top             =   75
               Width           =   1560
            End
         End
         Begin SizerOneLibCtl.ElasticOne ElasticOne2 
            Height          =   1050
            Left            =   30
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   375
            Width           =   11655
            _cx             =   20558
            _cy             =   1852
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
            Appearance      =   4
            MousePointer    =   0
            _ConvInfo       =   1
            Version         =   700
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   8
            BorderWidth     =   2
            ChildSpacing    =   2
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   1
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   2
            PicturePos      =   4
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   1
            GridCols        =   2
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmCronoProduccion.frx":0136
            Begin VB.Frame Frame2 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   990
               Left            =   30
               TabIndex        =   16
               Top             =   30
               Width           =   9060
               Begin VB.CommandButton CmdBusTip 
                  Enabled         =   0   'False
                  Height          =   240
                  Left            =   1740
                  Picture         =   "FrmCronoProduccion.frx":0178
                  Style           =   1  'Graphical
                  TabIndex        =   18
                  Top             =   720
                  Width           =   225
               End
               Begin VB.CommandButton CmdBusSup 
                  Enabled         =   0   'False
                  Height          =   240
                  Left            =   1740
                  Picture         =   "FrmCronoProduccion.frx":02AA
                  Style           =   1  'Graphical
                  TabIndex        =   17
                  Top             =   90
                  Width           =   225
               End
               Begin VB.TextBox TxtIdSup 
                  Height          =   300
                  Left            =   990
                  Locked          =   -1  'True
                  TabIndex        =   0
                  Text            =   "TxtIdSup"
                  Top             =   60
                  Width           =   1000
               End
               Begin VB.TextBox TxtTipPro 
                  Height          =   300
                  Left            =   990
                  Locked          =   -1  'True
                  TabIndex        =   3
                  Text            =   "TxtTipPro"
                  Top             =   690
                  Width           =   1000
               End
               Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
                  Height          =   300
                  Left            =   3405
                  TabIndex        =   2
                  Top             =   375
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
               Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
                  Height          =   300
                  Left            =   990
                  TabIndex        =   1
                  Top             =   375
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
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  Caption         =   "Tipo Prod."
                  Height          =   195
                  Index           =   1
                  Left            =   30
                  TabIndex        =   24
                  Top             =   750
                  Width           =   735
               End
               Begin VB.Label LblTipoProd 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblTipoProd"
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
                  Left            =   2025
                  TabIndex        =   23
                  Top             =   690
                  Width           =   4215
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Supervisor"
                  Height          =   195
                  Left            =   30
                  TabIndex        =   22
                  Top             =   105
                  Width           =   750
               End
               Begin VB.Label LblSupervisor 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblSupervisor"
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
                  Left            =   2025
                  TabIndex        =   21
                  Top             =   60
                  Width           =   4215
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  Caption         =   "Fch. Final"
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   20
                  Top             =   435
                  Width           =   690
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Fch. Inicio"
                  Height          =   195
                  Left            =   30
                  TabIndex        =   19
                  Top             =   435
                  Width           =   735
               End
            End
            Begin VB.Frame Frame5 
               BackColor       =   &H00C0E0FF&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   990
               Left            =   9120
               TabIndex        =   11
               Top             =   30
               Width           =   2505
               Begin VB.CommandButton CmdProcesar 
                  Caption         =   "&Procesar"
                  Height          =   570
                  Left            =   675
                  TabIndex        =   4
                  Top             =   180
                  Width           =   1110
               End
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   2685
            Left            =   30
            TabIndex        =   12
            Top             =   1455
            Width           =   11655
            _cx             =   20558
            _cy             =   4736
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
            Cols            =   10
            FixedRows       =   2
            FixedCols       =   2
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmCronoProduccion.frx":03DC
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
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Cronograma"
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
            Height          =   315
            Left            =   30
            TabIndex        =   13
            Top             =   30
            Width           =   11655
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8250
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion.frx":04B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion.frx":09F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion.frx":0B79
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion.frx":0FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion.frx":10E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion.frx":1629
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion.frx":1B6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion.frx":1C81
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion.frx":1D95
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion.frx":21E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion.frx":2355
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   5
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
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   11
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Recetas del producto"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Productos "
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Consulta de Producción"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   1350
      TabIndex        =   6
      Top             =   2295
      Width           =   11610
   End
   Begin VB.Menu menu_01 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu_01_01 
         Caption         =   "Programar Productos"
      End
   End
End
Attribute VB_Name = "FrmCronoProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'*                             COMENTARIOS DEL SUPER GENIO
'****************************************************************************************
'* PARA QUE ESTE FORMULARIO CAMINE BIEN SE DEBE DE CONFIGURAR LA HORA EN CONFIGURACION
'* REGIONAL DEL SISTEMA OPERATIVO A HH:mm:ss
'*
'*
'****************************************************************************************
Option Explicit

Dim xNomMatPriPro As String
Dim QueHace As Integer
Dim Agregando As Boolean
Dim RstLis As New ADODB.Recordset
Dim RstMatPro As New ADODB.Recordset
Dim xIdMatPri As Integer
Dim xFchPro, xHorPro As Date

Dim oPDF As cPDF
Dim xNumPag As Integer
Dim xFilaInicial As Integer
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO
Dim fOrdenLista As Boolean                                ' especfica el orden de la lista de la consulta
Dim SeEjecuto As Boolean


Private Sub CmAcepta_Click()
    If NulosN(TxtTotal.Text) > NulosN(TxtCan.Text) Then
        MsgBox "El cantidad a procesar en productos es mayor a la cantidad de materia prima", vbInformation + vbOKOnly + vbDefaultButton1
        TxtTotal.SetFocus
        Exit Sub
    End If
    
    Dim B As Integer
    
    For B = 1 To Fg2.Rows - 1
        RstMatPro.Filter = adFilterNone
        If Abs(NulosN(Fg2.TextMatrix(B, 3))) = 1 Then
            RstMatPro.Filter = "iditem = " & xIdMatPri & " AND fchpro = " & xFchPro & " AND horpro = " & xHorPro & " AND idpro = " & NulosN(Fg2.TextMatrix(B, 4)) & ""
            If RstMatPro.RecordCount = 0 Then
                RstMatPro.AddNew
                RstMatPro("id") = 0
                RstMatPro("iditem") = xIdMatPri
                RstMatPro("fchpro") = xFchPro
                RstMatPro("horpro") = xHorPro
                RstMatPro("idpro") = Fg2.TextMatrix(B, 4)
                RstMatPro("cantidad") = NulosN(Fg2.TextMatrix(B, 2))
            Else
                RstMatPro("cantidad") = NulosN(Fg2.TextMatrix(B, 2))
            End If
        Else
            RstMatPro.Filter = "iditem = " & xIdMatPri & " AND fchpro = " & xFchPro & " AND horpro = " & xHorPro & " AND idpro = " & NulosN(Fg2.TextMatrix(B, 4)) & ""
            If RstMatPro.RecordCount <> 0 Then
                RstMatPro.Delete
            End If
        End If
    Next B
    
    CmdCancelar_Click
End Sub

Private Sub CmdBusSup_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Apellidos y Nombres":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":                xCampos(1, 1) = "id":        xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT pro_emp.*, pla_empleados.nombre FROM (pro_emp LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) " _
        & "  LEFT JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id Where (((pro_empdet.idfun) = 2)) ORDER BY pla_empleados.nombre"
    
    'SELECT pro_emp.*, UCase([pla_empleados]![apepat]) & ' ' & UCase([pla_empleados]![apemat]) & ', ' & [pla_empleados]![nom] AS apenom " _
        & " FROM pla_empleados LEFT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp WHERE (((pro_emp.prog)=-1))"
    
    xform.titulo = "Buscando Supervisores"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdSup.Text = xRs("id")
            LblSupervisor.Caption = xRs("nombre")
            TxtFchIni.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing

End Sub

Private Sub CmdBusTip_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_tipoproducto.id, mae_tipoproducto.descripcion FROM mae_tipoproducto"
    
    xform.titulo = "Buscando Tipo de Item"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtTipPro.Text = xRs("id")
            LblTipoProd.Caption = xRs("descripcion")
            CmdProcesar.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdCancelar_Click()
    TabOne1.Enabled = True
    Toolbar1.Enabled = True
    
    Frame3.Visible = False
End Sub

Private Sub CmdProcesar_Click()
    If TxtFchIni.valor = "" Then
        MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    If TxtFchFin.valor = "" Then
        MsgBox "No ha especificado la fecha final", vbInformation + vbOKOnly + vbDefaultButton1
        TxtFchFin.SetFocus
        Exit Sub
    End If
    
    If NulosN(TxtTipPro.Text) = 0 Then
        MsgBox "No ha especificado el tipo de producto a procesar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipPro.SetFocus
        Exit Sub
    End If
    
    Fg1.Editable = flexEDKbdMouse
    Fg1.Cols = 2
    Dim xFchIni, xFchFin As Date
    Dim A, xCol As Integer
    
    xFchIni = TxtFchIni.valor
    xFchFin = TxtFchFin.valor
    xCol = 2
    Agregando = True
    
    For A = xFchIni To xFchFin
        Fg1.Cols = Fg1.Cols + 1
        Fg1.TextMatrix(1, xCol) = "Materia Prima/Producto"
        Fg1.ColWidth(Fg1.Cols - 1) = 2500
        
        Fg1.ColComboList(Fg1.Cols - 1) = "|..."
        
        xCol = xCol + 1
        
        Fg1.Cols = Fg1.Cols + 1
        Fg1.TextMatrix(1, xCol) = "Materia Prima/Producto"
        
        GRID_COMBINAR Fg1, 0, Fg1.Cols - 2, 0, Fg1.Cols - 1, Format(A, "dddd") & " -" & Format(A, "dd"), flexAlignCenterCenter, True, flexMergeFree, , &H8000000F, True
        
        Fg1.TextMatrix(1, xCol) = "Cantidad"
        Fg1.ColWidth(Fg1.Cols - 1) = 900
        xCol = xCol + 1
    Next A
    Agregando = False
End Sub

Function Grabar() As Boolean
    Dim xId As Double
    Dim xCampos(4, 5) As String
    Dim xCampos2(4, 5) As String
    Dim xCampos3(5, 5) As String
On Error GoTo LaCague

    xCon.BeginTrans
    
    If QueHace = 1 Then
        xId = HallaCodigoTabla("pro_cronograma", xCon, "id")
    Else
        xId = RstLis("id")
        xCon.Execute "DELETE * FROM pro_cronogramadet WHERE id = " & xId & ""
        xCon.Execute "DELETE * FROM pro_cronogramadetprod WHERE id = " & xId & ""
    End If
    
    'Columna    | Descripcion
    '------------------------
    '0          | campo
    '1          | Valor
    '2          | requerido
    '3          | tipo
    '4          | msj error
    
    '--------------------------------
    'GRABAMOS LA CABECERA DEL CRONOGRAMA
    xCampos(0, 0) = "id":           xCampos(0, 1) = Str(xId):             xCampos(0, 2) = "S":    xCampos(0, 3) = "N":     xCampos(0, 4) = "":                                         xCampos(0, 5) = "S"
    xCampos(1, 0) = "idsup":        xCampos(1, 1) = TxtIdSup.Text:        xCampos(1, 2) = "S":    xCampos(1, 3) = "N":     xCampos(1, 4) = "No ha especificado el supervisor":         xCampos(1, 5) = ""
    xCampos(2, 0) = "fchini":       xCampos(2, 1) = TxtFchIni.valor:      xCampos(2, 2) = "S":    xCampos(2, 3) = "F":     xCampos(2, 4) = "No ha especificado la fecha de inicio":    xCampos(2, 5) = ""
    xCampos(3, 0) = "fchfin":       xCampos(3, 1) = TxtFchFin.valor:      xCampos(3, 2) = "S":    xCampos(3, 3) = "F":     xCampos(3, 4) = "No ha especificado la fecha final":        xCampos(3, 5) = ""
    xCampos(4, 0) = "idtippro":     xCampos(4, 1) = TxtTipPro.Text:       xCampos(4, 2) = "S":    xCampos(4, 3) = "N":     xCampos(4, 4) = "No ha especificado el tipo de producto":   xCampos(4, 5) = ""
    
    If QueHace = 1 Then
        If EscribirNuevoRegistro(xCampos, "pro_cronograma", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
    Else
        If ModificarRegistro(xCampos, "pro_cronograma", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
    End If
    
    Dim xFchIni, xFchFin As Date
    Dim A, B, xCol As Integer
    
    xFchIni = TxtFchIni.valor
    xFchFin = TxtFchFin.valor
    
    '--------------------------------
    'GRABAMOS EL DETALLE DEL CRONOGRAMA
    xCol = 2
    Dim xCodPro As Double
    Dim xCanPro As Double
    Dim xHorPro As Date
    Dim xFchPro As Date
    
    For A = xFchIni To xFchFin
        For B = 2 To Fg1.Rows - 1
            If NulosC(Fg1.TextMatrix(B, xCol)) <> "" Then
                xCodPro = Busca_Codigo(Fg1.TextMatrix(B, xCol), "descripcion", "id", "alm_inventario", "C", xCon)
                xCanPro = NulosN(Fg1.TextMatrix(B, xCol + 1))
                xHorPro = Fg1.TextMatrix(B, 1)
                xFchPro = A
                
                xCampos2(0, 0) = "id":           xCampos2(0, 1) = xId:        xCampos2(0, 2) = "S":    xCampos2(0, 3) = "N":     xCampos2(0, 4) = "":                                        xCampos2(0, 5) = ""
                xCampos2(1, 0) = "fchpro":       xCampos2(1, 1) = xFchPro:    xCampos2(1, 2) = "S":    xCampos2(1, 3) = "F":     xCampos2(1, 4) = "No ha especificado el supervisor":        xCampos2(1, 5) = ""
                xCampos2(2, 0) = "horpro":       xCampos2(2, 1) = Fg1.TextMatrix(B, 1):    xCampos2(2, 2) = "S":    xCampos2(2, 3) = "F":     xCampos2(2, 4) = "No ha especificado la fecha de inicio":   xCampos2(2, 5) = ""
                xCampos2(3, 0) = "iditem":       xCampos2(3, 1) = xCodPro:    xCampos2(3, 2) = "S":    xCampos2(3, 3) = "N":     xCampos2(3, 4) = "No ha especificado la fecha final":       xCampos2(3, 5) = ""
                xCampos2(4, 0) = "cantidad":     xCampos2(4, 1) = xCanPro:    xCampos2(4, 2) = "S":    xCampos2(4, 3) = "N":     xCampos2(4, 4) = "No ha especificado el tipo de producto":  xCampos2(4, 5) = ""
                
                If EscribirNuevoRegistro(xCampos2, "pro_cronogramadet", xCon) = False Then
                    xCon.RollbackTrans
                    Exit Function
                End If
            End If
        Next B
        xCol = xCol + 2
    Next A
    
    RstMatPro.Filter = adFilterNone
    
    
    If RstMatPro.RecordCount <> 0 Then
        RstMatPro.MoveFirst
        For A = 1 To RstMatPro.RecordCount

            xCampos3(0, 0) = "id":           xCampos3(0, 1) = xId:                    xCampos3(0, 2) = "S":    xCampos3(0, 3) = "N":     xCampos3(0, 4) = "":   xCampos3(0, 5) = ""
            xCampos3(1, 0) = "iditem":       xCampos3(1, 1) = RstMatPro("iditem"):    xCampos3(1, 2) = "S":    xCampos3(1, 3) = "N":     xCampos3(1, 4) = "":   xCampos3(1, 5) = ""
            xCampos3(2, 0) = "fchpro":       xCampos3(2, 1) = RstMatPro("fchpro"):    xCampos3(2, 2) = "S":    xCampos3(2, 3) = "F":     xCampos3(2, 4) = "":   xCampos3(2, 5) = ""
            xCampos3(3, 0) = "horpro":       xCampos3(3, 1) = RstMatPro("horpro"):    xCampos3(3, 2) = "S":    xCampos3(3, 3) = "F":     xCampos3(3, 4) = "":   xCampos3(3, 5) = ""
            xCampos3(4, 0) = "idpro":        xCampos3(4, 1) = RstMatPro("idpro"):     xCampos3(4, 2) = "S":    xCampos3(4, 3) = "N":     xCampos3(4, 4) = "":   xCampos3(4, 5) = ""
            xCampos3(5, 0) = "cantidad":     xCampos3(5, 1) = RstMatPro("cantidad"):  xCampos3(5, 2) = "S":    xCampos3(5, 3) = "N":     xCampos3(5, 4) = "":   xCampos3(5, 5) = ""

            If EscribirNuevoRegistro(xCampos3, "pro_cronogramadetprod", xCon) = False Then
                xCon.RollbackTrans
                Exit Function
            End If

            RstMatPro.MoveNext
            If RstMatPro.EOF = True Then Exit For
        Next A
    End If
    
'    For A = xFchIni To xFchFin
'        For B = 2 To Fg1.Rows - 1
'            If NulosC(Fg1.TextMatrix(B, xCol)) <> "" Then
'                xCodPro = Busca_Codigo(Fg1.TextMatrix(B, xCol), "descripcion", "id", "alm_inventario", "C", xCon)
'                xCanPro = Fg1.TextMatrix(B, xCol + 1)
'                xHorPro = Fg1.TextMatrix(B, 1)
'                xFchPro = A
'
'                xCampos2(0, 0) = "id":           xCampos2(0, 1) = xId:        xCampos2(0, 2) = "S":    xCampos2(0, 3) = "N":     xCampos2(0, 4) = "":                                        xCampos2(0, 5) = ""
'                xCampos2(1, 0) = "fchpro":       xCampos2(1, 1) = xFchPro:    xCampos2(1, 2) = "S":    xCampos2(1, 3) = "F":     xCampos2(1, 4) = "No ha especificado el supervisor":        xCampos2(1, 5) = ""
'                xCampos2(2, 0) = "horpro":       xCampos2(2, 1) = Fg1.TextMatrix(B, 1):    xCampos2(2, 2) = "S":    xCampos2(2, 3) = "F":     xCampos2(2, 4) = "No ha especificado la fecha de inicio":   xCampos2(2, 5) = ""
'                xCampos2(3, 0) = "iditem":       xCampos2(3, 1) = xCodPro:    xCampos2(3, 2) = "S":    xCampos2(3, 3) = "N":     xCampos2(3, 4) = "No ha especificado la fecha final":       xCampos2(3, 5) = ""
'                xCampos2(4, 0) = "cantidad":     xCampos2(4, 1) = xCanPro:    xCampos2(4, 2) = "S":    xCampos2(4, 3) = "N":     xCampos2(4, 4) = "No ha especificado el tipo de producto":  xCampos2(4, 5) = ""
'
'                If EscribirNuevoRegistro(xCampos2, "pro_cronogramadet", xCon) = False Then
'                    xCon.RollbackTrans
'                    Exit Function
'                End If
'            End If
'        Next B
'        xCol = xCol + 2
'    Next A

    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId


    xCon.CommitTrans
    MsgBox "El registro se guardo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Exit Function

LaCague:
    'Resume
    xCon.RollbackTrans
    MsgBox "No se pudo guardar el registro por el siguiente error : " & Trim(Err.Description)
    Grabar = False
End Function

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstLis
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA EN FORMA ASCENDENTE O DECENDETE LAS COLUMNAS DEL CONTROL Dg3
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstLis.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstLis("id")), xCon
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(3, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Uni. Med.":     xCampos(2, 1) = "abrev":          xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
    
    If TxtTipPro.Text = "1" Then
        xform.SQLCad = "SELECT alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev " _
            & " FROM mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed " _
            & " Where (((alm_inventario.tippro) = 1)) ORDER BY alm_inventario.descripcion"
        
        xform.titulo = "Buscando Materia Prima"
    End If
    
    If TxtTipPro.Text = "3" Then
        xform.SQLCad = "SELECT alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.activo" _
            & " FROM mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed " _
            & " Where (((alm_inventario.tippro) = 3) And ((alm_inventario.activo) = -1)) " _
            & " ORDER BY alm_inventario.descripcion"
        xform.titulo = "Buscando Productos"
    End If
    
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            Fg1.TextMatrix(Fg1.Row, Fg1.Col) = xRs("descripcion")
            'menu_01_01_Click
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    Select Case Col
        Case 3, 5, 7, 9, 11, 13, 15, 17, 19, 21
            Fg1.TextMatrix(Fg1.Row, Fg1.Col) = Format(Fg1.TextMatrix(Fg1.Row, Fg1.Col), "0.00")
    
        Case 2, 4, 6, 8, 19, 12, 14, 16, 18, 20
            If Fg1.TextMatrix(Fg1.Row, Fg1.Col) = "" Then
                Fg1.TextMatrix(Fg1.Row, Fg1.Col + 1) = ""
                RstMatPro.Filter = adFilterNone
                
                ' hallamos el id de la materia prima
                xIdMatPri = Busca_Codigo(xNomMatPriPro, "descripcion", "id", "alm_inventario", "C", xCon)
                
                ' hallamos la hora de programacion
                xHorPro = CDate(Format(Fg1.TextMatrix(Fg1.Row, 1), "hh:mm"))
    
                ' hallamos el dia de programacion
                Dim xNum As Integer
                Dim xFchIni, xFchFin, xDia As Date
                
                xFchIni = TxtFchIni.valor
                xFchFin = TxtFchFin.valor
    
                xNum = NulosN(Mid(Trim(Fg1.TextMatrix(0, Fg1.Col)), Len(Trim(Fg1.TextMatrix(0, Fg1.Col))) - 1, 2))
            
                For xDia = xFchIni To xFchFin
                    If NulosN(Mid(Trim(Str(xDia)), 1, 2)) = xNum Then
                        xFchPro = xDia
                        Exit For
                    End If
                Next xDia
                
                
                RstMatPro.Filter = "iditem = " & xIdMatPri & " AND fchpro = " & xFchPro & " AND horpro = " & xHorPro & " "
                Dim X As Integer
                If RstMatPro.RecordCount <> 0 Then
                    For X = 1 To RstMatPro.RecordCount
                        RstMatPro.Delete
                        RstMatPro.MoveNext
                        If RstMatPro.EOF = True Then Exit For
                    Next X
                End If
            End If
    End Select

End Sub

Private Sub Fg1_EnterCell()
    
    If Fg1.Col = 2 Or Fg1.Col = 4 Or Fg1.Col = 6 Or Fg1.Col = 8 Or Fg1.Col = 10 Or Fg1.Col = 12 Or Fg1.Col = 14 Or Fg1.Col = 16 Then
        xNomMatPriPro = Fg1.TextMatrix(Fg1.Row, Fg1.Col)
        LblProducto.Caption = Fg1.TextMatrix(Fg1.Row, Fg1.Col)
    Else
        LblProducto.Caption = Fg1.TextMatrix(Fg1.Row, Fg1.Col - 1)
    End If
    
    'If Agregando = True Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    Fg1.Editable = flexEDKbdMouse
    Select Case Fg1.Col
        Case 3, 5, 7, 9, 11, 13, 15, 17, 19, 21
            If NulosC(Fg1.TextMatrix(Fg1.Row, Fg1.Col - 1)) = "" Then
                Fg1.Editable = flexEDNone
            Else
                Fg1.Editable = flexEDKbdMouse
            End If
    
    End Select
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If QueHace = 3 Then KeyAscii = 0
    
    If KeyAscii = 13 Then Exit Sub
    ' validar los caracteres que se ingresan
    Select Case Col
        Case 2, 4, 6, 8, 10, 12, 14, 16, 18, 20
            KeyAscii = 0
            
        Case 3, 5, 7, 9, 11, 13, 15, 17, 19, 21 ' canpro,preunibru,valdes,preuni,imptot
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End Select
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA Y DESACTIVA LOS BOTONES DE LA BARRA DE HERRAMIENTAS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If NulosC(Fg1.TextMatrix(Fg1.Row, Fg1.Col)) = "" Then
            Exit Sub
        Else
            If NulosN(TxtTipPro.Text) = 1 Then
                PopupMenu menu_01
            End If
        End If
    End If
End Sub

Private Sub Fg2_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Fg2.Col = 2 Then
        Fg2.TextMatrix(Row, Col) = Format(Fg2.TextMatrix(Row, Col), "0.00")
        
        TxtTotal.Text = Format(GRID_SUMAR_COL(Fg2, 2, 1, Fg2.Rows - 1), "0.00")
        If NulosN(Fg2.TextMatrix(Row, Col)) <> 0 Then
            Fg2.TextMatrix(Row, 3) = 1
        Else
            Fg2.TextMatrix(Row, 3) = 0
        End If
    End If
    If Fg2.Col = 3 Then
        If NulosN(Fg2.TextMatrix(Row, Col)) = 0 Then
            Fg2.TextMatrix(Row, 2) = ""
        End If
        TxtTotal.Text = Format(GRID_SUMAR_COL(Fg2, 2, 1, Fg2.Rows - 1), "0.00")
    End If
End Sub

Private Sub Fg2_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If QueHace = 3 Then KeyAscii = 0
    
    If KeyAscii = 13 Then Exit Sub
    ' validar los caracteres que se ingresan
    Select Case Col
        Case 1, 3
            KeyAscii = 0
            
        Case 2
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End Select
End Sub

Private Sub Form_Activate()
    Dim SeEjecuto As Boolean
    Dim Rpta As Integer
    
    If SeEjecuto = False Then
    
        SeEjecuto = True
    
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
    
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
        RST_Busq RstLis, "SELECT pro_cronograma.*, mae_tipoproducto.descripcion AS destippro, " _
            & " [pla_empleados]![apepat] & ' ' & [pla_empleados]![apemat] & ', ' & [pla_empleados]![nom] AS apenom " _
            & " FROM (pla_empleados RIGHT JOIN (pro_cronograma LEFT JOIN pro_emp ON pro_cronograma.idsup = pro_emp.id) ON pla_empleados.id = pro_emp.idemp) " _
            & " LEFT JOIN mae_tipoproducto ON pro_cronograma.idtippro = mae_tipoproducto.id ORDER BY pro_cronograma.fchini", xCon
        
        Set Dg1.DataSource = RstLis
        
    End If
End Sub

Sub MuestraSegundoTab()
    TxtIdSup.Text = RstLis("idsup")
    TxtIdSup_Validate True
    TxtFchIni.valor = RstLis("fchini")
    TxtFchFin.valor = RstLis("fchfin")
    TxtTipPro.Text = RstLis("idtippro")
    TxtTipPro_Validate True
    
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT pro_cronogramadet.*, alm_inventario.descripcion FROM pro_cronogramadet LEFT JOIN alm_inventario ON pro_cronogramadet.iditem = alm_inventario.id " _
        & " WHERE (((pro_cronogramadet.id)=" & RstLis("id") & "))", xCon

    Fg1.Cols = 2
    Fg1.Rows = Fg1.Rows - 1
    Dim xFchIni, xFchFin As Date
    Dim A, B, xCol, xFil As Integer
    
    xFchIni = TxtFchIni.valor
    xFchFin = TxtFchFin.valor
    xCol = 2
    Agregando = True
    
    GRID_COMBINAR Fg1, 0, 1, 1, 1, "Hora", flexAlignCenterCenter, False, flexMergeFree, , &H8000000F, True
    
    'CREAMOS LA CUADRICULA
    For A = xFchIni To xFchFin
        Fg1.Cols = Fg1.Cols + 1
        Fg1.TextMatrix(1, xCol) = "Materia Prima/Producto"
        Fg1.ColWidth(Fg1.Cols - 1) = 2500
        
        Fg1.ColComboList(Fg1.Cols - 1) = "|..."
        
        xCol = xCol + 1
        
        Fg1.Cols = Fg1.Cols + 1
        Fg1.TextMatrix(1, xCol) = "Materia Prima/Producto"
        
        GRID_COMBINAR Fg1, 0, Fg1.Cols - 2, 0, Fg1.Cols - 1, Format(A, "dddd") & " -" & Format(A, "dd"), flexAlignCenterCenter, True, flexMergeFree, , &H8000000F, True
        
        Fg1.TextMatrix(1, xCol) = "Cantidad"
        Fg1.ColWidth(Fg1.Cols - 1) = 900
        xCol = xCol + 1
    Next A
    
    'LLENAMOS DATOS EN LA CUADRICULA
    xCol = 2
    
    For A = xFchIni To xFchFin
        xFil = 2
        For B = 2 To Fg1.Rows - 1
            Rst.Filter = adFilterNone
            'Rst.Filter = "fchpro = '" & Format(A, "dd/mm/yy") & "' AND horpro = '" & CDate(Format(Fg1.TextMatrix(xFil, 1), "hh:mm")) & "'"
            Rst.Filter = "fchpro = '" & Format(A, "dd/mm/yy") & "' AND horpro = '" & (Format(Fg1.TextMatrix(xFil, 1), "hh:mm")) & "'"
            If Rst.RecordCount = 1 Then
                Fg1.TextMatrix(xFil, xCol) = Rst("descripcion")
                Fg1.TextMatrix(xFil, xCol + 1) = Format(Rst("cantidad"), "0.00")
            End If
            xFil = xFil + 1
        Next B
        xCol = xCol + 2
    Next A
    
    ' SUMAMOS LAS COLUMNAS DE LA CUADRICULA
    Fg1.Rows = Fg1.Rows + 1
    
    GRID_COMBINAR Fg1, Fg1.Rows - 1, 2, Fg1.Rows - 1, 2, "TOTALES", flexAlignCenterCenter, True, flexMergeFree, , &H8000000F, True
    
    xCol = 2
    Dim xTotal As Double
    
    For A = xFchIni To xFchFin
        xTotal = GRID_SUMAR_COL(Fg1, xCol + 1, 2, Fg1.Rows - 2)
        Fg1.TextMatrix(Fg1.Rows - 1, xCol + 1) = Format(xTotal, "0.00")
        
        xCol = xCol + 2
        xTotal = xTotal
    Next A
    
    Agregando = False
    Fg1.Select 2, 2
    LblProducto.Caption = Fg1.TextMatrix(Fg1.Row, Fg1.Col)
       
    RST_Busq RstMatPro, "SELECT pro_cronogramadetprod.*, alm_inventario.descripcion AS descpro FROM pro_cronogramadetprod " _
        & " LEFT JOIN alm_inventario ON pro_cronogramadetprod.idpro = alm_inventario.id WHERE (((pro_cronogramadetprod.id)=" & RstLis("id") & "))", xCon
    
    RstMatPro.ActiveConnection = Nothing
End Sub

Private Sub Form_Load()
    configurarGrid
    Agregando = False
    SeEjecuto = False
    QueHace = 3
    TabOne1.Top = 360
    TabOne1.Left = 15
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    Frame5.BackColor = &H8000000F
    TabOne1.CurrTab = 0
    Me.Height = 8000
    Me.Width = 12000
    'Frame5
End Sub

Sub Modificar()
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Modificando Cronograma de Produccion"
    QueHace = 2
    xHorIni = Time
    ActivaTool
    Bloquea
    Fg1.Editable = flexEDKbdMouse
    configurarGrid
    MuestraSegundoTab
    'Fg1.Cols = 2
    'Fg1.Rows = 2
    Fg1.Rows = Fg1.Rows - 1
    LblProducto.Caption = ""
    TxtIdSup.SetFocus
End Sub

Sub Nuevo()
    QueHace = 1
    xHorIni = Time
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Agregando Cronograma de Produccion"
    ActivaTool
    Bloquea
    Blanquea
    Fg1.Editable = flexEDKbdMouse
    Fg1.Cols = 2
    Fg1.Rows = 2
    configurarGrid
    LblProducto.Caption = ""
    
    
    RST_Busq RstMatPro, "SELECT pro_cronogramadetprod.*, alm_inventario.descripcion AS descpro FROM pro_cronogramadetprod " _
        & " LEFT JOIN alm_inventario ON pro_cronogramadetprod.idpro = alm_inventario.id WHERE (((pro_cronogramadetprod.id)=99999))", xCon
    
    RstMatPro.ActiveConnection = Nothing
    
    TxtIdSup.SetFocus
End Sub

Sub Bloquea()
    TxtIdSup.Locked = Not TxtIdSup.Locked
    TxtFchIni.Locked = Not TxtFchIni.Locked
    TxtFchFin.Locked = Not TxtFchFin.Locked
    TxtTipPro.Locked = Not TxtTipPro.Locked
    
    CmdBusSup.Enabled = Not CmdBusSup.Enabled
    CmdBusTip.Enabled = Not CmdBusTip.Enabled
End Sub

Sub Blanquea()
    TxtIdSup.Text = ""
    TxtFchIni.valor = ""
    TxtFchFin.valor = ""
    TxtTipPro.Text = ""
    LblSupervisor.Caption = ""
    LblTipoProd.Caption = ""
End Sub

Sub configurarGrid()
    Dim A As Integer
    
    Fg1.Rows = 2
    Fg1.Cols = 2
    Fg1.ColWidth(1) = 500
    Fg2.ColWidth(4) = 0
    Dim Hora As Date
    Dim xFila As Integer
    xFila = 2
    Hora = "07:00"
    For A = 1 To 48
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(xFila, 1) = Format(Hora, "hh:mm")
        Hora = Hora + "00:30"
        xFila = xFila + 1
        
        If Hora >= "23:59" Then Exit For
    Next A
End Sub

Private Sub Form_Resize()
    Dim TopEO As Integer
    
    TopEO = 400
    
    If Me.WindowState = 1 Then Exit Sub
    'EO.Width = Me.Width - 130
    
    If Me.Height <= (TopEO + 2375) Then
        Me.Height = (TopEO + 2375)
    Else
        'EO.Height = (Me.Height - (TopEO + 400))
    End If
    
    TabOne1.Height = (Me.Height - 760)
    TabOne1.Width = Me.Width - 130
End Sub

Sub Cancelar()
    QueHace = 3
    Bloquea
    Label5.Caption = "Consultando Cronograma de Produccion"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    ActivaTool
End Sub

Private Sub menu_01_01_Click()
    If QueHace = 3 Then
        CmAcepta.Enabled = False
        Fg2.SelectionMode = flexSelectionByRow
        Fg2.Editable = flexEDNone
    Else
        CmAcepta.Enabled = True
        Fg2.SelectionMode = flexSelectionFree
        Fg2.Editable = flexEDKbdMouse
    End If
    
    If TxtTipPro.Text <> 1 Then Exit Sub
    Dim Rst As New ADODB.Recordset
    Dim A, B As Integer
    Dim xMatPri As String
        
    TabOne1.Enabled = False
    Toolbar1.Enabled = False
    Fg2.Rows = 1
    Frame3.Left = ((Me.Width - Frame3.Width) / 2)
    Frame3.Top = ((Me.Height - Frame3.Height) / 2)
    
    If Fg1.Col = 2 Or Fg1.Col = 4 Or Fg1.Col = 6 Or Fg1.Col = 8 Or Fg1.Col = 10 Or Fg1.Col = 12 Or Fg1.Col = 14 Then
        If Fg1.TextMatrix(Fg1.Row, Fg1.Col + 1) = "" Then
            MsgBox "No ha especiticado la cantidad de materia prima a procesar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            CmdCancelar_Click
            Exit Sub
        End If
        
        TxtMP.Text = Fg1.TextMatrix(Fg1.Row, Fg1.Col)
        TxtCan.Text = Fg1.TextMatrix(Fg1.Row, Fg1.Col + 1)
        TxtCan.Text = Format(TxtCan.Text, "0.00")
        xMatPri = TxtMP.Text
    Else
        If Fg1.TextMatrix(Fg1.Row, Fg1.Col) = "" Then
            MsgBox "No ha especiticado la cantidad de materia prima a procesar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            CmdCancelar_Click
            Exit Sub
        End If
        
        TxtMP.Text = Fg1.TextMatrix(Fg1.Row, Fg1.Col - 1)
        TxtCan.Text = Fg1.TextMatrix(Fg1.Row, Fg1.Col)
        TxtCan.Text = Format(TxtCan.Text, "0.00")
        xMatPri = TxtMP.Text
    End If
    
    Dim xNum As Integer
    Dim xFchIni, xFchFin, xDia As Date
    
    xFchIni = TxtFchIni.valor
    xFchFin = TxtFchFin.valor
    
    xNum = NulosN(Mid(Trim(Fg1.TextMatrix(0, Fg1.Col)), Len(Trim(Fg1.TextMatrix(0, Fg1.Col))) - 1, 2))
    xHorPro = CDate(Format(Fg1.TextMatrix(Fg1.Row, 1), "hh:mm"))
    
    For xDia = xFchIni To xFchFin
        If NulosN(Mid(Trim(Str(xDia)), 1, 2)) = xNum Then
            xFchPro = xDia
            Exit For
        End If
    Next xDia
    
    xIdMatPri = Busca_Codigo(xMatPri, "descripcion", "id", "alm_inventario", "C", xCon)
    
    If xIdMatPri = 0 Then
        MsgBox "La materia prima especificada no existe", vbInformation + vbOKOnly + vbDefaultButton1
        Exit Sub
    End If
    
    ' MOSTRAMOS TODOS LOS PRODUCTOS DE LA MATERIA PRIMA
    RST_Busq Rst, "SELECT pro_redimiento.iditem, pro_redimiento.idpro, alm_inventario.descripcion " _
        & " FROM pro_redimiento LEFT JOIN alm_inventario ON pro_redimiento.idpro = alm_inventario.id " _
        & " WHERE (((pro_redimiento.iditem)=" & xIdMatPri & "))", xCon
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        Fg2.Rows = 1
        For A = 1 To Rst.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(A, 1) = Rst("descripcion")
            Fg2.TextMatrix(A, 2) = ""
            Fg2.TextMatrix(A, 3) = 0
            Fg2.TextMatrix(A, 4) = Rst("idpro")
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
        
        If Rst.RecordCount = 1 Then
            'Fg2.SelectionMode = flexSelectionByRow
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = Format(TxtCan.Text, "0.00")
            If QueHace = 3 Then
                Fg2.TextMatrix(Fg2.Rows - 1, 3) = 0
                Fg2.TextMatrix(Fg2.Rows - 1, 2) = ""
            Else
                Fg2.TextMatrix(Fg2.Rows - 1, 3) = 1
            End If
            Fg2.Editable = flexEDNone
        Else
            'Fg2.SelectionMode = flexSelectionFree
            'Fg2.Editable = flexEDKbdMouse
        End If
    End If
    Frame3.Visible = True
    
    ' MOSTRAMOS EL CHECK DE LOS PRODUCTOS QUE SE VAYAN A DEFINIR
    
    RstMatPro.Filter = adFilterNone
    If RstMatPro.RecordCount <> 0 Then
        RstMatPro.Filter = "iditem =" & xIdMatPri & " AND fchpro = " & xFchPro & " AND horpro = " & xHorPro & ""
        If RstMatPro.RecordCount <> 0 Then
            RstMatPro.MoveFirst
            For A = 1 To RstMatPro.RecordCount
                For B = 1 To Fg2.Rows - 1
                    If NulosN(Fg2.TextMatrix(B, 4)) = RstMatPro("idpro") Then
                        Fg2.TextMatrix(B, 3) = 1
                        Fg2.TextMatrix(B, 2) = Format(RstMatPro("cantidad"), "0.00")
                        Exit For
                    End If
                Next B
                RstMatPro.MoveNext
                If RstMatPro.EOF = True Then Exit For
            Next A
        End If
    End If
    TxtTotal.Text = Format(GRID_SUMAR_COL(Fg2, 2, 1, Fg2.Rows - 1), "0.00")
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then MuestraSegundoTab
    End If
End Sub

Sub Eliminar()
    TabOne1.CurrTab = 0
    Dim Rpta As Integer
    
    Rpta = MsgBox("¿Esta seguro de eliminar el cronograma seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM pro_cronograma WHERE id = " & RstLis("id") & ""
        xCon.Execute "DELETE * FROM pro_cronogramadet WHERE id = " & RstLis("id") & ""
        xCon.Execute "DELETE * FROM pro_cronogramatarea WHERE id = " & RstLis("id") & ""
        xCon.Execute "DELETE * FROM pro_cronogramadetprod WHERE id = " & RstLis("id") & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstLis("id") & " AND idform = " & IdMenuActivo
        
        
        RstLis.Requery
        Dg1.Refresh
        MsgBox "El cronograma se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then Modificar
    
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstLis.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then
        TabOne1.CurrTab = 0
        RstLis.Filter = ""
        TDB_FiltroLimpiar Dg1
    End If
    
    If Button.Index = 12 Then imprimir
    
    If Button.Index = 14 Then
        Set RstLis = Nothing
        Unload Me
    End If
End Sub

Sub imprimir()
    Dim Rst As New ADODB.Recordset
    
    If NulosN(RstLis("idtippro")) = 1 Then
        RST_Busq Rst, "TRANSFORM sum(pro_cronogramadetprod.cantidad) AS PromedioDecantidad SELECT pro_cronogramadetprod.iditem, alm_inventario_1.descripcion AS desmatpri, " _
            & " mae_unidades.abrev, alm_inventario.descripcion AS descprod, Sum(pro_cronogramadetprod.cantidad) AS [TotalFila]" _
            & " FROM ((pro_cronogramadetprod LEFT JOIN alm_inventario ON pro_cronogramadetprod.idpro = alm_inventario.id) LEFT JOIN alm_inventario AS alm_inventario_1 " _
            & " ON pro_cronogramadetprod.iditem = alm_inventario_1.id) LEFT JOIN mae_unidades ON alm_inventario_1.idunimed = mae_unidades.id " _
            & " Where (((pro_cronogramadetprod.ID) = " & RstLis("id") & ")) GROUP BY pro_cronogramadetprod.iditem, alm_inventario_1.descripcion, mae_unidades.abrev, " _
            & " alm_inventario.descripcion, pro_cronogramadetprod.id ORDER BY alm_inventario_1.descripcion, alm_inventario.descripcion " _
            & " PIVOT Format([fchpro],'dd-mm-yy')", xCon
    Else
        RST_Busq Rst, "TRANSFORM Sum(pro_cronogramadet.cantidad) AS SumaDecantidad SELECT pro_cronogramadet.iditem, alm_inventario.descripcion, mae_unidades.abrev, " _
            & " Sum(pro_cronogramadet.cantidad) AS TotalFila FROM (pro_cronogramadet LEFT JOIN alm_inventario ON pro_cronogramadet.iditem = alm_inventario.id) " _
            & " LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id Where (((pro_cronogramadet.ID) = " & RstLis("id") & ")) " _
            & " GROUP BY pro_cronogramadet.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_cronogramadet.id PIVOT Format([fchpro],'dd-mm-yy')", xCon
    End If
    
    Dim Li As Integer
    Dim strSource As String
    Dim xArea, xEmp, xDir, xCuerpo, xCad  As String
    Dim xEmpleado As String
    Dim Pagina As Integer
    Dim Lineas As Integer
    
    Set oPDF = New cPDF
    Dim A, B, C As Integer
    xNumPag = 0
    Dim xTipPro As String
    
On Error GoTo Cerrado
    
    If oPDF.PDFCreate(App.Path & "\pro00001.pdf") = True Then
        
        oPDF.Fonts.Add "Tit", Times_BoldItalic, WinAnsiEncoding
        oPDF.Fonts.Add "Head", Times_Italic, WinAnsiEncoding
        oPDF.Fonts.Add "Cont", Courier, WinAnsiEncoding
        oPDF.Fonts.Add "CB", Courier_Bold, WinAnsiEncoding
        oPDF.Fonts.Add "Time", Times_Roman, WinAnsiEncoding
        
        CrearCabecera
        Dim xFilaAct As Integer
        Dim xPosX As Integer
        Dim xFch As Date
        
        oPDF.WTextBox 40, 30, 10, 750, "CRONOGRAMA DE PRODUCCION (" & RstLis("destippro") & ")", "CB", 10, hCenter, vMiddle, vbBlack, 0, vbRed
        oPDF.WTextBox 52, 30, 10, 750, "DEL " & RstLis("fchini") & " AL " & RstLis("fchfin"), "CB", 10, hCenter, vMiddle, vbBlack, 0, vbRed
        
        If NulosN(RstLis("idtippro")) = 1 Then
            oPDF.WTextBox 68, 30, 19, 150, "MATERIA PRIMA", "CB", 8, hCenter, vMiddle, vbBlack, 1, vbBlack
            oPDF.WTextBox 68, 180, 19, 30, "UNI. MED.", "CB", 8, hCenter, vMiddle, vbBlack, 1, vbBlack
            oPDF.WTextBox 68, 210, 19, 250, "PRODUCTO", "CB", 8, hCenter, vMiddle, vbBlack, 1, vbBlack
            xPosX = 460
        Else
            oPDF.WTextBox 68, 30, 19, 250, "PRODUCTO", "CB", 8, hCenter, vMiddle, vbBlack, 1, vbBlack
            oPDF.WTextBox 68, 280, 19, 30, "UNI. MED.", "CB", 8, hCenter, vMiddle, vbBlack, 1, vbBlack
            xPosX = 310
        End If
        
        ' IMPRIMIMOS EL ROTULO DE LAS FECHAS
        For xFch = RstLis("fchini") To RstLis("fchfin")
            oPDF.WTextBox 68, xPosX, 19, 45, Format(xFch, "dd/mm/yy"), "CB", 8, hCenter, vMiddle, vbBlack, 1, vbBlack
            
            xPosX = xPosX + 45
        Next xFch
        
        ' IMPRIMIMOS EL ROTULO DEL TOTAL
        oPDF.WTextBox 68, xPosX, 19, 45, "TOTAL", "CB", 8, hCenter, vMiddle, vbBlack, 1, vbBlack
                 
        xFilaInicial = 88
        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst
            For A = 1 To Rst.RecordCount
                If NulosN(RstLis("idtippro")) = 1 Then
                    oPDF.WTextBox xFilaInicial, 30, 10, 150, Rst("desmatpri"), "CB", 8, hLeft, vMiddle, vbBlack, 0, vbBlack
                    oPDF.WTextBox xFilaInicial, 180, 10, 30, Rst("abrev"), "CB", 8, hCenter, vMiddle, vbBlack, 0, vbRed
                    oPDF.WTextBox xFilaInicial, 210, 10, 250, Rst("descprod"), "CB", 8, hLeft, vMiddle, vbBlack, 0, vbRed
                    xPosX = 460
                Else
                    oPDF.WTextBox xFilaInicial, 30, 10, 250, Rst("descripcion"), "CB", 8, hLeft, vMiddle, vbBlack, 0, vbBlack
                    oPDF.WTextBox xFilaInicial, 280, 10, 30, Rst("abrev"), "CB", 8, hCenter, vMiddle, vbBlack, 0, vbRed
                    xPosX = 310
                End If
                
                For xFch = RstLis("fchini") To RstLis("fchfin")
                    If RstRegistroBuscaCampo(Rst, Format(xFch, "dd-mm-yy")) = True Then
                        oPDF.WTextBox xFilaInicial, xPosX, 10, 45, Format(NulosN(Rst(Format(xFch, "dd-mm-yy"))), "0.00"), "CB", 8, hRight, vMiddle, vbBlack, 0, vbBlack
                    End If
                    xPosX = xPosX + 45
                Next xFch
                
                ' IMPRIMIMOS EL TOTAL DE LA FILA
                oPDF.WTextBox xFilaInicial, xPosX, 10, 45, Format(NulosN(Rst("TotalFila")), "0.00"), "CB", 8, hRight, vMiddle, vbBlack, 0, vbBlack
                xFilaInicial = xFilaInicial + 10
                
                Rst.MoveNext
                If Rst.EOF = True Then Exit For
            Next A
        End If
        
        oPDF.PDFClose
        Set oPDF = Nothing
        Shell ("rundll32.exe url.dll,FileProtocolHandler " & Trim(App.Path) & ("\pro00001.pdf")), vbMaximizedFocus
    Else
        Set oPDF = Nothing
        MsgBox "No se Puede Mostrar Documento pro00001.pdf, psoblemente el archivo ya se encuentra abierto", vbCritical, "Error"
    End If
    Exit Sub
    
Cerrado:
    'Resume
    If Err.Number = 1 Then
    End If
End Sub

Sub CrearCabecera()
    Dim xTelEmp, xNumDoc As String
    
    'oPDF.NewPage A4_Vertical ', 525, 675
    oPDF.NewPage A4_Horizontal  ', 525, 675
    xNumPag = xNumPag + 1
    
    oPDF.WTextBox 15, 30, 8, 50, "EMPRESA", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 15, 105, 8, 5, ":", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 15, 111, 8, 150, NomEmp, "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    
    oPDF.WTextBox 23, 30, 8, 50, "Nº R.U.C.", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 23, 105, 8, 5, ":", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 23, 111, 8, 100, NumRUC, "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    
    oPDF.WTextBox 15, 700, 8, 50, "Nº PAGINA", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 15, 750, 8, 5, ":", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 15, 753, 8, 50, Format(xNumPag, "000"), "CB", 8, hRight, vTop, RGB(0, 0, 128), , vbRed
    
    oPDF.WTextBox 23, 700, 8, 50, "FCH. IMPR", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 23, 750, 8, 5, ":", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 23, 753, 8, 50, Format(Date, "dd/mm/yy"), "CB", 8, hRight, vTop, RGB(0, 0, 128), , vbRed
    
    oPDF.WLineTo 30, 36, 800, 36
    oPDF.LineStroke
End Sub

Private Sub TxtIdSup_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtIdSup_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusSup_Click
    End If
End Sub

Private Sub TxtIdSup_Validate(Cancel As Boolean)
    If NulosN(TxtIdSup.Text) = 0 Then
        TxtIdSup.Text = ""
        Exit Sub
    Else
        Dim Rst As New ADODB.Recordset
        Dim xSqlCad As String
        xSqlCad = "SELECT pro_emp.*, pla_empleados.nombre, pro_emp.id " _
            & " FROM (pro_emp LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id " _
            & " Where (((pro_empdet.idfun) = 2) And ((pro_emp.ID) = " & Val(TxtIdSup.Text) & ")) ORDER BY pla_empleados.nombre"

        Set Rst = BuscaConCriterio(xSqlCad, xCon)
        
        If Rst.RecordCount <> 0 Then
            LblSupervisor.Caption = Rst("nombre")
        Else
            TxtIdSup.Text = ""
            LblSupervisor.Caption = ""
        End If
        
        Set Rst = Nothing
    End If
End Sub

Private Sub TxtTipPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtTipPro_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTip_Click
    End If
End Sub

Private Sub TxtTipPro_Validate(Cancel As Boolean)
    If NulosN(TxtTipPro.Text) = 0 Then
        TxtTipPro.Text = ""
        Exit Sub
    Else
        LblTipoProd.Caption = Busca_Codigo(TxtTipPro.Text, "id", "descripcion", " mae_tipoproducto", "N", xCon)
        If NulosC(LblTipoProd.Caption) = "" Then
            TxtTipPro.Text = ""
        End If
    End If
End Sub

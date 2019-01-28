VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmCronoTarea 
   Caption         =   "Produccion - Programacion de Tareas"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6465
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   5745
      Left            =   15
      TabIndex        =   0
      Top             =   360
      Width           =   11805
      _cx             =   20823
      _cy             =   10134
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
      Caption         =   "  Consulta  |   Detalle  "
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
      Begin SizerOneLibCtl.ElasticOne ElasticOne2 
         Height          =   5310
         Left            =   45
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   390
         Width           =   11715
         _cx             =   20664
         _cy             =   9366
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
         GridRows        =   3
         GridCols        =   1
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmCronoTarea.frx":0000
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Height          =   645
            Left            =   30
            TabIndex        =   8
            Top             =   4635
            Width           =   11655
            Begin VB.Label Label4 
               Caption         =   "Tarea =>"
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
               Left            =   60
               TabIndex        =   12
               Top             =   375
               Width           =   1560
            End
            Begin VB.Label LblTarea 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblTarea"
               Height          =   300
               Left            =   1695
               TabIndex        =   11
               Top             =   330
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
               Left            =   60
               TabIndex        =   10
               Top             =   60
               Width           =   1560
            End
            Begin VB.Label LblProducto 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblProducto"
               Height          =   300
               Left            =   1695
               TabIndex        =   9
               Top             =   15
               Width           =   9900
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   4230
            Left            =   30
            TabIndex        =   3
            Top             =   375
            Width           =   11655
            _cx             =   20558
            _cy             =   7461
            _ConvInfo       =   1
            Appearance      =   1
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
            Cols            =   20
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmCronoTarea.frx":004F
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
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Programacion"
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
            Height          =   315
            Left            =   30
            TabIndex        =   5
            Top             =   30
            Width           =   11655
         End
      End
      Begin SizerOneLibCtl.ElasticOne ElasticOne1 
         Height          =   5310
         Left            =   -12360
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   390
         Width           =   11715
         _cx             =   20664
         _cy             =   9366
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
         _GridInfo       =   $"FrmCronoTarea.frx":029B
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   4905
            Left            =   30
            TabIndex        =   7
            Top             =   375
            Width           =   11655
            _ExtentX        =   20558
            _ExtentY        =   8652
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
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Consulta"
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
            Height          =   315
            Left            =   30
            TabIndex        =   6
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
            Picture         =   "FrmCronoTarea.frx":02DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoTarea.frx":0821
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoTarea.frx":09A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoTarea.frx":0DF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoTarea.frx":0F11
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoTarea.frx":1455
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoTarea.frx":1999
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoTarea.frx":1AAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoTarea.frx":1BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoTarea.frx":2015
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoTarea.frx":2181
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   4
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
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
            Style           =   5
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
End
Attribute VB_Name = "FrmCronoTarea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim Agregando As Boolean
Dim RstLis As New ADODB.Recordset
Dim fOrdenLista As Boolean                                ' especfica el orden de la lista de la consulta
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO


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

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    'If QueHace <> 3 Then Exit Sub
    If Agregando = True Then Exit Sub
    Dim xTiempo As Double
    Dim xHorEst As String
    Dim xPorcentaje As Double
    Dim xTotal As Double
    
    If Col = 11 Then
        Fg1.TextMatrix(Fg1.Row, 11) = Format(Fg1.TextMatrix(Fg1.Row, 11), "00")
        ' ((factor * totprocs)/numpers)
        xTiempo = (Fg1.TextMatrix(Fg1.Row, 4) * Fg1.TextMatrix(Fg1.Row, 15)) / NulosN(Fg1.TextMatrix(Fg1.Row, 11))
        
        xHorEst = Format(Int(xTiempo), "00")
        xHorEst = xHorEst & ":" & Format(((xTiempo * 60) Mod 60), "00")
        
        Fg1.TextMatrix(Fg1.Row, 16) = Format(xHorEst, "hh:mm")
    End If
    
    If Col = 14 Then
        If NulosN(Fg1.TextMatrix(Fg1.Row, 14)) <> 0 Then
            Fg1.TextMatrix(Fg1.Row, 14) = Format(Fg1.TextMatrix(Fg1.Row, 14), "0.00")
            xPorcentaje = NulosN(Fg1.TextMatrix(Fg1.Row, 14)) / 100
            xTotal = NulosN(Fg1.TextMatrix(Fg1.Row, 13)) * xPorcentaje
            Fg1.TextMatrix(Fg1.Row, 15) = Format(xTotal, "0.00")
        Else
            Fg1.TextMatrix(Fg1.Row, 15) = Fg1.TextMatrix(Fg1.Row, 13)
        End If
        Fg1_CellChanged Fg1.Row, 11
    End If
    
    'Dim xHorEst   As String
    xHorEst = Fg1.TextMatrix(Fg1.Row, 16)
    If Val(Mid(xHorEst, 1, 2)) <= 8 Then
        Fg1.TextMatrix(Fg1.Row, 18) = Format(CDate(xHorEst) + CDate(Fg1.TextMatrix(Fg1.Row, 17)), "HH:MM")
    Else
        Fg1.TextMatrix(Fg1.Row, 18) = ""
    End If

End Sub

Private Sub Fg1_EnterCell()
    Fg1.Editable = flexEDNone
    If QueHace = 3 Then Exit Sub
    
    If Fg1.Col = 9 Or Fg1.Col = 11 Or Fg1.Col = 14 Or Fg1.Col = 17 Or Fg1.Col = 19 Then
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Fg1_RowColChange()
    LblProducto.Caption = Fg1.TextMatrix(Fg1.Row, 7)
    LblTarea.Caption = Fg1.TextMatrix(Fg1.Row, 8)
End Sub

Private Sub Form_Activate()
    
    If SeEjecuto = False Then
    
        SeEjecuto = True
            
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
    
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
    
        '--ocultar el boton a agregar
        Toolbar1.Buttons(1).Visible = False
        
        RST_Busq RstLis, "SELECT pro_cronograma.*, mae_tipoproducto.descripcion AS destippro, " _
            & " [pla_empleados]![apepat] & ' ' & [pla_empleados]![apemat] & ', ' & [pla_empleados]![nom] AS apenom " _
            & " FROM (pla_empleados RIGHT JOIN (pro_cronograma LEFT JOIN pro_emp ON pro_cronograma.idsup = pro_emp.id) ON pla_empleados.id = pro_emp.idemp) " _
            & " LEFT JOIN mae_tipoproducto ON pro_cronograma.idtippro = mae_tipoproducto.id ORDER BY pro_cronograma.fchini", xCon
        
        Set Dg1.DataSource = RstLis

    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    QueHace = 3
    ConfiguraDrid
    Frame1.BackColor = &H8000000F
    TabOne1.CurrTab = 0
    Fg1.Editable = flexEDNone
    Me.Height = 8000
    Me.Width = 12000
End Sub

Sub ConfiguraDrid()
    Fg1.ColWidth(1) = 0
    Fg1.ColWidth(2) = 0
    Fg1.ColWidth(3) = 0
    Fg1.ColWidth(4) = 0
    Fg1.ColWidth(5) = 0
    
    Fg1.ColEditMask(9) = "##/##/##"   'HOR. PRO.
    Fg1.ColEditMask(19) = "##/##/##"   'FCH. INI.
    
    Fg1.ColEditMask(10) = "##:##"   'HOR. PRO.
    Fg1.ColEditMask(12) = "##:##"   'EMP. EN.
    Fg1.ColEditMask(16) = "##:##"   'TIEMPO EST.
    Fg1.ColEditMask(17) = "##:##"   'HOR. INI. TAR.
    Fg1.ColEditMask(18) = "##:##"   'HOR. TER. TAR.
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

Sub Cancelar()
    QueHace = 3
    TabOne1.TabEnabled(0) = True
    ActivaTool
    Label1.Caption = "Detalle de la Programacion de Tareas"
    Fg1.Editable = flexEDNone
    TabOne1.CurrTab = 0
End Sub

Sub Bloquea()

End Sub

Sub MostrarSegundoTab()
    Dim RstDet As New ADODB.Recordset
    Dim xCadSQL As String
    
    If RstLis("idtippro") = 1 Then
        Fg1.ColWidth(6) = 1740
    Else
        Fg1.ColWidth(6) = 0
    End If
    
    Fg1.Rows = 1
    
    xCadSQL = "SELECT pro_cronogramatarea.*, alm_inventario.descripcion AS matpri, alm_inventario_1.descripcion AS despro, pro_tareas.descripcion AS destar " _
        & " FROM ((pro_cronogramatarea LEFT JOIN alm_inventario ON pro_cronogramatarea.iditem = alm_inventario.id) LEFT JOIN alm_inventario AS alm_inventario_1 " _
        & " ON pro_cronogramatarea.idpro = alm_inventario_1.id) LEFT JOIN pro_tareas ON pro_cronogramatarea.idtar = pro_tareas.id Where (((pro_cronogramatarea.id) = " & RstLis("id") & ")) " _
        & " ORDER BY pro_cronogramatarea.fchpro, pro_cronogramatarea.horpro, alm_inventario.descripcion, alm_inventario_1.descripcion, pro_cronogramatarea.orden"

    RST_Busq RstDet, xCadSQL, xCon
    
    If RstDet.RecordCount = 0 Then
        MsgBox "No se ha procesado el cronograma actual, haga clic en el modificar para procesar el cronograma", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstDet.Close
        Set RstDet = Nothing
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    
    
    Dim A As Integer
    If RstDet.RecordCount <> 0 Then
        RstDet.MoveFirst
        Agregando = True
        Dim xCadLlave As String
        Dim xCadLlave2 As String
        Dim xColor As Long
        Dim xPintar As Boolean
        
        xCadLlave = Format(RstDet("iditem"), "0000") & Format(RstDet("idpro"), "0000") & Format(RstDet("fchpro"), "dd/mm/yy") & Format(RstDet("horpro"), "hh:mm")
        xColor = &HE0FEFE
        xPintar = True
        
        For A = 1 To RstDet.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = RstDet("iditem")
            Fg1.TextMatrix(A, 2) = RstDet("idpro")
            Fg1.TextMatrix(A, 3) = RstDet("idtar")
            Fg1.TextMatrix(A, 4) = RstDet("factor")
            Fg1.TextMatrix(A, 5) = RstDet("orden")
            Fg1.TextMatrix(A, 6) = NulosC(RstDet("matpri"))
            Fg1.TextMatrix(A, 7) = RstDet("despro")
            Fg1.TextMatrix(A, 8) = RstDet("destar")
            Fg1.TextMatrix(A, 9) = Format(RstDet("fchpro"), "dd/mm/yy")
            Fg1.TextMatrix(A, 10) = Format(RstDet("horpro"), "hh:mm")
            Fg1.TextMatrix(A, 11) = Format(RstDet("numper"), "00")
            Fg1.TextMatrix(A, 12) = Format(RstDet("horarr"), "hh:mm")
            Fg1.TextMatrix(A, 13) = Format(RstDet("cantidad"), "0.00")
            Fg1.TextMatrix(A, 14) = Format(RstDet("aplpor"), "0.00")
            
            Fg1.TextMatrix(A, 19) = Format(RstDet("fchini"), "dd/mm/yy")
            
            Dim xTiempo As Double
            Dim xHorEst As String
            xTiempo = 1
            
            If NulosN(RstDet("aplpor")) <> 0 Then
                Fg1.TextMatrix(A, 15) = (RstDet("cantidad") * ((RstDet("aplpor") / 100)))
                Fg1.TextMatrix(A, 15) = Format(Fg1.TextMatrix(A, 15), "0.00")
                
                xTiempo = (RstDet("factor") * Fg1.TextMatrix(A, 15)) / RstDet("numper")
            Else
                Fg1.TextMatrix(A, 15) = Format(RstDet("cantidad"), "0.00")
                xTiempo = (RstDet("factor") * RstDet("cantidad")) / RstDet("numper")
            End If
            xHorEst = ""
            xHorEst = Format(Int(xTiempo), "00")
            xHorEst = xHorEst & ":" & Format(((xTiempo * 60) Mod 60), "00")
            Fg1.TextMatrix(A, 16) = xHorEst
            
            Fg1.TextMatrix(A, 17) = Format(RstDet("horinitar"), "hh:mm")
            
            If Val(Mid(xHorEst, 1, 2)) <= 8 Then
                Fg1.TextMatrix(A, 18) = Format(CDate(xHorEst) + CDate(Fg1.TextMatrix(A, 17)), "HH:MM")
            Else
                Fg1.TextMatrix(A, 18) = ""
            End If
            If xPintar = True Then
                GRID_COLOR_FONDO Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 19, xColor, flexFillRepeat
            End If
            
            RstDet.MoveNext
            If RstDet.EOF = True Then
                Exit For
            End If
            xCadLlave2 = Format(RstDet("iditem"), "0000") & Format(RstDet("idpro"), "0000") & Format(RstDet("fchpro"), "dd/mm/yy") & Format(RstDet("horpro"), "hh:mm")
            If xCadLlave2 <> xCadLlave Then
                xCadLlave = xCadLlave2
                xPintar = Not xPintar
            End If
        Next A
        
        Agregando = False
    End If
    
    Fg1.Select 1, 7
    LblProducto.Caption = Fg1.TextMatrix(Fg1.Row, 7)
    LblTarea.Caption = Fg1.TextMatrix(Fg1.Row, 8)

End Sub

Private Sub Form_Resize()
    Dim TopEO As Integer
    
    TopEO = 400
    
    If Me.WindowState = 1 Then Exit Sub
    
    If Me.Height <= (TopEO + 2375) Then
        Me.Height = (TopEO + 2375)
    Else
        'EO.Height = (Me.Height - (TopEO + 400))
    End If
    
    TabOne1.Height = (Me.Height - 760)
    TabOne1.Width = Me.Width - 130
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 3 Then
            MostrarSegundoTab
        End If
    End If
End Sub

Function Grabar() As Boolean
    
    Dim xId As Double
    Dim xCampos(4, 4) As String
    Dim xCampos2(14, 4) As String
    Dim A As Integer
    
On Error GoTo LaCague
    xCon.BeginTrans
    
    xId = RstLis("id")
    
    xCon.Execute "DELETE * FROM pro_cronogramatarea WHERE id = " & xId & ""
    
    For A = 1 To Fg1.Rows - 1
        xCampos2(0, 0) = "id":           xCampos2(0, 1) = xId:            xCampos2(0, 2) = "S":    xCampos2(0, 3) = "N":     xCampos2(0, 4) = "":
        xCampos2(1, 0) = "fchpro":       xCampos2(1, 1) = Fg1.TextMatrix(A, 9):    xCampos2(1, 2) = "S":    xCampos2(1, 3) = "F":     xCampos2(1, 4) = "No ha especificado la fecha de programacion"
        xCampos2(2, 0) = "horpro":       xCampos2(2, 1) = Fg1.TextMatrix(A, 10):   xCampos2(2, 2) = "S":    xCampos2(2, 3) = "F":     xCampos2(2, 4) = "No ha especificado la hora de recepcion"
        xCampos2(3, 0) = "iditem":       xCampos2(3, 1) = Fg1.TextMatrix(A, 1):    xCampos2(3, 2) = "S":    xCampos2(3, 3) = "N":     xCampos2(3, 4) = "No ha especificado la materia prima"
        xCampos2(4, 0) = "idpro":        xCampos2(4, 1) = Fg1.TextMatrix(A, 2):    xCampos2(4, 2) = "S":    xCampos2(4, 3) = "N":     xCampos2(4, 4) = "No ha especificado el producto"
        xCampos2(5, 0) = "idtar":        xCampos2(5, 1) = Fg1.TextMatrix(A, 3):    xCampos2(5, 2) = "S":    xCampos2(5, 3) = "N":     xCampos2(5, 4) = "No ha especificado la tarea"
        xCampos2(6, 0) = "cantidad":     xCampos2(6, 1) = Fg1.TextMatrix(A, 13):   xCampos2(6, 2) = "S":    xCampos2(6, 3) = "N":     xCampos2(6, 4) = "No ha especificado la cantidad"
        xCampos2(7, 0) = "factor":       xCampos2(7, 1) = Fg1.TextMatrix(A, 4):    xCampos2(7, 2) = "S":    xCampos2(7, 3) = "N":     xCampos2(7, 4) = "No ha especificado el factor"
        xCampos2(8, 0) = "costokg":      xCampos2(8, 1) = 0:                       xCampos2(8, 2) = "S":    xCampos2(8, 3) = "N":     xCampos2(8, 4) = "No ha especificado el costo por kilo"
        xCampos2(9, 0) = "numper":       xCampos2(9, 1) = Fg1.TextMatrix(A, 11):   xCampos2(9, 2) = "S":    xCampos2(9, 3) = "N":     xCampos2(9, 4) = "No ha especificado el numero de personas"
        xCampos2(10, 0) = "horarr":      xCampos2(10, 1) = Fg1.TextMatrix(A, 12):  xCampos2(10, 2) = "S":    xCampos2(10, 3) = "F":     xCampos2(10, 4) = "No ha especificado tiempo en que enpieza cada tarea"
        xCampos2(11, 0) = "aplpor":      xCampos2(11, 1) = Fg1.TextMatrix(A, 14):  xCampos2(11, 2) = "S":    xCampos2(11, 3) = "N":     xCampos2(11, 4) = "No ha especificado el porcentaje de rendimiento"
        xCampos2(12, 0) = "orden":       xCampos2(12, 1) = Fg1.TextMatrix(A, 5):   xCampos2(12, 2) = "S":    xCampos2(12, 3) = "N":     xCampos2(12, 4) = "No ha especificado el orden"
        xCampos2(13, 0) = "horinitar":   xCampos2(13, 1) = Fg1.TextMatrix(A, 17):  xCampos2(13, 2) = "S":    xCampos2(13, 3) = "F":     xCampos2(13, 4) = "No ha especificado la hora de inicio de la tarea"
        xCampos2(14, 0) = "fchini":      xCampos2(14, 1) = Fg1.TextMatrix(A, 19):  xCampos2(14, 2) = "S":    xCampos2(14, 3) = "F":     xCampos2(14, 4) = "No ha especificado la fecha de inicio"
        
        If EscribirNuevoRegistro(xCampos2, "pro_cronogramatarea", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
    Next A
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    
    xCon.CommitTrans
    MsgBox "El cronograma se guardo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Exit Function
    
LaCague:
    xCon.RollbackTrans
    MsgBox "No se pudo guardar el registro por el siguiente error : " & Trim(Err.Description)
    Grabar = False
End Function

Sub Modificar()
    Dim xRs As New ADODB.Recordset
    Dim RstTar As New ADODB.Recordset
    Dim SQLCad As String

    RST_Busq xRs, "SELECT * FROM pro_cronogramatarea WHERE id = " & RstLis("id") & "", xCon
    
    If xRs.RecordCount = 0 Then
        If NulosN(RstLis("idtippro")) = 3 Then
            ' SI SE ESTAN PROCESANDO PRODUCTOS
            SQLCad = "SELECT pro_cronogramadet.id, pro_cronogramadet.fchpro, pro_cronogramadet.Horpro, 0 AS iditem, pro_cronogramadet.iditem AS idpro, " _
                & " pro_receta.descripcion AS nomrec, '' AS nommatpri, alm_inventario.descripcion AS nompro, pro_recetatar.idtar, pro_tareas.codigo AS codtar, " _
                & " pro_tareas.descripcion AS destar, pro_cronogramadet.cantidad, pro_recetatar.factor, pro_recetatar.costokg, pro_recetatar.numper, pro_recetatar.horarr, " _
                & " pro_recetatar.aplpor, pro_recetatar.orden, IIf(pro_recetatar.aplpor=0,pro_recetatar.factor*pro_cronogramadet.cantidad,(pro_recetatar.factor*pro_cronogramadet.cantidad)*(pro_recetatar.aplpor/100)) AS tiempoesttotal, " _
                & " [tiempoesttotal]/[numper] AS tiempoesttotalper, [Horpro]+pro_recetatar.horarr AS horinitar, Format(Int(([tiempoesttotal]*60)/60),'00') & ':' & Format(([tiempoesttotal]*60) Mod 60,'00') AS tiempoesttotalhor, " _
                & " Format(Int(([tiempoesttotalper]*60)/60),'00') & ':' & Format(([tiempoesttotalper]*60) Mod 60,'00') AS tiempoesttotalhorperr, pro_cronogramadet.fchpro AS fchini " _
                & " FROM pro_tareas RIGHT JOIN (((pro_cronogramadet LEFT JOIN pro_receta ON pro_cronogramadet.iditem = pro_receta.iditem) LEFT JOIN alm_inventario " _
                & " ON pro_receta.iditem = alm_inventario.id) LEFT JOIN pro_recetatar ON pro_receta.id = pro_recetatar.idrec) ON pro_tareas.id = pro_recetatar.idtar" _
                & " Where (((pro_cronogramadet.id) = " & RstLis("id") & ") And ((pro_recetatar.factor) <> 0) And ((pro_recetatar.numper) <> 0) And ((pro_receta.prirec) = 1)) " _
                & " ORDER BY pro_cronogramadet.fchpro, pro_cronogramadet.Horpro, pro_receta.descripcion, pro_recetatar.orden"
        Else
            ' SI SE ESTA PROCESANDO MATERIA PRIMA
            SQLCad = "SELECT pro_cronogramadetprod.id, pro_cronogramadetprod.fchpro, pro_cronogramadetprod.horpro, pro_cronogramadetprod.iditem, pro_cronogramadetprod.idpro, " _
                & " pro_receta.descripcion AS nomrec, alm_inventario_1.descripcion AS nommatpri, alm_inventario.descripcion AS nompro, pro_recetatar.idtar, pro_tareas.codigo AS codtar, " _
                & " pro_tareas.descripcion AS destar, pro_cronogramadetprod.cantidad, pro_recetatar.factor, pro_recetatar.costokg, pro_recetatar.numper, pro_recetatar.horarr, " _
                & " pro_recetatar.aplpor, pro_recetatar.orden, IIf([pro_recetatar].[aplpor]=0,[pro_recetatar].[factor]*[pro_cronogramadetprod].[cantidad],([pro_recetatar].[factor]*[pro_cronogramadetprod].[cantidad])*([pro_recetatar].[aplpor]/100)) AS tiempoesttotal, " _
                & " [tiempoesttotal]/[numper] AS tiempoesttotalper, [pro_cronogramadetprod].[Horpro]+[pro_recetatar].[horarr] AS horinitar, " _
                & " Format(Int(([tiempoesttotal]*60)/60),'00') & ':' & Format(([tiempoesttotal]*60) Mod 60,'00') AS tiempoesttotalhor, " _
                & " Format(Int(([tiempoesttotalper]*60)/60),'00') & ':' & Format(([tiempoesttotalper]*60) Mod 60,'00') AS tiempoesttotalhorperr, pro_cronogramadetprod.fchpro AS fchini, " _
                & " pro_receta.prirec FROM ((((alm_inventario AS alm_inventario_1 RIGHT JOIN pro_cronogramadetprod ON alm_inventario_1.id = pro_cronogramadetprod.iditem) " _
                & " LEFT JOIN pro_receta ON pro_cronogramadetprod.idpro = pro_receta.iditem) LEFT JOIN alm_inventario ON pro_cronogramadetprod.idpro = alm_inventario.id) " _
                & " LEFT JOIN pro_recetatar ON pro_receta.id = pro_recetatar.idrec) LEFT JOIN pro_tareas ON pro_recetatar.idtar = pro_tareas.id " _
                & " Where (((pro_cronogramadetprod.id) = " & RstLis("id") & ") And ((pro_recetatar.factor) <> 0) And ((pro_recetatar.numper) <> 0) And ((pro_receta.prirec) = 1)) " _
                & " ORDER BY pro_cronogramadetprod.fchpro, pro_cronogramadetprod.horpro, alm_inventario.descripcion, pro_recetatar.orden"
        End If
        
        RST_Busq RstTar, SQLCad, xCon
        
        Fg1.Rows = 1
        Dim A As Integer
        If RstTar.RecordCount <> 0 Then
            RstTar.MoveFirst
            Agregando = True
            Dim xCadLlave As String
            Dim xCadLlave2 As String
            Dim xColor As Long
            Dim xPintar As Boolean
            
            xCadLlave = Format(RstTar("iditem"), "0000") & Format(RstTar("idpro"), "0000") & Format(RstTar("fchpro"), "dd/mm/yy") & Format(RstTar("horpro"), "hh:mm")
            xColor = &HE0FEFE
            xPintar = True
            
            For A = 1 To RstTar.RecordCount
                Fg1.Rows = Fg1.Rows + 1
                Fg1.TextMatrix(A, 1) = RstTar("iditem")
                Fg1.TextMatrix(A, 2) = RstTar("idpro")
                Fg1.TextMatrix(A, 3) = RstTar("idtar")
                Fg1.TextMatrix(A, 4) = RstTar("factor")
                Fg1.TextMatrix(A, 5) = RstTar("orden")
                Fg1.TextMatrix(A, 6) = RstTar("nommatpri")
                Fg1.TextMatrix(A, 7) = RstTar("nompro")
                Fg1.TextMatrix(A, 8) = RstTar("destar")
                Fg1.TextMatrix(A, 9) = Format(RstTar("fchpro"), "dd/mm/yy")
                Fg1.TextMatrix(A, 10) = Format(RstTar("horpro"), "hh:mm")
                Fg1.TextMatrix(A, 11) = Format(RstTar("numper"), "00")
                Fg1.TextMatrix(A, 12) = Format(RstTar("horarr"), "hh:mm")
                Fg1.TextMatrix(A, 13) = Format(RstTar("cantidad"), "0.00")
                Fg1.TextMatrix(A, 14) = Format(RstTar("aplpor"), "0.00")
                
                Fg1.TextMatrix(A, 19) = Format(RstTar("fchini"), "dd/mm/yy")
                
                Dim xTiempo As Double
                Dim xHorEst As String
                xTiempo = 1
                
                If NulosN(RstTar("aplpor")) <> 0 Then
                    Fg1.TextMatrix(A, 15) = (RstTar("cantidad") * ((RstTar("aplpor") / 100)))
                    Fg1.TextMatrix(A, 15) = Format(Fg1.TextMatrix(A, 15), "0.00")
                    
                    xTiempo = (RstTar("factor") * Fg1.TextMatrix(A, 15)) / RstTar("numper")
                Else
                    Fg1.TextMatrix(A, 15) = Format(RstTar("cantidad"), "0.00")
                    xTiempo = (RstTar("factor") * RstTar("cantidad")) / RstTar("numper")
                End If
                xHorEst = ""
                xHorEst = Format(Int(xTiempo), "00")
                xHorEst = xHorEst & ":" & Format(((xTiempo * 60) Mod 60), "00")
                Fg1.TextMatrix(A, 16) = xHorEst
                
                Fg1.TextMatrix(A, 17) = Format(RstTar("horinitar"), "hh:mm")
                
                If Val(Mid(xHorEst, 1, 2)) <= 8 Then
                    Fg1.TextMatrix(A, 18) = Format(CDate(xHorEst) + CDate(Fg1.TextMatrix(A, 17)), "HH:MM")
                Else
                    Fg1.TextMatrix(A, 18) = ""
                End If
                If xPintar = True Then
                    GRID_COLOR_FONDO Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 19, xColor, flexFillRepeat
                End If
                
                RstTar.MoveNext
                If RstTar.EOF = True Then
                    Exit For
                End If
                xCadLlave2 = Format(RstTar("iditem"), "0000") & Format(RstTar("idpro"), "0000") & Format(RstTar("fchpro"), "dd/mm/yy") & Format(RstTar("horpro"), "hh:mm")
                If xCadLlave2 <> xCadLlave Then
                    xCadLlave = xCadLlave2
                    xPintar = Not xPintar
                End If
            Next A
            
            Agregando = False
        End If
        
    Else
        MostrarSegundoTab
    End If
    
    'MostrarSegundoTab RstTar
    
    QueHace = 2
    
    xHorIni = Time
    
    TabOne1.TabEnabled(0) = False
    TabOne1.CurrTab = 1
    Label1.Caption = "Modificando Programacion de Tareas"
    ActivaTool
    Bloquea
    Fg1.Editable = flexEDKbdMouse
End Sub

Sub Eliminar()
    TabOne1.CurrTab = 0
    
    If RstLis.RecordCount = 0 Then
        MsgBox "No hay registro para eliminar", vbInformation, xTitulo
        Exit Sub
    End If
    
    
    Dim Rpta As Integer
    Rpta = MsgBox("¿Esta seguro de eliminar las tareas del cronograma seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM pro_cronogramatarea WHERE id = " & RstLis("id") & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstLis("id") & " AND idform = " & IdMenuActivo
                
        MsgBox "El cronograma de tareas se elimino con exito"
        RstLis.Requery
        Dg1.Refresh
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
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
    
    If Button.Index = 14 Then
        RstLis.Close
        Set RstLis = Nothing
        Unload Me
    End If
End Sub

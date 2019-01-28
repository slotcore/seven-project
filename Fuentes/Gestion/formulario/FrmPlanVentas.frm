VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmPlanVentas 
   Caption         =   "Sistema de Ventas - Plan de Ventas"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7695
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   11895
      _cx             =   20981
      _cy             =   12938
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
         Height          =   6915
         Left            =   45
         TabIndex        =   17
         Top             =   375
         Width           =   11805
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6570
            Left            =   30
            TabIndex        =   18
            Top             =   375
            Width           =   11760
            _ExtentX        =   20743
            _ExtentY        =   11589
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nº Proyecto"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Descripcion"
            Columns(1).DataField=   "descripcion"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Ini"
            Columns(2).DataField=   "fchini"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Fch. Fin"
            Columns(3).DataField=   "fchfin"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Estado"
            Columns(4).DataField=   "estado"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2381"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2302"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=8202"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=8123"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1826"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1746"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=1799"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1720"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=1667"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1588"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HDBFDFD&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H400000&"
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
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
            Caption         =   "Consulta Plan de Ventas"
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
            Left            =   105
            TabIndex        =   19
            Top             =   30
            Width           =   11595
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   6915
         Left            =   12540
         TabIndex        =   1
         Top             =   375
         Width           =   11805
         Begin VB.Frame Frame3 
            Height          =   645
            Left            =   8265
            TabIndex        =   22
            Top             =   825
            Width           =   3540
            Begin VB.CommandButton CmdDelPro 
               Caption         =   "&Eliminar Producto"
               Enabled         =   0   'False
               Height          =   390
               Left            =   1770
               TabIndex        =   24
               Top             =   180
               Width           =   1560
            End
            Begin VB.CommandButton CmdAddPro 
               Caption         =   "&Agregar Producto"
               Enabled         =   0   'False
               Height          =   390
               Left            =   195
               TabIndex        =   23
               Top             =   180
               Width           =   1560
            End
         End
         Begin VB.TextBox TxtDesc 
            Height          =   300
            Left            =   1155
            Locked          =   -1  'True
            TabIndex        =   2
            Text            =   "TxtDesc"
            Top             =   510
            Width           =   9780
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
            Height          =   300
            Left            =   1155
            TabIndex        =   3
            Top             =   825
            Width           =   1365
            _ExtentX        =   2408
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
            Enabled         =   0   'False
            Valor           =   "06/02/2006"
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
            Height          =   300
            Left            =   5070
            TabIndex        =   4
            Top             =   825
            Width           =   1365
            _ExtentX        =   2408
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
            Enabled         =   0   'False
            Valor           =   "06/02/2006"
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   2640
            Left            =   45
            TabIndex        =   5
            Top             =   1485
            Width           =   11745
            _cx             =   20717
            _cy             =   4657
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
            Rows            =   1
            Cols            =   17
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmPlanVentas.frx":0000
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
            Height          =   1830
            Left            =   45
            TabIndex        =   21
            Top             =   4380
            Visible         =   0   'False
            Width           =   11745
            _cx             =   20717
            _cy             =   3228
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
            Rows            =   1
            Cols            =   15
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmPlanVentas.frx":01D2
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
         Begin VB.Label Label10 
            Caption         =   "Nº de Productos : "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   5790
            TabIndex        =   26
            Top             =   1185
            Width           =   1545
         End
         Begin VB.Label LblNumReg 
            Alignment       =   1  'Right Justify
            Caption         =   "LblNumReg"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7395
            TabIndex        =   25
            Top             =   1185
            Width           =   795
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Termino"
            Height          =   195
            Left            =   3900
            TabIndex        =   16
            Top             =   855
            Width           =   930
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Productos "
            Height          =   195
            Left            =   60
            TabIndex        =   15
            Top             =   1230
            Width           =   765
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle Plan de Ventas"
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
            Left            =   105
            TabIndex        =   14
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Descripcion"
            Height          =   195
            Left            =   60
            TabIndex        =   13
            Top             =   540
            Width           =   840
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Inicio"
            Height          =   195
            Left            =   60
            TabIndex        =   12
            Top             =   855
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Descripcion"
            Height          =   255
            Left            =   60
            TabIndex        =   11
            Top             =   6615
            Width           =   1005
         End
         Begin VB.Label LblDesc 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDesc"
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   1155
            TabIndex        =   10
            Top             =   6585
            Width           =   10635
         End
         Begin VB.Label Label7 
            Caption         =   "Codigo"
            Height          =   255
            Left            =   60
            TabIndex        =   9
            Top             =   6285
            Width           =   1005
         End
         Begin VB.Label LblCodigo 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblCodigo"
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   1155
            TabIndex        =   8
            Top             =   6270
            Width           =   2160
         End
         Begin VB.Label LblUniMed 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblUniMed"
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   6000
            TabIndex        =   7
            Top             =   6270
            Width           =   1620
         End
         Begin VB.Label Label9 
            Caption         =   "Unidad Medida"
            Height          =   255
            Left            =   4725
            TabIndex        =   6
            Top             =   6285
            Width           =   1200
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
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanVentas.frx":03AF
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanVentas.frx":08F3
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanVentas.frx":0A77
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanVentas.frx":0ECB
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanVentas.frx":0FE3
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanVentas.frx":1527
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanVentas.frx":1A6B
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanVentas.frx":1B7F
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanVentas.frx":1C93
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanVentas.frx":20E7
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanVentas.frx":2253
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   20
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
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar Plan de Ventas"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Activar Plan de Ventas"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Eliminar Plan de Ventas"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Desactivar Plan de Ventas"
               EndProperty
            EndProperty
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
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu menu1 
      Caption         =   "menus"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Agregar               "
      End
      Begin VB.Menu menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Eliminar"
      End
   End
End
Attribute VB_Name = "FrmPlanVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMPLANVENTAS
'* Tipo             : FORMULARIO
'* Descripcion      : FORMULARIO PARA EL INGRESO Y EDICION DEL PLAN DE VENTAS
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 10/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim RstFuente As New ADODB.Recordset
Dim RstPlan As New ADODB.Recordset

Private Sub CmdAddPro_Click()
    ' EJECUTA LA BUSQUEDA DE UNA PROYECCION DE VENTA
    If Fg1.TextMatrix(Fg1.Rows - 1, 1) <> "" Then
        Dim xfrm As New eps_librerias.FormSeleccion
        Dim xCampos(3, 5) As String
        Dim xRs As New ADODB.Recordset
        
        xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "8000":    xCampos(0, 3) = "C":     xCampos(0, 4) = "N"
        xCampos(1, 0) = "Codigo":       xCampos(1, 1) = "codpro":        xCampos(1, 2) = "2000":    xCampos(1, 3) = "C":     xCampos(1, 4) = "N"
        xCampos(2, 0) = "id":           xCampos(2, 1) = "idpro":            xCampos(2, 2) = "0":       xCampos(2, 3) = "N":     xCampos(2, 4) = "S"
        
        xfrm.SQLCad = "SELECT DISTINCT 0 AS xsel, ges_ventaproydet.idpro, alm_inventario.codpro, alm_inventario.descripcion " _
            & " FROM (ges_ventaproydet LEFT JOIN alm_inventario ON ges_ventaproydet.idpro = alm_inventario.id) RIGHT JOIN ges_ventaproy " _
            & " ON ges_ventaproydet.id = ges_ventaproy.id WHERE (((ges_ventaproy.activo)=-1))"

        xfrm.Titulo = "Buscando Productos"
        Set xfrm.Coneccion = xCon
        Set xRs = xfrm.Seleccionar(xCampos)
        If xRs.State = 1 Then
            Dim A As Integer
            Fg1.Rows = 1
            PreparaRST
            For A = 1 To xRs.RecordCount
                Fg1.Rows = Fg1.Rows + 1
                
                Fg1.TextMatrix(A, 0) = xRs("idpro")
                Fg1.TextMatrix(A, 1) = NulosC(xRs("descripcion"))
                'MuestraContratosProducto xRs("idpro")
                MuestraProyeccionVentas xRs("idpro"), Fg1.Rows - 1
                
                HallarTotal A
                xRs.MoveNext
                If xRs.EOF = True Then
                    Exit For
                End If
            Next A
        End If
        Set xfrm = Nothing
    Else
        MsgBox "No ha especificado un producto en el ultimo item", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

Private Sub CmdDelPro_Click()
    If Fg1.Rows = 1 Then
        MsgBox "No ha productos para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Dim Rpta As Integer
    Rpta = MsgBox("Desea eliminar el item seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        Fg1.RemoveItem (Fg1.Row)
        LblNumReg.Caption = Val(LblNumReg.Caption) - 1
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'    Dim xform As New eps_librerias.FormBuscar
'    Dim xRs As New ADODB.Recordset
'    Dim xCampos(2, 4) As String
'
'    If Col = 1 Then
'        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
'        xCampos(0, 0) = "Producto":   xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "6300":     xCampos(0, 3) = "C"
'        xCampos(1, 0) = "Codigo":     xCampos(1, 1) = "codpro":        xCampos(1, 2) = "1900":     xCampos(1, 3) = "C"
'
'        'xform.SQLCad = "SELECT DISTINCT ges_pvventaproydet.idpro, MAE_Producto.Descripcion, MAE_Producto.Cod_Item" _
'            & " FROM ges_pvventaproydet LEFT JOIN MAE_Producto ON ges_pvventaproydet.idpro = MAE_Producto.id " _
'            & " Union " _
'            & " SELECT DISTINCT contratosdet.idpro, MAE_Producto.Descripcion, MAE_Producto.Cod_Item " _
'            & " FROM MAE_Producto RIGHT JOIN contratosdet ON MAE_Producto.id = contratosdet.idpro"
'
'        xform.SQLCad = "SELECT DISTINCT ges_ventaproydet.idpro, alm_inventario.descripcion, alm_inventario.id, alm_inventario.codpro " _
'            & " FROM ges_ventaproy LEFT JOIN (ges_ventaproydet LEFT JOIN alm_inventario ON ges_ventaproydet.idpro = alm_inventario.id) " _
'            & " ON ges_ventaproy.id = ges_ventaproydet.id Where (((ges_ventaproy.activo) = -1)) ORDER BY alm_inventario.descripcion"
'
'        xform.Titulo = "Buscando Productos"
'        xform.FormaBusca = Principio
'        xform.Criterio = ""
'        xform.Ordenado = "descripcion"
'        xform.CampoBusca = "descripcion"
'        Set xform.Coneccion = xCon
'        Set xRs = xform.BuscarReg(xCampos)
'        If xRs.State = 1 Then
'            Dim A As Integer
'            Dim salir As Boolean
'
'            For A = 1 To Fg1.Rows - 1
'                If Val(Fg1.TextMatrix(A, 0)) = xRs("idpro") Then
'                    MsgBox "El producto " + NulosC(xRs("descripcion")) & Chr(13) _
'                        & "ya fue agregado al plan de ventas", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'                    salir = True
'                    Exit For
'                End If
'            Next A
'            If salir = False Then
'                Fg1.TextMatrix(Fg1.Row, 0) = xRs("idpro")
'                Fg1.TextMatrix(Fg1.Row, 1) = NulosC(xRs("descripcion"))
'                MuestraContratosProducto xRs("idpro")
'                MuestraProyeccionVentas xRs("idpro"), Fg1.Row
'                Set xRs = Nothing
'
'                RST_Busq xRs, "SELECT MAE_Producto.Cod_Item, MAE_Unid_Med.Descripcion, MAE_Unid_Med.DescAbrevia " _
'                    & " FROM MAE_Producto LEFT JOIN MAE_Unid_Med ON MAE_Producto.Cod_Unidad = MAE_Unid_Med.Cod_Unidad " _
'                    & " WHERE (((MAE_Producto.id)=" & Val(Fg1.TextMatrix(Fg1.Row, 0)) & "))", xCon
'                Fg1.TextMatrix(Fg1.Row, 15) = xRs("cod_item")
'                Fg1.TextMatrix(Fg1.Row, 16) = xRs("descripcion")
'                HallarTotal Fg1.Row
'            End If
'        End If
'        Set xform = Nothing
'        Set xRs = Nothing
'    End If
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : HallarTotal
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : SUMA LOS TOTALES DE LA FILA DEL CONTROL Fg1
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    xFila        |  Integer   |  ESPECIFICA LA FILA DEL CONTROL Fg1
'* DEVUELVE         :
'*****************************************************************************************************
Sub HallarTotal(xFila As Integer)
    Dim xTotal  As Double
    Dim A, xCol As Integer
    
    xCol = 2
    For A = 1 To Fg1.Cols - 5
        xTotal = xTotal + Val(Fg1.TextMatrix(xFila, xCol))
        xCol = xCol + 1
    Next A
    
    Fg1.TextMatrix(xFila, 14) = Format(xTotal, "0.00")
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If QueHace <> 3 Then
        If Button = 2 Then
            PopupMenu menu1
        End If
    End If
End Sub

Private Sub Fg1_RowColChange()
    If Fg1.TextMatrix(Fg1.Row, 0) = "" Then Exit Sub
    
    LblCodigo.Caption = Trim(Fg1.TextMatrix(Fg1.Row, 15))
    LblUniMed.Caption = Trim(Fg1.TextMatrix(Fg1.Row, 16))
    LblDesc.Caption = Trim(Fg1.TextMatrix(Fg1.Row, 1))
    
    Fg2.Rows = 1
    If RstFuente.State = 1 Then
        RstFuente.Filter = adFilterNone
        If RstFuente.RecordCount <> 0 Then
            RstFuente.MoveFirst
            RstFuente.Filter = "idpro = " & Val(Fg1.TextMatrix(Fg1.Row, 0)) & ""
            
            If RstFuente.RecordCount <> 0 Then
                Dim A As Integer
                Dim xTotal As Double
                
                RstFuente.MoveFirst
                For A = 1 To RstFuente.RecordCount
                    Fg2.Rows = Fg2.Rows + 1
                    Fg2.TextMatrix(A, 1) = RstFuente("numero")
                    Fg2.TextMatrix(A, 2) = NulosN(RstFuente("ene"))
                    Fg2.TextMatrix(A, 3) = NulosN(RstFuente("feb"))
                    Fg2.TextMatrix(A, 4) = NulosN(RstFuente("mar"))
                    Fg2.TextMatrix(A, 5) = NulosN(RstFuente("abr"))
                    Fg2.TextMatrix(A, 6) = NulosN(RstFuente("may"))
                    Fg2.TextMatrix(A, 7) = NulosN(RstFuente("jun"))
                    Fg2.TextMatrix(A, 8) = NulosN(RstFuente("jul"))
                    Fg2.TextMatrix(A, 9) = NulosN(RstFuente("ago"))
                    Fg2.TextMatrix(A, 10) = NulosN(RstFuente("set"))
                    Fg2.TextMatrix(A, 11) = NulosN(RstFuente("oct"))
                    Fg2.TextMatrix(A, 12) = NulosN(RstFuente("nov"))
                    Fg2.TextMatrix(A, 13) = NulosN(RstFuente("dic"))
                
                    xTotal = NulosN(RstFuente("ene")) + NulosN(RstFuente("feb")) + NulosN(RstFuente("mar")) + NulosN(RstFuente("abr"))
                    xTotal = xTotal + NulosN(RstFuente("may")) + NulosN(RstFuente("jun")) + NulosN(RstFuente("jul")) + NulosN(RstFuente("ago"))
                    xTotal = xTotal + NulosN(RstFuente("set")) + NulosN(RstFuente("oct")) + NulosN(RstFuente("nov")) + NulosN(RstFuente("dic"))
                    
                    Fg2.TextMatrix(A, 14) = Format(xTotal, "0.00")
                    
                    RstFuente.MoveNext
                    If RstFuente.EOF = True Then
                        Exit For
                    End If
                Next
            End If
        End If
    End If
End Sub

Private Sub Form_Activate()
'Modificado: 08/01/11 Johan Castro
'            Agregar linea de codigo para bloquear accesos de usuarios


    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = False Then

        
        SeEjecuto = True
        
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, 194, Toolbar1, xCon
        '----------------------------------------------
        
        
        RST_Busq RstPlan, "SELECT ges_planventas.*, IIf([ges_planventas].[activo]=-1,'Activo','No Activo') AS estado " _
            & " FROM ges_planventas ORDER BY id DESC", xCon

        Set Dg1.DataSource = RstPlan
        Dg1.SetFocus
        
        
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    SeEjecuto = False
    QueHace = 3
    TabOne1.CurrTab = 0
    
    If QueHace = 3 Then
        Fg1.SelectionMode = flexSelectionByRow
        Fg1.BackColorSel = &H80&
        Fg1.Height = 5415
    End If
    
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    Fg1.ColWidth(15) = 0
    Fg1.ColWidth(16) = 0
    
    Fg1.FrozenCols = 1
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Toolbar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS BOTONES DE LA BARRA DE HERRAMIENTAS DEL FORMULARIO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Toolbar()
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

'*****************************************************************************************************
'* Nombre Archivo   : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL PROCESO DE INGRESAR O MODIFICAR UN REGISTRO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    Toolbar
    Bloquea
    Label1.Caption = "Detalle Plan de Ventas"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Dg1.SetFocus
    Fg1.Height = 5415
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA EL INGRESO DE UN REGISTRO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Nuevo()
    QueHace = 1
    Toolbar
    Blanquea
    Bloquea
    Fg1.Height = 2625
    Fg2.Visible = True
    TabOne1.CurrTab = 1
    Label1.Caption = "Añadiendo Plan de Ventas"
    TabOne1.TabEnabled(0) = False
    Fg1.ColComboList(1) = "|..."
    Fg1.Rows = Fg1.Rows + 1
    Fg1.Editable = flexEDKbdMouse
    
    Fg1.SelectionMode = flexSelectionFree
    Fg1.BackColorSel = &H80&
    Fg1.Rows = 1
    Fg2.Rows = 1
    PreparaRST
    TxtDesc.SetFocus
    LblNumReg = 0
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA LA EDICION DE UN REGISTRO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Modificar()
    QueHace = 2
    Toolbar
    Blanquea
    Bloquea
    Fg1.Height = 2625
    Fg2.Visible = True
    TabOne1.CurrTab = 1
    Label1.Caption = "Modificando Plan de Ventas"
    TabOne1.TabEnabled(0) = False
    Fg1.ColComboList(1) = "|..."
    Fg1.Rows = Fg1.Rows + 1
    Fg1.Editable = flexEDKbdMouse
    TxtDesc.SetFocus
    Fg1.SelectionMode = flexSelectionFree
    Fg1.BackColorSel = &H80&
    PreparaRST
    MuestraSegundoTab
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA LOS CONTROLES TEXTBOX DEL FORMULARIO PARA EL INGRESO DE UN REGISTRO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Blanquea()
    TxtDesc.Text = ""
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    LblCodigo.Caption = ""
    LblUniMed.Caption = ""
    LblDesc.Caption = ""
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Bloquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA Y DESACTIVA LOS CONTROLES TEXTOBOX DEL FORMULARIO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Bloquea()
    TxtDesc.Locked = Not TxtDesc.Locked
    TxtFchIni.Enabled = Not TxtFchIni.Enabled
    TxtFchFin.Enabled = Not TxtFchFin.Enabled
    CmdAddPro.Enabled = Not CmdAddPro.Enabled
    CmdDelPro.Enabled = Not CmdDelPro.Enabled
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : MuestraProyeccionVentas
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EL DETALLE DE LA PROYECCION DE VENTAS SELECCIONADA
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    IdPro        |  Integer   |  ESPECIFICA EL ID DEL PRODUCTO
'*                    xFila        |  Integer   |  ESPECIFICA EL ID LA FILA DEL CONTROL Fg2
'* DEVUELVE         :
'*****************************************************************************************************
Sub MuestraProyeccionVentas(IdPro As Integer, xFila As Integer)
    Dim xCad As String
    Dim RstSel As New ADODB.Recordset
    Dim A As Integer
    
    xCad = "SELECT ges_ventaproydet.idpro, ges_ventaproydet.id, ges_ventaproydet.ene, ges_ventaproydet.feb, ges_ventaproydet.mar, " _
        & " ges_ventaproydet.abr, ges_ventaproydet.may, ges_ventaproydet.jun, ges_ventaproydet.jul, ges_ventaproydet.ago, " _
        & " ges_ventaproydet.set, ges_ventaproydet.oct, ges_ventaproydet.nov, ges_ventaproydet.dic " _
        & " FROM (ges_ventaproydet LEFT JOIN alm_inventario ON ges_ventaproydet.idpro = alm_inventario.id) RIGHT JOIN ges_ventaproy " _
        & " ON ges_ventaproydet.id = ges_ventaproy.id Where (((ges_ventaproydet.IdPro) = " & IdPro & ") And ((ges_ventaproy.activo) = -1)) " _
        & " ORDER BY alm_inventario.Descripcion"

    RST_Busq RstSel, xCad, xCon
    
    If RstSel.RecordCount <> 0 Then
        RstSel.MoveFirst
        Fg2.Rows = 1
        For A = 1 To RstSel.RecordCount
            RstFuente.AddNew
            RstFuente("id") = 9999
            RstFuente("idpro") = IdPro
            RstFuente("tipo") = 1
            RstFuente("numero") = "Pro. Venta." & Format(RstSel("id"), "0000")
            RstFuente("ene") = NulosN(RstSel("ene"))
            RstFuente("feb") = NulosN(RstSel("feb"))
            RstFuente("mar") = NulosN(RstSel("mar"))
            RstFuente("abr") = NulosN(RstSel("abr"))
            RstFuente("may") = NulosN(RstSel("may"))
            RstFuente("jun") = NulosN(RstSel("jun"))
            RstFuente("jul") = NulosN(RstSel("jul"))
            RstFuente("ago") = NulosN(RstSel("ago"))
            RstFuente("set") = NulosN(RstSel("set"))
            RstFuente("oct") = NulosN(RstSel("oct"))
            RstFuente("nov") = NulosN(RstSel("nov"))
            RstFuente("dic") = NulosN(RstSel("dic"))
            
            Fg2.Rows = Fg2.Rows + 1
            
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = RstSel("id")
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = NulosN(RstSel("ene"))
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = NulosN(RstSel("feb"))
            Fg2.TextMatrix(Fg2.Rows - 1, 4) = NulosN(RstSel("mar"))
            Fg2.TextMatrix(Fg2.Rows - 1, 5) = NulosN(RstSel("abr"))
            Fg2.TextMatrix(Fg2.Rows - 1, 6) = NulosN(RstSel("may"))
            Fg2.TextMatrix(Fg2.Rows - 1, 7) = NulosN(RstSel("jun"))
            Fg2.TextMatrix(Fg2.Rows - 1, 8) = NulosN(RstSel("jul"))
            Fg2.TextMatrix(Fg2.Rows - 1, 9) = NulosN(RstSel("ago"))
            Fg2.TextMatrix(Fg2.Rows - 1, 10) = NulosN(RstSel("set"))
            Fg2.TextMatrix(Fg2.Rows - 1, 11) = NulosN(RstSel("oct"))
            Fg2.TextMatrix(Fg2.Rows - 1, 12) = NulosN(RstSel("nov"))
            Fg2.TextMatrix(Fg2.Rows - 1, 13) = NulosN(RstSel("dic"))
            
            Fg2.TextMatrix(Fg2.Rows - 1, 14) = NulosN(RstSel("ene")) + NulosN(RstSel("feb")) + NulosN(RstSel("mar")) + NulosN(RstSel("abr"))
            Fg2.TextMatrix(Fg2.Rows - 1, 14) = Val(Fg2.TextMatrix(Fg2.Rows - 1, 14)) + NulosN(RstSel("may")) + NulosN(RstSel("jun")) + NulosN(RstSel("jul"))
            Fg2.TextMatrix(Fg2.Rows - 1, 14) = Val(Fg2.TextMatrix(Fg2.Rows - 1, 14)) + NulosN(RstSel("ago")) + NulosN(RstSel("set")) + NulosN(RstSel("oct"))
            Fg2.TextMatrix(Fg2.Rows - 1, 14) = Val(Fg2.TextMatrix(Fg2.Rows - 1, 14)) + NulosN(RstSel("nov")) + NulosN(RstSel("dic"))
            Fg2.TextMatrix(Fg2.Rows - 1, 14) = Format(NulosN(Fg2.TextMatrix(Fg2.Rows - 1, 14)), "0.00")
            
            RstSel.MoveNext
            If RstSel.EOF = True Then
                Exit For
            End If
            
        Next A
    End If
    
    Dim B As Integer
    Dim Total As Double
    
    'recorremos las columnas
    For A = 2 To Fg2.Cols - 1
        Total = 0
        'recorremos las filas
        For B = 1 To Fg2.Rows - 1
            Total = Total + Val(Fg2.TextMatrix(B, A))
            
            If Fg2.Rows = (Fg2.Rows - 1) Then
                Exit For
            End If
        Next B
        Fg1.TextMatrix(xFila, A) = Format(Total, "0.00")
        If Fg1.Col = (Fg1.Cols - 1) Then
            Exit For
        End If
    Next A
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : MuestraContratosProducto
'* Tipo             : PROCEDIMIENTO
'* Descripcion      :
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    IdPro        |  Integer   |  ESPECIFICA EL ID DEL PRODUCTO
'* DEVUELVE         :
'*****************************************************************************************************
Sub MuestraContratosProducto(IdPro As Integer)
    Dim xCad As String
    Dim RstSel As New ADODB.Recordset
    
    xCad = "SELECT DISTINCT contratos.id, contratos.numord, contratosdetent.idpro, " _
    & " (SELECT Sum([canent]) AS totent From contratosdetent WHERE (((contratosdetent.idcon)=contratos.id) " _
    & " AND ((contratosdetent.idpro)=contratosdetent.idpro) AND ((contratosdetent.fchent)>=CDate('01/01/06') " _
    & " And (contratosdetent.fchent)<=CDate('31/01/06')))) AS totene, " _
    & " (SELECT Sum([canent]) AS totent From contratosdetent WHERE (((contratosdetent.idcon)=contratos.id) " _
    & " AND ((contratosdetent.idpro)=contratosdetent.idpro) AND ((contratosdetent.fchent)>=CDate('01/02/06') " _
    & " And (contratosdetent.fchent)<=CDate('28/02/06')))) AS totfeb, " _
    & " (SELECT Sum([canent]) AS totent From contratosdetent WHERE (((contratosdetent.idcon)=contratos.id) " _
    & " AND ((contratosdetent.idpro)=contratosdetent.idpro) AND ((contratosdetent.fchent)>=CDate('01/03/06') " _
    & " And (contratosdetent.fchent)<=CDate('31/03/06')))) AS totmar, " _
    & " (SELECT Sum([canent]) AS totent From contratosdetent WHERE (((contratosdetent.idcon)=contratos.id) " _
    & " AND ((contratosdetent.idpro)=contratosdetent.idpro) AND ((contratosdetent.fchent)>=CDate('01/04/06') " _
    & " And (contratosdetent.fchent)<=CDate('30/04/06')))) AS totabr, " _
    & " (SELECT Sum([canent]) AS totent From contratosdetent WHERE (((contratosdetent.idcon)=contratos.id) " _
    & " AND ((contratosdetent.idpro)=contratosdetent.idpro) AND ((contratosdetent.fchent)>=CDate('01/05/06') " _
    & " And (contratosdetent.fchent)<=CDate('31/05/06')))) AS totmay, " _
    & " (SELECT Sum([canent]) AS totent From contratosdetent WHERE (((contratosdetent.idcon)=contratos.id) " _
    & " AND ((contratosdetent.idpro)=contratosdetent.idpro) AND ((contratosdetent.fchent)>=CDate('01/06/06') " _
    & " And (contratosdetent.fchent)<=CDate('30/06/06')))) AS totjun, " _
    & " (SELECT Sum([canent]) AS totent From contratosdetent WHERE (((contratosdetent.idcon)=contratos.id) " _
    & " AND ((contratosdetent.idpro)=contratosdetent.idpro) AND ((contratosdetent.fchent)>=CDate('01/07/06') " _
    & " And (contratosdetent.fchent)<=CDate('31/07/06')))) AS totjul, " _
    & " (SELECT Sum([canent]) AS totent From contratosdetent WHERE (((contratosdetent.idcon)=contratos.id) " _
    & " AND ((contratosdetent.idpro)=contratosdetent.idpro) AND ((contratosdetent.fchent)>=CDate('01/08/06') " _
    & " And (contratosdetent.fchent)<=CDate('31/08/06')))) AS totago,"
    
    xCad = xCad + "(SELECT Sum([canent]) AS totent From contratosdetent WHERE (((contratosdetent.idcon)=contratos.id) " _
    & " AND ((contratosdetent.idpro)=contratosdetent.idpro) AND ((contratosdetent.fchent)>=CDate('01/09/06') " _
    & " And (contratosdetent.fchent)<=CDate('30/09/06')))) AS totset, " _
    & " (SELECT Sum([canent]) AS totent From contratosdetent WHERE (((contratosdetent.idcon)=contratos.id) " _
    & " AND ((contratosdetent.idpro)=contratosdetent.idpro) AND ((contratosdetent.fchent)>=CDate('01/10/06') " _
    & " And (contratosdetent.fchent)<=CDate('31/10/06')))) AS totoct, " _
    & " (SELECT Sum([canent]) AS totent From contratosdetent WHERE (((contratosdetent.idcon)=contratos.id) " _
    & " AND ((contratosdetent.idpro)=contratosdetent.idpro) AND ((contratosdetent.fchent)>=CDate('01/11/06') " _
    & " And (contratosdetent.fchent)<=CDate('30/11/06')))) AS totnov, " _
    & " (SELECT Sum([canent]) AS totent From contratosdetent WHERE (((contratosdetent.idcon)=contratos.id) " _
    & " AND ((contratosdetent.idpro)=contratosdetent.idpro) AND ((contratosdetent.fchent)>=CDate('01/12/06') " _
    & " And (contratosdetent.fchent)<=CDate('31/12/06')))) AS totdic " _
    & " FROM contratos RIGHT JOIN contratosdetent ON contratos.id = contratosdetent.idcon " _
    & " WHERE (((contratosdetent.idpro) = " & IdPro & "))"

    RST_Busq RstSel, xCad, xCon
    
    Fg2.Rows = 1
    Dim A As Integer
    
    If RstSel.RecordCount <> 0 Then
        RstSel.MoveFirst
        
        For A = 1 To RstSel.RecordCount
            RstFuente.AddNew
            RstFuente("id") = 9999
            RstFuente("idpro") = IdPro
            RstFuente("tipo") = 2
            RstFuente("numero") = "O.C.-" + Trim(RstSel("numord"))
            RstFuente("ene") = NulosN(RstSel("totene"))
            RstFuente("feb") = NulosN(RstSel("totfeb"))
            RstFuente("mar") = NulosN(RstSel("totmar"))
            RstFuente("abr") = NulosN(RstSel("totabr"))
            RstFuente("may") = NulosN(RstSel("totmay"))
            RstFuente("jun") = NulosN(RstSel("totjun"))
            RstFuente("jul") = NulosN(RstSel("totjul"))
            RstFuente("ago") = NulosN(RstSel("totago"))
            RstFuente("set") = NulosN(RstSel("totset"))
            RstFuente("oct") = NulosN(RstSel("totoct"))
            RstFuente("nov") = NulosN(RstSel("totnov"))
            RstFuente("dic") = NulosN(RstSel("totdic"))
            
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(A, 1) = RstSel("numord")
            Fg2.TextMatrix(A, 2) = NulosN(RstSel("totene"))
            Fg2.TextMatrix(A, 3) = NulosN(RstSel("totfeb"))
            Fg2.TextMatrix(A, 4) = NulosN(RstSel("totmar"))
            Fg2.TextMatrix(A, 5) = NulosN(RstSel("totabr"))
            Fg2.TextMatrix(A, 6) = NulosN(RstSel("totmay"))
            Fg2.TextMatrix(A, 7) = NulosN(RstSel("totjun"))
            Fg2.TextMatrix(A, 8) = NulosN(RstSel("totjul"))
            Fg2.TextMatrix(A, 9) = NulosN(RstSel("totago"))
            Fg2.TextMatrix(A, 10) = NulosN(RstSel("totset"))
            Fg2.TextMatrix(A, 11) = NulosN(RstSel("totoct"))
            Fg2.TextMatrix(A, 12) = NulosN(RstSel("totnov"))
            Fg2.TextMatrix(A, 13) = NulosN(RstSel("totdic"))
            
            Fg2.TextMatrix(A, 14) = NulosN(RstSel("totene")) + NulosN(RstSel("totfeb")) + NulosN(RstSel("totmar")) + NulosN(RstSel("totabr"))
            Fg2.TextMatrix(A, 14) = Val(Fg2.TextMatrix(A, 14)) + NulosN(RstSel("totmay")) + NulosN(RstSel("totjun")) + NulosN(RstSel("totjul"))
            Fg2.TextMatrix(A, 14) = Val(Fg2.TextMatrix(A, 14)) + NulosN(RstSel("totago")) + NulosN(RstSel("totset")) + NulosN(RstSel("totoct"))
            Fg2.TextMatrix(A, 14) = Val(Fg2.TextMatrix(A, 14)) + NulosN(RstSel("totnov")) + NulosN(RstSel("totdic"))
            Fg2.TextMatrix(A, 14) = Format(Fg2.TextMatrix(A, 14), "0.00")
            
            RstSel.MoveNext
            If RstSel.EOF = True Then
                Exit For
            End If
        Next A
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 1
        Exit Sub
    End If
End Sub

Private Sub menu1_1_Click()
    If Fg1.TextMatrix(Fg1.Rows - 1, 1) <> "" Then
        CmdAddPro_Click
    Else
        MsgBox "No ha especificado un producto en el ultimo item", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

Private Sub menu1_3_Click()
    If Fg1.Rows = 1 Then
        MsgBox "No ha productos para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    CmdDelPro_Click
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        Nuevo
    End If

    If Button.Index = 2 Then
        Modificar
    End If

    If Button.Index = 3 Then
        Eliminar
    End If
    
    If Button.Index = 5 Then
        If Grabar = True Then
            RstPlan.Requery
            Dg1.Refresh
            Cancelar
        End If
    End If
    If Button.Index = 6 Then
        Cancelar
    End If
    
    If Button.Index = 14 Then
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO DE LA TABLA ges_planventas
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Eliminar()
    Dim Rpta As Integer
    
    Rpta = MsgBox("Esta seguro de eliminar el plan de ventas seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM ges_planventas WHERE id =" & RstPlan("id") & ""
        MsgBox "El plan de ventas se elimino con exito", vbInformation + vbQuestion + vbDefaultButton1, xTitulo
        RstPlan.Requery
        Dg1.Refresh
        Exit Sub
    End If
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : PreparaRST
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA UN RECORDSET TEMPORAL
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub PreparaRST()
    Dim xFun As New eps_librerias.FuncionesData
    
    Dim xCampos(16, 3) As String

    xCampos(0, 0) = "id":         xCampos(0, 1) = "C":      xCampos(0, 2) = "4"
    xCampos(1, 0) = "idpro":      xCampos(1, 1) = "C":      xCampos(1, 2) = "4"
    xCampos(2, 0) = "tipo":       xCampos(2, 1) = "N":      xCampos(2, 2) = "2"
    xCampos(3, 0) = "ene":        xCampos(3, 1) = "N":      xCampos(3, 2) = "2"
    xCampos(4, 0) = "feb":        xCampos(4, 1) = "N":      xCampos(4, 2) = "2"
    xCampos(5, 0) = "mar":        xCampos(5, 1) = "N":      xCampos(5, 2) = "2"
    xCampos(6, 0) = "abr":        xCampos(6, 1) = "N":      xCampos(6, 2) = "2"
    xCampos(7, 0) = "may":        xCampos(7, 1) = "N":      xCampos(7, 2) = "2"
    xCampos(8, 0) = "jun":        xCampos(8, 1) = "N":      xCampos(8, 2) = "2"
    xCampos(9, 0) = "jul":        xCampos(9, 1) = "N":      xCampos(9, 2) = "2"
    xCampos(10, 0) = "ago":        xCampos(10, 1) = "N":      xCampos(10, 2) = "2"
    xCampos(11, 0) = "set":        xCampos(11, 1) = "N":      xCampos(11, 2) = "2"
    xCampos(12, 0) = "oct":        xCampos(12, 1) = "N":      xCampos(12, 2) = "2"
    xCampos(13, 0) = "nov":        xCampos(13, 1) = "N":      xCampos(13, 2) = "2"
    xCampos(14, 0) = "dic":        xCampos(14, 1) = "N":      xCampos(14, 2) = "2"
    xCampos(15, 0) = "numero":     xCampos(15, 1) = "C":      xCampos(15, 2) = "15"
    Set RstFuente = xFun.CrearRstTMP(xCampos)
    RstFuente.Open
End Sub


'*****************************************************************************************************
'* Nombre Archivo   : Grabar
'* Tipo             : FUNCION
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA ges_planventas, ESTA FUNCION DEVUELVE VERDADERO
'*                    CUANDO TIENE EXITO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         : Boolean
'*****************************************************************************************************
Function Grabar() As Boolean
    If NulosC(TxtDesc.Text) = "" Then
        MsgBox "No ha especificado la descripcion para la proyeccion de ventas", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtDesc.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtFchIni.Valor) = "" Then
        MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Function
    End If

    If NulosC(TxtFchFin.Valor) = "" Then
        MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchFin.SetFocus
        Exit Function
    End If
    
    'comprobamos si se han agregado items al primer grid
    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado un producto para el plan de ventas", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Function
    End If

    Dim A As Integer
    'comprobamos si el grid tiene filas en blanco
    For A = 1 To Fg1.Rows - 1
        If Fg1.TextMatrix(A, 1) = "" Then
            Fg1.RemoveItem A
        End If
        
        If A = (Fg1.Rows - 1) Then
            Exit For
        End If
    Next A
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstFue As New ADODB.Recordset
    Dim xId As Integer
    
    On Error GoTo LaCague
    
    xCon.BeginTrans
    
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT * FROM ges_planventas", xCon
        RST_Busq RstDet, "SELECT * FROM ges_planventasdet", xCon
        
        xId = HallaCodigoTabla("ges_planventas", xCon, "id")
        RstCab.AddNew
        RstCab("id") = xId
    Else
        RST_Busq RstCab, "SELECT * FROM ges_planventas WHERE id=" & RstPlan("id") & " ", xCon
        xCon.Execute "DELETE * FROM ges_planventasdet WHERE idpv = " & RstPlan("id") & ""
        
        RST_Busq RstDet, "SELECT * FROM ges_planventasdet", xCon
        xId = RstPlan("id")
    End If
    
    RstCab("descripcion") = TxtDesc.Text
    RstCab("fchini") = TxtFchIni.Valor
    RstCab("fchfin") = TxtFchFin.Valor
    RstCab.Update
    
    Dim xFila, xCol, xMes As Integer
    
    For xFila = 1 To Fg1.Rows - 1
        xMes = 1
        For xCol = 2 To Fg1.Cols - 4
            RstDet.AddNew
            RstDet("idpv") = xId
            RstDet("codpro") = Fg1.TextMatrix(xFila, 0)
            RstDet("idmes") = xMes
            RstDet("cantidad") = Fg1.TextMatrix(xFila, xCol)
            
            RstDet.Update
            xMes = xMes + 1
        Next xCol
    Next xFila
    
    xCon.CommitTrans
    MsgBox "El plan de Ventas se grabo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    Grabar = True
    Exit Function

LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function

'*****************************************************************************************************
'* Nombre Archivo   : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DETALLADA DEL PLAN DE VENTAS SELECCIONADO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    TxtDesc.Text = RstPlan("descripcion")
    TxtFchIni.Valor = RstPlan("fchini")
    TxtFchFin.Valor = RstPlan("fchfin")
    
    Dim Rst As New ADODB.Recordset
    Dim Rst2 As New ADODB.Recordset
    
    Dim A, xCol, B As Integer
    Dim Total As Double
    
    RST_Busq Rst, "SELECT DISTINCT alm_inventario.descripcion, ges_planventasdet.idpv, ges_planventasdet.codpro FROM ges_planventasdet LEFT JOIN alm_inventario " _
        & " ON ges_planventasdet.codpro = alm_inventario.id Where (((ges_planventasdet.idpv) = " & RstPlan("id") & ")) ORDER BY alm_inventario.descripcion", xCon

    Fg1.Rows = 1
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        
        LblNumReg.Caption = Format(Rst.RecordCount, "000")
        
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Total = 0
            Fg1.TextMatrix(A, 0) = Rst("codpro")
            Fg1.TextMatrix(A, 1) = NulosC(Rst("descripcion"))
                    
            RST_Busq Rst2, "SELECT ges_planventasdet.codpro, ges_planventasdet.idpv, ges_planventasdet.cantidad, ges_planventasdet.idmes " _
                & " From ges_planventasdet Where (((ges_planventasdet.CodPro) = " & Val(Fg1.TextMatrix(A, 0)) & ") And ((ges_planventasdet.idpv) = " & RstPlan("id") & ")) " _
                & " ORDER BY ges_planventasdet.idmes", xCon
            xCol = 2
            Rst2.MoveFirst
            
            For B = 1 To Rst2.RecordCount
                Fg1.TextMatrix(A, xCol) = Format(Rst2("cantidad"), "0.00")
                Total = Total + Rst2("cantidad")
                
                Rst2.MoveNext
                xCol = xCol + 1
                If Rst2.EOF = True Then
                    Exit For
                End If
            Next B
            
            Fg1.TextMatrix(A, xCol) = Format(Total, "0.00")
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
    End If
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : CambiarEstado
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA UN REGISTRO DE LA TABLA ges_planventas
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Activado     |  Boolean   |  ESPECIFICA SI SE ACTIVA O DESACTIVA EL REGISTRO
'* DEVUELVE         :
'*****************************************************************************************************
Sub CambiarEstado(Activado As Boolean)
    Dim Rpta As Integer
    TabOne1.CurrTab = 0
    If Activado = False Then
        Rpta = MsgBox("Esta seguro de desactivar el plan de ventas seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    Else
        Rpta = MsgBox("Esta seguro de activar el plan de ventas seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    End If
    
    If Rpta = vbYes Then
        If Activado = False Then
            xCon.Execute "UPDATE ges_planventas SET ges_planventas.activo = 0 Where (((ges_planventas.id) = " & RstPlan("id") & "))"
            MsgBox "El plan de ventas se desactivo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Else
            xCon.Execute "UPDATE ges_planventas SET ges_planventas.activo = -1 Where (((ges_planventas.id) = " & RstPlan("id") & "))"
            MsgBox "El plan de ventas se activo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
    End If
    RstPlan.Requery
    Dg1.Refresh
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.Index = 2 Then
        If ButtonMenu.Index = 1 Then Modificar
        If ButtonMenu.Index = 2 Then CambiarEstado True
    End If
    If ButtonMenu.Parent.Index = 3 Then
        If ButtonMenu.Index = 1 Then Eliminar
        If ButtonMenu.Index = 2 Then CambiarEstado False
    End If
End Sub

Private Sub TxtDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

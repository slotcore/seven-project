VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SIZERONE.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmComprasDiversas 
   Caption         =   "Compras - Compras Diversas"
   ClientHeight    =   7605
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7230
      Left            =   0
      TabIndex        =   1
      Top             =   375
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
      FrontTabForeColor=   -2147483630
      Caption         =   "  Consulta  |  Detalle   "
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6810
         Left            =   12525
         TabIndex        =   3
         Top             =   375
         Width           =   11790
         Begin VB.CommandButton CmdDelDoc 
            Caption         =   "Eliminar Documento"
            Height          =   350
            Left            =   1815
            TabIndex        =   10
            Top             =   6435
            Width           =   1695
         End
         Begin VB.CommandButton CmdAddDoc 
            Caption         =   "Agregar Documento"
            Height          =   350
            Left            =   90
            TabIndex        =   9
            Top             =   6435
            Width           =   1695
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmi 
            Height          =   300
            Left            =   1380
            TabIndex        =   6
            Top             =   480
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
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   5100
            Left            =   15
            TabIndex        =   8
            Top             =   1275
            Width           =   11775
            _cx             =   20770
            _cy             =   8996
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
            Rows            =   1
            Cols            =   13
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmComprasDiversas.frx":0000
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
            AutoSize        =   -1  'True
            Caption         =   "Fecha Ingreso"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   525
            Width           =   1020
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Detalle Compras"
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
            Index           =   1
            Left            =   105
            TabIndex        =   5
            Top             =   30
            Width           =   11595
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6810
         Left            =   45
         TabIndex        =   2
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6420
            Left            =   30
            TabIndex        =   11
            Top             =   360
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11324
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nº Movimiento"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Fch. Mov."
            Columns(1).DataField=   "fchmov"
            Columns(1).NumberFormat=   "dd/mm/yy"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Importe"
            Columns(2).DataField=   "imptot"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2619"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2540"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=513"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=1746"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1667"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2566"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2487"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=514"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=1"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(48)  =   "Named:id=33:Normal"
            _StyleDefs(49)  =   ":id=33,.parent=0"
            _StyleDefs(50)  =   "Named:id=34:Heading"
            _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(52)  =   ":id=34,.wraptext=-1"
            _StyleDefs(53)  =   "Named:id=35:Footing"
            _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(55)  =   "Named:id=36:Selected"
            _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(57)  =   "Named:id=37:Caption"
            _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(59)  =   "Named:id=38:HighlightRow"
            _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(61)  =   "Named:id=39:EvenRow"
            _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(63)  =   "Named:id=40:OddRow"
            _StyleDefs(64)  =   ":id=40,.parent=33"
            _StyleDefs(65)  =   "Named:id=41:RecordSelector"
            _StyleDefs(66)  =   ":id=41,.parent=34"
            _StyleDefs(67)  =   "Named:id=42:FilterBar"
            _StyleDefs(68)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta Compras"
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
            Left            =   105
            TabIndex        =   4
            Top             =   30
            Width           =   11595
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
            Picture         =   "FrmComprasDiversas.frx":0188
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmComprasDiversas.frx":06CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmComprasDiversas.frx":0A5E
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmComprasDiversas.frx":0BE2
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmComprasDiversas.frx":1036
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmComprasDiversas.frx":114E
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmComprasDiversas.frx":1692
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmComprasDiversas.frx":1BD6
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmComprasDiversas.frx":1CEA
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmComprasDiversas.frx":1DFE
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmComprasDiversas.frx":2252
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmComprasDiversas.frx":23BE
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
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
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Guia"
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
Attribute VB_Name = "FrmComprasDiversas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RstMov As New ADODB.Recordset
Dim QueHace As Integer
Dim SeEjecuto As Boolean
    
Private Sub Command1_Click()

End Sub

Private Sub CmdAddDoc_Click()
    Fg1.Rows = Fg1.Rows + 1
End Sub

Private Sub CmdDelDoc_Click()
    Fg1.RemoveItem Fg1.Row
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    'Dim xForm As New eps_librerias.FormBuscar
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    If Col = 1 Then
        Dim xCampos(2, 4) As String
        
        xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5500":    xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":       xCampos(1, 1) = "id":             xCampos(1, 2) = "1000":    xCampos(1, 3) = "N"
        
        xform.SQLCad = "SELECT * FROM mae_documento ORDER BY descripcion"
    
        xform.Titulo = "Buscando Documentos"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "descripcion"
        xform.CampoBusca = "descripcion"
        
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        
        If xRs.State = 1 Then
            Fg1.TextMatrix(Fg1.Row, 1) = xRs("abrev")
            Fg1.TextMatrix(Fg1.Row, 9) = xRs("id")
        End If
        Set xform = Nothing
        Set xRs = Nothing
    End If

    If Col = 2 Then
        Dim xCampos1(2, 4) As String
        
        xCampos1(0, 0) = "Proveedor":  xCampos1(0, 1) = "nombre":    xCampos1(0, 2) = "5500":    xCampos1(0, 3) = "C"
        xCampos1(1, 0) = "Nº Ruc":     xCampos1(1, 1) = "numruc":    xCampos1(1, 2) = "1200":    xCampos1(1, 3) = "C"
        
        xform.SQLCad = "SELECT * FROM mae_prov ORDER BY nombre"
    
        xform.Titulo = "Buscando Proveedores"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "nombre"
        xform.CampoBusca = "nombre"
        
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos1)
        
        If xRs.State = 1 Then
            Fg1.TextMatrix(Fg1.Row, 2) = xRs("numruc")
            Fg1.TextMatrix(Fg1.Row, 3) = xRs("nombre")
            Fg1.TextMatrix(Fg1.Row, 10) = xRs("id")
        End If
        Set xform = Nothing
        Set xRs = Nothing
    End If
    
    If Col = 8 Then
        Dim xCampos2(2, 4) As String
        
        xCampos2(0, 0) = "Descripcion":  xCampos2(0, 1) = "descripcion":    xCampos2(0, 2) = "5500":    xCampos2(0, 3) = "C"
        xCampos2(1, 0) = "Codigo":       xCampos2(1, 1) = "codpro":         xCampos2(1, 2) = "2000":    xCampos2(1, 3) = "C"
        
        xform.SQLCad = "SELECT * FROM alm_inventario WHERE tippro = 5 ORDER BY descripcion"
    
        xform.Titulo = "Buscando Servicios"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "descripcion"
        xform.CampoBusca = "descripcion"
        
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos2)
        
        If xRs.State = 1 Then
            Fg1.TextMatrix(Fg1.Row, 8) = xRs("descripcion")
            Fg1.TextMatrix(Fg1.Row, 11) = xRs("id")
            Fg1.TextMatrix(Fg1.Row, 12) = xRs("idcuenta")
        End If
        Set xform = Nothing
        Set xRs = Nothing
    End If
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 5 Then
        Fg1.TextMatrix(Fg1.Row, 5) = Format(Fg1.TextMatrix(Fg1.Row, 5), "0000")
    End If

    If Col = 6 Then
        Fg1.TextMatrix(Fg1.Row, 6) = Format(Fg1.TextMatrix(Fg1.Row, 6), "0000000000")
    End If
    
    If Col = 7 Then
        Fg1.TextMatrix(Fg1.Row, 7) = Format(Fg1.TextMatrix(Fg1.Row, 7), "0.00")
    End If

End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        Dim Rpta As Integer
        
        SeEjecuto = True
        RST_Busq RstMov, "SELECT com_diversos.id, com_diversos.fchmov, com_diversos.imptot " _
            & " From com_diversos ORDER BY com_diversos.fchmov", xCon
        
        Set Dg1.DataSource = RstMov
        If RstMov.RecordCount = 0 Then
            Rpta = MsgBox("No se ha registrado ninguna compra, ¿Desea agregar una ahora?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                Nuevo
            Else
                Set RstMov = Nothing
                Unload Me
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    QueHace = 3
    SeEjecuto = False
    TabOne1.CurrTab = 0
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    Fg1.ColWidth(9) = 0
    Fg1.ColWidth(10) = 0
    Fg1.ColWidth(11) = 0
    Fg1.ColWidth(12) = 0
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.Editable = flexEDNone
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
            RstMov.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 15 Then
        Set RstMov = Nothing
        Unload Me
    End If
End Sub

Sub Cancelar()
    Bloquea
    ActivaTool
    Label4(1).Caption = "Detalle de Compras"
    QueHace = 3
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
End Sub

Function Grabar() As Boolean
    If TxtFchEmi.Valor = "" Then
        MsgBox "No ha especificado la fecha de ingreso", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchEmi.SetFocus
        Exit Function
    End If
    
    Dim A, xId As Integer
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    
    For A = 1 To Fg1.Rows - 1
        If Fg1.TextMatrix(A, 1) = "" Then
            MsgBox "No ha especificado el tipo de documento para la compra Nº " + Format(Str(A), "000")
            Exit Function
        End If
        If Fg1.TextMatrix(A, 2) = "" Then
            MsgBox "No ha especificado el proveedor para la compra Nº " + Format(Str(A), "000")
            Exit Function
        End If
    
        If Fg1.TextMatrix(A, 4) = "" Then
            MsgBox "No ha especificado la fecha de adquisicion para la compra Nº " + Format(Str(A), "000")
            Exit Function
        End If
    
        If Fg1.TextMatrix(A, 5) = "" Then
            MsgBox "No ha especificado el numero de serie para la compra Nº " + Format(Str(A), "000")
            Exit Function
        End If
    
        If Fg1.TextMatrix(A, 6) = "" Then
            MsgBox "No ha especificado el numero de documento para la compra Nº " + Format(Str(A), "000")
            Exit Function
        End If
    
        If Fg1.TextMatrix(A, 7) = "" Then
            MsgBox "No ha especificado el importe para la compra Nº " + Format(Str(A), "000")
            Exit Function
        End If
    
        If Fg1.TextMatrix(A, 8) = "" Then
            MsgBox "No ha especificado el concepto para la compra Nº " + Format(Str(A), "000")
            Exit Function
        End If
    Next A
    
    If QueHace = 1 Then
        xId = HallaCodigoTabla("com_diversos", xCon, "id")
        RST_Busq RstCab, "SELECT * FROM com_diversos", xCon
        RST_Busq RstDet, "SELECT * FROM com_diversosdet", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else
    
    End If
    
    RstCab("fchmov") = CDate(TxtFchEmi.Valor)
    RstCab.Update
    
    For A = 1 To Fg1.Rows - 1
        RstDet.AddNew
        RstDet("id") = xId
        RstDet("idtipdoc") = Val(Fg1.TextMatrix(A, 9))
        If NulosN(Fg1.TextMatrix(A, 10)) <> 0 Then RstDet("idpro") = Fg1.TextMatrix(A, 10)
        RstDet("numruc") = Fg1.TextMatrix(A, 2)
        RstDet("nompro") = Fg1.TextMatrix(A, 3)
        RstDet("fchdoc") = CDate(Fg1.TextMatrix(A, 4))
        RstDet("numser") = Fg1.TextMatrix(A, 5)
        RstDet("numdoc") = Fg1.TextMatrix(A, 6)
        RstDet("impdoc") = Fg1.TextMatrix(A, 7)
        RstDet("idcon") = Fg1.TextMatrix(A, 11)
        RstDet.Update
    Next A
    
    MsgBox "La compra se registro con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
End Function

Sub Modificar()

End Sub

Sub Eliminar()

End Sub

Sub Blanquea()

End Sub

Sub Bloquea()

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
    Toolbar1.Buttons(11).Enabled = Not Toolbar1.Buttons(11).Enabled
    
    Toolbar1.Buttons(13).Enabled = Not Toolbar1.Buttons(13).Enabled
    Toolbar1.Buttons(15).Enabled = Not Toolbar1.Buttons(15).Enabled
End Sub

Sub Nuevo()
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Blanquea
    Bloquea
    ActivaTool
    QueHace = 1
    Fg1.ColComboList(1) = "|..."
    Fg1.ColComboList(2) = "|..."
    Fg1.ColComboList(8) = "|..."
    Label4(1).Caption = "Agregando Compras"
    Fg1.Editable = flexEDKbdMouse
    Fg1.Rows = Fg1.Rows + 1
    Fg1.SetFocus
End Sub

Sub MuestraSegundoTab()
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    
    RST_Busq Rst, "SELECT com_diversosdet.*, [mae_documento]![abrev] AS descdoc, alm_inventario.descripcion AS desccon " _
        & " FROM alm_inventario RIGHT JOIN (mae_documento RIGHT JOIN com_diversosdet ON mae_documento.id = com_diversosdet.idtipdoc) " _
        & " ON alm_inventario.id = com_diversosdet.idcon Where (((com_diversosdet.id) = " & RstMov("id") & ")) " _
        & " ORDER BY com_diversosdet.numser, com_diversosdet.numdoc", xCon
    Fg1.Rows = 1
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = Rst("descdoc")
            Fg1.TextMatrix(A, 2) = Rst("numruc")
            Fg1.TextMatrix(A, 3) = Rst("nompro")
            Fg1.TextMatrix(A, 4) = Rst("fchdoc")
            Fg1.TextMatrix(A, 5) = Rst("numser")
            Fg1.TextMatrix(A, 6) = Rst("numdoc")
            Fg1.TextMatrix(A, 7) = Format(Rst("impdoc"), "0.00")
            Fg1.TextMatrix(A, 8) = Rst("desccon")
            Fg1.TextMatrix(A, 9) = Rst("idtipdoc")
            Fg1.TextMatrix(A, 10) = Rst("idpro")
            Fg1.TextMatrix(A, 11) = Rst("idcon")
            'Fg1.TextMatrix(A, 12) = Rst("") 'cuenta contable
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
    End If
End Sub


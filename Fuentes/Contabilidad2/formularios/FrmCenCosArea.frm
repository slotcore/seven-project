VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmCenCosArea 
   Caption         =   "Asignar Centro de Costo a Areas"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7605
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8010
      Top             =   15
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
            Picture         =   "FrmCenCosArea.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCenCosArea.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCenCosArea.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCenCosArea.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCenCosArea.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCenCosArea.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCenCosArea.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCenCosArea.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCenCosArea.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCenCosArea.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCenCosArea.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCenCosArea.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCenCosArea.frx":277E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   609
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
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Enviar por Correo"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir a PDF"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7230
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   7605
      _cx             =   13414
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6810
         Left            =   8250
         TabIndex        =   7
         Top             =   375
         Width           =   7515
         Begin VB.CommandButton CmdBusArea 
            Height          =   240
            Left            =   1620
            Picture         =   "FrmCenCosArea.frx":2B10
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   660
            Width           =   240
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   4830
            Left            =   60
            TabIndex        =   10
            Top             =   1230
            Width           =   7365
            _cx             =   12991
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
            BackColor       =   14614269
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   128
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   14614269
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
            FormatString    =   $"FrmCenCosArea.frx":2C42
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
         Begin VB.TextBox TxtIdArea 
            Height          =   300
            Left            =   1110
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   9
            Text            =   "TxtIdArea"
            Top             =   630
            Width           =   780
         End
         Begin VB.Frame Frame3 
            Height          =   795
            Left            =   60
            TabIndex        =   15
            Top             =   6000
            Width           =   7395
            Begin VB.CommandButton CmdDelCenCos 
               Caption         =   "&Eliminar Centro de Costo"
               Enabled         =   0   'False
               Height          =   480
               Left            =   3720
               TabIndex        =   17
               Top             =   195
               Width           =   1440
            End
            Begin VB.CommandButton CmdAddCenCos 
               Caption         =   "&Agregar Centro Costo"
               Enabled         =   0   'False
               Height          =   495
               Left            =   2235
               TabIndex        =   16
               Top             =   195
               Width           =   1440
            End
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "[  Centros de Costo Asignados  ]"
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
            Left            =   75
            TabIndex        =   14
            Top             =   990
            Width           =   2760
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Centros de Costo por Area"
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
            TabIndex        =   13
            Top             =   30
            Width           =   7320
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Area"
            Height          =   195
            Left            =   105
            TabIndex        =   12
            Top             =   675
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
            Left            =   1950
            TabIndex        =   11
            Top             =   630
            Width           =   5490
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6810
         Left            =   45
         TabIndex        =   2
         Top             =   375
         Width           =   7515
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6480
            Left            =   30
            TabIndex        =   3
            Top             =   315
            Width           =   7470
            _ExtentX        =   13176
            _ExtentY        =   11430
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo"
            Columns(0).DataField=   "idarea"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Descripcion"
            Columns(1).DataField=   "descripcion"
            Columns(1).NumberFormat=   "0.00"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   397
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2011"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1931"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=10186"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=10107"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=13,.alignment=0"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
            _StyleDefs(44)  =   "Named:id=33:Normal"
            _StyleDefs(45)  =   ":id=33,.parent=0"
            _StyleDefs(46)  =   "Named:id=34:Heading"
            _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(48)  =   ":id=34,.wraptext=-1"
            _StyleDefs(49)  =   "Named:id=35:Footing"
            _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(51)  =   "Named:id=36:Selected"
            _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=37:Caption"
            _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(55)  =   "Named:id=38:HighlightRow"
            _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(57)  =   "Named:id=39:EvenRow"
            _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(59)  =   "Named:id=40:OddRow"
            _StyleDefs(60)  =   ":id=40,.parent=33"
            _StyleDefs(61)  =   "Named:id=41:RecordSelector"
            _StyleDefs(62)  =   ":id=41,.parent=34"
            _StyleDefs(63)  =   "Named:id=42:FilterBar"
            _StyleDefs(64)  =   ":id=42,.parent=33"
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
            TabIndex        =   6
            Top             =   30
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Centros de Costos por Area"
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
            TabIndex        =   5
            Top             =   30
            Width           =   7305
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
            TabIndex        =   4
            Top             =   0
            Visible         =   0   'False
            Width           =   1860
         End
      End
   End
End
Attribute VB_Name = "FrmCenCosArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstLista As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim QueHace As Integer
Dim CaracteresNumericos As String

Sub ActivaTool()
    Dim A As Integer
    For A = 1 To 16
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Private Sub CmdAddVar_Click()
    If NulosN(TxtIdArea.Text) = "" Then
        MsgBox "No ha especificado el area al que se le asignaran los centros de costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdArea.SetFocus
        Exit Sub
    End If
End Sub

Private Sub CmdAddCenCos_Click()
    If TxtIdArea.Text = "" Then
        MsgBox "No ha especificado el area para los centros de costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdArea.SetFocus
        Exit Sub
    End If
    Dim xFun As New SGI2_funciones.formularios
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    Set xRs = xFun.CentroCostoSeleBloques(xCon)
    
    If xRs.State <> 0 Then
        If xRs.RecordCount <> 0 Then
            xRs.MoveFirst
            For A = 1 To xRs.RecordCount
                Fg1.Rows = Fg1.Rows + 1
                Fg1.TextMatrix(Fg1.Rows - 1, 3) = xRs("id")
                Fg1.TextMatrix(Fg1.Rows - 1, 4) = xRs("tipo")
                If xRs("tipo") = 1 Then
                    FlexFormatoCelda Fg1, Fg1.Rows - 1, 1, &H800000, True, &HE0FEFE, xRs("codigo")
                    FlexFormatoCelda Fg1, Fg1.Rows - 1, 2, &H800000, True, &HE0FEFE, xRs("descripcion")
                Else
                    FlexFormatoCelda Fg1, Fg1.Rows - 1, 1, &H80000012, False, &HE0FEFE, xRs("codigo")
                    FlexFormatoCelda Fg1, Fg1.Rows - 1, 2, &H80000012, False, &HE0FEFE, xRs("descripcion")
                End If
                xRs.MoveNext
            Next A
        End If
    End If
    
    Set xFun = Nothing
End Sub

Private Sub CmdBusArea_Click()
    If QueHace <> 1 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_area ORDER BY descripcion"
    
    xform.Titulo = "Buscando Area"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdArea.Text = xRs("id")
            LblArea.Caption = xRs("descripcion")
            Fg1.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdDelCenCos_Click()
    If Fg1.Rows = 1 Then Exit Sub
    If NulosN(Fg1.TextMatrix(Fg1.Row, 4)) = 1 Then
        Dim Rpta, A As Integer
        Dim xLongitud As Integer
        Dim xCad As String
        Rpta = MsgBox("¿Esta seguro de eliminar este centro de costo?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
        If Rpta = vbYes Then
            xCad = Fg1.TextMatrix(Fg1.Row, 1)
            xLongitud = Len(Trim(Fg1.TextMatrix(Fg1.Row, 1)))
            
            For A = 1 To Fg1.Rows - 1
                If xCad = Mid(Trim(Fg1.TextMatrix(A, 1)), 1, xLongitud) Then
                    Fg1.RemoveItem Fg1.Row
                    A = A - 1
                End If
                If A = Fg1.Rows - 1 Then Exit Sub
            Next A
        End If
        
    Else
        MsgBox "No puede eliminar un centro de costo hijo, para eliminar este centro de costo elimine el centro de costo padre", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 1 Then
        Dim Rst As New ADODB.Recordset
        Dim A, B As Integer
        Dim Encontro As Boolean
        Dim xFrm As New SGI2_funciones.formularios
        Set Rst = xFrm.SeleCentroCosto(xCon)
        
        If Rst.State = 1 Then
            If Rst.RecordCount <> 0 Then
                Rst.MoveFirst
                For A = 1 To Rst.RecordCount
                    Fg1.Rows = Fg1.Rows + 1
                    Fg1.TextMatrix(Fg1.Rows - 1, 1) = Trim(Rst("codigo"))
                    Fg1.TextMatrix(Fg1.Rows - 1, 2) = Rst("descripcion")
                    Fg1.TextMatrix(Fg1.Rows - 1, 3) = Rst("idcencos")
                    
                    Rst.MoveNext
                    If Rst.EOF = True Then Exit For
                Next A
            End If
        End If
        Set xFrm = Nothing
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        Dim Rpta As Integer
        SeEjecuto = True
        
        RST_Busq RstLista, "SELECT DISTINCT con_centocostoarea.idarea, mae_area.descripcion, mae_area.abrev " _
            & " FROM con_centocostoarea LEFT JOIN mae_area ON con_centocostoarea.idarea = mae_area.id", xCon
        
        Set Dg1.DataSource = RstLista
        
        If RstLista.RecordCount = 0 Then
            Rpta = MsgBox("No se ha especificado centros de costo a las areas, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                Nuevo
            Else
                Set RstLista = Nothing
                Unload Me
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    QueHace = 3
    CaracteresNumericos = "0123456789." & Chr(8)
    TabOne1.CurrTab = 0
    Fg1.Editable = flexEDNone
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.ColWidth(3) = 0
    Fg1.ColWidth(4) = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then Modificar
    
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstLista.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 16 Then
        Set RstLista = Nothing
        Unload Me
    End If
End Sub

Sub Blanquea()
    TxtIdArea.Text = ""
    Fg1.Rows = 1
    LblArea.Caption = ""
End Sub

Sub Bloquea(Valor As Boolean)
    TxtIdArea.Locked = Valor
End Sub

Sub xEnabled(Valor As Boolean)
    CmdAddCenCos.Enabled = Valor
    CmdDelCenCos.Enabled = Valor
End Sub

Sub Cancelar()
    QueHace = 3
    Label5.Caption = "Detalle de Centros de Costo por Area"
    ActivaTool
    Bloquea True
    xEnabled False
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Fg1.Editable = flexEDNone
    Fg1.SelectionMode = flexSelectionByRow
End Sub

Function Grabar() As Boolean
    If NulosN(TxtIdArea.Text) = 0 Then
        MsgBox "No se ha especificado el area para el centro de costos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdArea.SetFocus
        Exit Function
    End If
    
    If Fg1.Rows = 1 Then
        MsgBox "No se han especificado centros de costo para el area especificada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdArea.SetFocus
        Exit Function
    End If
    
    Dim A As Integer
On Error GoTo LaCague
    
    xCon.BeginTrans
    
    xCon.Execute "DELETE con_centocostoarea.idarea, con_centocostoarea.* From con_centocostoarea WHERE (((con_centocostoarea.idarea)=" & NulosN(TxtIdArea.Text) & "))"

    For A = 1 To Fg1.Rows - 1
        xCon.Execute " INSERT INTO con_centocostoarea (idarea, idcencos) SELECT " & NulosN(TxtIdArea.Text) & " AS idarea, " & NulosN(Fg1.TextMatrix(A, 3)) & " AS idcencos"
    Next A
    
    xCon.CommitTrans
    Grabar = True
    Exit Function
    
LaCague:
    xCon.RollbackTrans
    MsgBox "No se pudo guardar el registro por el siguiente motivo: " & Err.Description, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = False
End Function

Sub Modificar()
    QueHace = 2
    ActivaTool
    Blanquea
    Bloquea True
    xEnabled True
    TabOne1.CurrTab = 1
    MuestraSegundoTab
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Modificando Centros de Costos al Area"
    Fg1.Editable = flexEDKbdMouse
    TxtIdArea.SetFocus
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    
    Rpta = MsgBox("Esta seguro de eliminar el area y sus centros de costos asignados", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM con_centocostoarea WHERE idarea = " & RstLista("idarea") & ""
        RstLista.Requery
        Dg1.Refresh
        If RstLista.RecordCount = 0 Then
            Rpta = MsgBox("No hay registros que mostrar, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                Nuevo
            Else
                Set RstLista = Nothing
                Unload Me
            End If
        End If
    End If
End Sub

Sub Nuevo()
    QueHace = 1
    ActivaTool
    Blanquea
    Bloquea False
    xEnabled True
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Asignando Centros de Costos al Area"
    Fg1.Editable = flexEDKbdMouse
    TxtIdArea.SetFocus
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
        LblArea.Caption = ""
        Exit Sub
    End If
    
    LblArea.Caption = Busca_Codigo(TxtIdArea.Text, "id", "descripcion", "mae_area", "N", xCon)
    If NulosC(LblArea.Caption) = "" Then
        TxtIdArea.Text = ""
        LblArea.Caption = ""
    End If
End Sub

Sub MuestraSegundoTab()
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    
    TxtIdArea.Text = NulosN(RstLista("idarea"))
    LblArea.Caption = NulosC(RstLista("descripcion"))
    
    RST_Busq xRs, "SELECT con_centrocosto.* FROM con_centocostoarea LEFT JOIN con_centrocosto ON con_centocostoarea.idcencos = con_centrocosto.id " _
        & " Where (((con_centocostoarea.idarea) = " & RstLista("idarea") & ")) ORDER BY con_centrocosto.codigo", xCon
    Fg1.Rows = Fg1.Rows + 1
    
    Fg1.Rows = 1
    If xRs.RecordCount <> 0 Then
        xRs.MoveFirst
        For A = 1 To xRs.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            
            If xRs("tipo") = 1 Then
                FlexFormatoCelda Fg1, Fg1.Rows - 1, 1, &H800000, True, &HE0FEFE, xRs("codigo")
                FlexFormatoCelda Fg1, Fg1.Rows - 1, 2, &H800000, True, &HE0FEFE, xRs("descripcion")
            Else
                FlexFormatoCelda Fg1, Fg1.Rows - 1, 1, &H80000012, False, &HE0FEFE, NulosC(xRs("codigo"))
                FlexFormatoCelda Fg1, Fg1.Rows - 1, 2, &H80000012, False, &HE0FEFE, NulosC(xRs("descripcion"))
            End If
            
            Fg1.TextMatrix(A, 3) = NulosN(xRs("id"))
            Fg1.TextMatrix(A, 4) = NulosN(xRs("tipo"))
            xRs.MoveNext
            If xRs.EOF = True Then Exit For
        Next A
    End If
    Set xRs = Nothing
End Sub

VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmManEquipos 
   Caption         =   "Mantenimiento de equipos"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7230
      Left            =   0
      TabIndex        =   9
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6810
         Left            =   12525
         TabIndex        =   13
         Top             =   375
         Width           =   11790
         Begin MSComDlg.CommonDialog CDAbrir 
            Left            =   6075
            Top             =   1695
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton CmdAgregar 
            Caption         =   "&Agregar Imagen"
            Enabled         =   0   'False
            Height          =   300
            Left            =   6720
            TabIndex        =   34
            Tag             =   "b"
            Top             =   6465
            Width           =   2505
         End
         Begin VB.CommandButton CmdQuitar 
            Caption         =   "&Eliminar Imagen"
            Enabled         =   0   'False
            Height          =   300
            Left            =   9255
            TabIndex        =   33
            Tag             =   "b"
            Top             =   6465
            Width           =   2505
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   1020
            Left            =   6705
            TabIndex        =   32
            Top             =   5430
            Width           =   5070
            _cx             =   8943
            _cy             =   1799
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
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmManEquipos.frx":0000
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
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchAdqui 
            Height          =   300
            Left            =   1320
            TabIndex        =   5
            Top             =   5370
            Width           =   1275
            _ExtentX        =   2249
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
         Begin VB.CommandButton CmdBusCenCos 
            Height          =   240
            Left            =   2160
            Picture         =   "FrmManEquipos.frx":0079
            Style           =   1  'Graphical
            TabIndex        =   29
            Tag             =   "b"
            Top             =   6030
            Width           =   250
         End
         Begin VB.CommandButton CmdBusArea 
            Height          =   240
            Left            =   1950
            Picture         =   "FrmManEquipos.frx":01AB
            Style           =   1  'Graphical
            TabIndex        =   28
            Tag             =   "b"
            Top             =   5715
            Width           =   250
         End
         Begin VB.CommandButton CmdBusPrio 
            Height          =   240
            Left            =   1950
            Picture         =   "FrmManEquipos.frx":02DD
            Style           =   1  'Graphical
            TabIndex        =   14
            Tag             =   "b"
            Top             =   6345
            Width           =   250
         End
         Begin VB.TextBox TxtNombre 
            Height          =   300
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   0
            Tag             =   "a"
            Text            =   "TxtNombre"
            Top             =   540
            Width           =   5040
         End
         Begin VB.TextBox TxtPrio 
            Height          =   300
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   8
            Tag             =   "a"
            Text            =   "TxtPrio"
            Top             =   6315
            Width           =   915
         End
         Begin VB.TextBox TxtCenCos 
            Height          =   300
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   7
            Tag             =   "a"
            Text            =   "TxtCenCos"
            Top             =   6000
            Width           =   1125
         End
         Begin VB.TextBox TxtIdArea 
            Height          =   300
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   6
            Tag             =   "a"
            Text            =   "TxtIdArea"
            Top             =   5685
            Width           =   915
         End
         Begin VB.TextBox TxtCaracteristicas 
            Height          =   3240
            Left            =   75
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   4
            Tag             =   "a"
            Text            =   "FrmManEquipos.frx":040F
            Top             =   2085
            Width           =   6495
         End
         Begin VB.TextBox TxtModelo 
            Height          =   300
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   2
            Tag             =   "a"
            Text            =   "TxtModelo"
            Top             =   1170
            Width           =   5040
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   3
            Tag             =   "a"
            Text            =   "TxtNumSer"
            Top             =   1485
            Width           =   5040
         End
         Begin VB.TextBox TxtMarca 
            Height          =   300
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   1
            Tag             =   "a"
            Text            =   "TxtMarca"
            Top             =   855
            Width           =   5040
         End
         Begin VB.Label LblIdCenCos 
            Caption         =   "LblIdCenCos"
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
            Height          =   225
            Left            =   3750
            TabIndex        =   35
            Top             =   5400
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Adquisision"
            Height          =   195
            Index           =   8
            Left            =   60
            TabIndex        =   31
            Top             =   5400
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "Lista de Fotos"
            Height          =   255
            Left            =   6705
            TabIndex        =   30
            Top             =   5205
            Width           =   3135
         End
         Begin VB.Image ImgFoto 
            BorderStyle     =   1  'Fixed Single
            Height          =   4740
            Left            =   6705
            Stretch         =   -1  'True
            Top             =   390
            Width           =   5040
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            Height          =   195
            Index           =   7
            Left            =   60
            TabIndex        =   27
            Top             =   585
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Prioridad"
            Height          =   195
            Index           =   4
            Left            =   60
            TabIndex        =   26
            Top             =   6345
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Centro de Costo"
            Height          =   195
            Index           =   3
            Left            =   60
            TabIndex        =   25
            Top             =   6045
            Width           =   1140
         End
         Begin VB.Label LblCentroCosto 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblCentroCosto"
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
            Left            =   2460
            TabIndex        =   24
            Tag             =   "a"
            Top             =   6000
            Width           =   4095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Area"
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   23
            Top             =   5715
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
            Left            =   2250
            TabIndex        =   22
            Tag             =   "a"
            Top             =   5685
            Width           =   4305
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Características"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   21
            Top             =   1845
            Width           =   1065
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Modelo"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   20
            Top             =   1200
            Width           =   525
         End
         Begin VB.Label LblPrioridad 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblPrioridad"
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
            Left            =   2250
            TabIndex        =   19
            Tag             =   "a"
            Top             =   6315
            Width           =   4305
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Marca"
            Height          =   195
            Index           =   6
            Left            =   60
            TabIndex        =   17
            Top             =   900
            Width           =   450
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle Orden de Compra"
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
            TabIndex        =   16
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Serie"
            Height          =   195
            Index           =   5
            Left            =   60
            TabIndex        =   15
            Top             =   1530
            Width           =   585
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6810
         Left            =   45
         TabIndex        =   10
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6420
            Left            =   30
            TabIndex        =   11
            ToolTipText     =   "Click derecho para Aceptar o Rechazar"
            Top             =   345
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11324
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
            Columns(1).Caption=   "Descripción"
            Columns(1).DataField=   "nombre"
            Columns(1).NumberFormat=   "Short Date"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Marca"
            Columns(2).DataField=   "marca"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Modelo"
            Columns(3).DataField=   "modelo"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Nº Serie"
            Columns(4).DataField=   "serie"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Caracteristica"
            Columns(5).DataField=   "caracteristicas"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Area"
            Columns(6).DataField=   "area"
            Columns(6).NumberFormat=   "0.00"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1244"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1164"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=3704"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3625"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=3307"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3228"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=3519"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=3440"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=2910"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2831"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=2778"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2699"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=2223"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2143"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=512"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=13,.alignment=0"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=66,.parent=13,.alignment=0"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
            _StyleDefs(64)  =   "Named:id=33:Normal"
            _StyleDefs(65)  =   ":id=33,.parent=0"
            _StyleDefs(66)  =   "Named:id=34:Heading"
            _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(68)  =   ":id=34,.wraptext=-1"
            _StyleDefs(69)  =   "Named:id=35:Footing"
            _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(71)  =   "Named:id=36:Selected"
            _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=37:Caption"
            _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(75)  =   "Named:id=38:HighlightRow"
            _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(77)  =   "Named:id=39:EvenRow"
            _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(79)  =   "Named:id=40:OddRow"
            _StyleDefs(80)  =   ":id=40,.parent=33"
            _StyleDefs(81)  =   "Named:id=41:RecordSelector"
            _StyleDefs(82)  =   ":id=41,.parent=34"
            _StyleDefs(83)  =   "Named:id=42:FilterBar"
            _StyleDefs(84)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Equipos"
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
            TabIndex        =   12
            Top             =   30
            Width           =   11610
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
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
            Picture         =   "FrmManEquipos.frx":0421
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManEquipos.frx":0965
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManEquipos.frx":0CF7
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManEquipos.frx":0E7B
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManEquipos.frx":12CF
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManEquipos.frx":13E7
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManEquipos.frx":192B
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManEquipos.frx":1E6F
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManEquipos.frx":1F83
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManEquipos.frx":2097
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManEquipos.frx":24EB
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManEquipos.frx":2657
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   18
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
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Ficha del Equipo"
            ImageIndex      =   12
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
Attribute VB_Name = "FrmManEquipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SeEjecuto As Boolean
Dim CaracteresNumericos As String, CaracteresNumericos2 As String, vMant As String
Dim Quehace As Integer
Dim RsEquipo As New ADODB.Recordset
Dim xRsFotosTemp As New ADODB.Recordset
Dim RstTmp As New ADODB.Recordset
Dim vPuntero As Integer
Dim Agregando As Boolean
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO


Sub Buscar()
    TabOne1.CurrTab = 0
    
    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":         xCampos(0, 1) = "descripcion":          xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id":            xCampos(1, 1) = "id":           xCampos(1, 2) = "1500":         xCampos(1, 3) = "N"
    
    xForm.SQLCad = "SELECT man_equipo.descripcion, man_equipo.id, man_equipo.capacidad " _
        & "FROM man_equipo"

    xForm.Titulo = "Buscando Equipo"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "descripcion"
    xForm.CampoBusca = "descripcion"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        RsEquipo.MoveFirst
        RsEquipo.Find "id = " & xRs("id") & ""
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub LimpiarTxtFoto()
    'Limpiar
    'LblIdFoto.Caption = ""
    'TxtRuta.Text = ""
'    TxtDesFoto.Text = ""
    'fin Limpiar
End Sub

Sub MuestraSegundoTab()
    Blanquea
    
    TxtNombre.Text = RsEquipo("nombre")
    TxtMarca.Text = NulosC(RsEquipo("marca"))
    TxtModelo.Text = NulosC(RsEquipo("modelo"))
    TxtNumSer.Text = NulosC(RsEquipo("serie"))
    
    
    TxtCenCos.Text = NulosC(RsEquipo("codcemcos"))
    LblCentroCosto.Caption = NulosC(RsEquipo("desccencos"))
    LblIdCenCos.Caption = RsEquipo("idcencos")
    
    TxtCaracteristicas.Text = NulosC(RsEquipo("caracteristicas"))
    If IsNull(RsEquipo("fchadq")) = True Then
        TxtFchAdqui.Valor = ""
    Else
        TxtFchAdqui.Valor = RsEquipo("fchadq")
    End If
    TxtIdArea.Text = RsEquipo("idarea")
    'TxtCenCos.Text = RsEquipo("idcencos")
    TxtPrio.Text = RsEquipo("idprio")
    
    TxtIdArea_Validate True
    TxtPrio_Validate True
    
    Fg1.Rows = 1
    Dim RstTmp As New ADODB.Recordset
    Dim xRuta As String
    Dim A As Integer
    Dim rst As New ADODB.Recordset
    
    xRuta = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTAAR", "RUTAS")
    xRuta = xRuta + "0002\"
    Agregando = True
    RST_Busq RstTmp, "SELECT * FROM man_equiposfoto WHERE idequipo = " & Val(RsEquipo("id")) & " ORDER BY id", xCon
    If RstTmp.RecordCount <> 0 Then
        RstTmp.MoveFirst
        For A = 1 To RstTmp.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = RstTmp("descripcion")
            Fg1.TextMatrix(A, 2) = xRuta + NulosC(RstTmp("archivo"))
            Fg1.TextMatrix(A, 3) = RstTmp("id")
            RstTmp.MoveNext
            If RstTmp.EOF = True Then
                Exit For
            End If
        Next A
    End If
    Set rst = Nothing
    Dim xFile As New FileSystemObject
    
    If Fg1.Rows >= 2 Then
        If xFile.FileExists(Fg1.TextMatrix(1, 2)) = True Then
            ImgFoto = LoadPicture(NulosC(Fg1.TextMatrix(1, 2)))
        Else
            MsgBox "El archivo de imagen asociado al registro ha sido eliminado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Fg1.Rows = 1
            ImgFoto = LoadPicture("")
        End If
    Else
        ImgFoto = LoadPicture("")
    End If
    Agregando = False
End Sub

Sub Cancelar()
    Quehace = 3
    ActivaTool
    Bloquea
    Label5.Caption = "Detalle de Equipos"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Dg1.SetFocus
End Sub

Function Grabar() As Boolean
    Dim xId As Double
    Dim xCampos(10, 5) As String
    Dim xCampos2(4, 5) As String
    Dim A As Integer
    
On Error GoTo LaCague

    xCon.BeginTrans
    
    If Quehace = 1 Then
        xId = HallaCodigoTabla("man_equipos", xCon, "id")
    Else
        xId = RsEquipo("id")
        'xCon.Execute "DELETE * FROM  pro_cronogramadet WHERE id = " & RstLis("id") & ""
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
    xCampos(0, 0) = "id":              xCampos(0, 1) = Str(xId):                     xCampos(0, 2) = "S":    xCampos(0, 3) = "N":     xCampos(0, 4) = "":                                                       xCampos(0, 5) = "S"
    xCampos(1, 0) = "fchadq":          xCampos(1, 1) = TxtFchAdqui.Valor:            xCampos(1, 2) = "S":    xCampos(1, 3) = "F":     xCampos(1, 4) = "No ha especificado la fecha de adquisicion":             xCampos(1, 5) = ""
    xCampos(2, 0) = "nombre":          xCampos(2, 1) = TxtNombre.Text:               xCampos(2, 2) = "S":    xCampos(2, 3) = "C":     xCampos(2, 4) = "No ha especificado el nombre del equipo":                xCampos(2, 5) = ""
    xCampos(3, 0) = "marca":           xCampos(3, 1) = TxtMarca.Text:                xCampos(3, 2) = "S":    xCampos(3, 3) = "C":     xCampos(3, 4) = "No ha especificado la marca del equipo":                 xCampos(3, 5) = ""
    xCampos(4, 0) = "modelo":          xCampos(4, 1) = TxtModelo.Text:               xCampos(4, 2) = "S":    xCampos(4, 3) = "C":     xCampos(4, 4) = "No ha especificado el modelo del equipo":                xCampos(4, 5) = ""
    xCampos(5, 0) = "serie":           xCampos(5, 1) = TxtNumSer.Text:               xCampos(5, 2) = "S":    xCampos(5, 3) = "C":     xCampos(5, 4) = "No ha especificado el numero de serie del equipo":       xCampos(5, 5) = ""
    xCampos(6, 0) = "caracteristicas": xCampos(6, 1) = TxtCaracteristicas.Text:      xCampos(6, 2) = "S":    xCampos(6, 3) = "C":     xCampos(6, 4) = "No ha especificado las caracteristicas del equipo":      xCampos(6, 5) = ""
    xCampos(7, 0) = "observaciones":   xCampos(7, 1) = "":                           xCampos(7, 2) = "N":    xCampos(7, 3) = "C":     xCampos(7, 4) = "No ha especificado el tipo de producto":                 xCampos(7, 5) = ""
    xCampos(8, 0) = "idarea":          xCampos(8, 1) = TxtIdArea.Text:               xCampos(8, 2) = "S":    xCampos(8, 3) = "N":     xCampos(8, 4) = "No ha especificado el area al que pertenece el equipo":  xCampos(8, 5) = ""
    xCampos(9, 0) = "idprio":          xCampos(9, 1) = TxtPrio.Text:                 xCampos(9, 2) = "S":    xCampos(9, 3) = "N":     xCampos(9, 4) = "No ha especificado la prioridad del equipo":             xCampos(9, 5) = ""
    xCampos(10, 0) = "idcencos":       xCampos(10, 1) = NulosC(LblIdCenCos.Caption): xCampos(10, 2) = "S":   xCampos(10, 3) = "N":    xCampos(10, 4) = "No ha especificado el centro de costo":                 xCampos(10, 5) = ""
    
    If Quehace = 1 Then
        If EscribirNuevoRegistro(xCampos, "man_equipos", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
    Else
        If ModificarRegistro(xCampos, "man_equipos", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
    End If
    
    'GRABAR FOTOS
    Dim rst As New ADODB.Recordset
    Dim xArchivo As String
    Dim xRuta As String
    Dim xItemFoto As Integer
    
    If Fg1.Rows > 1 Then
        RST_Busq rst, "SELECT * FROM man_equiposfoto", xCon
        Dim xFile As New FileSystemObject
        
        For A = 1 To Fg1.Rows - 1
            xArchivo = ""
            xRuta = ""
            If NulosN(Fg1.TextMatrix(A, 3)) = 0 Then
                rst.AddNew
                rst("idequipo") = xId
                xItemFoto = xItemFoto + 1
                Fg1.TextMatrix(A, 3) = xItemFoto
                rst("id") = xItemFoto
                rst("descripcion") = NulosC(Fg1.TextMatrix(A, 1))
                
                xArchivo = Format(xId, "0000") + "-" + Format(NulosN(Fg1.TextMatrix(A, 3)), "0000") + ".jpg"
                xRuta = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTAAR", "RUTAS")
                ' 0002 = indica la carpeta donde se almacenaran las imagenes de los equipos
                xRuta = xRuta + "0002\" + xArchivo
                xFile.CopyFile NulosC(Fg1.TextMatrix(A, 2)), xRuta
                rst("archivo") = NulosC(xArchivo)
                rst.Update
            Else
                xItemFoto = Fg1.TextMatrix(A, 3)
            End If
        Next A
    End If
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, Quehace, xHorIni, Time, Date, xCon, xId
    
    xCon.CommitTrans
    
    MsgBox "El equipo se grabo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Exit Function
    
LaCague:
    Resume
    xCon.RollbackTrans
    'Set RsFotos = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function

Sub Eliminar()
    Dim Rpta, A As Integer
    Dim xRuta As String
    Dim xFile As New FileSystemObject
    
    xRuta = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTAAR", "RUTAS")
    xRuta = xRuta + "0002\"
    
    Rpta = MsgBox("¿Esta seguro de eliminar el equipo seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        TabOne1.CurrTab = 0
        
        ' eliminamos las fotos
        Dim Rs As New ADODB.Recordset
        RST_Busq Rs, "SELECT * FROM man_equiposfoto WHERE idequipo = " & RsEquipo("id"), xCon
        
        If Rs.RecordCount <> 0 Then
            Rs.MoveFirst
            For A = 1 To Rs.RecordCount
                ' NOS ASEGURAMOS QUE EL ARCHIVO EXISTE
                If xFile.FileExists(xRuta + NulosC(Rs("archivo"))) = True Then
                    ' SI EXISTE LO BORRRAMOS
                    xFile.DeleteFile xRuta + NulosC(Rs("archivo"))
                End If
                Rs.MoveNext
                If Rs.EOF = True Then Exit For
            Next A
        End If
        
        ' BORRAMOS EL EQUIPO DE LA bd
        xCon.Execute "DELETE * FROM man_equipos WHERE id = " & RsEquipo("id") & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RsEquipo("id") & " AND idform = " & IdMenuActivo

        
        MsgBox "El equipo se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RsEquipo.Requery
        Dg1.Refresh
        Dg1.SetFocus
    End If
End Sub

Sub Modificar()
    Quehace = 2
    xHorIni = Time
    ActivaTool
    Blanquea
    Bloquea
    Label5.Caption = "Modificando Equipos"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    MuestraSegundoTab
    
    Fg1.SelectionMode = flexSelectionFree
    Fg1.ColComboList(1) = "|..."
    Fg1.Editable = flexEDKbdMouse
    
    TxtNombre.SetFocus
End Sub

Sub Blanquea()
    TxtNombre.Text = ""
    TxtMarca.Text = ""
    TxtModelo.Text = ""
    TxtNumSer.Text = ""
    TxtCaracteristicas.Text = ""
    TxtFchAdqui.Valor = ""
    TxtIdArea.Text = ""
    TxtCenCos.Text = ""
    TxtPrio.Text = ""
    LblArea.Caption = ""
    LblCentroCosto.Caption = ""
    LblPrioridad.Caption = ""
    LblIdCenCos.Caption = ""
End Sub

Sub Bloquea()
    TxtNombre.Locked = Not TxtNombre.Locked
    TxtMarca.Locked = Not TxtMarca.Locked
    TxtModelo.Locked = Not TxtModelo.Locked
    TxtNumSer.Locked = Not TxtNumSer.Locked
    TxtCaracteristicas.Locked = Not TxtCaracteristicas.Locked
    TxtFchAdqui.Locked = Not TxtFchAdqui.Locked
    TxtIdArea.Locked = Not TxtIdArea.Locked
    'TxtCenCos.Locked = Not TxtCenCos.Locked
    TxtPrio.Locked = Not TxtPrio.Locked
    
    CmdAgregar.Enabled = Not CmdAgregar.Enabled
    CmdQuitar.Enabled = Not CmdQuitar.Enabled
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

Sub Nuevo()
    Quehace = 1
    xHorIni = Time
    ActivaTool
    Blanquea
    Bloquea
    Label5.Caption = "Agregando Equipos"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Fg1.Rows = 1
    Fg1.SelectionMode = flexSelectionFree
    Fg1.ColComboList(1) = "|..."
    Fg1.Editable = flexEDKbdMouse
    TxtNombre.SetFocus
End Sub

Private Sub CmdAgregar_Click()
    Fg1.Rows = Fg1.Rows + 1
End Sub

Private Sub CmdBusArea_Click()
    If Quehace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xForm.SQLCad = "SELECT * FROM pla_area"
    
    xForm.Titulo = "Buscando Area"
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
            TxtCenCos.SetFocus
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusCenCos_Click()
    If Quehace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Codigo":        xCampos(0, 1) = "codigo":       xCampos(0, 2) = "2000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Descripcion":   xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "5000":         xCampos(1, 3) = "C"
    
    xForm.SQLCad = "SELECT con_centrocosto.codigo, con_centrocosto.descripcion, con_centrocosto.id From con_centrocosto ORDER BY con_centrocosto.codigo"
    
    xForm.Titulo = "Buscando Centro de Costo"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "codigo"
    xForm.CampoBusca = "codigo"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            LblIdCenCos.Caption = xRs("id")
            TxtCenCos.Text = xRs("codigo")
            LblCentroCosto.Caption = xRs("descripcion")
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing

End Sub

Private Sub CmdBusPrio_Click()
    If Quehace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xForm.SQLCad = "SELECT * FROM mae_prioridad"
    
    xForm.Titulo = "Buscando Prioridad"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "descripcion"
    xForm.CampoBusca = "descripcion"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtPrio.Text = xRs("id")
            LblPrioridad.Caption = xRs("descripcion")
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusPrio_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
    
    End If
End Sub

Private Sub CmdBusUniMed_Click()

End Sub

Private Sub CmdQuitar_Click()
    If Fg1.Rows = 1 Then
        MsgBox "No ha imagenes para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    Dim Rpta As Integer
    Rpta = MsgBox("¿ Esta seguro de eliminar la imagen selecciona ?, la imagen eliminada no podra ser recuperada", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        Dim xFile As New FileSystemObject
        
        If NulosN(Fg1.TextMatrix(Fg1.Row, 3)) <> 0 Then
            If xFile.FileExists(NulosC(Fg1.TextMatrix(Fg1.Row, 2))) = True Then
                xFile.DeleteFile NulosC(Fg1.TextMatrix(Fg1.Row, 2))
            End If
            xCon.Execute "DELETE * FROM man_equipofoto WHERE idequipo = " & RsEquipo("id") & " AND id = " & NulosC(Fg1.TextMatrix(Fg1.Row, 3)) & ""
            Fg1.RemoveItem Fg1.Row
            MsgBox "La imagen se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Else
            Fg1.RemoveItem Fg1.Row
        End If
        
        If Fg1.Rows > 1 Then
            Fg1.Select 1, 1, 1, 1
            ImgFoto.Picture = LoadPicture(Fg1.TextMatrix(Fg1.Row, 2))
        Else
            ImgFoto.Picture = LoadPicture("")
        End If
    End If
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
    MuestraSegundoTab
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
  
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RsEquipo("id")), xCon
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 1 Then
        CDAbrir.Filter = "Archivos jpg (*.jpg)|*.jpg"
        CDAbrir.ShowOpen
        Fg1.TextMatrix(Fg1.Row, 2) = CDAbrir.FileName
        ImgFoto = LoadPicture(Fg1.TextMatrix(Fg1.Row, 2))
    End If
End Sub

Private Sub Fg1_RowColChange()
    Dim Archivo As String
    Dim xFile As New FileSystemObject
    If Agregando = True Then Exit Sub
    If NulosC(Fg1.TextMatrix(Fg1.Row, 2)) <> "" Then
        Archivo = Fg1.TextMatrix(Fg1.Row, 2)
        If xFile.FileExists(Archivo) = True Then
            ImgFoto = LoadPicture(Fg1.TextMatrix(Fg1.Row, 2))
        Else
            ImgFoto = LoadPicture("")
        End If
    Else
        ImgFoto = LoadPicture("")
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        Dim Rpta As Integer
        Blanquea
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
    
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
        
        RST_Busq RsEquipo, "SELECT man_equipos.*, pla_area.descripcion AS area, con_centrocosto.codigo AS codcemcos, con_centrocosto.descripcion AS desccencos" _
            & " FROM (man_equipos LEFT JOIN pla_area ON man_equipos.idarea = pla_area.id) LEFT JOIN con_centrocosto ON man_equipos.idcencos = con_centrocosto.id " _
            & " ORDER BY man_equipos.nombre", xCon

        'SELECT man_equipos.*, pla_area.descripcion AS area FROM man_equipos LEFT JOIN pla_area " _
            & " ON man_equipos.idarea = pla_area.id ORDER BY man_equipos.nombre", xCon

        Set Dg1.DataSource = RsEquipo

    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    Quehace = 3
    SeEjecuto = False
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    TabOne1.CurrTab = 0
    Fg1.ColWidth(2) = 0
    Fg1.ColWidth(3) = 0
    
    CaracteresNumericos = "0123456789." & Chr(8)
    CaracteresNumericos2 = "0123456789." & Chr(8) & Chr(13)
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If Quehace <> 1 Then MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then
        Modificar
    End If
    
    If Button.Index = 3 Then
        Eliminar
    End If
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RsEquipo.Requery
            Dg1.Refresh
            Dg1.SetFocus
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 10 Then Buscar
'
'    If Button.Index = 12 Then
'        FrmPrintEquiFicha.propFormulario = "FormFicEqu"
'        FrmPrintEquiFicha.propId = RsEquipo("id")
'        FrmPrintEquiFicha.Show
'    End If
    
    If Button.Index = 14 Then
        Set RsEquipo = Nothing
        Unload Me
    End If
End Sub

Private Sub TxtCenCos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtCenCos_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCenCos_Click
    End If
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
    If NulosC(TxtIdArea.Text) <> "" Then
        Set RstTmp = Nothing
        Set RstTmp = BuscaConCriterio("SELECT * FROM pla_area WHERE id = " & Val(TxtIdArea.Text) & "", xCon)
        
        If RstTmp.RecordCount <> 0 Then
            LblArea.Caption = Trim(RstTmp("descripcion"))
        Else
            TxtIdArea.Text = ""
            LblArea.Caption = ""
        End If
    End If
    Set RstTmp = Nothing
End Sub

Private Sub TxtMarca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtModelo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumSer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtPrio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtPrio_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusPrio_Click
    End If
End Sub

Private Sub TxtPrio_Validate(Cancel As Boolean)
    If NulosC(TxtPrio.Text) = "" Then Exit Sub
    
    LblPrioridad.Caption = Busca_Codigo(NulosN(TxtPrio.Text), "Id", "descripcion", "mae_prioridad", "N", xCon)
    If NulosC(LblPrioridad.Caption) = "" Then
        TxtPrio.Text = ""
    End If
End Sub

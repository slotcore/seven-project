VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmPVEstimado 
   Caption         =   "Sistena de ventas - Estimado de Ventas"
   ClientHeight    =   7710
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7335
      Left            =   0
      TabIndex        =   6
      Top             =   375
      Width           =   11880
      _cx             =   20955
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   6915
         Left            =   45
         TabIndex        =   10
         Top             =   375
         Width           =   11790
         Begin VB.CommandButton CmdVerGraf 
            Caption         =   "Ver Grafico"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4155
            TabIndex        =   30
            Top             =   4635
            Width           =   1665
         End
         Begin VB.CommandButton CmdMax 
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   11250
            TabIndex        =   29
            ToolTipText     =   "Aumentar ancho de columna"
            Top             =   4620
            Width           =   375
         End
         Begin VB.CommandButton CmdMin 
            Caption         =   "<<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10860
            TabIndex        =   28
            ToolTipText     =   "Disminuir ancho de columna"
            Top             =   4620
            Width           =   375
         End
         Begin VB.TextBox TxtPorcentaje 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   9810
            Locked          =   -1  'True
            TabIndex        =   26
            Text            =   "TxtPorcentaje"
            Top             =   4635
            Width           =   765
         End
         Begin VB.CommandButton CmdHistVenta 
            Caption         =   "Ver Historico Ventas"
            Enabled         =   0   'False
            Height          =   300
            Left            =   5835
            TabIndex        =   5
            Top             =   4635
            Width           =   1665
         End
         Begin VB.TextBox TxtDesc 
            Height          =   300
            Left            =   1155
            Locked          =   -1  'True
            TabIndex        =   0
            Text            =   "TxtDesc"
            Top             =   390
            Width           =   9780
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
            Height          =   300
            Left            =   1155
            TabIndex        =   1
            Top             =   705
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
         Begin VSFlex7Ctl.VSFlexGrid Fg2 
            Height          =   1965
            Left            =   45
            TabIndex        =   4
            Top             =   4950
            Width           =   11745
            _cx             =   20717
            _cy             =   3466
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
            Cols            =   16
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmPVEstimado.frx":0000
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
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
            Height          =   300
            Left            =   5070
            TabIndex        =   2
            Top             =   705
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
            Height          =   2670
            Left            =   45
            TabIndex        =   3
            Top             =   1260
            Width           =   11745
            _cx             =   20717
            _cy             =   4710
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
            Cols            =   16
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmPVEstimado.frx":01B5
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
            AutoSize        =   -1  'True
            Caption         =   "Aplicar Porcentaje"
            Height          =   195
            Left            =   8415
            TabIndex        =   27
            Top             =   4680
            Width           =   1290
         End
         Begin VB.Label LblNumItem 
            Alignment       =   1  'Right Justify
            Caption         =   "LblNumItem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   11085
            TabIndex        =   25
            Top             =   3990
            Width           =   675
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nº Productos Procesados  :"
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
            Height          =   195
            Left            =   8640
            TabIndex        =   24
            Top             =   3990
            Width           =   2370
         End
         Begin VB.Label Label9 
            Caption         =   "Unidad Medida"
            Height          =   255
            Left            =   4725
            TabIndex        =   23
            Top             =   3960
            Width           =   1200
         End
         Begin VB.Label LblUniMed 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblUniMed"
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   6000
            TabIndex        =   22
            Top             =   3945
            Width           =   1125
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000003&
            X1              =   90
            X2              =   11760
            Y1              =   4590
            Y2              =   4590
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            X1              =   75
            X2              =   11745
            Y1              =   4605
            Y2              =   4605
         End
         Begin VB.Label LblCodigo 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblCodigo"
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   1155
            TabIndex        =   21
            Top             =   3945
            Width           =   2160
         End
         Begin VB.Label Label7 
            Caption         =   "Codigo"
            Height          =   255
            Left            =   60
            TabIndex        =   20
            Top             =   3960
            Width           =   1005
         End
         Begin VB.Label LblDesc 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDesc"
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   1155
            TabIndex        =   19
            Top             =   4245
            Width           =   10620
         End
         Begin VB.Label Label5 
            Caption         =   "Descripcion"
            Height          =   255
            Left            =   60
            TabIndex        =   18
            Top             =   4260
            Width           =   1005
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Inicio"
            Height          =   195
            Left            =   60
            TabIndex        =   16
            Top             =   735
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Descripcion"
            Height          =   195
            Left            =   60
            TabIndex        =   15
            Top             =   420
            Width           =   840
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle Proyeccion de Ventas"
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
            Top             =   45
            Width           =   11610
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Productos "
            Height          =   195
            Left            =   60
            TabIndex        =   13
            Top             =   1035
            Width           =   765
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Cronograma de Entegra"
            Height          =   225
            Left            =   60
            TabIndex        =   12
            Top             =   4680
            Width           =   1680
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Termino"
            Height          =   195
            Left            =   3900
            TabIndex        =   11
            Top             =   735
            Width           =   930
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6915
         Left            =   -12435
         TabIndex        =   7
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6570
            Left            =   30
            TabIndex        =   8
            Top             =   375
            Width           =   11790
            _ExtentX        =   20796
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
            Caption         =   "Consulta proyeccion de Ventas"
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
            Left            =   120
            TabIndex        =   9
            Top             =   45
            Width           =   11610
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
            Picture         =   "FrmPVEstimado.frx":03A1
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPVEstimado.frx":08E5
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPVEstimado.frx":0A69
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPVEstimado.frx":0EBD
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPVEstimado.frx":0FD5
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPVEstimado.frx":1519
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPVEstimado.frx":1A5D
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPVEstimado.frx":1B71
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPVEstimado.frx":1C85
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPVEstimado.frx":20D9
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPVEstimado.frx":2245
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
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
                  Text            =   "Modificar Proyeccion de Ventas"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Activar Proyeccion de Ventas"
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
                  Text            =   "Eliminar Proyeccion de Ventas"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Desactivar Proyeccion de Ventas"
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
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Menu1_1 
         Caption         =   "Agregar Promedio - Mes Actual"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "Agregar Promedio - Todos los meses"
      End
      Begin VB.Menu Menu1_5 
         Caption         =   "Agregar año seleccionado y aplicar porcentaje"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_4 
         Caption         =   "Limpiar varianzas"
      End
      Begin VB.Menu Menu1_6 
         Caption         =   "Exportar a Excel"
      End
   End
   Begin VB.Menu Menu2 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu Menu2_1 
         Caption         =   "Agregar Producto"
      End
      Begin VB.Menu Menu2_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu2_3 
         Caption         =   "Eliminar Producto"
      End
   End
End
Attribute VB_Name = "FrmPVEstimado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMPVESTIMADO
'* Tipo             : FORMULARIO
'* Descripcion      : PERMITE GENERAR UNA PROYECCION DE LAS VENTAS, EN FUNCION A LOS DATOS HISTORICOS
'*                    DE VENTA
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 10/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim RstPlanes As New ADODB.Recordset   ' RECORSET QUE ALMACENRA LOS REGISTRO DE LA TABLA ges_ventaproy
Dim QueHace As Integer                 ' INDICA EN QUE MODO SE ENCUENTRA EL FORMULARIO
Dim SeEjecuto As Boolean               ' INDICA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim xTitulo As String                  ' ALAMCENA EL TITULO DEL FORMULARIO
Dim xFilaActual As Integer             ' INDICA LA FILA ACTUAL PARA EL CONTROL FLEXGRID
Dim Agregando  As Boolean              ' VARIABLE QUE INDICA QUE SE ESTA AGREGANDO UNA FILA AL CONTROL FLEXGRID

'*****************************************************************************************************
'* Nombre Archivo   : Bloquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES TEXTBOX DEL FORMULARIO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Bloquea()
    TxtDesc.Locked = Not TxtDesc.Locked
    TxtFchIni.Enabled = Not TxtFchIni.Enabled
    TxtFchFin.Enabled = Not TxtFchFin.Enabled
    CmdHistVenta.Enabled = Not CmdHistVenta.Enabled
    TxtPorcentaje.Locked = Not TxtPorcentaje.Locked
    Fg1.Rows = 1
    Fg2.Rows = 1
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UNA REGISTRO DE LA TABLA ges_ventaproy
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Eliminar()
    Dim Rpta As Integer
    Rpta = MsgBox("¿Esta seguro de eliminar la proyeccion de ventas seleccionada?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM ges_ventaproy WHERE id = " & RstPlanes("id") & ""
        xCon.Execute "DELETE * FROM ges_ventaproydet WHERE id = " & RstPlanes("id") & ""
        MsgBox "La proyeccion de ventas se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstPlanes.Requery
        Dg1.Refresh
    End If
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA LOS CONTROLES TEXTOBOX PARA EL INGRESO DE UN NUEVO REGISTRO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Blanquea()
    TxtDesc.Text = ""
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    TxtPorcentaje.Text = ""
    LblUniMed.Caption = ""
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Grabar
'* Tipo             : FUNCION
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA ges_ventaproy, ESTA FUNCION DEVUELVE VERDADERO
'*                    CUANDO TIENE EXITO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         : Boolean
'*****************************************************************************************************
Function Grabar() As Boolean
    If TxtDesc.Text = "" Then
        MsgBox "No ha especificado la descripcion del plan", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtDesc.SetFocus
        Exit Function
    End If
    
    If TxtFchIni.Valor = "" Then
        MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Function
    End If

    If TxtFchFin.Valor = "" Then
        MsgBox "No ha especificado la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchFin.SetFocus
        Exit Function
    End If

    Dim A As Integer
    
    'eliminar las filas que esten vacias
    For A = 1 To Fg1.Rows
        If Fg1.TextMatrix(A, 1) = "" Then
            Fg1.RemoveItem (A)
            A = A - 1
        End If
        
        If A = Fg1.Rows - 1 Then
            Exit For
        End If
    Next A

    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xId As Integer
    
    On Error GoTo LaCague

    xCon.BeginTrans
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT * FROM ges_ventaproy", xCon
        RST_Busq RstDet, "SELECT * FROM ges_ventaproydet", xCon
        
        xId = HallaCodigoTabla("ges_ventaproy", xCon, "id")
        RstCab.AddNew
        
        RstCab("id") = xId
    Else
        RST_Busq RstCab, "SELECT * FROM ges_ventaproy WHERE id = " & RstPlanes("id") & "", xCon
        xCon.Execute "DELETE * FROM ges_ventaproydet WHERE id = " & RstPlanes("id") & ""
        RST_Busq RstDet, "SELECT * FROM ges_ventaproydet", xCon
        
        xId = RstPlanes("id")
    End If
    
    RstCab("descripcion") = NulosC(TxtDesc.Text)
    RstCab("fchini") = TxtFchIni.Valor
    RstCab("fchfin") = TxtFchFin.Valor
    RstCab.Update
    
    For A = 1 To Fg1.Rows
        RstDet.AddNew
        RstDet("id") = xId
        RstDet("idpro") = NulosN(Fg1.TextMatrix(A, 0))
        RstDet("ene") = NulosN(Fg1.TextMatrix(A, 3))
        RstDet("feb") = NulosN(Fg1.TextMatrix(A, 4))
        RstDet("mar") = NulosN(Fg1.TextMatrix(A, 5))
        RstDet("abr") = NulosN(Fg1.TextMatrix(A, 6))
        RstDet("may") = NulosN(Fg1.TextMatrix(A, 7))
        RstDet("jun") = NulosN(Fg1.TextMatrix(A, 8))
        RstDet("jul") = NulosN(Fg1.TextMatrix(A, 9))
        RstDet("ago") = NulosN(Fg1.TextMatrix(A, 10))
        RstDet("set") = NulosN(Fg1.TextMatrix(A, 11))
        RstDet("oct") = NulosN(Fg1.TextMatrix(A, 12))
        RstDet("nov") = NulosN(Fg1.TextMatrix(A, 13))
        RstDet("dic") = NulosN(Fg1.TextMatrix(A, 14))
        RstDet.Update
        
        If A = Fg1.Rows - 1 Then
            Exit For
        End If
    Next A
    
    xCon.CommitTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    
    MsgBox "El plan proyectado de ventas se guardo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
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
'* Nombre Archivo   : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL PROCESO DE INGRESAR O MODIFICAR UN REGISTRO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    Bloquea
    Toolbar
    Label1.Caption = "Detalle Proyeccion de Ventas"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Fg1.ColComboList(1) = ""
    Fg1.Editable = flexEDNone
    
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.BackColorSel = &H80&
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA EL INGRESO DE UN NUEVO REGISTRO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Nuevo()
    QueHace = 1
    Label1.Caption = "Agregando Proyeccion de Ventas"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Toolbar
    Bloquea
    Blanquea
    Fg1.ColComboList(1) = "|..."
    Fg1.Rows = Fg1.Rows + 1
    Fg1.Editable = flexEDKbdMouse
    Fg2.Editable = flexEDNone
    
    TxtDesc.SetFocus
    
    Fg1.SelectionMode = flexSelectionFree
    Fg1.BackColorSel = &H80&
    LblNumItem.Caption = 0
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
    Label1.Caption = "Modificando Proyeccion de Ventas"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Toolbar
    Bloquea
    Blanquea
    Fg1.ColComboList(1) = "|..."
    Fg1.Rows = Fg1.Rows + 1
    Fg1.Editable = flexEDKbdMouse
    Fg2.Editable = flexEDNone
    MuestraSegundoTab
    TxtDesc.SetFocus

    Fg1.SelectionMode = flexSelectionFree
    Fg1.BackColorSel = &H80&
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
'* Nombre Archivo   : CmdHistVenta_Click
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EL HISTORICO DE LAS VENTAS BUSCANDO EN LA DATA DE OTROS AÑOS
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Private Sub CmdHistVenta_Click()
    Dim RstRutas As New ADODB.Recordset
    Dim RstHisVenta As New ADODB.Recordset
    Dim A, B As Integer
    Dim NumAños As Integer
    
    Fg2.Rows = 1
    RST_Busq RstRutas, "SELECT * FROM ges_rutahis WHERE activo = -1 ORDER BY año", xCon
    
    NumAños = RstRutas.RecordCount
    
    If RstRutas.RecordCount <> 0 Then
        RstRutas.MoveFirst
        
        For A = 1 To RstRutas.RecordCount
            Set RstHisVenta = MostrarAños(RstRutas("año"), Val(Fg1.TextMatrix(Fg1.Row, 0)), RstRutas("ruta"))
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = RstRutas("año")
            If RstHisVenta.RecordCount <> 0 Then
                Fg2.TextMatrix(Fg2.Rows - 1, 2) = Format(RstHisVenta("ene"), "0.00")
                Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(RstHisVenta("feb"), "0.00")
                Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(RstHisVenta("mar"), "0.00")
                Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(RstHisVenta("abr"), "0.00")
                Fg2.TextMatrix(Fg2.Rows - 1, 6) = Format(RstHisVenta("may"), "0.00")
                Fg2.TextMatrix(Fg2.Rows - 1, 7) = Format(RstHisVenta("jun"), "0.00")
                Fg2.TextMatrix(Fg2.Rows - 1, 8) = Format(RstHisVenta("jul"), "0.00")
                Fg2.TextMatrix(Fg2.Rows - 1, 9) = Format(RstHisVenta("ago"), "0.00")
                Fg2.TextMatrix(Fg2.Rows - 1, 10) = Format(RstHisVenta("sep"), "0.00")
                Fg2.TextMatrix(Fg2.Rows - 1, 11) = Format(RstHisVenta("oct"), "0.00")
                Fg2.TextMatrix(Fg2.Rows - 1, 12) = Format(RstHisVenta("nov"), "0.00")
                Fg2.TextMatrix(Fg2.Rows - 1, 13) = Format(RstHisVenta("dic"), "0.00")
                Fg2.TextMatrix(Fg2.Rows - 1, 14) = Format(RstHisVenta("total"), "0.00")
            End If
            RstRutas.MoveNext
            If RstRutas.EOF = True Then
                Exit For
            End If
        Next A
    End If
    
    Fg2.Rows = Fg2.Rows + 2
    Fg2.TextMatrix(Fg2.Rows - 1, 1) = "Total ==>"
    
    Dim xTotal As Double
    For A = 2 To 14
        xTotal = 0
        For B = 1 To Fg2.Rows - 2
            xTotal = NulosN(Fg2.TextMatrix(B, A)) + xTotal
            
            If B = Fg2.Rows - 2 Then
                Exit For
            End If
        Next B
        Fg2.TextMatrix(Fg2.Rows - 1, A) = Format(xTotal, "0.00")
    Next A
    
    Fg2.Rows = Fg2.Rows + 1
    Fg2.TextMatrix(Fg2.Rows - 1, 1) = "Promedio"
    For A = 2 To 14
        Fg2.TextMatrix(Fg2.Rows - 1, A) = (Val(Fg2.TextMatrix(Fg2.Rows - 2, A)) / NumAños)
        Fg2.TextMatrix(Fg2.Rows - 1, A) = Format(Fg2.TextMatrix(Fg2.Rows - 1, A), "0.00")
    Next A
        
    Fg2.Rows = Fg2.Rows + 1
    Fg2.TextMatrix(Fg2.Rows - 1, 1) = "Varianza"
    
    With Fg2
        .Select 1, 1, Fg2.Rows - 1, 1
        .FillStyle = flexFillRepeat
        .CellBackColor = &HDDFFFF
        .Select Fg2.Rows - 3, 1, Fg2.Rows - 2, 15
        .FillStyle = flexFillRepeat
        .CellBackColor = &HEBD7BC
        .Select Fg2.Rows - 1, 2, Fg2.Rows - 1, 2
    End With
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : MostrarAños
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA LA INFORMACION DE VENTAS DE TODOS LOS AÑOS
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Año          |  String    |  ESPECIFICA EL AÑO DE TRABAJO
'*                    CodProducto  |  Integer   |  ESPECIFICA EL ID DEL PRODUCTO
'*                    RutaData     |  String    |  ESPECIFICA LA RUTA DE LA BASE DE DATOS
'* DEVUELVE         :
'*****************************************************************************************************
Function MostrarAños(Año As String, CodProducto As Integer, RutaData As String) As ADODB.Recordset
    Dim RstAño As New ADODB.Recordset
    Dim xCad As String
    
    Dim xFun As New eps_librerias.FuncionesData
    Dim xRutaData As String
    Dim xRst As New ADODB.Recordset
    Dim xCon2 As New ADODB.Connection
    
    xCad = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")
    
    xFun.F_BASEDATOS = xCad + RutaData
    xFun.F_GRUPOTRABAJO = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS") + "seven.mdw"
    xFun.F_PASSWORD = Eps_Pass
    xFun.F_USUARIO = Eps_User
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"
    
    Set xCon2 = xFun.AbrirConeccion
    
    RST_Busq RstAño, "TRANSFORM Sum(vta_ventasdet.canpro) AS SumaDecanpro SELECT vta_ventasdet.iditem, alm_inventario.descripcion, Sum(vta_ventasdet.canpro) AS total" _
        & " FROM vta_ventas INNER JOIN (vta_ventasdet INNER JOIN alm_inventario ON vta_ventasdet.iditem = alm_inventario.id) ON vta_ventas.id = vta_ventasdet.idvta " _
        & " Where (((vta_ventasdet.iditem) = " & CodProducto & ")) " _
        & " GROUP BY vta_ventasdet.iditem, alm_inventario.descripcion " _
        & " PIVOT Format([fchdoc],'mmm') In ('ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic')", xCon2

    Set MostrarAños = RstAño
End Function

Private Sub CmdMax_Click()
    ' INCREMENTA EN 10 PIXEL EL ANCHO DE LAS COLUMNAS DEL CONTROL Fg1, Fg2
    Fg1.ColWidth(1) = Fg1.ColWidth(1) + 10
    Fg1.ColWidth(3) = Fg1.ColWidth(3) + 10
    Fg1.ColWidth(4) = Fg1.ColWidth(4) + 10
    Fg1.ColWidth(5) = Fg1.ColWidth(5) + 10
    Fg1.ColWidth(6) = Fg1.ColWidth(6) + 10
    Fg1.ColWidth(7) = Fg1.ColWidth(7) + 10
    Fg1.ColWidth(8) = Fg1.ColWidth(8) + 10
    Fg1.ColWidth(9) = Fg1.ColWidth(9) + 10
    Fg1.ColWidth(10) = Fg1.ColWidth(10) + 10
    Fg1.ColWidth(11) = Fg1.ColWidth(11) + 10
    Fg1.ColWidth(12) = Fg1.ColWidth(12) + 10
    Fg1.ColWidth(13) = Fg1.ColWidth(13) + 10
    Fg1.ColWidth(14) = Fg1.ColWidth(14) + 10
    
    Fg2.ColWidth(1) = Fg2.ColWidth(1) + 10
    Fg2.ColWidth(2) = Fg2.ColWidth(2) + 10
    Fg2.ColWidth(3) = Fg2.ColWidth(3) + 10
    Fg2.ColWidth(4) = Fg2.ColWidth(4) + 10
    Fg2.ColWidth(5) = Fg2.ColWidth(5) + 10
    Fg2.ColWidth(6) = Fg2.ColWidth(6) + 10
    Fg2.ColWidth(7) = Fg2.ColWidth(7) + 10
    Fg2.ColWidth(8) = Fg2.ColWidth(8) + 10
    Fg2.ColWidth(9) = Fg2.ColWidth(9) + 10
    Fg2.ColWidth(10) = Fg2.ColWidth(10) + 10
    Fg2.ColWidth(11) = Fg2.ColWidth(11) + 10
    Fg2.ColWidth(12) = Fg2.ColWidth(12) + 10
End Sub

Private Sub CmdMin_Click()
    ' DECREMENTA EN 10 PIXEL EL ANCHO DE LAS COLUMNAS DEL CONTROL Fg1, Fg2
    Fg1.ColWidth(1) = Fg1.ColWidth(1) - 10
    Fg1.ColWidth(3) = Fg1.ColWidth(3) - 10
    Fg1.ColWidth(4) = Fg1.ColWidth(4) - 10
    Fg1.ColWidth(5) = Fg1.ColWidth(5) - 10
    Fg1.ColWidth(6) = Fg1.ColWidth(6) - 10
    Fg1.ColWidth(7) = Fg1.ColWidth(7) - 10
    Fg1.ColWidth(8) = Fg1.ColWidth(8) - 10
    Fg1.ColWidth(9) = Fg1.ColWidth(9) - 10
    Fg1.ColWidth(10) = Fg1.ColWidth(10) - 10
    Fg1.ColWidth(11) = Fg1.ColWidth(11) - 10
    Fg1.ColWidth(12) = Fg1.ColWidth(12) - 10
    Fg1.ColWidth(13) = Fg1.ColWidth(13) - 10
    Fg1.ColWidth(14) = Fg1.ColWidth(14) - 10
    
    
    Fg2.ColWidth(1) = Fg2.ColWidth(1) - 10
    Fg2.ColWidth(2) = Fg2.ColWidth(2) - 10
    Fg2.ColWidth(3) = Fg2.ColWidth(3) - 10
    Fg2.ColWidth(4) = Fg2.ColWidth(4) - 10
    Fg2.ColWidth(5) = Fg2.ColWidth(5) - 10
    Fg2.ColWidth(6) = Fg2.ColWidth(6) - 10
    Fg2.ColWidth(7) = Fg2.ColWidth(7) - 10
    Fg2.ColWidth(8) = Fg2.ColWidth(8) - 10
    Fg2.ColWidth(9) = Fg2.ColWidth(9) - 10
    Fg2.ColWidth(10) = Fg2.ColWidth(10) - 10
    Fg2.ColWidth(11) = Fg2.ColWidth(11) - 10
    Fg2.ColWidth(12) = Fg2.ColWidth(12) - 10
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    ' EJECUTA LA BUSQUEDA DE UN PRODUCTO
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(3, 4) As String
    
    If Col = 1 Then
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Producto":   xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5700":     xCampos(0, 3) = "C"
        xCampos(1, 0) = "Unidad":     xCampos(1, 1) = "abrev":         xCampos(1, 2) = "800":      xCampos(1, 3) = "C"
        xCampos(2, 0) = "Codigo":     xCampos(2, 1) = "codpro":        xCampos(2, 2) = "1700":     xCampos(2, 3) = "C"
        
        xform.SQLCad = "SELECT alm_inventario.descripcion, alm_inventario.codpro, mae_unidades.abrev, alm_inventario.idunimed, alm_inventario.id " _
            & " FROM mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed Where (((alm_inventario.tippro) = 3)) " _
            & " ORDER BY alm_inventario.descripcion"
        
        xform.Titulo = "Buscando Productos"
        xform.FormaBusca = CualquierParte
        xform.Criterio = ""
        xform.Ordenado = "descripcion"
        xform.CampoBusca = "descripcion"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        If xRs.State = 1 Then
            If BuscaItemGrid(xRs("codpro")) = False Then
                LblCodigo.Caption = xRs("codpro")
                LblDesc.Caption = xRs("descripcion")
                LblUniMed.Caption = Busca_Codigo(xRs("idunimed"), "id", "descripcion", "mae_unidades", "N", xCon)
                
                Fg1.TextMatrix(Fg1.Row, 0) = xRs("id")
                Fg1.TextMatrix(Fg1.Row, 1) = xRs("descripcion")
                Fg1.TextMatrix(Fg1.Row, 2) = xRs("codpro")
                Fg1.TextMatrix(Fg1.Row, 3) = ""
                'PREGUNTAMOS SI LA ULTIMA FILA ESTA VACIA PARA AGREGARLE OTRO ITEM
                If Fg1.TextMatrix(Fg1.Rows - 1, 1) <> "" Then
                    Fg1.Rows = Fg1.Rows + 1
                End If
            End If
        End If
        Set xform = Nothing
        Set xRs = Nothing
    End If
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace = 3 Then Exit Sub
        PopupMenu Menu2
    End If
End Sub

Private Sub Fg1_RowColChange()
    If Agregando = True Then Exit Sub
    If Fg1.Rows = 1 Then Exit Sub
    
    If Fg1.Row <> xFilaActual Then
        Fg2.Rows = 1
        xFilaActual = Fg1.Row
    End If
    LblDesc.Caption = Fg1.TextMatrix(Fg1.Row, 1)
    LblCodigo.Caption = Fg1.TextMatrix(Fg1.Row, 2)
    Dim xIdUniMed As Integer
    If NulosN(Fg1.TextMatrix(Fg1.Row, 0)) <> 0 Then
        xIdUniMed = Busca_Codigo(NulosN(Fg1.TextMatrix(Fg1.Row, 0)), "id", "idunimed", "alm_inventario", "N", xCon)
        LblUniMed.Caption = Busca_Codigo(xIdUniMed, "id", "descripcion", "mae_unidades", "N", xCon)
    End If
End Sub

Private Sub Fg2_CellChanged(ByVal Row As Long, ByVal Col As Long)
    CopiarValores
End Sub

Private Sub Fg2_EnterCell()
    If QueHace = 3 Then Fg1.Editable = flexEDNone: Exit Sub
    If Fg2.Row = Fg2.Rows - 1 Then
        If Fg2.Col = 1 Or Fg2.Col >= 14 Then
            Fg2.Editable = flexEDNone
            Exit Sub
        End If
        Fg2.Editable = flexEDKbdMouse
    Else
        Fg2.Editable = flexEDNone
    End If
End Sub

Private Sub Fg2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CopiarValores
End Sub

Private Sub Fg2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If QueHace = 3 Then Exit Sub
    If Button = 2 Then
        PopupMenu Menu1
    End If
End Sub

Private Sub Form_Activate()
'Modificado: 08/01/11 Johan Castro
'            Agregar linea de codigo para bloquear accesos de usuarios


    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, 193, Toolbar1, xCon
        '----------------------------------------------
        
        Dim Rpta As Integer
        RST_Busq RstPlanes, "SELECT ges_ventaproy.*, IIf([ges_ventaproy].[activo]=-1,'Activo','No Activo') AS estado " _
            & " FROM ges_ventaproy ORDER BY id DESC", xCon

        Set Dg1.DataSource = RstPlanes
        
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    TabOne1.CurrTab = 0
    QueHace = 3
    SeEjecuto = False
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Fg1.ColWidth(2) = 0
    If QueHace = 3 Then
        Fg1.SelectionMode = flexSelectionByRow
        Fg1.BackColorSel = &H80&
    End If
    
    Fg1.FrozenCols = 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 1
        Exit Sub
    End If
End Sub

Private Sub menu1_1_Click()
    CopiarValores
    HallarTotales
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : CopiarValores
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : COPIA LOS VALORES DE UNA CELDA DEL CONTRO FLEXGRID Fg1, Fg2
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub CopiarValores()
    If Fg2.Col = 2 Then
        If Val(Fg2.TextMatrix(Fg2.Rows - 1, 2)) <> 0 Then
            Fg1.TextMatrix(Fg1.Row, 3) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 2)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 2)) / 100) + 1)
        Else
            Fg1.TextMatrix(Fg1.Row, 3) = Fg2.TextMatrix(Fg2.Rows - 2, 2)
        End If
        Fg1.TextMatrix(Fg1.Row, 3) = Format(Fg1.TextMatrix(Fg1.Row, 3), "0.00")
    End If
    
    If Fg2.Col = 3 Then
        If Val(Fg2.TextMatrix(Fg2.Rows - 1, 3)) <> 0 Then
            Fg1.TextMatrix(Fg1.Row, 4) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 3)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 3)) / 100) + 1)
        Else
            Fg1.TextMatrix(Fg1.Row, 4) = Fg2.TextMatrix(Fg2.Rows - 2, 3)
        End If
        Fg1.TextMatrix(Fg1.Row, 4) = Format(Fg1.TextMatrix(Fg1.Row, 4), "0.00")
    End If
    
    If Fg2.Col = 4 Then
        If Val(Fg2.TextMatrix(Fg2.Rows - 1, 4)) <> 0 Then
            Fg1.TextMatrix(Fg1.Row, 5) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 4)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 4)) / 100) + 1)
        Else
            Fg1.TextMatrix(Fg1.Row, 5) = Fg2.TextMatrix(Fg2.Rows - 2, 4)
        End If
        Fg1.TextMatrix(Fg1.Row, 5) = Format(Fg1.TextMatrix(Fg1.Row, 5), "0.00")
    End If
    
    If Fg2.Col = 5 Then
        If Val(Fg2.TextMatrix(Fg2.Rows - 1, 5)) <> 0 Then
            Fg1.TextMatrix(Fg1.Row, 6) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 5)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 5)) / 100) + 1)
        Else
            Fg1.TextMatrix(Fg1.Row, 6) = Fg2.TextMatrix(Fg2.Rows - 2, 5)
        End If
        Fg1.TextMatrix(Fg1.Row, 6) = Format(Fg1.TextMatrix(Fg1.Row, 6), "0.00")
    End If
    
    If Fg2.Col = 6 Then
        If Val(Fg2.TextMatrix(Fg2.Rows - 1, 6)) <> 0 Then
            Fg1.TextMatrix(Fg1.Row, 7) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 6)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 6)) / 100) + 1)
        Else
            Fg1.TextMatrix(Fg1.Row, 7) = Fg2.TextMatrix(Fg2.Rows - 2, 6)
        End If
        Fg1.TextMatrix(Fg1.Row, 7) = Format(Fg1.TextMatrix(Fg1.Row, 7), "0.00")
    End If
    
    If Fg2.Col = 7 Then
        If Val(Fg2.TextMatrix(Fg2.Rows - 1, 7)) <> 0 Then
            Fg1.TextMatrix(Fg1.Row, 8) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 7)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 7)) / 100) + 1)
        Else
            Fg1.TextMatrix(Fg1.Row, 8) = Fg2.TextMatrix(Fg2.Rows - 2, 7)
        End If
        Fg1.TextMatrix(Fg1.Row, 8) = Format(Fg1.TextMatrix(Fg1.Row, 8), "0.00")
    End If
    
    If Fg2.Col = 8 Then
        If Val(Fg2.TextMatrix(Fg2.Rows - 1, 8)) <> 0 Then
            Fg1.TextMatrix(Fg1.Row, 9) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 8)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 8)) / 100) + 1)
        Else
            Fg1.TextMatrix(Fg1.Row, 9) = Fg2.TextMatrix(Fg2.Rows - 2, 8)
        End If
        Fg1.TextMatrix(Fg1.Row, 9) = Format(Fg1.TextMatrix(Fg1.Row, 9), "0.00")
    End If
    
    If Fg2.Col = 9 Then
        If Val(Fg2.TextMatrix(Fg2.Rows - 1, 9)) <> 0 Then
            Fg1.TextMatrix(Fg1.Row, 10) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 9)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 9)) / 100) + 1)
        Else
            Fg1.TextMatrix(Fg1.Row, 10) = Fg2.TextMatrix(Fg2.Rows - 2, 9)
        End If
        Fg1.TextMatrix(Fg1.Row, 10) = Format(Fg1.TextMatrix(Fg1.Row, 10), "0.00")
    End If
    
    If Fg2.Col = 10 Then
        If Val(Fg2.TextMatrix(Fg2.Rows - 1, 10)) <> 0 Then
            Fg1.TextMatrix(Fg1.Row, 11) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 10)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 10)) / 100) + 1)
        Else
            Fg1.TextMatrix(Fg1.Row, 11) = Fg2.TextMatrix(Fg2.Rows - 2, 10)
        End If
        Fg1.TextMatrix(Fg1.Row, 11) = Format(Fg1.TextMatrix(Fg1.Row, 11), "0.00")
    End If
    
    If Fg2.Col = 11 Then
        If Val(Fg2.TextMatrix(Fg2.Rows - 1, 11)) <> 0 Then
            Fg1.TextMatrix(Fg1.Row, 12) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 11)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 11)) / 100) + 1)
        Else
            Fg1.TextMatrix(Fg1.Row, 12) = Fg2.TextMatrix(Fg2.Rows - 2, 11)
        End If
        Fg1.TextMatrix(Fg1.Row, 12) = Format(Fg1.TextMatrix(Fg1.Row, 12), "0.00")
    End If
    
    If Fg2.Col = 12 Then
        If Val(Fg2.TextMatrix(Fg2.Rows - 1, 12)) <> 0 Then
            Fg1.TextMatrix(Fg1.Row, 13) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 12)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 12)) / 100) + 1)
        Else
            Fg1.TextMatrix(Fg1.Row, 13) = Fg2.TextMatrix(Fg2.Rows - 2, 12)
        End If
        Fg1.TextMatrix(Fg1.Row, 13) = Format(Fg1.TextMatrix(Fg1.Row, 13), "0.00")
    End If
    
    If Fg2.Col = 13 Then
        If Val(Fg2.TextMatrix(Fg2.Rows - 1, 13)) <> 0 Then
            Fg1.TextMatrix(Fg1.Row, 14) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 13)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 13)) / 100) + 1)
        Else
            Fg1.TextMatrix(Fg1.Row, 14) = Fg2.TextMatrix(Fg2.Rows - 2, 13)
        End If
        Fg1.TextMatrix(Fg1.Row, 14) = Format(Fg1.TextMatrix(Fg1.Row, 14), "0.00")
    End If
End Sub

Private Sub menu1_2_Click()
    'ENERO
    If Val(Fg2.TextMatrix(Fg2.Rows - 1, 2)) <> 0 Then
        Fg1.TextMatrix(Fg1.Row, 3) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 2)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 2)) / 100) + 1)
    Else
        Fg1.TextMatrix(Fg1.Row, 3) = Fg2.TextMatrix(Fg2.Rows - 2, 2)
    End If
    Fg1.TextMatrix(Fg1.Row, 3) = Format(Fg1.TextMatrix(Fg1.Row, 3), "0.00")
    
    'FEBRERO
    If Val(Fg2.TextMatrix(Fg2.Rows - 1, 3)) <> 0 Then
        Fg1.TextMatrix(Fg1.Row, 4) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 3)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 3)) / 100) + 1)
    Else
        Fg1.TextMatrix(Fg1.Row, 4) = Fg2.TextMatrix(Fg2.Rows - 2, 3)
    End If
    Fg1.TextMatrix(Fg1.Row, 4) = Format(Fg1.TextMatrix(Fg1.Row, 4), "0.00")
    
    'MARZO
    If Val(Fg2.TextMatrix(Fg2.Rows - 1, 4)) <> 0 Then
        Fg1.TextMatrix(Fg1.Row, 5) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 4)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 4)) / 100) + 1)
    Else
        Fg1.TextMatrix(Fg1.Row, 5) = Fg2.TextMatrix(Fg2.Rows - 2, 4)
    End If
    Fg1.TextMatrix(Fg1.Row, 5) = Format(Fg1.TextMatrix(Fg1.Row, 5), "0.00")
    
    'ABRIL
    If Val(Fg2.TextMatrix(Fg2.Rows - 1, 5)) <> 0 Then
        Fg1.TextMatrix(Fg1.Row, 6) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 5)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 5)) / 100) + 1)
    Else
        Fg1.TextMatrix(Fg1.Row, 6) = Fg2.TextMatrix(Fg2.Rows - 2, 5)
    End If
    Fg1.TextMatrix(Fg1.Row, 6) = Format(Fg1.TextMatrix(Fg1.Row, 6), "0.00")
    
    'MAYO
    If Val(Fg2.TextMatrix(Fg2.Rows - 1, 6)) <> 0 Then
        Fg1.TextMatrix(Fg1.Row, 7) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 6)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 6)) / 100) + 1)
    Else
        Fg1.TextMatrix(Fg1.Row, 7) = Fg2.TextMatrix(Fg2.Rows - 2, 6)
    End If
    Fg1.TextMatrix(Fg1.Row, 7) = Format(Fg1.TextMatrix(Fg1.Row, 7), "0.00")
    
    'JUNIO
    If Val(Fg2.TextMatrix(Fg2.Rows - 1, 7)) <> 0 Then
        Fg1.TextMatrix(Fg1.Row, 8) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 7)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 7)) / 100) + 1)
    Else
        Fg1.TextMatrix(Fg1.Row, 8) = Fg2.TextMatrix(Fg2.Rows - 2, 7)
    End If
    Fg1.TextMatrix(Fg1.Row, 8) = Format(Fg1.TextMatrix(Fg1.Row, 8), "0.00")
    
    'JULIO
    If Val(Fg2.TextMatrix(Fg2.Rows - 1, 8)) <> 0 Then
        Fg1.TextMatrix(Fg1.Row, 9) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 8)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 8)) / 100) + 1)
    Else
        Fg1.TextMatrix(Fg1.Row, 9) = Fg2.TextMatrix(Fg2.Rows - 2, 8)
    End If
    Fg1.TextMatrix(Fg1.Row, 9) = Format(Fg1.TextMatrix(Fg1.Row, 9), "0.00")
    
    'AGOSTO
    If Val(Fg2.TextMatrix(Fg2.Rows - 1, 9)) <> 0 Then
        Fg1.TextMatrix(Fg1.Row, 10) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 9)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 9)) / 100) + 1)
    Else
        Fg1.TextMatrix(Fg1.Row, 10) = Fg2.TextMatrix(Fg2.Rows - 2, 9)
    End If
    Fg1.TextMatrix(Fg1.Row, 10) = Format(Fg1.TextMatrix(Fg1.Row, 10), "0.00")
    
    'SETIEMBRE
    If Val(Fg2.TextMatrix(Fg2.Rows - 1, 10)) <> 0 Then
        Fg1.TextMatrix(Fg1.Row, 11) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 10)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 10)) / 100) + 1)
    Else
        Fg1.TextMatrix(Fg1.Row, 11) = Fg2.TextMatrix(Fg2.Rows - 2, 10)
    End If
    Fg1.TextMatrix(Fg1.Row, 11) = Format(Fg1.TextMatrix(Fg1.Row, 11), "0.00")
    
    'OCTUBRE
    If Val(Fg2.TextMatrix(Fg2.Rows - 1, 11)) <> 0 Then
        Fg1.TextMatrix(Fg1.Row, 12) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 11)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 11)) / 100) + 1)
    Else
        Fg1.TextMatrix(Fg1.Row, 12) = Fg2.TextMatrix(Fg2.Rows - 2, 11)
    End If
    Fg1.TextMatrix(Fg1.Row, 12) = Format(Fg1.TextMatrix(Fg1.Row, 12), "0.00")
    
    'NOVIEMBRE
    If Val(Fg2.TextMatrix(Fg2.Rows - 1, 12)) <> 0 Then
        Fg1.TextMatrix(Fg1.Row, 13) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 12)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 12)) / 100) + 1)
    Else
        Fg1.TextMatrix(Fg1.Row, 13) = Fg2.TextMatrix(Fg2.Rows - 2, 12)
    End If
    Fg1.TextMatrix(Fg1.Row, 13) = Format(Fg1.TextMatrix(Fg1.Row, 13), "0.00")
    
    'DICIEMBRE
    If Val(Fg2.TextMatrix(Fg2.Rows - 1, 13)) <> 0 Then
        Fg1.TextMatrix(Fg1.Row, 14) = Val(Fg2.TextMatrix(Fg2.Rows - 2, 13)) * ((Val(Fg2.TextMatrix(Fg2.Rows - 1, 13)) / 100) + 1)
    Else
        Fg1.TextMatrix(Fg1.Row, 14) = Fg2.TextMatrix(Fg2.Rows - 2, 13)
    End If
    Fg1.TextMatrix(Fg1.Row, 14) = Format(Fg1.TextMatrix(Fg1.Row, 14), "0.00")

    HallarTotales
End Sub

Private Sub Menu1_4_Click()
    Fg2.TextMatrix(Fg2.Rows - 1, 2) = ""
    Fg2.TextMatrix(Fg2.Rows - 1, 3) = ""
    Fg2.TextMatrix(Fg2.Rows - 1, 4) = ""
    Fg2.TextMatrix(Fg2.Rows - 1, 5) = ""
    Fg2.TextMatrix(Fg2.Rows - 1, 6) = ""
    Fg2.TextMatrix(Fg2.Rows - 1, 7) = ""
    Fg2.TextMatrix(Fg2.Rows - 1, 8) = ""
    Fg2.TextMatrix(Fg2.Rows - 1, 9) = ""
    Fg2.TextMatrix(Fg2.Rows - 1, 10) = ""
    Fg2.TextMatrix(Fg2.Rows - 1, 11) = ""
    Fg2.TextMatrix(Fg2.Rows - 1, 12) = ""
    Fg2.TextMatrix(Fg2.Rows - 1, 13) = ""
    
    Fg1.TextMatrix(Fg1.Row, 2) = Fg2.TextMatrix(Fg2.Rows - 2, 2)
    Fg1.TextMatrix(Fg1.Row, 3) = Fg2.TextMatrix(Fg2.Rows - 2, 3)
    Fg1.TextMatrix(Fg1.Row, 4) = Fg2.TextMatrix(Fg2.Rows - 2, 4)
    Fg1.TextMatrix(Fg1.Row, 5) = Fg2.TextMatrix(Fg2.Rows - 2, 5)
    Fg1.TextMatrix(Fg1.Row, 6) = Fg2.TextMatrix(Fg2.Rows - 2, 6)
    Fg1.TextMatrix(Fg1.Row, 7) = Fg2.TextMatrix(Fg2.Rows - 2, 7)
    Fg1.TextMatrix(Fg1.Row, 8) = Fg2.TextMatrix(Fg2.Rows - 2, 8)
    Fg1.TextMatrix(Fg1.Row, 9) = Fg2.TextMatrix(Fg2.Rows - 2, 9)
    Fg1.TextMatrix(Fg1.Row, 10) = Fg2.TextMatrix(Fg2.Rows - 2, 10)
    Fg1.TextMatrix(Fg1.Row, 11) = Fg2.TextMatrix(Fg2.Rows - 2, 11)
    Fg1.TextMatrix(Fg1.Row, 12) = Fg2.TextMatrix(Fg2.Rows - 2, 12)
    Fg1.TextMatrix(Fg1.Row, 13) = Fg2.TextMatrix(Fg2.Rows - 2, 13)
End Sub

Private Sub Menu1_5_Click()
    CompiarValoresAplicarPorcentaje
    HallarTotales
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : HallarTotales
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : HALLA LOS TOTALES DEL CONTROL Fg2
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub HallarTotales()
    Dim A, xCol As Integer
    Dim xTotal As Double
    
    xCol = 3
    For A = 1 To Fg1.Cols - 2
        xTotal = xTotal + NulosN(Fg1.TextMatrix(Fg1.Row, A))
        xCol = xCol + 1
    Next A
    Fg1.TextMatrix(Fg1.Row, Fg1.Cols - 1) = Format(xTotal, "0.00")
End Sub

Private Sub Menu1_6_Click()
    ExportarExcel
End Sub

Private Sub Menu2_1_Click()
    If Fg1.TextMatrix(Fg1.Rows - 1, 1) = "" Then Exit Sub
    
    Fg1.Rows = Fg1.Rows + 1
    Fg1.SetFocus
    Fg1.Select Fg1.Rows - 1, 1
    Fg1_CellButtonClick Fg1.Rows - 1, 1
End Sub

Private Sub Menu2_3_Click()
    ' ELIMINA UNA FILA DEL CONTROL Fg1
    Dim Rpta As Integer
    Rpta = MsgBox("¿Esta seguro de eliminar el producto seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        Fg1.RemoveItem (Fg1.Row)
    End If
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 1 Then Exit Sub
        MuestraSegundoTab
    End If
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : CambiarEstado
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA UN REGISTRO DE LA TABLA ges_ventaproy
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Activado     |  Boolean   |  INDICA SI SE ACTIVA O DESACTIVA EL REGISTRO
'* DEVUELVE         :
'*****************************************************************************************************
Sub CambiarEstado(Activado As Boolean)
    Dim Rpta As Integer
    TabOne1.CurrTab = 0
    If Activado = False Then
        Rpta = MsgBox("Esta seguro de desactivar la proyeccion de ventas seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    Else
        Rpta = MsgBox("Esta seguro de activar la proyeccion de ventas seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    End If
    
    If Rpta = vbYes Then
        If Activado = False Then
            xCon.Execute "UPDATE ges_ventaproy SET ges_ventaproy.activo = 0 Where (((ges_ventaproy.id) = " & RstPlanes("id") & "))"
            MsgBox "La proyeccion de ventas se desactivo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Else
            xCon.Execute "UPDATE ges_ventaproy SET ges_ventaproy.activo = -1 Where (((ges_ventaproy.id) = " & RstPlanes("id") & "))"
            MsgBox "La proyeccion de ventas se activo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
    End If
    RstPlanes.Requery
    Dg1.Refresh
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
            Cancelar
            RstPlanes.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then
        Cancelar
    End If
    
    If Button.Index = 14 Then
        Unload Me
    End If
End Sub

Function BuscaItemGrid(CodigoProducto As String) As Boolean
    Dim A As Integer
    
    If Fg1.Rows > 2 Then
        For A = 1 To Fg1.Rows
            If CodigoProducto = Fg1.TextMatrix(A, 1) Then
                MsgBox "El producto ya fue seleccionado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                BuscaItemGrid = True
                Exit Function
            End If
            If A = Fg1.Rows - 1 Then
                Exit For
            End If
        Next A
    End If
    BuscaItemGrid = False
End Function

'*****************************************************************************************************
'* Nombre Archivo   : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DETALLADA DEL REGISTRO EN LA PESTAÑA DETALLE DEL FORMULARIO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    
    Blanquea
    
    TxtDesc.Text = RstPlanes("descripcion")
    TxtFchIni.Valor = RstPlanes("fchini")
    TxtFchFin.Valor = RstPlanes("fchfin")
    
    Fg1.Rows = 1
    
     RST_Busq Rst, "SELECT ges_ventaproydet.*, alm_inventario.descripcion, alm_inventario.codpro FROM ges_ventaproydet LEFT JOIN alm_inventario " _
        & " ON ges_ventaproydet.idpro = alm_inventario.id Where (((ges_ventaproydet.id) = " & RstPlanes("id") & ")) ORDER BY alm_inventario.descripcion", xCon
    
    LblNumItem.Caption = Rst.RecordCount
    Dim xTotal As Double
    
    Agregando = True
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            xTotal = 0
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 0) = NulosC(Rst("idpro"))
            Fg1.TextMatrix(A, 1) = NulosC(Rst("descripcion"))
            Fg1.TextMatrix(A, 2) = NulosC(Rst("codpro"))
            Fg1.TextMatrix(A, 3) = Format(Rst("ene"), "0.00")
            xTotal = xTotal + NulosN(Rst("ene"))
            Fg1.TextMatrix(A, 4) = Format(Rst("feb"), "0.00")
            xTotal = xTotal + NulosN(Rst("feb"))
            Fg1.TextMatrix(A, 5) = Format(Rst("mar"), "0.00")
            xTotal = xTotal + NulosN(Rst("mar"))
            Fg1.TextMatrix(A, 6) = Format(Rst("abr"), "0.00")
            xTotal = xTotal + NulosN(Rst("abr"))
            Fg1.TextMatrix(A, 7) = Format(Rst("may"), "0.00")
            xTotal = xTotal + NulosN(Rst("may"))
            Fg1.TextMatrix(A, 8) = Format(Rst("jun"), "0.00")
            xTotal = xTotal + NulosN(Rst("jun"))
            Fg1.TextMatrix(A, 9) = Format(Rst("jul"), "0.00")
            xTotal = xTotal + NulosN(Rst("jul"))
            Fg1.TextMatrix(A, 10) = Format(Rst("ago"), "0.00")
            xTotal = xTotal + NulosN(Rst("ago"))
            Fg1.TextMatrix(A, 11) = Format(Rst("set"), "0.00")
            xTotal = xTotal + NulosN(Rst("set"))
            Fg1.TextMatrix(A, 12) = Format(Rst("oct"), "0.00")
            xTotal = xTotal + NulosN(Rst("oct"))
            Fg1.TextMatrix(A, 13) = Format(Rst("nov"), "0.00")
            xTotal = xTotal + NulosN(Rst("nov"))
            Fg1.TextMatrix(A, 14) = Format(Rst("dic"), "0.00")
            xTotal = xTotal + NulosN(Rst("dic"))
            Fg1.TextMatrix(A, 15) = Format(xTotal, "0.00")
            
            Rst.MoveNext
            
            If Rst.EOF = True Then
                Exit For
            End If
        Next
    End If
    Agregando = False
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

Private Sub TxtPorcentaje_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtPorcentaje.Text = Format(TxtPorcentaje.Text, "0.00")
    End If
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : CompiarValoresAplicarPorcentaje
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : COPIA LOS VALORES DE LA CELDA DE LOS CONTROLES Fg1 Y Fg2
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub CompiarValoresAplicarPorcentaje()
    If Fg2.TextMatrix(Fg2.Row, 1) <> "2004" And Fg2.TextMatrix(Fg2.Row, 1) <> "2005" And Fg2.TextMatrix(Fg2.Row, 1) <> "2006" And Fg2.TextMatrix(Fg2.Row, 1) <> "2007" Then
        MsgBox "No ha seleccionado una fila valida del historico de ventas, " & Chr(13) _
            & "seleccione una fila valida para aplicar esta opcion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    If NulosC(TxtPorcentaje.Text) = "" Then
        MsgBox "No ha especificado el porcentaje de aumento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtPorcentaje.SetFocus
        Exit Sub
    End If
    
'    If Fg2.Col = 2 Then
        Fg1.TextMatrix(Fg1.Row, 3) = Val(Fg2.TextMatrix(Fg2.Row, 2)) * ((Val(TxtPorcentaje.Text) / 100) + 1)
        Fg1.TextMatrix(Fg1.Row, 3) = Format(Fg1.TextMatrix(Fg1.Row, 3), "0.00")
'    End If
    
    'If Fg2.Col = 3 Then
        Fg1.TextMatrix(Fg1.Row, 4) = Val(Fg2.TextMatrix(Fg2.Row, 3)) * ((Val(TxtPorcentaje.Text) / 100) + 1)
        Fg1.TextMatrix(Fg1.Row, 4) = Format(Fg1.TextMatrix(Fg1.Row, 4), "0.00")
    'End If
    
    'If Fg2.Col = 4 Then
        Fg1.TextMatrix(Fg1.Row, 5) = Val(Fg2.TextMatrix(Fg2.Row, 4)) * ((Val(TxtPorcentaje.Text) / 100) + 1)
        Fg1.TextMatrix(Fg1.Row, 5) = Format(Fg1.TextMatrix(Fg1.Row, 5), "0.00")
    'End If
    
    'If Fg2.Col = 5 Then
        Fg1.TextMatrix(Fg1.Row, 6) = Val(Fg2.TextMatrix(Fg2.Row, 5)) * ((Val(TxtPorcentaje.Text) / 100) + 1)
        Fg1.TextMatrix(Fg1.Row, 6) = Format(Fg1.TextMatrix(Fg1.Row, 6), "0.00")
    'End If
    
    'If Fg2.Col = 6 Then
        Fg1.TextMatrix(Fg1.Row, 7) = Val(Fg2.TextMatrix(Fg2.Row, 6)) * ((Val(TxtPorcentaje.Text) / 100) + 1)
        Fg1.TextMatrix(Fg1.Row, 7) = Format(Fg1.TextMatrix(Fg1.Row, 7), "0.00")
    'End If
    
    'If Fg2.Col = 7 Then
        Fg1.TextMatrix(Fg1.Row, 8) = Val(Fg2.TextMatrix(Fg2.Row, 7)) * ((Val(TxtPorcentaje.Text) / 100) + 1)
        Fg1.TextMatrix(Fg1.Row, 8) = Format(Fg1.TextMatrix(Fg1.Row, 8), "0.00")
    'End If
    
    'If Fg2.Col = 8 Then
        Fg1.TextMatrix(Fg1.Row, 9) = Val(Fg2.TextMatrix(Fg2.Row, 8)) * ((Val(TxtPorcentaje.Text) / 100) + 1)
        Fg1.TextMatrix(Fg1.Row, 9) = Format(Fg1.TextMatrix(Fg1.Row, 9), "0.00")
    'End If
    
    'If Fg2.Col = 9 Then
        Fg1.TextMatrix(Fg1.Row, 10) = Val(Fg2.TextMatrix(Fg2.Row, 9)) * ((Val(TxtPorcentaje.Text) / 100) + 1)
        Fg1.TextMatrix(Fg1.Row, 10) = Format(Fg1.TextMatrix(Fg1.Row, 10), "0.00")
    'End If
    
    'If Fg2.Col = 10 Then
        Fg1.TextMatrix(Fg1.Row, 11) = Val(Fg2.TextMatrix(Fg2.Row, 10)) * ((Val(TxtPorcentaje.Text) / 100) + 1)
        Fg1.TextMatrix(Fg1.Row, 11) = Format(Fg1.TextMatrix(Fg1.Row, 11), "0.00")
    'End If
    
    'If Fg2.Col = 11 Then
        Fg1.TextMatrix(Fg1.Row, 12) = Val(Fg2.TextMatrix(Fg2.Row, 11)) * ((Val(TxtPorcentaje.Text) / 100) + 1)
        Fg1.TextMatrix(Fg1.Row, 12) = Format(Fg1.TextMatrix(Fg1.Row, 12), "0.00")
    'End If
    
    'If Fg2.Col = 12 Then
        Fg1.TextMatrix(Fg1.Row, 13) = Val(Fg2.TextMatrix(Fg2.Row, 12)) * ((Val(TxtPorcentaje.Text) / 100) + 1)
        Fg1.TextMatrix(Fg1.Row, 13) = Format(Fg1.TextMatrix(Fg1.Row, 13), "0.00")
    'End If
    
    'If Fg2.Col = 13 Then
        Fg1.TextMatrix(Fg1.Row, 14) = Val(Fg2.TextMatrix(Fg2.Row, 13)) * ((Val(TxtPorcentaje.Text) / 100) + 1)
        Fg1.TextMatrix(Fg1.Row, 14) = Format(Fg1.TextMatrix(Fg1.Row, 14), "0.00")
    'End If
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : ExportarExcel
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTA A MS EXCEL LOS DATOS DEL CONTROL Fg1
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub ExportarExcel()
    If Fg2.Rows = 1 Then
        MsgBox "No se ha registrado ventas para exportar", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
        Exit Sub
    End If
    
    Dim A As Integer
    Dim B As Integer
    Dim xFilas As Integer
    
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    
    objExcel.Visible = True
    'determina el numero de hojas que se mostrara en el Excel
    objExcel.SheetsInNewWorkbook = 1
    'abre el Libro
    objExcel.WindowState = 1
    objExcel.Workbooks.Open Trim(App.Path) + "\formatos\ComparacionVentas.xls"
    xFilas = 4
    
    With objExcel.ActiveSheet
        .Cells(2, 3) = Fg1.TextMatrix(Fg1.Row, 1)
        For A = 0 To Fg2.Rows - 1
            For B = 1 To Fg2.Cols - 1
                If A = 0 Then
                    .Cells(xFilas, B + 1) = "'" + Fg2.TextMatrix(A, B)
                Else
                    If B = 1 Then
                        .Cells(xFilas, B + 1) = "'" + Fg2.TextMatrix(A, B)
                    Else
                        .Cells(xFilas, B + 1) = Val(Fg2.TextMatrix(A, B))
                    End If
                End If
            Next B
            xFilas = xFilas + 1
        Next A
    End With
    
    MsgBox "El proceso de exportacion termino con exito", vbInformation + vbOKOnly + vbDefaultButton1, "Registro de Compras y Ventas"
    objExcel.WindowState = 1
    Set objExcel = Nothing
End Sub

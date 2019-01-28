VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCanjeLetras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caja y Bancos - Canje de Letras"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraInfLetra 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   4260
      Left            =   12030
      TabIndex        =   60
      Top             =   360
      Visible         =   0   'False
      Width           =   7395
      Begin VB.CommandButton CmdInfLet 
         Caption         =   "&Cancelar"
         Height          =   465
         Index           =   2
         Left            =   2760
         TabIndex        =   74
         Top             =   3690
         Width           =   1260
      End
      Begin VB.CommandButton CmdInfLet 
         Caption         =   "&Exportar"
         Height          =   465
         Index           =   1
         Left            =   1455
         TabIndex        =   73
         Top             =   3690
         Width           =   1260
      End
      Begin VB.CommandButton CmdInfLet 
         Caption         =   "&Imprimir"
         Height          =   465
         Index           =   0
         Left            =   150
         TabIndex        =   61
         Top             =   3690
         Width           =   1260
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg3 
         Height          =   2235
         Left            =   150
         TabIndex        =   63
         Top             =   1290
         Width           =   7155
         _cx             =   12621
         _cy             =   3942
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
         ForeColorSel    =   16777215
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmCanjeLetras.frx":0000
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
      Begin VB.Label lblInfLetra 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblInfLetra(8)"
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
         Height          =   270
         Index           =   8
         Left            =   2625
         TabIndex        =   72
         Top             =   975
         Width           =   1290
      End
      Begin VB.Label lblInfLetra 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblInfLetra(7)"
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
         Height          =   270
         Index           =   7
         Left            =   840
         TabIndex        =   71
         Top             =   960
         Width           =   1740
      End
      Begin VB.Label lblInfLetra 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblInfLetra(6)"
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
         Height          =   270
         Index           =   6
         Left            =   3360
         TabIndex        =   70
         Top             =   705
         Width           =   1335
      End
      Begin VB.Label lblInfLetra 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblInfLetra(5)"
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
         Height          =   270
         Index           =   5
         Left            =   840
         TabIndex        =   69
         Top             =   690
         Width           =   1335
      End
      Begin VB.Label lblInfLetra 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblInfLetra(4)"
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
         Height          =   270
         Index           =   4
         Left            =   840
         TabIndex        =   68
         Top             =   390
         Width           =   1740
      End
      Begin VB.Label lblInfLetra 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Importe"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   67
         Top             =   975
         Width           =   525
      End
      Begin VB.Label lblInfLetra 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fch.Ven."
         Height          =   195
         Index           =   2
         Left            =   2655
         TabIndex        =   66
         Top             =   765
         Width           =   645
      End
      Begin VB.Label lblInfLetra 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fch.Emi."
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   65
         Top             =   690
         Width           =   615
      End
      Begin VB.Label lblInfLetra 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Letra"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   64
         Top             =   420
         Width           =   585
      End
      Begin VB.Line ln 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   30
         X2              =   8500
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line ln 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   7425
         Y1              =   4245
         Y2              =   4245
      End
      Begin VB.Line ln 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   2
         X1              =   7380
         X2              =   7365
         Y1              =   -195
         Y2              =   5590
      End
      Begin VB.Line ln 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   3
         X1              =   0
         X2              =   15
         Y1              =   -1350
         Y2              =   4200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   150
         X2              =   7215
         Y1              =   3585
         Y2              =   3585
      End
      Begin VB.Label lbl_titulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Información Relacionada a la Letra:"
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
         Index           =   0
         Left            =   75
         TabIndex        =   62
         Top             =   60
         Width           =   3060
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   315
         Index           =   0
         Left            =   45
         Top             =   15
         Width           =   7305
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7215
      Left            =   0
      TabIndex        =   17
      Top             =   390
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
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6795
         Left            =   -12435
         TabIndex        =   24
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6435
            Left            =   15
            TabIndex        =   77
            Top             =   345
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11351
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Num.Reg."
            Columns(0).DataField=   "numreg"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Número Doc."
            Columns(1).DataField=   "numerodoc"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Emi."
            Columns(2).DataField=   "fchemi"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Cliente"
            Columns(3).DataField=   "nomcli"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Proveedor"
            Columns(4).DataField=   "nompro"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "M"
            Columns(5).DataField=   "monabrev"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Imp. Canjeado"
            Columns(6).DataField=   "impcan"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1720"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1640"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=513"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2434"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2355"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1773"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1693"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=5424"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=5345"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=512"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=5054"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=4974"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=847"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=767"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=2381"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2302"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=514"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.alignment=1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
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
            TabIndex        =   76
            Top             =   75
            Width           =   1980
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Canje de Letras"
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
            TabIndex        =   25
            Top             =   30
            Width           =   11610
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   6795
         Left            =   45
         TabIndex        =   18
         Top             =   375
         Width           =   11790
         Begin VB.CommandButton CmdIdLet 
            Height          =   240
            Left            =   2160
            Picture         =   "FrmCanjeLetras.frx":00F2
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   1560
            Width           =   240
         End
         Begin VB.TextBox TxtTotal3Pro 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   285
            Left            =   5010
            Locked          =   -1  'True
            TabIndex        =   50
            Text            =   "TxtTotal3Pro"
            Top             =   5670
            Width           =   1095
         End
         Begin VB.TextBox TxtTotal5Pro 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   285
            Left            =   10170
            Locked          =   -1  'True
            TabIndex        =   49
            Text            =   "TxtTotal5Pro"
            Top             =   5670
            Width           =   1200
         End
         Begin VB.Frame fra_letra 
            Caption         =   "[ Datos de la Emisión de la Letra ]"
            Enabled         =   0   'False
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
            Height          =   1290
            Left            =   75
            TabIndex        =   39
            Top             =   1905
            Width           =   11415
            Begin VB.TextBox TxtNumLet 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1470
               TabIndex        =   4
               Text            =   "TxtNumLet"
               Top             =   240
               Width           =   810
            End
            Begin VB.TextBox TxtNumDiaVen 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   4755
               TabIndex        =   5
               Text            =   "TxtNumDiaVen"
               Top             =   240
               Width           =   810
            End
            Begin VB.TextBox TxtVenDias 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   7905
               TabIndex        =   6
               Text            =   "TxtVenDias"
               Top             =   240
               Width           =   810
            End
            Begin VB.TextBox TxtGirado 
               Height          =   300
               Left            =   1470
               TabIndex        =   7
               Text            =   "TxtGirado"
               Top             =   555
               Width           =   4995
            End
            Begin VB.TextBox TxtNumDoc 
               Height          =   300
               Left            =   7905
               TabIndex        =   9
               Text            =   "TxtNumDoc"
               Top             =   870
               Width           =   1590
            End
            Begin VB.CommandButton CmdBusIdDoc 
               Height          =   240
               Left            =   2085
               Picture         =   "FrmCanjeLetras.frx":0224
               Style           =   1  'Graphical
               TabIndex        =   40
               Top             =   900
               Width           =   240
            End
            Begin VB.TextBox TxtIdDocIden 
               Height          =   300
               Left            =   1470
               TabIndex        =   8
               Text            =   "TxtIdDocIden"
               Top             =   870
               Width           =   900
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nº de Letras"
               Height          =   195
               Index           =   1
               Left            =   135
               TabIndex        =   48
               Top             =   330
               Width           =   885
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vence Cada"
               Height          =   195
               Index           =   2
               Left            =   3735
               TabIndex        =   47
               Top             =   330
               Width           =   885
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vence los Dias"
               Height          =   195
               Index           =   5
               Left            =   6690
               TabIndex        =   46
               Top             =   330
               Width           =   1065
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Girado A"
               Height          =   195
               Index           =   0
               Left            =   135
               TabIndex        =   45
               Top             =   615
               Width           =   615
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               Caption         =   "dias"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   6
               Left            =   5610
               TabIndex        =   44
               Top             =   330
               Width           =   285
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nº Documento"
               Height          =   195
               Index           =   7
               Left            =   6705
               TabIndex        =   43
               Top             =   945
               Width           =   1050
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Doc. Identidad"
               Height          =   195
               Index           =   8
               Left            =   135
               TabIndex        =   42
               Top             =   945
               Width           =   1050
            End
            Begin VB.Label LblDocIdentidad 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblDocIdentidad"
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
               Left            =   2385
               TabIndex        =   41
               Top             =   870
               Width           =   4080
            End
         End
         Begin VB.TextBox TxtTotal2Pro 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   285
            Left            =   3885
            Locked          =   -1  'True
            TabIndex        =   38
            Text            =   "TxtTotal2Pro"
            Top             =   5670
            Width           =   1095
         End
         Begin VB.CommandButton CmdBusProv 
            Height          =   240
            Left            =   2865
            Picture         =   "FrmCanjeLetras.frx":0356
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   870
            Width           =   240
         End
         Begin VB.Frame Frame3 
            Caption         =   "( Periodo )"
            Height          =   720
            Left            =   9480
            TabIndex        =   30
            Top             =   255
            Width           =   2010
            Begin VB.Label lblperiodo 
               Alignment       =   2  'Center
               Caption         =   "lblperiodo(1)"
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
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   31
               Top             =   330
               Width           =   1740
            End
         End
         Begin VB.CommandButton CmdBusMon 
            Height          =   240
            Left            =   5400
            Picture         =   "FrmCanjeLetras.frx":0488
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   1200
            Width           =   240
         End
         Begin VB.Frame fra_tipo 
            Enabled         =   0   'False
            Height          =   525
            Left            =   90
            TabIndex        =   26
            Top             =   255
            Width           =   5385
            Begin VB.OptionButton opt_tipo 
               Caption         =   "Recepción de Letras"
               Height          =   195
               Index           =   0
               Left            =   360
               TabIndex        =   28
               Top             =   255
               Value           =   -1  'True
               Width           =   1950
            End
            Begin VB.OptionButton opt_tipo 
               Caption         =   "Emisión de Letras"
               Height          =   195
               Index           =   1
               Left            =   2475
               TabIndex        =   27
               Top             =   255
               Width           =   1680
            End
         End
         Begin VB.TextBox TxtIdMon 
            Height          =   300
            Left            =   4845
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   2
            Text            =   "TxtIdMon"
            Top             =   1170
            Width           =   825
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmi 
            Height          =   300
            Left            =   1530
            TabIndex        =   1
            Top             =   1170
            Width           =   1260
            _ExtentX        =   2223
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
            Valor           =   "18/07/2008"
         End
         Begin VB.TextBox TxtRucPro 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   0
            Text            =   "TxtRucPro"
            Top             =   840
            Width           =   1620
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   2010
            Left            =   60
            TabIndex        =   13
            Top             =   3600
            Width           =   6270
            _cx             =   11060
            _cy             =   3545
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
            ForeColorSel    =   16777215
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
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   9
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmCanjeLetras.frx":05BA
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
            Height          =   2010
            Left            =   6480
            TabIndex        =   16
            Top             =   3600
            Width           =   5205
            _cx             =   9181
            _cy             =   3545
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
            ForeColorSel    =   16777215
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
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmCanjeLetras.frx":06C7
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
         Begin VB.TextBox TxtIdLet 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   3
            Text            =   "TxtIdLet"
            Top             =   1530
            Width           =   900
         End
         Begin VB.Frame Frame8 
            Height          =   705
            Left            =   60
            TabIndex        =   51
            Top             =   5925
            Width           =   6270
            Begin VB.CommandButton CmdDoc 
               Caption         =   "Agregar Documentos"
               Enabled         =   0   'False
               Height          =   400
               Index           =   0
               Left            =   45
               TabIndex        =   10
               Top             =   195
               Width           =   2025
            End
            Begin VB.CommandButton CmdDoc 
               Caption         =   "Eliminar Documento"
               Enabled         =   0   'False
               Height          =   400
               Index           =   2
               Left            =   4140
               TabIndex        =   12
               Top             =   195
               Width           =   2025
            End
            Begin VB.CommandButton CmdDoc 
               Caption         =   "Seleccionar Documentos"
               Enabled         =   0   'False
               Height          =   400
               Index           =   1
               Left            =   2070
               TabIndex        =   11
               Top             =   195
               Width           =   2025
            End
         End
         Begin VB.Frame Frame9 
            Height          =   705
            Left            =   6465
            TabIndex        =   52
            Top             =   5925
            Width           =   5205
            Begin VB.CommandButton CmdLetra 
               Caption         =   "Información Letra"
               Enabled         =   0   'False
               Height          =   400
               Index           =   2
               Left            =   3660
               TabIndex        =   75
               Top             =   210
               Width           =   1410
            End
            Begin VB.CommandButton CmdLetra 
               Caption         =   "Agregar Letra"
               Enabled         =   0   'False
               Height          =   400
               Index           =   0
               Left            =   60
               TabIndex        =   14
               Top             =   210
               Width           =   1410
            End
            Begin VB.CommandButton CmdLetra 
               Caption         =   "Eliminar Letra"
               Enabled         =   0   'False
               Height          =   400
               Index           =   1
               Left            =   1515
               TabIndex        =   15
               Top             =   210
               Width           =   1410
            End
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Letra"
            Height          =   195
            Index           =   9
            Left            =   135
            TabIndex        =   59
            Top             =   1560
            Width           =   360
         End
         Begin VB.Label LblLetra 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblLetra"
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
            TabIndex        =   58
            Top             =   1530
            Width           =   4080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Documentos del Proveedor"
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
            Height          =   195
            Left            =   60
            TabIndex        =   56
            Top             =   3360
            Width           =   2310
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total ==>"
            Height          =   195
            Left            =   9300
            TabIndex        =   55
            Top             =   5760
            Width           =   675
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total ==>"
            Height          =   195
            Left            =   3045
            TabIndex        =   54
            Top             =   5760
            Width           =   675
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Letras del Proveedor"
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
            Height          =   195
            Left            =   6465
            TabIndex        =   53
            Top             =   3360
            Width           =   1785
         End
         Begin VB.Label LblTipoCambio 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipoCambio"
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
            Left            =   9225
            TabIndex        =   37
            Top             =   1170
            Width           =   1080
         End
         Begin VB.Label LblTipCam2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Cambio"
            Height          =   195
            Left            =   7965
            TabIndex        =   36
            Top             =   1275
            Width           =   1110
         End
         Begin VB.Label LblIdProveedor 
            AutoSize        =   -1  'True
            Caption         =   "LblIdProveedor"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   5730
            TabIndex        =   35
            Top             =   555
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label LblTitulo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   135
            TabIndex        =   34
            Top             =   915
            Width           =   735
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
            Left            =   3180
            TabIndex        =   33
            Top             =   840
            Width           =   4680
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Emisión"
            Height          =   195
            Index           =   3
            Left            =   135
            TabIndex        =   23
            Top             =   1275
            Width           =   1260
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Canje de Letras"
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
            Left            =   60
            TabIndex        =   22
            Top             =   30
            Width           =   11610
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
            Left            =   5685
            TabIndex        =   21
            Top             =   1170
            Width           =   2175
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda"
            Height          =   195
            Index           =   4
            Left            =   3855
            TabIndex        =   20
            Top             =   1275
            Width           =   585
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
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
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Listado"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Registro"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7410
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
               Picture         =   "FrmCanjeLetras.frx":0783
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCanjeLetras.frx":0CC7
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCanjeLetras.frx":1059
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCanjeLetras.frx":11DD
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCanjeLetras.frx":1631
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCanjeLetras.frx":1749
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCanjeLetras.frx":1C8D
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCanjeLetras.frx":21D1
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCanjeLetras.frx":22E5
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCanjeLetras.frx":23F9
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCanjeLetras.frx":284D
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCanjeLetras.frx":29B9
               Key             =   "IMG11"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Menu1_1 
         Caption         =   "Agregar Documentos"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "Seleccionar Documentos"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_4 
         Caption         =   "Eliminar Documento"
      End
   End
   Begin VB.Menu Menu2 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu Menu2_1 
         Caption         =   "Agregar Letra"
      End
      Begin VB.Menu Menu2_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu2_3 
         Caption         =   "Eliminar Letra"
      End
   End
   Begin VB.Menu Menu3 
      Caption         =   "Menu3"
      Visible         =   0   'False
      Begin VB.Menu Menu3_1 
         Caption         =   "Información de Letra"
      End
   End
End
Attribute VB_Name = "FrmCanjeLetras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMCANJELETRAS.FRM
'* Tipo             : FORMULARIO
'* Descripcion      :
'*                  :
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 12/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim RstFrm As New ADODB.Recordset
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim Agregando As Boolean
Dim xCuenLetra As Integer


Private Sub pTotalizarDocumento()
    TxtTotal2Pro.Text = Format(GRID_SUMAR_COL(Fg1, 5), FORMAT_MONTO)
    TxtTotal3Pro.Text = Format(GRID_SUMAR_COL(Fg1, 6), FORMAT_MONTO)
End Sub


Private Sub CmdBusIdDoc_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Nombre":      xCampos(0, 1) = "descripcion":     xCampos(0, 2) = "4500":      xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":   xCampos(1, 1) = "id":              xCampos(1, 2) = "1400":      xCampos(1, 3) = "C"

    xForm.SQLCad = "SELECT * FROM mae_dociden ORDER BY descripcion"
    xForm.Titulo = "Buscando Documento de Identidad"

    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "descripcion"
    xForm.CampoBusca = "descripcion"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdDocIden.Text = xRs("id") & ""
        LblDocIdentidad.Caption = xRs("descripcion") & ""
        TxtNumDoc.SetFocus
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
    xCampos(0, 0) = "Nombre":      xCampos(0, 1) = "descripcion":     xCampos(0, 2) = "4500":      xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id":   xCampos(1, 1) = "id":              xCampos(1, 2) = "500":      xCampos(1, 3) = "N"

    xForm.SQLCad = "SELECT * FROM mae_moneda ORDER BY descripcion"
    xForm.Titulo = "Buscando Moneda"

    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "descripcion"
    xForm.CampoBusca = "descripcion"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdMon.Text = xRs("id") & ""
        LblMoneda.Caption = xRs("descripcion") & ""
        TxtNumLet.SetFocus
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdDoc_Click(Index As Integer)
    Select Case Index
        Case 0 '--agregar
            pRegistroAdd False
        Case 1 '--selecc
            pRegistroAdd True
        Case 2 '--eliminar
                If Fg1.Rows = 1 Then
                MsgBox "No hay documentos a eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
            
            Fg1.RemoveItem Fg1.Row
            pTotalizarDocumento
        End Select
End Sub

Private Sub CmdInfLet_Click(Index As Integer)
    Select Case Index
        Case 0 '--imprimir
            pLetraInfImprimir
        Case 1 '--exportar
            pLetraInfExportar
        Case 2 '--cancelar
            Me.fraInfLetra.Visible = False
            Me.Toolbar1.Enabled = True
            Me.TabOne1.Enabled = True
            Fg2.Col = 1
            Fg2.SetFocus
    End Select
End Sub

Private Sub CmdLetra_Click(Index As Integer)
    Select Case Index
        Case 0 '--agregar
            pAddLetra
        Case 1 '--eliminar
            pDelLetra
        Case 2 '--documentos asociados a una letra especifica
            pMostrarInfLetra
    End Select
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    RstFrm.Sort = CStr(Dg1.Columns(ColIndex).DataField)
    Err.Clear
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Row < 1 Then Exit Sub
    If Col <> 2 Then Exit Sub
    If NulosC(TxtRucPro.Text) = "" Then
        MsgBox "Seleccione el " + LblTitulo.Caption, vbExclamation, xTitulo
        CmdBusProv.SetFocus
        Exit Sub
    End If
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "Seleccione la Moneda", vbExclamation, xTitulo
        TxtIdMon.SetFocus
        Exit Sub
    End If
    Dim xCampos(6, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim nSQLId As String '--almacenara los codigos de documentos ya seleccionados
    Dim nSQL As String
    Dim nTitulo As String

    xCampos(0, 0) = "Tip.Doc.":        xCampos(0, 1) = "abrev":     xCampos(0, 2) = "500":     xCampos(0, 3) = "C":     xCampos(0, 4) = "N"
    xCampos(1, 0) = "Nº Documento":    xCampos(1, 1) = "numdoc":    xCampos(1, 2) = "2000":    xCampos(1, 3) = "C":     xCampos(1, 4) = "S"
    xCampos(2, 0) = "Fch.Emi.":        xCampos(2, 1) = "fchdoc":    xCampos(2, 2) = "1200":    xCampos(2, 3) = "C":     xCampos(2, 4) = "N"
    xCampos(3, 0) = "M":               xCampos(3, 1) = "simbolo":   xCampos(3, 2) = "500":     xCampos(3, 3) = "C":     xCampos(4, 4) = "N"
    xCampos(4, 0) = "Importe":         xCampos(4, 1) = "imptotdoc": xCampos(4, 2) = "1000":    xCampos(4, 3) = "N":     xCampos(5, 4) = "N"
    xCampos(5, 0) = "Saldo":           xCampos(5, 1) = "impsal":    xCampos(5, 2) = "1000":    xCampos(5, 3) = "N":     xCampos(6, 4) = "N"
        
        
    '*************************************************************
    nSQLId = GRID_GENERAR_SQL_ID(Fg1, 7, "com_compras.id", " NOT IN ")
    '*************************************************************
    
    nSQL = "SELECT mae_documento.abrev, com_compras!numser+'-'+com_compras!numdoc AS numdoc, com_compras.impsal, format(com_compras.fchdoc,'dd/mm/yy') as fchdoc ,com_compras.fchven, com_compras.idpro, com_compras.imptot AS imptotdoc, com_compras.id, con_diario.idcue, con_planctas.cuenta, con_diario.idlib, mae_moneda.simbolo " _
        + vbCr + " FROM mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (con_planctas RIGHT JOIN (com_compras LEFT JOIN con_diario ON com_compras.id = con_diario.idmov) ON con_planctas.id = con_diario.idcue) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon " _
        + vbCr + " WHERE (((com_compras.impsal)>0) AND ((com_compras.idpro)=" & NulosN(LblIdProveedor.Caption) & ") AND ((con_planctas.cuenta) Like '42%') AND ((con_diario.idlib)=1)) AND com_compras.idmon= " & NulosN(TxtIdMon.Text) & " " _
        + vbCr + IIf(nSQLId = "", "", " AND " + nSQLId) _
        + vbCr + " ORDER BY com_compras.numser+'-'+com_compras.numdoc"
       
    'nSQL = nSQL + IIf(nSQLId = "", "", " AND " + nSQLId)
    
    If opt_tipo(0).Value = True Then '---compras
            
    Else '--ventas
        nSQL = Replace(nSQL, "42%", "12%")
        nSQL = Replace(nSQL, "com_compras.idpro", "vta_ventas.idcli")
        nSQL = Replace(nSQL, "com_compras.imptot", "vta_ventas.imptotdoc")
        nSQL = Replace(nSQL, "(con_diario.idlib)=1", "(con_diario.idlib)=2")
        nSQL = Replace(nSQL, "com_compras.id", "vta_ventas.id")
        nSQL = Replace(nSQL, "com_compras", "vta_ventas")

    End If
    
    nTitulo = "Buscando Documentos del " + LblTitulo.Caption + ": " + StrConv(LblProveedor.Caption, 3)
    '*************************************************************

    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "numdoc", "numdoc", CualquierParte
    If xRs.State = 0 Then GoTo Salir
    If xRs.RecordCount = 0 Then GoTo Salir
    Agregando = True
    Do While Not xRs.EOF
        Fg1.TextMatrix(Row, 1) = NulosC(xRs("abrev"))
        Fg1.TextMatrix(Row, 2) = NulosC(xRs("numdoc"))
        Fg1.TextMatrix(Row, 3) = NulosC(xRs("fchdoc"))
        Fg1.TextMatrix(Row, 4) = NulosC(xRs("simbolo"))
        Fg1.TextMatrix(Row, 5) = Format(NulosN(xRs("imptotdoc")), FORMAT_MONTO)
        Fg1.TextMatrix(Row, 6) = Format(NulosN(xRs("impsal")), FORMAT_MONTO)
        Fg1.TextMatrix(Row, 7) = NulosN(xRs("id"))
        Fg1.TextMatrix(Row, 8) = NulosN(xRs("idcue"))
        xRs.MoveNext
    Loop
    Agregando = False
    pTotalizarDocumento
    Fg1.Row = Row: Fg1.Col = 2:  Fg1.SetFocus
Salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    pTotalizarDocumento
    Agregando = False
    SHOW_ERROR Me.Name, "pRegistroAdd"
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone: Exit Sub
    End If
    If Fg1.Col = 2 Then
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub
Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = 45 Then
        CmdDoc_Click 0
    End If
    If KeyCode = 46 Then
        CmdDoc_Click 2
    End If
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace = 1 Then PopupMenu Menu1
    End If
End Sub

Private Sub Fg2_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    Select Case Col
        Case 1
            Fg2.TextMatrix(Row, Col) = Format(Fg2.TextMatrix(Row, Col), "00000000")
        Case 2, 3
            Fg2.TextMatrix(Row, Col) = Format(Fg2.TextMatrix(Row, Col), "dd/mm/yyyy")
        Case 4
            Fg2.TextMatrix(Row, Col) = Format(Fg2.TextMatrix(Row, Col), FORMAT_MONTO)
            pTotalizarLetra
    End Select
End Sub

Private Sub pTotalizarLetra()
    TxtTotal5Pro.Text = Format(GRID_SUMAR_COL(Fg2, 4), FORMAT_MONTO)
End Sub

Private Sub Fg2_EnterCell()

    If QueHace = 3 Then
        Fg2.Editable = flexEDNone: Exit Sub
    End If

    If Fg2.Col >= 1 And Fg2.Col <= 4 Then
        Fg2.Editable = flexEDKbdMouse
    Else
        Fg2.Editable = flexEDNone
    End If
End Sub

Private Sub Fg2_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If opt_tipo(1).Value = True Then Exit Sub '--SI ES CLIENTE NO HACER NADA
    If KeyCode = 45 Then
        CmdLetra_Click 0
    End If
    If KeyCode = 46 Then
        CmdLetra_Click 1
    End If
End Sub


Private Sub Fg2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace = 1 Then
            If opt_tipo(0).Value = True Then PopupMenu menu2
        ElseIf QueHace = 3 Then
            If Fg2.Rows >= 2 Then PopupMenu Menu3
        End If
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
   
    SeEjecuto = False
    
    pCargarGrid
    
    SeEjecuto = True
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.RecordCount = 0 Then
        If MsgBox("No se ha registrado Canje de Letra, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
            Nuevo
        End If
    End If
    
    
End Sub


Private Sub MuestraSegundoTab()
    Dim RstTmp As New ADODB.Recordset
    
    Fg1.Cols = 9
    
    Blanquea
    CmdLetra(2).Enabled = True
    With RstFrm
        If .State = 0 Then Exit Sub
        If .EOF = True Or .BOF = True Or .RecordCount = 0 Then Exit Sub
    
        
        If .Fields("tiplet") = "1" Then opt_tipo(0).Value = True
        If .Fields("tiplet") = "2" Then opt_tipo(1).Value = True
        
        TxtFchEmi.Valor = .Fields("fchemi") & ""
        TxtFchEmi_Validate True
        TxtIdMon.Text = .Fields("idmon") & ""
        LblMoneda.Caption = .Fields("mondesc") & ""
        
        TxtRucPro.Text = .Fields("numruc") & ""
        LblIdProveedor.Caption = .Fields("idclipro") & ""
        LblProveedor.Caption = .Fields("nombre") & ""
        TxtNumLet.Text = .Fields("numlet") & ""
        If NulosN(.Fields("vendia")) <> 0 Then
            TxtNumDiaVen.Text = NulosN(.Fields("vendia"))
        Else
            TxtVenDias.Text = NulosN(.Fields("venlosdia"))
        End If
        TxtGirado.Text = .Fields("girado") & ""
        TxtIdDocIden.Text = .Fields("iddocgir") & ""
        LblDocIdentidad = .Fields("idendesc") & ""
        TxtNumDoc.Text = .Fields("numdocgir") & ""
        '------
        TxtIdLet.Text = .Fields("idlet") & ""
        LblLetra.Caption = .Fields("letdesc") & ""
        If opt_tipo(0).Value = True Then    '--PROVEEDOR
            xCuenLetra = NulosN(.Fields("idcuencom"))
        Else                                '--CLIENTE
            xCuenLetra = NulosN(.Fields("idcuenven"))
        End If
        '------
        'BUSCAMOS LOS DOCUMENTOS ASIGNADOS A LA LETRA
        Dim nSQL As String
        If opt_tipo(0).Value = True Then
            nSQL = "SELECT con_letradoc.iddoc, mae_documento.abrev, com_compras!numser+'-'+com_compras!numdoc AS numdoc, com_compras.fchdoc, mae_moneda.simbolo, com_compras.imptot AS imptotdoc, con_letradoc.impcan, con_diario.idcue " _
                + vbCr + " FROM (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (com_compras INNER JOIN con_letradoc ON com_compras.id = con_letradoc.iddoc) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) LEFT JOIN (con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) ON com_compras.id = con_diario.idmov " _
                + vbCr + " WHERE con_letradoc.idlet =" & .Fields("id") & " AND con_diario.idlib=1 AND con_planctas.cuenta Like '42%';"
        Else
            nSQL = "SELECT con_letradoc.iddoc, mae_documento.abrev, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, vta_ventas.fchdoc, mae_moneda.simbolo, vta_ventas.imptotdoc, con_letradoc.impcan, con_diario.idcue " _
                + vbCr + " FROM (((mae_documento RIGHT JOIN (vta_ventas INNER JOIN con_letradoc ON (con_letradoc.iddoc = vta_ventas.id) AND (vta_ventas.id = con_letradoc.iddoc)) ON mae_documento.id = vta_ventas.tipdoc) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN con_diario ON vta_ventas.id = con_diario.idmov) LEFT JOIN con_planctas ON con_diario.idcue = con_planctas.id " _
                + vbCr + " WHERE con_letradoc.idlet =" & .Fields("id") & " AND con_diario.idlib=2 AND con_planctas.cuenta Like '12%';"
        End If
        
        RST_Busq RstTmp, nSQL, xCon
        Fg1.Rows = 1
        If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(RstTmp("abrev"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstTmp("numdoc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(RstTmp("fchdoc"), "dd/mm/yy")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(RstTmp("simbolo"))
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(NulosN(RstTmp("imptotdoc")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(NulosN(RstTmp("impcan")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosN(RstTmp("iddoc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosN(RstTmp("idcue"))
            
            RstTmp.MoveNext
        Loop
        
        'BUSCAMOS LAS LETRAS ASIGNADAS
        Set RstTmp = Nothing
        
        RST_Busq RstTmp, "SELECT con_letradet.* From con_letradet WHERE (((con_letradet.idlet)=" & .Fields("id") & "))", xCon
    
        Fg2.Rows = 1
        If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosC(RstTmp("numlet"))
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = Format(RstTmp("fchemi"), "dd/mm/yyyy")
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(RstTmp("fchven"), "dd/mm/yyyy")
            Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(RstTmp("implet"), FORMAT_MONTO)
            Fg2.TextMatrix(Fg2.Rows - 1, 5) = NulosC(RstTmp("corr"))
            RstTmp.MoveNext
        Loop
        
        Set RstTmp = Nothing
        
    End With
    pTotalizarDocumento
    pTotalizarLetra
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If fraInfLetra.Visible = True Then CmdInfLet_Click 2
    End If
End Sub

Private Sub Form_Load()
    Dg1.Columns("fchemi").NumberFormat = FORMAT_DATE:
    Dg1.Columns("fchven").NumberFormat = FORMAT_DATE:
    Dg1.Columns("letraimp").NumberFormat = FORMAT_MONTO:

    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    
    
    Fg1.ColWidth(7) = 0
    Fg1.ColWidth(8) = 0
    
    Fg2.ColWidth(5) = 0
    
    Fg1.SelectionMode = flexSelectionByRow
    
    Fg2.ColEditMask(2) = "##/##/####"
    Fg2.ColEditMask(3) = "##/##/####"
    QueHace = 3
    SeEjecuto = False
    TabOne1.CurrTab = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation, xTitulo
        Cancel = 1
        Exit Sub
    End If
    Set RstFrm = Nothing
    Set Dg1.DataSource = Nothing
End Sub

Private Sub Menu1_1_Click()
    '--AGREGAR DOCUMENTOS
    CmdDoc_Click 0
End Sub

Private Sub Menu1_2_Click()
    '--SELECCIONAR DOCUMENTOS
    CmdDoc_Click 1
End Sub

Private Sub menu1_4_Click()
    '--ELIMINAR DOCUMENTOS
    CmdDoc_Click 2
End Sub

Private Sub Menu2_1_Click()
    '--AGREGAR LETRA
    CmdLetra_Click 0
End Sub

Private Sub Menu2_3_Click()
    '--ELIMINAR LETRA
    CmdLetra_Click 1
End Sub

Private Sub Menu3_1_Click()
    '--INFORMACION DE LETRA
    CmdLetra_Click 2
End Sub

Private Sub opt_tipo_Click(Index As Integer)
    If Index = 0 Then '--proveedor
        fra_letra.Enabled = False
        LblTitulo.Caption = "Proveedor"
        Label6.Caption = "Documentos del Proveedor"
        Label9.Caption = "Letras del Proveedor"
        TxtRucPro.Text = ""
        '*********
        CmdLetra(0).Caption = "Agregar Letra"
        CmdLetra(1).Visible = True
        '---
        TxtNumLet.Text = ""
        TxtNumDiaVen.Text = ""
        TxtVenDias.Text = ""
        TxtGirado.Text = ""
        TxtIdDocIden.Text = ""
    Else '--cliente
        fra_letra.Enabled = True
        LblTitulo.Caption = "Cliente"
        Label6.Caption = "Documentos del Cliente"
        Label9.Caption = "Letras del Cliente"
        TxtRucPro.Text = ""
        '************
        CmdLetra(0).Caption = "Generar Letra"
        CmdLetra(1).Visible = False
    End If
End Sub

Private Sub opt_tipo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then TxtRucPro.SetFocus
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
            RstFrm.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
       
    If Button.Index = 8 Then Filtrar
   
   If Button.Index = 9 Then
        If RstFrm.State = 0 Then Exit Sub
        RstFrm.Filter = adFilterNone
        RstFrm.Requery
        Dg1.Refresh
    End If
    
    If Button.Index = 10 Then Buscar
    If Button.Index = 11 Then CambiarMes
    If Button.Index = 15 Then
        Set RstFrm = Nothing
        Unload Me
    End If

End Sub

Sub Cancelar()
    QueHace = 3
    ActivaTool
    Bloquea
    Label5.Caption = "Detalle de Canje de Letras"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    
    Fg1.SelectionMode = flexSelectionByRow
    Fg2.SelectionMode = flexSelectionByRow
    
    Dg1.SetFocus
End Sub

Function Grabar() As Boolean
    
    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "Grabar", "Modificar") + " el canje de la letra", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstDoc As New ADODB.Recordset
    Dim RstDia As New ADODB.Recordset
    Dim xNumAsiento As String
    Dim xId, A As Integer
    On Error GoTo LaCague
    Me.MousePointer = vbHourglass
    xCon.BeginTrans
    If QueHace = 1 Then
        xId = HallaCodigoTabla("con_letra", xCon, "id")
        xNumAsiento = NuevoNumAsiento(37, xMes, xCon)
        
        RST_Busq RstCab, "SELECT TOP 1 * FROM con_letra", xCon
        RST_Busq RstDet, "SELECT TOP 1 * FROM con_letradet", xCon
        RST_Busq RstDoc, "SELECT TOP 1 * FROM con_letradoc", xCon
        RST_Busq RstDia, "SELECT TOP 1 * FROM con_diario", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else
        RST_Busq RstCab, "SELECT * FROM con_letra WHERE id = " & RstFrm("id") & ";", xCon
        '--AGREGANDO EL SALDO AL DOCUMENTO DEL PROEVEEDOR O CLIENTE
        Dim RstTmp As New ADODB.Recordset
        RST_Busq RstTmp, "SELECT iddoc, impcan From con_letradoc WHERE (((idlet)=" & RstFrm("id") & "));", xCon
        If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            If RstFrm.Fields("tiplet") = 1 Then
                xCon.Execute "UPDATE com_compras SET impsal = impsal + " & NulosN(RstTmp.Fields("impcan")) & " WHERE id = " & RstTmp.Fields("iddoc") & ";"
            Else
                xCon.Execute "UPDATE vta_ventas SET impsal = impsal + " & NulosN(RstTmp.Fields("impcan")) & " WHERE id = " & RstTmp.Fields("iddoc") & ";"
            End If
            RstTmp.MoveNext
        Loop
        Set RstTmp = Nothing
        'ELIMINAMOS LOS DETALLES
        xCon.Execute "DELETE * FROM con_letradet WHERE idlet = " & RstFrm("id") & ""
        xCon.Execute "DELETE * FROM con_letradoc WHERE idlet = " & RstFrm("id") & ""
        
        xNumAsiento = DevuelveNumAsiento(37, RstFrm("id"), xMes, xCon)
        
        If xNumAsiento = "" Then xNumAsiento = NuevoNumAsiento(37, xMes, xCon)
        
        'ELIMINAMOS EL ASIENTO REGISTRADO EN EL DIARIO
        xCon.Execute "DELETE * FROM con_diario WHERE ((idmes = " & xMes & ") and (idlib = 37) AND (idmov = " & RstFrm("id") & ")) ;"

        RST_Busq RstDet, "SELECT TOP 1 * FROM con_letradet", xCon
        RST_Busq RstDoc, "SELECT TOP 1 * FROM con_letradoc", xCon
        RST_Busq RstDia, "SELECT TOP 1 * FROM con_diario", xCon
        
        xId = RstCab("id")
        
    End If
    
    RstCab("ano") = AnoTra
    RstCab("idmes") = xMes
    RstCab("idlib") = 37
    RstCab("numreg") = Format(xMes, "00") + xNumAsiento
    If xMes <> 0 And xMes <> 13 Then
        RstCab("fchreg") = CDate("01/" + Format(xMes, "00") + "/" + AnoTra)
    End If
    
    RstCab("fchemi") = TxtFchEmi.Valor
    RstCab("idmon") = NulosN(TxtIdMon.Text)
    RstCab("idclipro") = NulosN(LblIdProveedor.Caption)
    
    If opt_tipo(0).Value = True Then RstCab("tiplet") = 1 'proveedor
    If opt_tipo(1).Value = True Then RstCab("tiplet") = 2 'cliente
    
    RstCab("idlet") = NulosN(TxtIdLet.Text)
    
    RstCab("numlet") = NulosN(TxtNumLet.Text)
    RstCab("vendia") = NulosN(TxtNumDiaVen.Text)
    RstCab("venlosdia") = NulosN(TxtVenDias.Text)
    RstCab("girado") = NulosC(TxtGirado.Text)
    
    RstCab("iddocgir") = NulosN(TxtIdDocIden.Text)
    RstCab("numdocgir") = NulosC(TxtNumDoc.Text)
    RstCab("implet") = NulosC(TxtTotal5Pro.Text)
    
    RstCab.Update
    
    For A = 1 To Fg2.Rows - 1
        RstDet.AddNew
        RstDet("idlet") = xId
        RstDet("corr") = A
        RstDet("numlet") = NulosC(Fg2.TextMatrix(A, 1))
        RstDet("fchemi") = Fg2.TextMatrix(A, 2)
        RstDet("fchven") = Fg2.TextMatrix(A, 3)
        RstDet("implet") = NulosN(Fg2.TextMatrix(A, 4))
        If QueHace = 1 Then
            
        End If
         
        RstDet.Update
    Next A

    For A = 1 To Fg1.Rows - 1
        RstDoc.AddNew
        RstDoc("idlet") = xId
        RstDoc("iddoc") = NulosN(Fg1.TextMatrix(A, 7))
        RstDoc("impcan") = NulosN(Fg1.TextMatrix(A, 6))
        RstDoc.Update
    Next A
    '**********************************************************************************************************
    '***********GRABAMOS EL DIARIO
    Dim xIdCuen As Integer
    Dim mSaldoDoc As Double '--SALDO DEL DOCUMENTO
    Dim mImporteLetra As Double
    Dim mIdDoc As Long
    Dim fEsHaber As Boolean
    Dim mRowLetra As Long
    '--CARGANDO LOS SALDOS PARA GENERAR EL ASIENTO (ESTA COLUMNA SERA UTIL PARA VALIDAR)
    Fg1.Cols = Fg1.Cols + 1
    Fg1.ColHidden(Fg1.Cols - 1) = True
    For A = 1 To Fg1.Rows - 1
        Fg1.TextMatrix(A, Fg1.Cols - 1) = NulosN(Fg1.TextMatrix(A, 6))
    Next A
        
    For mRowLetra = 1 To Fg2.Rows - 1
        If opt_tipo(0).Value = True Then fEsHaber = False   '--DEL PROVEEDOR
        If opt_tipo(1).Value = True Then fEsHaber = True    '--DEL PROVEEDOR
        mImporteLetra = NulosN(Fg2.TextMatrix(mRowLetra, 4))
        'GRABAMOS LA LETRA
        pGenerarAsiento RstDia, AnoTra, xMes, 37, xId, 0, mRowLetra, xNumAsiento, NulosN(LblTipoCambio.Caption), CDate(TxtFchEmi.Valor), xCuenLetra, NulosN(TxtIdMon.Text), mImporteLetra, fEsHaber
        
        A = 1
        xIdCuen = NulosN(Fg1.TextMatrix(A, 8))
        For A = 1 To Fg1.Rows - 1
            If NulosN(Fg1.TextMatrix(A, Fg1.Cols - 1)) <> 0 Then
                If mImporteLetra >= NulosN(Fg1.TextMatrix(A, Fg1.Cols - 1)) Then
                    mImporteLetra = mImporteLetra - NulosN(Fg1.TextMatrix(A, Fg1.Cols - 1))
                    mSaldoDoc = NulosN(Fg1.TextMatrix(A, Fg1.Cols - 1))
                Else
                    mSaldoDoc = mImporteLetra
                    mImporteLetra = 0
                End If
                '--OBTENIENDO EL ULTIMO SALDO DEL DOCUMENTO
                Fg1.TextMatrix(A, Fg1.Cols - 1) = NulosN(Fg1.TextMatrix(A, Fg1.Cols - 1)) - mSaldoDoc
                
                mIdDoc = NulosN(Fg1.TextMatrix(A, 7))
                
                If opt_tipo(0).Value = True Then '--DEL PROVEEDOR
                    xCon.Execute "UPDATE com_compras SET impsal = impsal - " & mSaldoDoc & " WHERE id = " & mIdDoc & ";"
                Else                             '--DEL CLIENTE
                    xCon.Execute "UPDATE vta_ventas SET impsal = impsal - " & mSaldoDoc & " WHERE id = " & mIdDoc & ";"
                End If
                
                pGenerarAsiento RstDia, AnoTra, xMes, 37, xId, mIdDoc, mRowLetra, xNumAsiento, NulosN(LblTipoCambio.Caption), CDate(TxtFchEmi.Valor), xIdCuen, NulosN(TxtIdMon.Text), mSaldoDoc, Not fEsHaber
                
                If mImporteLetra = 0 Then Exit For '--SALIR DE CANCELAR LOS DOCUMENTOS
            End If
        Next A
    Next mRowLetra
    
    '**********************************************************************************************************
    
    xCon.CommitTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDoc = Nothing
    Set RstDia = Nothing
    Me.MousePointer = vbDefault
    MsgBox "El canje de la letra se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito" + vbCr + "Nun.Reg. " + Format(xMes, "00") + xNumAsiento, vbInformation, xTitulo

    Grabar = True
    Exit Function
    
LaCague:
'    Resume
    Me.MousePointer = vbDefault
    xCon.RollbackTrans
    MsgBox "No se pudo guardar el canje de la letra por el siguiente motivo: " + vbCr + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDoc = Nothing
    Set RstDia = Nothing
End Function

Sub Bloquea()
    TxtFchEmi.Locked = Not TxtFchEmi.Locked
    TxtIdMon.Locked = Not TxtIdMon.Locked
    TxtRucPro.Locked = Not TxtRucPro.Locked
    
    TxtIdLet.Locked = Not TxtIdLet.Locked
       
    fra_tipo.Enabled = Not fra_tipo.Enabled
    fra_letra.Enabled = Not fra_letra.Enabled
        
    habilitar CmdDoc, Not CmdDoc(0).Enabled
    habilitar CmdLetra, Not CmdLetra(0).Enabled
End Sub

Sub Blanquea()
    TxtFchEmi.Valor = ""
    TxtIdMon.Text = ""
    TxtRucPro.Text = ""
    
    TxtIdLet.Text = ""
    LblLetra.Caption = ""
    
    TxtNumLet.Text = ""
    TxtNumDiaVen.Text = ""
    TxtVenDias.Text = ""
    TxtGirado.Text = ""
    TxtIdDocIden.Text = ""
    TxtNumDoc.Text = ""
    
    TxtTotal3Pro.Text = ""
    TxtTotal5Pro.Text = ""
    TxtTotal2Pro.Text = ""
    
    LblMoneda.Caption = ""
    LblDocIdentidad.Caption = ""
    
    LblTipoCambio.Caption = ""
    
    Fg1.Rows = 1
    Fg2.Rows = 1
End Sub

Sub ActivaTool()
    Dim A As Integer
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Sub Nuevo()
    QueHace = 1
    ActivaTool
    Label5.Caption = "Agregando nuevo Canje de Letras"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Blanquea
    Bloquea
    CmdLetra(2).Enabled = False
    opt_tipo(0).Value = True
    
    Fg1.Rows = 1
    Fg2.Rows = 1
    Fg1.Editable = flexEDNone
    
    Fg1.SelectionMode = flexSelectionFree
    
    Fg2.SelectionMode = flexSelectionFree
    
    GRID_COMBOLIST Fg1, 2
    
    opt_tipo(0).SetFocus
    fra_letra.Enabled = False
End Sub

Private Sub TxtFchEmi_Validate(Cancel As Boolean)
    If IsDate(TxtFchEmi.Valor) = True Then
        LblTipoCambio.Caption = HallaTipoCambio(TxtFchEmi.Valor, 2, Venta, xCon)
    Else
        LblTipoCambio.Caption = ""
    End If
End Sub

Private Sub TxtGirado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdDocIden_Change()
    If Trim(TxtIdDocIden) = "" Then LblDocIdentidad.Caption = ""
End Sub

Private Sub TxtIdDocIden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtIdDocIden_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusIdDoc_Click
    End If
End Sub

Private Sub TxtIdDocIden_Validate(Cancel As Boolean)
    If NulosN(TxtIdDocIden.Text) <> 0 Then
        LblDocIdentidad.Caption = Busca_Codigo(NulosN(TxtIdDocIden.Text), "id", "descripcion", "mae_dociden", "N", xCon)
        If NulosC(LblDocIdentidad.Caption) = "" Then
            TxtIdDocIden.Text = ""
        End If
    End If
End Sub

Private Sub TxtIdMon_Change()
    If Trim(TxtIdMon.Text) = "" Then LblMoneda.Caption = ""
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtIdMon_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMon_Click
    End If
End Sub

Private Sub TxtIdMon_Validate(Cancel As Boolean)
    If NulosC(TxtIdMon.Text) <> "" Then
        LblMoneda.Caption = Busca_Codigo(TxtIdMon.Text, "id", "descripcion", "mae_moneda", "N", xCon)
        If NulosC(LblMoneda.Caption) = "" Then
            TxtIdMon.Text = ""
        End If
    End If
End Sub

Private Sub TxtNumDiaVen_Change()
    If Trim(TxtNumDiaVen.Text) <> "" Then
        TxtVenDias.Text = ""
    End If
End Sub

Private Sub TxtNumDiaVen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtNumLet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtVenDias_Change()
    If Trim(TxtVenDias.Text) <> "" Then
        TxtNumDiaVen.Text = ""
    End If
End Sub

Private Sub TxtVenDias_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0

End Sub

'********************
Private Sub pAddLetra()
    If opt_tipo(0).Value = True Then '--proveedor
        If Fg2.TextMatrix(Fg2.Rows - 1, 1) = "" Or Fg2.TextMatrix(Fg2.Rows - 1, 2) = "" Then
            If IsDate(Fg2.TextMatrix(Fg2.Rows - 1, 2)) = False Then Fg2.TextMatrix(Fg2.Rows - 1, 2) = TxtFchEmi.Valor
            Fg2.Row = Fg2.Rows - 1: Fg2.Col = 1
            Fg2.SetFocus
            Exit Sub
        End If
        Fg2.Rows = Fg2.Rows + 1
        If IsDate(Fg2.TextMatrix(Fg2.Rows - 1, 2)) = False Then Fg2.TextMatrix(Fg2.Rows - 1, 2) = TxtFchEmi.Valor
        Fg2.Row = Fg2.Rows - 1: Fg2.Col = 1
        Fg2.SetFocus
        Exit Sub
    Else '--cliente
        '--generar las letras
        If IsDate(TxtFchEmi.Valor) = False Then
            MsgBox "Ingrese la fecha de emisión", vbExclamation, xTitulo
            TxtFchEmi.SetFocus
            Exit Sub
        End If
        If NulosN(TxtNumLet.Text) = 0 Then
            MsgBox "Ingrese el Nº de letras", vbExclamation, xTitulo
            TxtNumLet.SetFocus
            Exit Sub
        End If
        Dim nLetra As Integer
        Dim dFecha As String
        Dim dFechaIni As Date
        Dim dFechaTmp As String
        Dim mDiasMes As Integer
        Fg2.Rows = 1
        dFechaIni = TxtFchEmi.Valor '--asignando la fecha inicial
        For nLetra = 1 To NulosN(TxtNumLet.Text)
            Fg2.Rows = Fg2.Rows + 1
            If NulosN(TxtNumDiaVen.Text) <> 0 Then
                dFecha = DateAdd("d", NulosN(TxtNumDiaVen.Text), dFechaIni)
            ElseIf NulosN(TxtVenDias.Text) <> 0 Then
                Select Case Month(dFechaIni)
                    Case 12:    dFechaTmp = "01/01/" + Format(CStr(Year(dFechaIni) + 1), "0000")
                    Case Else:  dFechaTmp = "01/" + Format(Month(dFechaIni) + 1, "00") + "/" + Format((Year(dFechaIni)), "0000")
                End Select
                mDiasMes = HallaDiasMes(CDate(dFechaTmp))
                If mDiasMes >= NulosN(TxtVenDias.Text) Then
                    mDiasMes = NulosN(TxtVenDias.Text)
                End If
                dFecha = Format(mDiasMes, "00") + Right(dFechaTmp, Len(dFechaTmp) - 2)
            Else
                '--SI NO INGRESA EL PERIODO DE VENCIMIENTO O LOS DIAS QUE VENCE, SE CONSIDERARA QUE VENCERA EL FIN DE MES
                Select Case Month(dFechaIni)
                    Case 12:    dFechaTmp = "01/01/" + Format(CStr(Year(dFechaIni) + 1), "0000")
                    Case Else:  dFechaTmp = "01/" + Format(Month(dFechaIni) + 1, "00") + "/" + Format((Year(dFechaIni)), "0000")
                End Select
                mDiasMes = HallaDiasMes(CDate(dFechaTmp))
                dFecha = Format(mDiasMes, "00") + Right(dFechaTmp, Len(dFechaTmp) - 2)
            End If
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = TxtFchEmi.Valor
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = CDate(dFecha)
            '----------
            dFechaIni = dFecha
        Next nLetra
        '-------------------
        pTotalizarLetra
        '-------------------
        Fg2.Row = 1: Fg2.Col = 1
        Fg2.SetFocus
    End If
End Sub

Private Sub pDelLetra()
    If Fg2.Row <= 0 Then Exit Sub
    If Fg2.Rows = 1 Then
        MsgBox "No hay letra a eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Fg2.RemoveItem Fg2.Row
    pTotalizarLetra
End Sub

Private Sub CambiarMes()
    xMes = SeleccionaMes(xCon)
    pCargarGrid
    TabOne1.CurrTab = 0
End Sub

Private Sub pCargarGrid()
    On Error GoTo error
    Dim nSQL As String
    LblPeriodo(0).Caption = Busca_Codigo(xMes, "id", "descripcion", "con_meses", "N", xCon)
    LblPeriodo(1).Caption = LblPeriodo(0).Caption
    
    nSQL = "SELECT con_letra.*, mae_dociden.descripcion AS idendesc, IIf(con_letra.tiplet=1,mae_prov.numruc,mae_cliente.numruc) AS numruc, IIf(con_letra.tiplet=1,mae_prov.nombre,mae_cliente.nombre) AS nombre, con_letradet.numlet AS letra, con_letradet.fchven, con_letradet.implet AS letraimp, mae_moneda.simbolo AS monabrev, mae_moneda.descripcion AS mondesc, IIf([tiplet]=1,'Proveedor','Cliente') AS tipo, mae_letra.descripcion AS letdesc, mae_letra.idcuencom, mae_letra.idcuenven " _
        + vbCr + " FROM ((((mae_cliente RIGHT JOIN ((mae_prov RIGHT JOIN con_letra ON mae_prov.id = con_letra.idclipro) LEFT JOIN mae_moneda ON con_letra.idmon = mae_moneda.id) ON mae_cliente.id = con_letra.idclipro) LEFT JOIN mae_letra ON con_letra.idlet = mae_letra.id) LEFT JOIN mae_dociden ON con_letra.iddocgir = mae_dociden.id) INNER JOIN mae_libros ON con_letra.idlib = mae_libros.id) LEFT JOIN con_letradet ON con_letra.id = con_letradet.idlet " _
        + vbCr + " WHERE (((con_letra.ano) = " & AnoTra & ") And ((con_letra.idmes) = " & xMes & ")) " _
        + vbCr + " ORDER BY con_letra.numreg,con_letradet.numlet;"

    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, nSQL, xCon
    
    Set Dg1.DataSource = RstFrm
    Me.MousePointer = vbDefault
Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pCargarGrid"
End Sub

Private Sub CmdBusProv_Click()
    On Error GoTo error
    If QueHace = 3 Then Exit Sub
    Dim xRs As New ADODB.Recordset
    Dim xCampos2(2, 4) As String
    Dim nSQL As String
    Dim nTitulo As String
    
    If opt_tipo(0).Value = True Then
        xCampos2(0, 0) = "Proveedor":   xCampos2(0, 1) = "nombre":       xCampos2(0, 2) = "6000":         xCampos2(0, 3) = "C"
        nSQL = "SELECT DISTINCT mae_prov.* FROM mae_prov INNER JOIN com_compras ON mae_prov.id = com_compras.idpro WHERE (((com_compras.impsal)<>0)) ORDER BY mae_prov.nombre; "
        nTitulo = "Buscando Proveedores"
    Else
        xCampos2(0, 0) = "Cliente":   xCampos2(0, 1) = "nombre":       xCampos2(0, 2) = "6000":         xCampos2(0, 3) = "C"
        nSQL = "SELECT DISTINCT mae_cliente.* FROM mae_cliente INNER JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli WHERE (((vta_ventas.impsal)<>0)) ORDER BY mae_cliente.nombre; "
        nTitulo = "Buscando Clientes"
    End If
    xCampos2(1, 0) = "Nº R.U.C.":   xCampos2(1, 1) = "numruc":       xCampos2(1, 2) = "1500":         xCampos2(1, 3) = "C"
      
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos2(), nTitulo, "nombre", "nombre", Principio
    
    If xRs.State = 0 Then GoTo Salir:
    If xRs.RecordCount = 0 Then GoTo Salir:
    '----------------
    LblIdProveedor.Tag = LblIdProveedor.Caption
    '----------------
    TxtRucPro.Text = NulosC(xRs("numruc"))
    LblProveedor.Caption = NulosC(xRs("nombre"))
    LblIdProveedor.Caption = NulosN(xRs("id"))
    '----------------
    If NulosN(LblIdProveedor.Tag) <> NulosN(LblIdProveedor.Caption) Then
        Fg1.Rows = 1:        pTotalizarDocumento
        Fg2.Rows = 1:        pTotalizarLetra
    End If

    TxtFchEmi.SetFocus
Salir:
    Set xRs = Nothing
Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "CmdBusProv_Click"
End Sub

Private Sub TxtRucPro_Change()
    If Trim(TxtRucPro) = "" Then
        LblProveedor.Caption = ""
        LblIdProveedor.Caption = ""
        Fg1.Rows = 1 '--documentos
        Fg2.Rows = 1 '--letras
        pTotalizarDocumento
        pTotalizarLetra
    End If
End Sub

Private Sub TxtRucPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtRucPro_Validate True
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0

End Sub

Private Sub TxtRucPro_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusProv_Click
    End If
End Sub

Private Sub TxtRucPro_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If TxtRucPro.Text <> "" Then
        Dim Rst As New ADODB.Recordset
        LblIdProveedor.Tag = LblIdProveedor.Caption
        If opt_tipo(0).Value = True Then
            RST_Busq Rst, "SELECT * FROM mae_prov WHERE numruc = '" & Trim(TxtRucPro.Text) & "'", xCon
            If Rst.RecordCount <> 0 Then
                Rst.MoveFirst
                TxtRucPro.Text = NulosC(Rst("numruc"))
                LblProveedor.Caption = NulosC(Rst("nombre"))
                LblIdProveedor.Caption = NulosN(Rst("id"))
            End If
        Else
            RST_Busq Rst, "SELECT * FROM mae_cliente WHERE numruc = '" & Trim(TxtRucPro.Text) & "'", xCon
            If Rst.RecordCount <> 0 Then
                Rst.MoveFirst
                TxtRucPro.Text = NulosC(Rst("numruc"))
                LblProveedor.Caption = NulosC(Rst("nombre"))
                LblIdProveedor.Caption = NulosN(Rst("id"))
            End If
        End If
        Set Rst = Nothing
        '----------------
        If NulosN(LblIdProveedor.Tag) <> NulosN(LblIdProveedor.Caption) Then
            Fg1.Rows = 1:        pTotalizarDocumento
            Fg2.Rows = 1:        pTotalizarLetra
        End If
        '----------------
    
    End If
End Sub

Private Sub pRegistroAdd(Optional F_SELECCION_VARIOS As Boolean = True)

    On Error GoTo error
    If NulosC(TxtRucPro.Text) = "" Then
        MsgBox "Seleccione el " + LblTitulo.Caption, vbExclamation, xTitulo
        CmdBusProv.SetFocus
        Exit Sub
    End If
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "Seleccione la Moneda", vbExclamation, xTitulo
        TxtIdMon.SetFocus
        Exit Sub
    End If
    
    
    Dim xCampos(6, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim nSQLId As String '--almacenara los codigos de documentos ya seleccionados
    Dim nSQL As String
    Dim nTitulo As String

    xCampos(0, 0) = "Tip.Doc.":        xCampos(0, 1) = "abrev":     xCampos(0, 2) = "500":    xCampos(0, 3) = "C":     xCampos(0, 4) = "N"
    xCampos(1, 0) = "Nº Documento":    xCampos(1, 1) = "numdoc":    xCampos(1, 2) = "2000":    xCampos(1, 3) = "C":     xCampos(1, 4) = "S"
    xCampos(2, 0) = "Fch.Emi.":        xCampos(2, 1) = "fchdoc":    xCampos(2, 2) = "1200":     xCampos(2, 3) = "C":     xCampos(2, 4) = "N"
    xCampos(3, 0) = "M":               xCampos(3, 1) = "simbolo":   xCampos(3, 2) = "500":     xCampos(3, 3) = "C":     xCampos(4, 4) = "N"
    xCampos(4, 0) = "Importe":         xCampos(4, 1) = "imptotdoc": xCampos(4, 2) = "1000":    xCampos(4, 3) = "N":     xCampos(5, 4) = "N"
    xCampos(5, 0) = "Saldo":           xCampos(5, 1) = "impsal":    xCampos(5, 2) = "1000":    xCampos(5, 3) = "N":     xCampos(6, 4) = "N"
           
    '*************************************************************************
    nSQLId = GRID_GENERAR_SQL_ID(Fg1, 7, "com_compras.id", " NOT IN ")
    '*************************************************************************
    
    nSQL = "SELECT mae_documento.abrev, com_compras!numser+'-'+com_compras!numdoc AS numdoc, com_compras.impsal, format(com_compras.fchdoc,'dd/mm/yy') as fchdoc ,com_compras.fchven, com_compras.idpro, com_compras.imptot AS imptotdoc, com_compras.id, con_diario.idcue, con_planctas.cuenta, con_diario.idlib, mae_moneda.simbolo " _
        + vbCr + " FROM mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (con_planctas RIGHT JOIN (com_compras LEFT JOIN con_diario ON com_compras.id = con_diario.idmov) ON con_planctas.id = con_diario.idcue) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon " _
        + vbCr + " WHERE (((com_compras.impsal)>0) AND ((com_compras.idpro)=" & NulosN(LblIdProveedor.Caption) & ") AND ((con_planctas.cuenta) Like '42%') AND ((con_diario.idlib)=1)) AND com_compras.idmon= " & NulosN(TxtIdMon.Text) & " " _
        + vbCr + IIf(nSQLId = "", "", " AND " + nSQLId) _
        + vbCr + " ORDER BY com_compras.numser+'-'+com_compras.numdoc"
       
    'nSQL = nSQL + IIf(nSQLId = "", "", " AND " + nSQLId)
    
    If opt_tipo(1).Value = True Then '--ventas
        nSQL = Replace(nSQL, "42%", "12%")
        nSQL = Replace(nSQL, "com_compras.idpro", "vta_ventas.idcli")
        nSQL = Replace(nSQL, "com_compras.imptot", "vta_ventas.imptotdoc")
        nSQL = Replace(nSQL, "(con_diario.idlib)=1", "(con_diario.idlib)=2")
        nSQL = Replace(nSQL, "com_compras.id", "vta_ventas.id")
        nSQL = Replace(nSQL, "com_compras", "vta_ventas")

        '*************************************************************
        nSQLId = Replace(nSQLId, "com_compras.id", "vta_ventas.id")
        '*************************************************************
    End If
    
    nTitulo = "Buscando Documento del " + LblTitulo.Caption + ": " + StrConv(LblProveedor.Caption, 3)
    '*************************************************************
    
    If F_SELECCION_VARIOS = True Then
        CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), nTitulo
    Else
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "numdoc", "numdoc", CualquierParte
    End If
    If xRs.State = 0 Then GoTo Salir
    If xRs.RecordCount = 0 Then GoTo Salir
    If F_SELECCION_VARIOS = True Then xRs.MoveFirst
    Agregando = True
    Do While Not xRs.EOF
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(xRs("abrev"))
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(xRs("numdoc"))
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(xRs("fchdoc"))
        Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(xRs("simbolo"))
        Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(NulosN(xRs("imptotdoc")), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(NulosN(xRs("impsal")), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosN(xRs("id"))
        Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosN(xRs("idcue"))
        
        If F_SELECCION_VARIOS = False Then Exit Do
        xRs.MoveNext
    Loop
    Agregando = False
    pTotalizarDocumento
    Fg1.Row = Fg1.Rows - 1: Fg1.Col = 2:  Fg1.SetFocus
Salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    pTotalizarDocumento
    Agregando = False
    SHOW_ERROR Me.Name, "pRegistroAdd"
End Sub

Private Sub Eliminar()
    On Error GoTo error
    Dim Rpta As Integer
    If RstFrm.RecordCount = 0 Or RstFrm.EOF = True Or RstFrm.BOF = True Then
        MsgBox "No hay registro para eliminar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    
    Rpta = MsgBox("Esta seguro de eliminar el Canje de la letra seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        '--AGREGANDO EL SALDO AL DOCUMENTO DEL PROEVEEDOR O CLIENTE
        Dim RstTmp As New ADODB.Recordset
        RST_Busq RstTmp, "SELECT iddoc, impcan From con_letradoc WHERE (((idlet)=" & RstFrm("id") & "));", xCon
        xCon.BeginTrans
        If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            If RstFrm.Fields("tiplet") = 1 Then
                xCon.Execute "UPDATE com_compras SET impsal = impsal + " & NulosN(RstTmp.Fields("impcan")) & " WHERE id = " & RstTmp.Fields("iddoc") & ";"
            Else
                xCon.Execute "UPDATE vta_ventas SET impsal = impsal + " & NulosN(RstTmp.Fields("impcan")) & " WHERE id = " & RstTmp.Fields("iddoc") & ";"
            End If
            RstTmp.MoveNext
        Loop
        Set RstTmp = Nothing
        'ELIMINAMOS LOS DETALLES
        xCon.Execute "DELETE * FROM con_letradet WHERE idlet = " & RstFrm("id") & ""
        xCon.Execute "DELETE * FROM con_letradoc WHERE idlet = " & RstFrm("id") & ""
        xCon.Execute "DELETE * FROM con_letra WHERE id = " & RstFrm("id") & ""
        'ELIMINAMOS EL ASIENTO REGISTRADO EN EL DIARIO
        xCon.Execute "DELETE * FROM con_diario WHERE ((idmes = " & xMes & ") and (idlib = 37) AND (idmov = " & RstFrm("id") & "))"

        xCon.CommitTrans
        RstFrm.Requery
        Dg1.Refresh
        MsgBox "El Canje de la letra se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
    TabOne1.CurrTab = 0
    
    Exit Sub
error:
        xCon.RollbackTrans
        SHOW_ERROR Me.Name, "Eliminar", True, "No pudo modificar por el siguiente motivo:"
End Sub


Private Sub Modificar()
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay registro para modificar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    
    QueHace = 2
    Label5.Caption = "Modificando Canje de Letra"
    Bloquea
    CmdLetra(2).Enabled = False
    ActivaTool
    
    If TabOne1.CurrTab = 0 Then
        TabOne1.CurrTab = 1
        MuestraSegundoTab
    End If
    
    TabOne1.TabEnabled(0) = False
    GRID_COMBOLIST Fg1, 2
    
    Fg1.Editable = flexEDKbdMouse
    Fg1.SelectionMode = flexSelectionFree
    Fg2.Editable = flexEDKbdMouse
    Fg2.SelectionMode = flexSelectionFree
    
    
    
    Agregando = False
    
    If opt_tipo(1).Value = True Then
        fra_letra.Enabled = True
    Else
        fra_letra.Enabled = False
    End If

    TxtRucPro.SetFocus
End Sub

Private Sub Filtrar()
    
    ReDim xCampos(6, 4) As String
    'xCampos(0, 0) = "Num.Reg.":             xCampos(0, 1) = "numreg":   xCampos(0, 2) = "900":   xCampos(0, 3) = "C"
    xCampos(0, 0) = "Tipo Letra":           xCampos(0, 1) = "tipo":     xCampos(0, 2) = "C":    xCampos(1, 3) = "800"
    xCampos(1, 0) = "Cliente / Proveedor":  xCampos(1, 1) = "nombre":   xCampos(1, 2) = "C":    xCampos(2, 3) = "3200"
    xCampos(2, 0) = "M":                    xCampos(2, 1) = "monabrev": xCampos(2, 2) = "C":    xCampos(3, 3) = "1000"
    xCampos(3, 0) = "Nº Letra":             xCampos(3, 1) = "letra":    xCampos(3, 2) = "C":    xCampos(4, 3) = "1500"
    xCampos(4, 0) = "Fch.Emi":              xCampos(4, 1) = "fchemi":   xCampos(4, 2) = "F":    xCampos(5, 3) = "900"
    xCampos(5, 0) = "Fch.Ven":              xCampos(5, 1) = "fchven":   xCampos(5, 2) = "F":    xCampos(5, 3) = "900"
    xCampos(6, 0) = "Importe":              xCampos(6, 1) = "letraimp": xCampos(6, 2) = "N":    xCampos(6, 3) = "1000"
    
    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg1
    Me.TabOne1.CurrTab = 0
End Sub

Private Function fValidarDatos() As Boolean
    If IsDate(TxtFchEmi.Valor) = False Then
        MsgBox "No ha especificado la fecha de emisión de la letra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchEmi.SetFocus
        Exit Function
    End If
    
    If NulosN(LblIdProveedor.Caption) = 0 Then
        MsgBox "No ha especificado el " + LblTitulo.Caption, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtRucPro.SetFocus
        Exit Function
    End If
    
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "No ha especificado la moneda en que se emite la letra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMon.SetFocus
        Exit Function
    End If
    
    If NulosN(TxtIdLet.Text) = 0 Then
        MsgBox "No ha especificado la letra que se esta aplicando", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdLet.SetFocus
        Exit Function
    End If

    
    If opt_tipo(1).Value = True Then
        If TxtNumLet.Text = "" Then
            MsgBox "No ha especificado el número de letras a emitir", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtNumLet.SetFocus
            Exit Function
        End If
        
        If NulosN(TxtNumDiaVen.Text) = 0 And NulosN(TxtVenDias.Text) = 0 Then
            MsgBox "No ha especificado cada cuantos dias vence una letra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtNumDiaVen.SetFocus
            Exit Function
        End If
        
        If TxtGirado.Text = "" Then
            MsgBox "No ha especificado a quién se girará la letra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtGirado.SetFocus
            Exit Function
        End If
        If TxtIdDocIden.Text = "" Then
            MsgBox "No ha especificado el tipo de documento de identidad", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtIdDocIden.SetFocus
            Exit Function
        End If
        If TxtNumDoc.Text = "" Then
            MsgBox "No ha especificado el número de documento de identidad", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtNumDoc.SetFocus
            Exit Function
        End If
    End If
    
    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado que documentos se cargaran a las letras", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Function
    End If

    If Fg2.Rows = 1 Then
        MsgBox "No ha especificado las letras que se asignaran a los documentos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Function
    End If
    
    '**********************************************
        '--VALIDAR EL INGRESO DE LOS DATOS
    Dim Q_ROW  As Long
    Dim Q_COL As Long '--COLUMNA A POSICIONAR SI FALTAN DATOS
    Q_COL = -1
    For Q_ROW = 1 To Fg2.Rows - 1
        If NulosC(Fg2.TextMatrix(Q_ROW, 1)) = "" Then
            MsgBox "Ingrese el número de la Letra", vbExclamation, xTitulo
            Q_COL = 1:          Exit For
        ElseIf IsDate(Fg2.TextMatrix(Q_ROW, 2)) = False Then
            MsgBox "Ingrese la Fecha de Emisión de la letra: " + Fg2.TextMatrix(Q_ROW, 1), vbExclamation, xTitulo
            Q_COL = 2:          Exit For
        ElseIf IsDate(Fg2.TextMatrix(Q_ROW, 3)) = False Then
            MsgBox "Ingrese la Fecha de Vencimiento de la letra: " + Fg2.TextMatrix(Q_ROW, 1), vbExclamation, xTitulo
            Q_COL = 3:          Exit For
        ElseIf IsNumeric(Fg2.TextMatrix(Q_ROW, 4)) = False Then
           MsgBox "Ingrese el Importe de la letra: " + Fg2.TextMatrix(Q_ROW, 1), vbExclamation, xTitulo
            Q_COL = 4:        Exit For
        End If
    Next Q_ROW
    If Q_COL <> -1 Then
        Agregando = True:  Fg2.Row = Q_ROW: Fg2.Col = Q_COL: Agregando = False
        Fg2.SetFocus
        Exit Function
    End If
    '---------------------------------------------------------------------------
    If Fg1.Rows = 2 Then
        If NulosN(TxtTotal5Pro.Text) <> NulosN(TxtTotal3Pro.Text) Then
            MsgBox "El Importe total de " + IIf(Fg2.Rows > 1, "las letras", "la letra") + " debe de ser igual o menor al saldo del documento", vbExclamation, xTitulo
            Exit Function
        End If
    Else
        If NulosN(TxtTotal5Pro.Text) <> NulosN(TxtTotal3Pro.Text) Then
            MsgBox "El Importe total de " + IIf(Fg2.Rows > 1, "las letras", "la letra") + " debe de ser igual saldo total de los documentos" + vbCr + _
            "Importe Total de " + IIf(Fg2.Rows > 1, "las letras", "la letra") + ":   " + TxtTotal5Pro.Text + vbCr + _
            "Saldo Total de los Documentos:  " + TxtTotal3Pro.Text + vbCr + _
            "Diferencia: " + CStr(Format(NulosN(TxtTotal5Pro.Text) - NulosN(TxtTotal3Pro.Text), FORMAT_MONTO)) + vbCr + _
            "Obs: Modifique los importes de las letra, cuya suma sea igual al saldo total de los documentos", vbExclamation, xTitulo
            Exit Function
        End If
    End If

    '**********************************************
    fValidarDatos = True
End Function

Private Sub Buscar()
    On Error GoTo error
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    ReDim xCampos(7, 4) As String
    
    xCampos(0, 0) = "Num.Reg.":             xCampos(0, 1) = "numreg":   xCampos(0, 2) = "900":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Tipo Letra":           xCampos(1, 1) = "tipo":     xCampos(1, 2) = "950":   xCampos(1, 3) = "C"
    xCampos(2, 0) = "Cliente / Proveedor":  xCampos(2, 1) = "nombre":   xCampos(2, 2) = "2800":  xCampos(2, 3) = "C"
    xCampos(3, 0) = "M":                    xCampos(3, 1) = "monabrev": xCampos(3, 2) = "450":  xCampos(3, 3) = "C"
    xCampos(4, 0) = "Nº Letra":             xCampos(4, 1) = "letra":    xCampos(4, 2) = "1200":  xCampos(4, 3) = "C"
    xCampos(5, 0) = "Fch.Emi":              xCampos(5, 1) = "fchemi":   xCampos(5, 2) = "950":   xCampos(5, 3) = "F"
    xCampos(6, 0) = "Importe":              xCampos(6, 1) = "letraimp": xCampos(6, 2) = "1000":  xCampos(6, 3) = "N"
    
    nSQL = "SELECT con_letra.*, mae_dociden.descripcion AS idendesc, IIf(con_letra.tiplet=1,mae_prov.numruc,mae_cliente.numruc) AS numruc, IIf(con_letra.tiplet=1,mae_prov.nombre,mae_cliente.nombre) AS nombre, con_letradet.numlet AS letra, con_letradet.fchven, con_letradet.implet AS letraimp, mae_moneda.simbolo AS monabrev, mae_moneda.descripcion AS mondesc, IIf([tiplet]=1,'Proveedor','Cliente') AS tipo, mae_letra.descripcion AS letdesc, mae_letra.idcuencom, mae_letra.idcuenven " _
        + vbCr + " FROM ((((mae_cliente RIGHT JOIN ((mae_prov RIGHT JOIN con_letra ON mae_prov.id = con_letra.idclipro) LEFT JOIN mae_moneda ON con_letra.idmon = mae_moneda.id) ON mae_cliente.id = con_letra.idclipro) LEFT JOIN mae_letra ON con_letra.idlet = mae_letra.id) LEFT JOIN mae_dociden ON con_letra.iddocgir = mae_dociden.id) INNER JOIN mae_libros ON con_letra.idlib = mae_libros.id) LEFT JOIN con_letradet ON con_letra.id = con_letradet.idlet " _
        + vbCr + " WHERE (((con_letra.ano) = " & AnoTra & ") And ((con_letra.idmes) = " & xMes & ")) " _
        + vbCr + " ORDER BY con_letra.numreg,con_letradet.numlet;"
            
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Canjes de Letra", "letra", "letra", Principio
    
    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir
    
    RstFrm.MoveFirst
    RstFrm.Find "id = " & xRs("id") & ""
Salir:
    Set xRs = Nothing
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub


Private Sub pGenerarAsiento(RstDiario As ADODB.Recordset, nAnoTrabajo, mMesActivo, IDLibro, IDMov, mIdDocPro, mCorr, nAsiento, mTipoCambio, FchDoc, IDcuenta, IDMoneda, mImporte, Optional EsDEBE As Boolean)
    '--mCorr por le general es igual a 0
    RstDiario.AddNew
    RstDiario("año") = nAnoTrabajo
    RstDiario("idmes") = mMesActivo  'CODIGO DEL MES
    RstDiario("idlib") = IDLibro     'CODIGO DEL LIBRO
    RstDiario("idmov") = IDMov       'CODIGO DEL MOVIMIENTO
    RstDiario("iddocpro") = mIdDocPro
    RstDiario("correlativo") = mCorr
    RstDiario("numasi") = nAsiento
    RstDiario("tc") = mTipoCambio
    If mMesActivo = 0 Then
        RstDiario("fchasi") = CDate("01/01/" + nAnoTrabajo)
    ElseIf mMesActivo = 13 Then
        RstDiario("fchasi") = CDate("31/12/" + nAnoTrabajo)
    Else
        RstDiario("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + nAnoTrabajo)
    End If
    RstDiario("fchdoc") = FchDoc
    RstDiario("idcue") = IDcuenta
    If EsDEBE = False Then
        If IDMoneda = 1 Then '--DE LA MONEDA
            RstDiario("imphabsol") = mImporte
            RstDiario("imphabdol") = 0
        Else
            RstDiario("imphabsol") = mImporte * mTipoCambio
            RstDiario("imphabdol") = mImporte
        End If
    Else
        If IDMoneda = 1 Then '--DE LA MONEDA
            RstDiario("impdebsol") = mImporte
            RstDiario("impdebdol") = 0
        Else
            RstDiario("impdebsol") = mImporte * mTipoCambio
            RstDiario("impdebdol") = mImporte
        End If
    End If

    RstDiario.Update
End Sub


Private Sub CmdIdLet_Click()
    If QueHace = 3 Then Exit Sub
    Dim xRs As New ADODB.Recordset
    ReDim xCampos(1, 4) As String
    On Error GoTo error
    xCampos(0, 0) = "Descripcion":     xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "6500":         xCampos(0, 3) = "C"
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, "SELECT * FROM mae_letra", xCampos(), "Buscando Letra", "descripcion", "descripcion", Principio
    
    If xRs.State = 0 Then GoTo Salir:
    If xRs.RecordCount = 0 Then GoTo Salir:
    TxtIdLet.Text = xRs("id") & ""
    LblLetra.Caption = xRs("descripcion") & ""
    If opt_tipo(0).Value = True Then
        xCuenLetra = NulosN(xRs("idcuencom"))
    Else
        xCuenLetra = NulosN(xRs("idcuenven"))
    End If
    If opt_tipo(0).Value = True Then
        CmdDoc(0).SetFocus
    Else
        TxtNumLet.SetFocus
    End If

Salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "CmdIdLet_Click"
End Sub

Private Sub TxtIdLet_Change()
    If Trim(TxtIdLet.Text) = "" Then LblLetra.Caption = ""
End Sub

Private Sub TxtIdLet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdLet_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then CmdIdLet_Click
End Sub

Private Sub TxtIdLet_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If Trim(TxtIdLet.Text) = "" Then Exit Sub
    
    LblLetra.Caption = Busca_Codigo(NulosN(TxtIdLet.Text), "id", "descripcion", "mae_letra", "N", xCon)
    
    If LblLetra.Caption <> "" Then
        If opt_tipo(0).Value = True Then
            xCuenLetra = Busca_Codigo(NulosN(TxtIdLet.Text), "id", "idcuencom", "mae_letra", "N", xCon)
        Else
            xCuenLetra = Busca_Codigo(NulosN(TxtIdLet.Text), "id", "idcuenven", "mae_letra", "N", xCon)
        End If
    Else
        TxtIdLet.Text = ""
    End If
End Sub


Private Sub pMostrarInfLetra()
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    '----------
    Me.Toolbar1.Enabled = False
    Me.TabOne1.Enabled = False
    '----------
    fraInfLetra.Visible = True
    fraInfLetra.Top = 1710
    fraInfLetra.Left = 2505
    
    With Fg2
        lblInfLetra(4).Caption = .TextMatrix(.Row, 1)
        lblInfLetra(5).Caption = .TextMatrix(.Row, 2)
        lblInfLetra(6).Caption = .TextMatrix(.Row, 3)
        lblInfLetra(7).Caption = .TextMatrix(.Row, 4)
        lblInfLetra(8).Caption = LblMoneda.Caption
    End With
    
    nSQL = "SELECT mae_documento.abrev, com_compras.numreg, [com_compras].[numser] & '-' & [com_compras].[numdoc] AS numdoc, com_compras.fchdoc , mae_moneda.simbolo, com_compras.imptot as imptotdoc, con_diario.impdebsol, con_diario.imphabsol, con_diario.impdebdol, con_diario.imphabdol " _
        + vbCr + " FROM con_letra INNER JOIN ((con_letradet INNER JOIN con_diario ON (con_letradet.idlet = con_diario.idmov) AND (con_letradet.corr = con_diario.correlativo)) INNER JOIN (mae_moneda RIGHT JOIN (mae_documento INNER JOIN vta_ventas ON (mae_documento.id = vta_ventas.tipdoc) AND (mae_documento.id = vta_ventas.tipdoc)) ON mae_moneda.id = vta_ventas.idmon) ON con_diario.iddocpro = vta_ventas.id) ON (con_letra.id = con_letradet.idlet) AND (con_letra.idclipro = vta_ventas.idcli) " _
        + vbCr + " WHERE (((con_diario.idlib)=8) AND ((con_diario.iddocpro)<>0) AND ((con_letradet.idlet)=" & RstFrm.Fields("id") & ") AND ((con_letradet.corr)=" & Fg2.TextMatrix(Fg2.Row, 5) & "));"


    If opt_tipo(1).Value = True Then
        nSQL = Replace(nSQL, "com_compras.imptot", "vta_ventas.imptotdoc")
        nSQL = Replace(nSQL, "com_compras", "vta_ventas")
    End If
    
    RST_Busq RstTmp, nSQL, xCon
    Fg3.Rows = 1
    If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
    Do While Not RstTmp.EOF
        Fg3.Rows = Fg3.Rows + 1
        Fg3.TextMatrix(Fg3.Rows - 1, 1) = NulosC(RstTmp("numreg"))
        Fg3.TextMatrix(Fg3.Rows - 1, 2) = NulosC(RstTmp("abrev"))
        Fg3.TextMatrix(Fg3.Rows - 1, 3) = NulosC(RstTmp("numdoc"))
        Fg3.TextMatrix(Fg3.Rows - 1, 4) = Format(RstTmp("fchdoc"), "dd/mm/yy")
        Fg3.TextMatrix(Fg3.Rows - 1, 5) = NulosC(RstTmp("simbolo"))
        Fg3.TextMatrix(Fg3.Rows - 1, 6) = Format(NulosN(RstTmp("imptotdoc")), FORMAT_MONTO)
        If opt_tipo(0).Value = True Then '--PROVEEDOR
            If NulosN(TxtIdMon.Text) = 1 Then Fg3.TextMatrix(Fg3.Rows - 1, 7) = Format(NulosN(RstTmp("impdebsol")), FORMAT_MONTO)
            If NulosN(TxtIdMon.Text) = 2 Then Fg3.TextMatrix(Fg3.Rows - 1, 7) = Format(NulosN(RstTmp("impdebdol")), FORMAT_MONTO)
        Else '--CLIENTE
            If NulosN(TxtIdMon.Text) = 1 Then Fg3.TextMatrix(Fg3.Rows - 1, 7) = Format(NulosN(RstTmp("imphabsol")), FORMAT_MONTO)
            If NulosN(TxtIdMon.Text) = 2 Then Fg3.TextMatrix(Fg3.Rows - 1, 7) = Format(NulosN(RstTmp("imphabdol")), FORMAT_MONTO)
        End If
        
        RstTmp.MoveNext
    Loop
    Dim mTotalImporte As Double
    Dim mTotalImpLetra As Double
    
    mTotalImporte = GRID_SUMAR_COL(Fg3, 6)
    mTotalImpLetra = GRID_SUMAR_COL(Fg3, 7)
    
    Fg3.Rows = Fg3.Rows + 1
    FORMATO_CELDA Fg3, Fg3.Rows - 1, 4, , True, , "Total"
    FORMATO_CELDA Fg3, Fg3.Rows - 1, 6, , True, , Format(mTotalImporte, FORMAT_MONTO)
    FORMATO_CELDA Fg3, Fg3.Rows - 1, 7, , True, , Format(mTotalImpLetra, FORMAT_MONTO)
    
    Set RstTmp = Nothing
    Exit Sub
error:
    Set RstTmp = Nothing
    CmdInfLet_Click 2
    SHOW_ERROR Me.Name, "pMostrarInfLetra"
End Sub


Private Sub pLetraInfExportar()
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios
    Dim nTitulo As String
    Dim nPeriodo As String
    nTitulo = "Información relacionada a la Letra Nª " + lblInfLetra(4).Caption
    nPeriodo = "Fecha Emisión: " + lblInfLetra(5).Caption
    Me.MousePointer = vbHourglass
    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, Fg3, nTitulo, nPeriodo, "", "Letra: " + lblInfLetra(4).Caption
    Set X_EXPORT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Set X_EXPORT = Nothing
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Exportar"
End Sub


Private Sub pLetraInfImprimir()
    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios
    Dim nTitulo As String
    Dim nPeriodo As String
    nTitulo = "Información relacionada a la Letra Nª " + lblInfLetra(4).Caption
    nPeriodo = "Fecha Emisión: " + lblInfLetra(5).Caption

    Me.MousePointer = vbHourglass
    X_PRINT.Imprimir_x_VSFlexGrid Fg3, nTitulo, "", nPeriodo, True, True
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Imprimir"
End Sub

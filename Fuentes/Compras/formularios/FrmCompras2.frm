VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCompras2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compras - Ingreso de Compras"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdApertura 
      Caption         =   "&Apertura"
      Height          =   315
      Left            =   10575
      TabIndex        =   128
      Top             =   375
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Frame Frame6 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   3645
      Left            =   12210
      TabIndex        =   29
      Top             =   3630
      Visible         =   0   'False
      Width           =   8610
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   390
         Left            =   4425
         TabIndex        =   35
         Top             =   3120
         Width           =   1320
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   390
         Left            =   5790
         TabIndex        =   34
         Top             =   3120
         Width           =   1320
      End
      Begin VB.CommandButton CmdAddCenCos 
         Caption         =   "&Agregar C.C."
         Height          =   390
         Left            =   1500
         TabIndex        =   33
         Top             =   3120
         Width           =   1320
      End
      Begin VB.CommandButton CmdDelCenCos 
         Caption         =   "&Eliminar C.C."
         Height          =   390
         Left            =   2865
         TabIndex        =   32
         Top             =   3120
         Width           =   1320
      End
      Begin VB.TextBox TxtTotPor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   6330
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "TxtTotPor"
         Top             =   2670
         Width           =   975
      End
      Begin VB.TextBox TxtTotImp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   7305
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "TxtTotImp"
         Top             =   2670
         Width           =   960
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg5 
         Height          =   2190
         Left            =   75
         TabIndex        =   36
         Top             =   465
         Width           =   8460
         _cx             =   14922
         _cy             =   3863
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmCompras2.frx":0000
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
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   8595
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   8595
         X2              =   8595
         Y1              =   15
         Y2              =   3645
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   3615
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   8595
         Y1              =   3630
         Y2              =   3630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detallar Centro de Costos"
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
         Left            =   255
         TabIndex        =   37
         Top             =   90
         Width           =   2190
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   300
         Left            =   45
         Top             =   45
         Width           =   8520
      End
   End
   Begin VB.Frame Frame11 
      BorderStyle     =   0  'None
      Caption         =   "Frame11"
      Height          =   2700
      Left            =   12030
      TabIndex        =   25
      Top             =   300
      Visible         =   0   'False
      Width           =   7320
      Begin VB.CommandButton CmdAcep 
         Caption         =   "&Aceptar"
         Height          =   405
         Left            =   2985
         TabIndex        =   26
         Top             =   2220
         Width           =   1305
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg4 
         Height          =   1710
         Left            =   195
         TabIndex        =   27
         Top             =   465
         Width           =   6900
         _cx             =   12171
         _cy             =   3016
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmCompras2.frx":00BB
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
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   7305
         Y1              =   2685
         Y2              =   2685
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   7290
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   15
         Y1              =   15
         Y2              =   2670
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   7305
         X2              =   7305
         Y1              =   15
         Y2              =   2685
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documentos Adjuntos"
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
         Left            =   180
         TabIndex        =   28
         Top             =   135
         Width           =   1860
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00400000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   330
         Left            =   45
         Top             =   60
         Width           =   7230
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7230
      Left            =   15
      TabIndex        =   38
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
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6810
         Left            =   -12435
         TabIndex        =   87
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6510
            Left            =   30
            TabIndex        =   88
            Top             =   300
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11483
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
            Columns(1).Caption=   "Nº Reg."
            Columns(1).DataField=   "numreg1"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "T.D."
            Columns(2).DataField=   "abrev"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Nº Documento"
            Columns(3).DataField=   "numerodoc"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fch. Emi"
            Columns(4).DataField=   "fchdoc1"
            Columns(4).NumberFormat=   "Short Date"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Fch. Ven."
            Columns(5).DataField=   "fchven1"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Proveedor"
            Columns(6).DataField=   "nombre"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "M"
            Columns(7).DataField=   "simbolo"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "T.C."
            Columns(8).DataField=   "impven1"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Importe"
            Columns(9).DataField=   "imptot1"
            Columns(9).NumberFormat=   "0.00"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Saldo"
            Columns(10).DataField=   "impsal1"
            Columns(10).NumberFormat=   "0.00"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   11
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=11"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1905"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1826"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=900"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=820"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=2461"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2381"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=1693"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1614"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=1720"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1640"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=5715"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=5636"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=512"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=767"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=688"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=513"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(50)=   "Column(8).Width=1138"
            Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=1058"
            Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=516"
            Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(56)=   "Column(9).Width=1640"
            Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=1561"
            Splits(0)._ColumnProps(59)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(60)=   "Column(9)._ColStyle=514"
            Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(62)=   "Column(10).Width=1667"
            Splits(0)._ColumnProps(63)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(64)=   "Column(10)._WidthInPix=1588"
            Splits(0)._ColumnProps(65)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(66)=   "Column(10)._ColStyle=514"
            Splits(0)._ColumnProps(67)=   "Column(10).Order=11"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=74,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=71,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=72,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=73,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=62,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=78,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=75,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=76,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=77,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=51,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=52,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=53,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=63,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=64,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=65,.parent=17"
            _StyleDefs(80)  =   "Named:id=33:Normal"
            _StyleDefs(81)  =   ":id=33,.parent=0"
            _StyleDefs(82)  =   "Named:id=34:Heading"
            _StyleDefs(83)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(84)  =   ":id=34,.wraptext=-1"
            _StyleDefs(85)  =   "Named:id=35:Footing"
            _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(87)  =   "Named:id=36:Selected"
            _StyleDefs(88)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(89)  =   "Named:id=37:Caption"
            _StyleDefs(90)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(91)  =   "Named:id=38:HighlightRow"
            _StyleDefs(92)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(93)  =   "Named:id=39:EvenRow"
            _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(95)  =   "Named:id=40:OddRow"
            _StyleDefs(96)  =   ":id=40,.parent=33"
            _StyleDefs(97)  =   "Named:id=41:RecordSelector"
            _StyleDefs(98)  =   ":id=41,.parent=34"
            _StyleDefs(99)  =   "Named:id=42:FilterBar"
            _StyleDefs(100) =   ":id=42,.parent=33"
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
            TabIndex        =   91
            Top             =   30
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
            TabIndex        =   90
            Top             =   30
            Width           =   11610
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
            TabIndex        =   89
            Top             =   0
            Width           =   1860
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6810
         Left            =   45
         TabIndex        =   39
         Top             =   375
         Width           =   11790
         Begin VB.CommandButton CmdBusDocRef2 
            Height          =   240
            Left            =   8010
            Picture         =   "FrmCompras2.frx":0198
            Style           =   1  'Graphical
            TabIndex        =   125
            Top             =   2340
            Width           =   240
         End
         Begin VB.CommandButton CmdBusTipDocRef 
            Height          =   240
            Left            =   2145
            Picture         =   "FrmCompras2.frx":02CA
            Style           =   1  'Graphical
            TabIndex        =   121
            Top             =   2340
            Width           =   240
         End
         Begin VB.CommandButton CmdBusDocRef 
            Height          =   240
            Left            =   8010
            Picture         =   "FrmCompras2.frx":03FC
            Style           =   1  'Graphical
            TabIndex        =   115
            Top             =   1710
            Width           =   240
         End
         Begin VB.TextBox TxtDocRef 
            Height          =   300
            Left            =   6255
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   10
            Text            =   "TxtDocRef"
            Top             =   1680
            Width           =   2025
         End
         Begin VB.CommandButton CmdBusTipDoc 
            Height          =   240
            Left            =   6735
            Picture         =   "FrmCompras2.frx":052E
            Style           =   1  'Graphical
            TabIndex        =   110
            Top             =   1395
            Width           =   240
         End
         Begin VB.CommandButton CmdBusProv 
            Height          =   240
            Left            =   3150
            Picture         =   "FrmCompras2.frx":0660
            Style           =   1  'Graphical
            TabIndex        =   107
            Top             =   450
            Width           =   240
         End
         Begin VB.TextBox TxtGlosa 
            Height          =   300
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "TxtGlosa"
            Top             =   1995
            Width           =   10050
         End
         Begin VB.CommandButton CmdBusMon 
            Height          =   240
            Left            =   2145
            Picture         =   "FrmCompras2.frx":0792
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   1410
            Width           =   240
         End
         Begin VB.CommandButton CmdBusCondicion 
            Height          =   240
            Left            =   6735
            Picture         =   "FrmCompras2.frx":08C4
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   765
            Width           =   240
         End
         Begin VB.Frame Frame5 
            Height          =   495
            Left            =   9675
            TabIndex        =   52
            Top             =   -90
            Width           =   2115
            Begin VB.Label LblPeriodo2 
               Alignment       =   2  'Center
               Caption         =   "LblPeriodo2"
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
               Left            =   120
               TabIndex        =   53
               Top             =   150
               Width           =   1860
            End
         End
         Begin VB.CommandButton CmdBusTipoCompra 
            Height          =   240
            Left            =   2145
            Picture         =   "FrmCompras2.frx":09F6
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   1080
            Width           =   240
         End
         Begin VB.Frame Frame3 
            Caption         =   "( Afecta :)"
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
            Height          =   600
            Left            =   1020
            TabIndex        =   48
            Top             =   4350
            Visible         =   0   'False
            Width           =   2805
            Begin VB.OptionButton OptNo 
               Caption         =   "No Afecto"
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
               Left            =   1320
               TabIndex        =   50
               Top             =   270
               Width           =   1440
            End
            Begin VB.OptionButton OptSi 
               Caption         =   "Afecto"
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
               Left            =   105
               TabIndex        =   49
               Top             =   285
               Width           =   1125
            End
         End
         Begin VB.CommandButton CmdBusAlm 
            Height          =   240
            Left            =   6735
            Picture         =   "FrmCompras2.frx":0B28
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   1095
            Width           =   240
         End
         Begin VB.Frame Frame7 
            Caption         =   "[ Opciones de Descuento]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   540
            Left            =   9135
            TabIndex        =   44
            Top             =   2625
            Width           =   2580
            Begin VB.OptionButton OptDes1 
               Caption         =   "Porcentaje"
               Height          =   195
               Left            =   165
               TabIndex        =   46
               Top             =   270
               Width           =   1215
            End
            Begin VB.OptionButton OptDes2 
               Caption         =   "Valor"
               Height          =   195
               Left            =   1590
               TabIndex        =   45
               Top             =   270
               Width           =   870
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "[ Rta 4ta Cat. ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   540
            Left            =   5385
            TabIndex        =   42
            Top             =   2625
            Visible         =   0   'False
            Width           =   1815
            Begin VB.CheckBox ChkImpRen4 
               Caption         =   "Aplicar Impuesto"
               Height          =   195
               Left            =   195
               TabIndex        =   43
               Top             =   270
               Width           =   1470
            End
         End
         Begin VB.Frame Frame10 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   540
            Left            =   7215
            TabIndex        =   40
            Top             =   2625
            Width           =   1905
            Begin VB.CheckBox Check1 
               Caption         =   "Ingresar Neto"
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
               Left            =   210
               TabIndex        =   41
               Top             =   270
               Width           =   1500
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   2190
            Left            =   105
            TabIndex        =   14
            Top             =   3150
            Width           =   11610
            _cx             =   20479
            _cy             =   3863
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
            Rows            =   20
            Cols            =   18
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmCompras2.frx":0C5A
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
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchDoc 
            Height          =   300
            Left            =   1665
            TabIndex        =   1
            Top             =   735
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
            Valor           =   "22/05/2008"
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchVen 
            Height          =   300
            Left            =   10485
            TabIndex        =   3
            Top             =   735
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
            Valor           =   "22/05/2008"
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchPago 
            Height          =   300
            Left            =   10485
            TabIndex        =   15
            Top             =   1680
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
            Valor           =   "22/05/2008"
         End
         Begin VB.TextBox TxtIdAlmacen 
            Height          =   300
            Left            =   6255
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   5
            Text            =   "TxtIdAlmacen"
            Top             =   1050
            Width           =   750
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2760
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   9
            Text            =   "TxtNumDoc"
            Top             =   1680
            Width           =   1440
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   8
            Text            =   "TxtNumSer"
            Top             =   1680
            Width           =   915
         End
         Begin VB.TextBox TxtIdMon 
            Height          =   300
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   6
            Text            =   "TxtIdMon"
            Top             =   1365
            Width           =   750
         End
         Begin VB.TextBox TxtConPag 
            Height          =   300
            Left            =   6255
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   2
            Text            =   "TxtConPag"
            Top             =   735
            Width           =   750
         End
         Begin VB.TextBox TxtTipCom 
            Height          =   300
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   4
            Text            =   "TxtTipCom"
            Top             =   1050
            Width           =   750
         End
         Begin VB.Frame Frame4 
            Height          =   1485
            Left            =   105
            TabIndex        =   56
            Top             =   5325
            Width           =   11610
            Begin VB.CommandButton CmdVerAsiento 
               Caption         =   "&Ver Asiento Contable"
               Height          =   390
               Left            =   8910
               TabIndex        =   129
               Top             =   465
               Width           =   2565
            End
            Begin VB.TextBox TxtOtros 
               Alignment       =   1  'Right Justify
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
               Left            =   7575
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   119
               TabStop         =   0   'False
               Text            =   "TxtOtros"
               Top             =   1110
               Width           =   1230
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Presupuesto"
               Height          =   345
               Left            =   1545
               Style           =   1  'Graphical
               TabIndex        =   106
               ToolTipText     =   "Presupuesto"
               Top             =   990
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.CommandButton CmdPreHist 
               Caption         =   "Ver His. Precios"
               Height          =   345
               Left            =   135
               Style           =   1  'Graphical
               TabIndex        =   105
               ToolTipText     =   "Historico de Precios"
               Top             =   990
               Width           =   1395
            End
            Begin VB.CheckBox ChkAjusta 
               Caption         =   "Ajustar Totales"
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
               Left            =   8835
               TabIndex        =   61
               Top             =   0
               Visible         =   0   'False
               Width           =   1605
            End
            Begin VB.TextBox TxtIGV3 
               Alignment       =   1  'Right Justify
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
               Left            =   6255
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   22
               TabStop         =   0   'False
               Text            =   "TxtIGV3"
               Top             =   1110
               Width           =   1230
            End
            Begin VB.TextBox TxtIGV2 
               Alignment       =   1  'Right Justify
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
               Left            =   4650
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   21
               TabStop         =   0   'False
               Text            =   "TxtIGV2"
               Top             =   1110
               Width           =   1230
            End
            Begin VB.TextBox TxtBruto3 
               Alignment       =   1  'Right Justify
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
               Left            =   6255
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   18
               TabStop         =   0   'False
               Text            =   "TxtBruto3"
               Top             =   555
               Width           =   1230
            End
            Begin VB.TextBox TxtBruto2 
               Alignment       =   1  'Right Justify
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
               Left            =   4650
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   17
               TabStop         =   0   'False
               Text            =   "TxtBruto2"
               Top             =   555
               Width           =   1230
            End
            Begin VB.TextBox TxtTotal 
               Alignment       =   1  'Right Justify
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
               Left            =   10215
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   24
               TabStop         =   0   'False
               Text            =   "TxtTotal"
               Top             =   1110
               Width           =   1230
            End
            Begin VB.TextBox TxtIGV 
               Alignment       =   1  'Right Justify
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
               Left            =   3285
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   20
               TabStop         =   0   'False
               Text            =   "TxtIGV"
               Top             =   1110
               Width           =   1230
            End
            Begin VB.TextBox TxtBruto 
               Alignment       =   1  'Right Justify
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
               Left            =   3285
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   16
               TabStop         =   0   'False
               Text            =   "TxtBruto"
               Top             =   555
               Width           =   1230
            End
            Begin VB.TextBox TxtInafecto 
               Alignment       =   1  'Right Justify
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
               Left            =   7575
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   19
               TabStop         =   0   'False
               Text            =   "TxtInafect"
               Top             =   555
               Width           =   1230
            End
            Begin VB.TextBox TxtISC 
               Alignment       =   1  'Right Justify
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
               Left            =   8895
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   23
               TabStop         =   0   'False
               Text            =   "TxtISC"
               Top             =   1110
               Width           =   1230
            End
            Begin VB.CommandButton CmdDelItem 
               Caption         =   "Eliminar Item"
               Enabled         =   0   'False
               Height          =   345
               Left            =   1545
               Style           =   1  'Graphical
               TabIndex        =   60
               ToolTipText     =   "Eliminar Item"
               Top             =   270
               Width           =   1395
            End
            Begin VB.CommandButton CmdAddItem 
               Caption         =   "Agregar Item"
               Enabled         =   0   'False
               Height          =   345
               Left            =   135
               Style           =   1  'Graphical
               TabIndex        =   59
               ToolTipText     =   "Agregar Item"
               Top             =   270
               Width           =   1395
            End
            Begin VB.CommandButton CmdDetCenCos 
               Caption         =   "Centro de Costo"
               Height          =   345
               Left            =   1545
               Style           =   1  'Graphical
               TabIndex        =   58
               ToolTipText     =   "Centro de Costos"
               Top             =   630
               Width           =   1395
            End
            Begin VB.CommandButton CmdSeleccionar 
               Caption         =   "Seleccionar Item"
               Height          =   345
               Left            =   135
               Style           =   1  'Graphical
               TabIndex        =   57
               ToolTipText     =   "Seleccionar Items "
               Top             =   630
               Width           =   1395
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Otros Cargos"
               Height          =   195
               Left            =   7620
               TabIndex        =   120
               Top             =   885
               Width           =   915
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Tasa del I.G.V."
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
               Left            =   8925
               TabIndex        =   113
               Top             =   225
               Width           =   1305
            End
            Begin VB.Label LblIgvTasa 
               Alignment       =   2  'Center
               Caption         =   "LblIgvTasa"
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
               Height          =   225
               Left            =   10335
               TabIndex        =   104
               Top             =   210
               Width           =   1110
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "No Gravadas"
               Height          =   195
               Index           =   9
               Left            =   7620
               TabIndex        =   103
               Top             =   360
               Width           =   945
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Credito Filscal"
               Height          =   195
               Index           =   8
               Left            =   6240
               TabIndex        =   102
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "de Exp. o no Grav"
               Height          =   195
               Index           =   7
               Left            =   4650
               TabIndex        =   101
               Top             =   360
               Width           =   1290
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Grav y Exp"
               Height          =   195
               Index           =   6
               Left            =   3270
               TabIndex        =   100
               Top             =   360
               Width           =   780
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "I.G.V. "
               Height          =   195
               Left            =   6240
               TabIndex        =   99
               Top             =   885
               Width           =   450
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "I.G.V. "
               Height          =   195
               Left            =   4650
               TabIndex        =   98
               Top             =   885
               Width           =   450
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "A.G. sin derecho"
               Height          =   195
               Index           =   5
               Left            =   6240
               TabIndex        =   97
               Top             =   180
               Width           =   1185
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base Imp. Ope. Grav."
               Height          =   195
               Index           =   4
               Left            =   4650
               TabIndex        =   96
               Top             =   180
               Width           =   1530
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total"
               Height          =   195
               Index           =   2
               Left            =   10215
               TabIndex        =   66
               Top             =   885
               Width           =   360
            End
            Begin VB.Label LblRotulo 
               AutoSize        =   -1  'True
               Caption         =   "I.G.V. "
               Height          =   195
               Left            =   3270
               TabIndex        =   65
               Top             =   885
               Width           =   450
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base. Imp. Ope."
               Height          =   195
               Index           =   0
               Left            =   3270
               TabIndex        =   64
               Top             =   180
               Width           =   1140
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Adquisiciones"
               Height          =   195
               Index           =   1
               Left            =   7620
               TabIndex        =   63
               Top             =   180
               Width           =   975
            End
            Begin VB.Line Line4 
               BorderColor     =   &H80000003&
               Index           =   0
               X1              =   3060
               X2              =   3060
               Y1              =   105
               Y2              =   1470
            End
            Begin VB.Line Line4 
               BorderColor     =   &H80000005&
               Index           =   1
               X1              =   3075
               X2              =   3075
               Y1              =   105
               Y2              =   1455
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "I.S.C."
               Height          =   195
               Index           =   3
               Left            =   9015
               TabIndex        =   62
               Top             =   885
               Width           =   390
            End
         End
         Begin VB.TextBox TxtNumRuc 
            Height          =   300
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   0
            Text            =   "TxtNumRuc"
            Top             =   420
            Width           =   1770
         End
         Begin VB.TextBox TxtTipDoc 
            Height          =   300
            Left            =   6255
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   7
            Text            =   "TxtTipDoc"
            Top             =   1365
            Width           =   750
         End
         Begin VB.Frame Frame9 
            Caption         =   "[ Opciones de Compra ]"
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
            ForeColor       =   &H00400000&
            Height          =   540
            Left            =   120
            TabIndex        =   67
            Top             =   2625
            Width           =   5250
            Begin VB.OptionButton OptOpera2 
               Caption         =   "Ord. de Compra"
               Height          =   195
               Left            =   2535
               TabIndex        =   71
               ToolTipText     =   "Orden de Compra"
               Top             =   270
               Visible         =   0   'False
               Width           =   1410
            End
            Begin VB.CommandButton CmdCargaDoc 
               Caption         =   "Adicionar"
               Height          =   300
               Left            =   4080
               TabIndex        =   70
               Top             =   165
               Width           =   1095
            End
            Begin VB.OptionButton OptOpera1 
               Caption         =   "Normal"
               Height          =   195
               Left            =   105
               TabIndex        =   69
               ToolTipText     =   "Operacion Normal"
               Top             =   270
               Width           =   825
            End
            Begin VB.OptionButton OptOpera3 
               Caption         =   "Doc. Entrada"
               Height          =   195
               Left            =   1110
               TabIndex        =   68
               ToolTipText     =   "Documentos de Entrada"
               Top             =   270
               Width           =   1260
            End
         End
         Begin VB.TextBox TxtDocRef2 
            Height          =   300
            Left            =   6255
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   13
            Text            =   "TxtDocRef2"
            Top             =   2310
            Width           =   2025
         End
         Begin VB.TextBox TxtTipDocRef 
            Height          =   300
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   12
            Text            =   "Txt"
            Top             =   2310
            Width           =   750
         End
         Begin VB.Label lblReg 
            Caption         =   "lblReg"
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
            Height          =   270
            Left            =   9540
            TabIndex        =   127
            Top             =   1050
            Width           =   2190
         End
         Begin VB.Label LblIdDocRef2 
            AutoSize        =   -1  'True
            Caption         =   "LblIdDocRef2"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   8370
            TabIndex        =   126
            Top             =   2355
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Doc. Referencia"
            Height          =   195
            Index           =   13
            Left            =   4800
            TabIndex        =   124
            Top             =   2355
            Width           =   1395
         End
         Begin VB.Label LblTipDocref 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipDocref"
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
            TabIndex        =   123
            Top             =   2310
            Width           =   2325
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tip de Doc. Ref."
            Height          =   195
            Index           =   12
            Left            =   150
            TabIndex        =   122
            Top             =   2355
            Width           =   1185
         End
         Begin VB.Label LblIdTipPer 
            AutoSize        =   -1  'True
            Caption         =   "LblIdTipPer"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   0
            TabIndex        =   118
            Top             =   225
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label Lbltipo 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Lbltipo"
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   9540
            TabIndex        =   117
            Top             =   420
            Width           =   2190
         End
         Begin VB.Label LblIdDocRef 
            AutoSize        =   -1  'True
            Caption         =   "LblIdDocRef"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   8310
            TabIndex        =   116
            Top             =   1725
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Referente al Documento"
            Height          =   195
            Index           =   9
            Left            =   4455
            TabIndex        =   114
            Top             =   1725
            Width           =   1740
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Documento"
            Height          =   195
            Index           =   1
            Left            =   4785
            TabIndex        =   112
            Top             =   1425
            Width           =   1410
         End
         Begin VB.Label LblNomDoc 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblNomDoc"
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
            Left            =   7020
            TabIndex        =   111
            Top             =   1365
            Width           =   2325
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            Height          =   195
            Index           =   7
            Left            =   150
            TabIndex        =   109
            Top             =   450
            Width           =   735
         End
         Begin VB.Label LblNomPro 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblNomPro"
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
            Left            =   3450
            TabIndex        =   108
            Top             =   420
            Width           =   5910
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Glosa"
            Height          =   195
            Index           =   10
            Left            =   150
            TabIndex        =   95
            Top             =   2025
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   5
            Left            =   150
            TabIndex        =   94
            Top             =   1425
            Width           =   585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Condición de Pago"
            Height          =   195
            Index           =   4
            Left            =   4845
            TabIndex        =   93
            Top             =   780
            Width           =   1350
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Almacén"
            Height          =   195
            Index           =   11
            Left            =   5580
            TabIndex        =   92
            Top             =   1110
            Width           =   615
         End
         Begin VB.Label LblIdProveedor 
            AutoSize        =   -1  'True
            Caption         =   "LblIdProveedor"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   1110
            TabIndex        =   86
            Top             =   270
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000001&
            BackStyle       =   1  'Opaque
            Height          =   90
            Left            =   2610
            Top             =   1800
            Width           =   105
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Documento"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   85
            Top             =   1725
            Width           =   1275
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Compras"
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
            TabIndex        =   84
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
            Left            =   2430
            TabIndex        =   83
            Top             =   1365
            Width           =   2325
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
            Left            =   7020
            TabIndex        =   82
            Top             =   735
            Width           =   2325
         End
         Begin VB.Label LblTipoCambio 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   10485
            TabIndex        =   81
            Top             =   1365
            Width           =   1260
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "T.C."
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
            Height          =   195
            Left            =   10005
            TabIndex        =   80
            Top             =   1425
            Width           =   375
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Emisión"
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   79
            Top             =   780
            Width           =   1260
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Venc."
            Height          =   195
            Index           =   3
            Left            =   9600
            TabIndex        =   78
            Top             =   780
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Item"
            Height          =   195
            Index           =   6
            Left            =   150
            TabIndex        =   77
            Top             =   1125
            Width           =   660
         End
         Begin VB.Label LblTipoCompra 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipoCompra"
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
            TabIndex        =   76
            Top             =   1050
            Width           =   2325
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Pago"
            Height          =   195
            Index           =   8
            Left            =   9285
            TabIndex        =   75
            Top             =   1725
            Width           =   1095
         End
         Begin VB.Label LblIdCenCos 
            AutoSize        =   -1  'True
            Caption         =   "LblIdCenCos"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   4050
            TabIndex        =   74
            Top             =   285
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label LblIdAlmacen 
            AutoSize        =   -1  'True
            Caption         =   "LblIdAlmacen"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   2655
            TabIndex        =   73
            Top             =   270
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label LblDesAlmacen 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDesAlmacen"
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
            Left            =   7020
            TabIndex        =   72
            Top             =   1050
            Width           =   2325
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8265
      Top             =   30
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
            Picture         =   "FrmCompras2.frx":0E73
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCompras2.frx":13B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCompras2.frx":1749
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCompras2.frx":18CD
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCompras2.frx":1D21
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCompras2.frx":1E39
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCompras2.frx":237D
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCompras2.frx":28C1
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCompras2.frx":29D5
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCompras2.frx":2AE9
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCompras2.frx":2F3D
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCompras2.frx":30A9
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCompras2.frx":35F1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   130
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   1058
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
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin VB.Menu menu1 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu1_1_1 
         Caption         =   "Agregar Item"
      End
      Begin VB.Menu menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Eliminar Item"
      End
      Begin VB.Menu menu1_4 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_5 
         Caption         =   "Ver Historico de Precios"
      End
   End
   Begin VB.Menu Opciones 
      Caption         =   "Opciones"
      Visible         =   0   'False
      Begin VB.Menu Opciones_1 
         Caption         =   "Agregar documentos de entrada del proveedor"
      End
      Begin VB.Menu Opciones_2 
         Caption         =   "Agregar documentos de entrada registrados del proveedor"
      End
      Begin VB.Menu Opciones_3 
         Caption         =   "-"
      End
      Begin VB.Menu Opciones_4 
         Caption         =   "Agregar documentos de entrada - Otros Proveedores"
      End
      Begin VB.Menu Opciones_5 
         Caption         =   "Agregar documentos de entrada registrados - Otros Proveedores"
      End
   End
End
Attribute VB_Name = "FrmCompras2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMCOMPRAS2.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : FORMULARIO DONDE SE REGISTRAN LAS COMPRAS, PERMITIENDO DETALLAR LOS ITEMS DE LA
'*                    COMPRA Y SU RESPECTIVO CENTRO DE COSTOS, ASI MISMO SE GENERA EL PROCESO CONTABLE
'                     PARA LA COMPRA
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 10/09/09
'* VERSION          : 1.0
'*****************************************************************************************************

Option Explicit
Dim RstComp As New ADODB.Recordset         ' RECORDSET PRINCIPAL EL FORMULARIO, ALMACENARA LAS COMPRAS DE UN PERIODO ESPECIFICADO
Dim QueHace As Integer                     ' VARIABLE QUE INDICA EN QUE MODO SE ENCUENTA EL FORMULARIO: 1 = NUEVO; 2 = MODIFICA; 3 = SOLO LECTURA
Dim TasaImpuesto As Double                 ' ALAMCENA LA TASA DEL I.G.V.
Dim CaracteresNumericos As String          ' ALMACENA LOS CARACTERES NUMERICOS PARA EL EVENTO KeyPress DE LOS CONTROLES TextBox
Dim CaracteresNumericos2 As String         ' ALMACENA LOS CARACTERES NUMERICOS PARA EL EVENTO KeyPress DE LOS CONTROLES TextBox
Dim SeEjecuto As Boolean                   ' VARIABLE QUE SE UTILIZA COMO SWITCH PARA CONTROLAR EL EVENTO Activate DEL FORMULARIO
Dim ValTipCam As Double                    ' ALMACENA EL VALOR DEL TIPO DE CAMBIO USADO PARA LA OPERACION
Dim xDescImp As String                     ' ALMACENA LA DESCRIPCION DEL IMPUESTO
Dim xIdCuenTasa As Integer                 ' codigo de la cuenta contable del impuesto
Dim xCuentaDoc As Integer                  ' codigo de la cuenta contable del documento
Dim Mostrando As Boolean                   ' VARIABLE QUE SE UTILIZA COMO SWITCH PARA SABER QUE SE ESTAN AGREGANDO DATOS A UN CONTROL FlexGrid
Dim RstTmp As New ADODB.Recordset          ' RECORDSET TEMPORAL PARA CARGAR DATOS TEMPORALES
Dim xFchFin, xFchIni, xFechaMes As String  ' ALMACENA LA FECHA DE INICIO Y LA FECHA FINAL PARA REALIZAR OPERACIONES
Dim RstTempISC As New ADODB.Recordset      ' RECORDSET TEMPORAL PARA CARGAR LOS VALORES DEL IMPUESTO SELECTIVO
Dim AgePer As Boolean                      ' ESPECIFICA SI ES UN AGENTE DE PERPCEPCION
Dim AgeRet As Boolean                      ' ESPECIFICA SI ES UN AGENTE DE RETENCION
Dim DetCenCos As Boolean                   ' especifica si se va a detallar el centro de costos
Dim CodSunatDoc As String                  ' especifica el codigo de la sunat del documento
Dim xPorIgv  As Double                     ' ALMACENA EL VALOR EN PORCENTAJE DEL IGV
Dim xHorIni As Date                        ' ALMACENA LA HORA DE INICIO EN QUE SE GENERA LA OPERACION

Dim fOrdenLista As Boolean                 ' --especfica el orden de la lista de la consulta
Dim mIdRegistro&                           ' --identificador del registro
Dim Agregando As Boolean
Dim mMesActivo As Integer                  ' --indica el mes activo
Dim fCierrePeriodo As Boolean              ' --indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)

Dim PermitirEdicion As Boolean             ' ESPECIFICA SI SE PODRA HACER MODIFICACIONES EN LOS CONTROLES DE SELECCION POR DEFECTO ES VERDADERO
Dim SeSeleTipDocRef As Boolean             ' ESPECIFICA SI SE SELECCIONO O DIGITO EL CODIGO DE UN TIPO DE DOCUMENTO DE REFERENCIA
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO
'Para mover el frame
Dim OrigFX As Long
Dim OrigFY As Long



'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PROCEDIMIENTO PARA ELIMINAR UN OPERACION DE COMPRA, ADEMAS ELIMINA TAMBIEN EL
'*                    REGISTRO CONTABLE GENREADO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    Dim Rpta As Integer
    TabOne1.CurrTab = 0
    If RstComp.State = 0 Then Exit Sub
    If RstComp.RecordCount = 0 Then
        MsgBox "No hay Registros de Compras para eliminar", vbInformation, xTitulo
        Exit Sub
    End If
    
    '**********************************************************************************************************************
    '---evaluar si el registro de compras esta vinculado con otros modulos
    Dim nSQL As String
    Dim Rst As New ADODB.Recordset
    Dim xId&
    
    xId = RstComp("id")
    '--generando la consulta
    '-----
    nSQL = "SELECT Left(tes_caja.[numreg], 2) & '01' & Right(tes_caja.[numreg],4) AS registro   " _
        + vbCr + " FROM tes_caja INNER JOIN tes_cajadestinodet ON tes_caja.id = tes_cajadestinodet.idtes " _
        + vbCr + " WHERE (((tes_cajadestinodet.iddoc)=" & xId & ") AND ((tes_cajadestinodet.idmod)=1) AND ((tes_caja.tipmov)=2));"
    RST_Busq Rst, nSQL, xCon
    If Rst.RecordCount <> 0 Then
        MsgBox "El registro de Compra esta vinculado con: " + vbCr + "Módulo: Tesoreria - Egresos" & vbCr & "Nº. Registro: " & NulosC(Rst("registro")) & vbCr & "Si desea continuar, Elimine primero el Registro " & NulosC(Rst("registro")) & " del módulo de Tesoreria - Egresos", vbInformation, xTitulo
        Set Rst = Nothing
        Exit Sub
    End If
    Set Rst = Nothing
    
    '-----
    nSQL = "SELECT Left([con_canjes].[numreg],2) & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & Right([con_canjes].[numreg],4) AS registro  " _
        + vbCr + " FROM (con_canjes LEFT JOIN mae_libros ON con_canjes.idlib = mae_libros.id) INNER JOIN con_canjesdet ON con_canjes.id = con_canjesdet.idcan " _
        + vbCr + " WHERE (((con_canjesdet.iddoc)=" & xId & ") AND ((con_canjesdet.tipo)=2)) "
    RST_Busq Rst, nSQL, xCon
    If Rst.RecordCount <> 0 Then
        MsgBox "El registro de Compra esta vinculado con: " + vbCr + "Módulo: Tesorería - Canje de documentos" & vbCr & "Nº. Registro: " & NulosC(Rst("registro")) & vbCr & "Si desea continuar, Elimine primero el Registro " & NulosC(Rst("registro")) & " del módulo de Tesorería - Canje de documentos", vbInformation, xTitulo
        Set Rst = Nothing
        Exit Sub
    End If
    Set Rst = Nothing
    '**********************************************************************************************************************
    
    Rpta = MsgBox("¿Esta seguro de eliminar la compra seleccionada?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        '--Eliminar referencia en transferencia de documento si lo hubiera, asimismo cambiar estado preparado para transferir
        xCon.Execute "UPDATE tra_documento SET tra_documento.idmod = 0, tra_documento.iddoc = 0, tra_documento.estado = 0 WHERE (((tra_documento.idmod)=1) AND ((tra_documento.iddoc)=" & xId & ")) "
        
        ' Actualizamos los documentos que esten vinculados con la compra Documentos de ingreso o Ordenes de compra
        
        ' actualizamos ingresos a almacen
        xCon.Execute "UPDATE alm_ingreso SET alm_ingreso.idfac = 0 WHERE (((alm_ingreso.idfac)=" & xId & "))"
        
        ' eliminar documentos relacionados a ingreso almacen
        xCon.Execute "DELETE * FROM alm_ingresodoc WHERE iddoc = " & xId & ""
        
        ' actualizamos orden de compra
        xCon.Execute "UPDATE com_ordencompra SET com_ordencompra.idfac = 0 WHERE (((com_ordencompra.idfac)=" & xId & "))"
        
        ' --eliminando el diario
        xCon.Execute "DELETE * FROM con_diario WHERE idmov = " & xId & " AND idlib = 1"
        
        xCon.Execute "DELETE * FROM com_comprascosto WHERE idcom = " & xId & ""
        xCon.Execute "DELETE * FROM com_comprasdet WHERE idcom = " & xId & ""
        xCon.Execute "DELETE * FROM com_compras WHERE id = " & xId & ""
        
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & xId & " AND idform = " & IdMenuActivo
        
        MsgBox "La compra se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstComp.Requery
        Dg1.Refresh
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LA BARRA DE HERRAMIENTAS DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL PROCESO DE AGREGAR O MODIFICAR UNA COMPRA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    Bloquea
    Fg1.ColComboList(1) = ""
    Label5.Caption = "Detalle de Compra"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.Editable = flexEDNone
    ActivaTool
    QueHace = 3
    Dg1.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FRMULARIO PARA EL INGRESO DE UNA COMPRA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Nuevo()
    
    If PuedeAgregarRegistro("COMPRAS", xCon) = False Then
        MsgBox "Esta utilizando una versión de prueba del maravilloso sistema SEVEN Soft, si desea la version comercial contactese con el " & Chr(13) _
            & " extraordinario programador Enrique Pollongo a eps_76@hotmail.com y solicite un número de licencia para esta PC", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If


    QueHace = 1
    Blanquea
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    Label5.Caption = "Agregando Compra"
    Fg1.ColComboList(1) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    OptSi.Value = True
    
    Fg1.Rows = 1
    If Fg1.Rows = 2 Then
        Fg1.TextMatrix(1, 4) = 0
        Fg1.TextMatrix(1, 5) = 0
    End If
    Fg5.Rows = 1
    Fg1.Rows = Fg1.Rows + 1
    OptDes1.Value = True
    OptDes1_Click
    OptOpera1.Value = True
    If xOrigen = 1 Then
        CargarValoresDefecto
    End If
    TxtIdAlmacen.Visible = True
    LblDesAlmacen.Visible = True
    CmdBusAlm.Visible = True
    Label3(11).Visible = True
    pGridConfigurar
    SeSeleTipDocRef = False
    xHorIni = Time
    TxtNumRuc.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : CargarValoresDefecto
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA ALGUNOS VALORES POR DEFECTO PARA EL INGRESO DE UN NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub CargarValoresDefecto()
    TxtFchDoc.Valor = Date
    TxtTipCom.Text = "1"
    TxtTipCom_Validate True
    TxtIdMon.Text = 1
    TxtIdMon_Validate True
    TxtTipDoc.Text = "1"
    TxtTipDoc_Validate True
    TxtConPag.Text = "1"
    TxtConPag_Validate True
    TxtIdAlmacen.Text = "1"
    TxtIdAlmacen_Validate True
    
    TxtFchVen.Valor = Date
    OptOpera1.Value = True
    OptOpera1_Click
End Sub

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Modificar()
    QueHace = 2
    Blanquea
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    MuestraSegundoTab
    Label5.Caption = "Modificando Compra"
    Fg1.ColComboList(1) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    xHorIni = Time
    SeSeleTipDocRef = False
    TxtFchDoc.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DETALLADA DEL REGISTRO ACTUAL
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    ' BLANQUEAMOS LA PESTAÑA DETALLE Y MOSTRAMOS LOS PRINCIPALES DATOS DE LA COMPRA
    Blanquea
    If RstComp.EOF = True Or RstComp.BOF = True Or RstComp.RecordCount = 0 Then Exit Sub
    Dim xRs As New ADODB.Recordset
    lblReg.Caption = "Nº Reg. " & NulosC(RstComp("numreg1"))
    
    TxtTipCom.Text = NulosN(RstComp("idtipo"))
    pGridConfigurar
    TxtTipDoc.Text = NulosN(RstComp("tipdoc"))
    TxtNumRuc.Text = NulosC(RstComp("numruc"))
    TxtNumSer.Text = NulosC(RstComp("numser"))
    TxtNumDoc.Text = NulosC(RstComp("numdoc"))
    If IsDate(RstComp("fchdoc")) = True Then TxtFchDoc.Valor = RstComp("fchdoc")
    If IsDate(RstComp("fchven")) = True Then TxtFchVen.Valor = RstComp("fchven")
    If IsDate(RstComp("fchpag")) = True Then TxtFchPago.Valor = RstComp("fchpag")
    
    ' mostramos el documento de referencia de la compra
    If NulosN(RstComp("idtipdocref")) = 0 Then
        TxtTipDocRef.Text = ""
    Else
        TxtTipDocRef.Text = NulosN(RstComp("idtipdocref"))
    End If
    TxtTipDocRef_Validate False
    LblIdDocRef2.Caption = NulosN(RstComp("iddocref2"))
    
    TxtConPag.Text = NulosN(RstComp("idconpag"))
    TxtIdMon.Text = NulosN(RstComp("idmon"))
    TxtGlosa.Text = NulosC(RstComp("glosa"))
    
    Dim Rst As New ADODB.Recordset
    
    ' BUSCAMOS SI LA COMPRA TIENE ALGUN DOCUMENTO DE REFERENCIA ASIGNADO
    If NulosN(TxtTipDocRef.Text) = 1 Then
        RST_Busq Rst, "SELECT com_ordencompra.id, [com_ordencompra]![numser] & '-' & [com_ordencompra]![numdoc] AS numdoc From com_ordencompra " _
            & " WHERE (((com_ordencompra.id)=" & NulosN(LblIdDocRef2.Caption) & "))", xCon
    End If
    If NulosN(TxtTipDocRef.Text) = 2 Then
    End If
    
    If NulosN(TxtTipDocRef.Text) = 3 Then
    End If
    
    If NulosN(TxtTipDocRef.Text) = 4 Then
        RST_Busq Rst, "SELECT var_ordendespacho.id, var_ordendespacho.numerodoc AS numdoc " _
            & " From var_ordendespacho WHERE (((var_ordendespacho.id)=" & NulosN(LblIdDocRef2.Caption) & "))", xCon
    
    End If
    
    If NulosN(TxtTipDocRef.Text) = 6 Then
        RST_Busq Rst, "SELECT com_ordenreq.id, [com_ordenreq]![numser] & '-' & [com_ordenreq]![numdoc] AS numdoc From com_ordenreq WHERE (((com_ordenreq.id)=" & NulosN(LblIdDocRef2.Caption) & "))", xCon
    End If
    
    If Rst.State <> 0 Then
        If Rst.RecordCount <> 0 Then
            TxtDocRef2.Text = NulosC(Rst("numdoc"))
            LblIdDocRef2.Caption = Rst("id")
        Else
            TxtDocRef2.Text = ""
            LblIdDocRef2.Caption = ""
        End If
    End If
    Set Rst = Nothing
    '--uso temporal
    If NulosN(LblIdDocRef2.Caption) = 0 Then TxtDocRef2.Text = NulosC(RstComp("numerodocref"))
    
    
    LblTipoCompra.Caption = RstComp("desctipcom")
    LblNomDoc.Caption = NulosC(RstComp("nomdoc"))
    LblNomPro.Caption = NulosC(RstComp("nombre"))
    LblCondPag.Caption = NulosC(RstComp("desccond"))
    TxtNumRuc.Text = NulosC(RstComp("numruc"))
    LblMoneda.Caption = NulosC(RstComp("descmon"))
    
    LblIdProveedor.Caption = NulosN(RstComp("idpro"))
    LblIdAlmacen.Caption = NulosN(RstComp("idalm"))
    TxtIdAlmacen.Text = NulosN(RstComp("idalm"))
    LblDesAlmacen.Caption = Busca_Codigo(RstComp("idalm"), "id", "descripcion", "alm_almacenes", "N", xCon)
    
    Dim xCambioQueHace As Boolean
    xCambioQueHace = False
    If QueHace = 3 Then
        xCambioQueHace = True
        QueHace = 2
    End If
    TxtNumRuc_Validate True
    TxtTipCom_Validate True
    If xCambioQueHace = True Then
        QueHace = 3
    End If
    
    If NulosN(TxtTipDoc.Text) = 7 Then   ' SI ES UNA NOTA DE CREDITO
        Label3(9).Visible = True
        TxtDocRef.Visible = True
        CmdBusDocRef.Visible = True
    Else
        Label3(9).Visible = False
        TxtDocRef.Visible = False
        CmdBusDocRef.Visible = False
    End If
    
    ' SOLO CUANDO SEA NOTA DE CREDITO
    ' CARGA EL DOCUMENTO DE COMPRA AL QUE HACE REFERENCIA LA NOTA DE CREDITO
    If NulosN(RstComp("iddocref")) <> 0 Then
        TxtDocRef.Text = Busca_Codigo(RstComp("iddocref"), "id", "numser", "com_compras", "N", xCon) + "-" + Busca_Codigo(RstComp("iddocref"), "id", "numdoc", "com_compras", "N", xCon)
    End If
    LblIdDocRef.Caption = NulosN(RstComp("iddocref"))
    
    If LblDesAlmacen.Caption = "" Then TxtIdAlmacen.Text = ""
    
    If RstComp("idmon") = 1 Then
        'LblTipoCambio.Visible = False
    Else
        'LblTipoCambio.Visible = True
        If mMesActivo = 0 Then
            LblTipoCambio.Caption = HallaTipoCambio("01/01/" + Trim(AnoTra), 2, Venta, xCon)
        Else
            LblTipoCambio.Caption = HallaTipoCambio(RstComp("fchdoc"), 2, Venta, xCon)
        End If
    End If
    
    ' mostramos el tipo de descuento que se le aplica a la compra
    Mostrando = True
    If RstComp("tipdes") = 1 Or NulosN(RstComp("tipdes")) = 0 Then
        OptDes1.Value = True
    End If
    
    If RstComp("tipdes") = 2 Then
        OptDes2.Value = True
    End If
    Mostrando = False
    
    ' Preguntamos en que contexto se realizo la compra
    If RstComp("tipcom") = 1 Then
        ' Se registro una compra normal
        OptOpera1.Value = True
        OptOpera1_Click
    End If
    
    If RstComp("tipcom") = 2 Then
        'Se registro una compra con documento de ingreso
        OptOpera3.Value = True
        OptOpera3_Click
        CargarIngresoAlmacen RstComp("id")
    End If
    
    If RstComp("tipcom") = 3 Then
        ' Se registro una compra con orden de compra
        OptOpera2.Value = True
        OptOpera2_Click
    End If
    
    ' --------------------------------------
    ' revisar si este pedaso de codigo sirve
    If RstComp("afecto") = -1 Then
        OptSi.Value = True
    Else
        OptNo.Value = True
    End If
    '--------------------------------------
    
    TxtTipDoc_Validate True
    Dim RstDet As New ADODB.Recordset
    Dim A As Integer
    
    ' cargamos la cuenta del igv
    RST_Busq RstDet, "SELECT mae_impuestos.idcuen, mae_impuestos.tasa, mae_documento.id FROM mae_documento LEFT JOIN mae_impuestos " _
        & " ON mae_documento.idimp = mae_impuestos.id WHERE (((mae_documento.id)=val(" & NulosN(TxtTipDoc.Text) & ")))", xCon

    If RstDet.RecordCount <> 0 Then
        xIdCuenTasa = NulosN(RstDet("idcuen"))
        TasaImpuesto = NulosN(RstDet("tasa"))
    End If
    Set RstDet = Nothing
    
    Set RstDet = Nothing
    Mostrando = True
    Fg1.Rows = 1
    
    ' CARGAMOS EL DETALLE DE LA COMPRA
    RST_Busq RstDet, "SELECT com_comprasdet.*, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.idcuenta, " _
        & " alm_inventario.idtipcom, con_planctas.ctadesdeb, con_planctas.ctadeshab,  alm_inventario.idnetonodomi, " _
        & " con_planctas_1.ctadesdeb AS ctadesdeb1, con_planctas_1.ctadeshab AS ctadeshab1 " _
        & " FROM (con_planctas RIGHT JOIN (mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN com_comprasdet ON alm_inventario.id = com_comprasdet.iditem)  " _
        & " ON mae_unidades.id = alm_inventario.idunimed) ON con_planctas.id = alm_inventario.idcuenta) LEFT JOIN con_planctas AS con_planctas_1  " _
        & " ON com_comprasdet.idcue = con_planctas_1.id WHERE (((com_comprasdet.idcom)=" & RstComp("id") & "))", xCon
                       
    If RstDet.RecordCount <> 0 Then
        RstDet.MoveFirst
        For A = 1 To RstDet.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(RstDet("descripcion"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstDet("abrev"))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(NulosN(RstDet("canpro")), "#,###,##0.0000")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(NulosN(RstDet("preunibru")), "0.000000")
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(NulosN(RstDet("preunibruina")), "0.000000")
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(NulosN(NulosN(RstDet("valdes"))), "0.000000")
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(NulosN(RstDet("preuni")), "0.000000")
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(NulosN(RstDet("imptot")), "#,###,##0.0000")
            
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosN(RstDet("iditem"))
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = NulosN(RstDet("idunimed"))
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = NulosN(RstDet("idtipcom"))
            
            '--verificar si hay cuenta en el detalle de la compra
            If NulosN(RstDet("idcue")) <> 0 Then
                Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosN(RstDet("idcue"))
                Fg1.TextMatrix(Fg1.Rows - 1, 13) = NulosN(RstDet("ctadesdeb"))
                Fg1.TextMatrix(Fg1.Rows - 1, 14) = NulosN(RstDet("ctadeshab"))
            Else
                '--asignar cuenta que esta configurada en item de compra
                Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosN(RstDet("idcuenta"))
                Fg1.TextMatrix(Fg1.Rows - 1, 13) = NulosN(RstDet("ctadesdeb1"))
                Fg1.TextMatrix(Fg1.Rows - 1, 14) = NulosN(RstDet("ctadeshab1"))
            End If
            
            Fg1.TextMatrix(Fg1.Rows - 1, 17) = NulosN(RstDet("idnetonodomi"))
            RstDet.MoveNext
            If RstDet.EOF = True Then
                Exit For
            End If
        Next A
    End If
    
    ' CARGAMOS LOS IMPUESTO APLICADOS A LA COMPRA
    BuscarImpuestos
    
    ' CARGAMOS LOS CENTROS DE COSTOS ASIGNADOS A LA COMPRA
    AgregarCentroCosto2 True, RstComp("id")
    
    If NulosN(TxtTipDoc.Text) = 2 Then
        ' recibo por honorarios
        TxtInafecto.Text = Format(NulosN(RstComp("impina")), FORMAT_MONTO)
        TxtBruto.Text = Format(NulosN(RstComp("impbru")), FORMAT_MONTO)
        TxtIGV.Text = Format(NulosN(RstComp("impigv")), FORMAT_MONTO)
        TxtTotal.Text = Format(NulosN(RstComp("imptot")), FORMAT_MONTO)
    Else
        TxtInafecto.Text = Format(NulosN(RstComp("impina")), FORMAT_MONTO)
        TxtBruto.Text = Format(NulosN(RstComp("impbru")), FORMAT_MONTO)
        TxtBruto2.Text = Format(NulosN(RstComp("impbru2")), FORMAT_MONTO)
        TxtBruto3.Text = Format(NulosN(RstComp("impbru3")), FORMAT_MONTO)
        TxtIGV.Text = Format(NulosN(RstComp("impigv")), FORMAT_MONTO)
        TxtIGV2.Text = Format(NulosN(RstComp("impigv2")), FORMAT_MONTO)
        TxtIGV3.Text = Format(NulosN(RstComp("impigv3")), FORMAT_MONTO)
        TxtOtros.Text = Format(NulosN(RstComp("otroscargos")), FORMAT_MONTO)
        TxtISC.Text = Format(NulosN(RstComp("impisc")), FORMAT_MONTO)
        TxtTotal.Text = Format(NulosN(RstComp("imptot")), FORMAT_MONTO)
    End If
    
    Set RstDet = Nothing
    Mostrando = False
        
    ' HALLA LA CUENTA CONTABLE DEL DOCUMENTO DE COMPRA ACTUAL
    xCuentaDoc = HallaNumCuenta(TxtTipDoc.Text, TxtIdMon.Text)
    Set RstDet = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : CargarIngresoAlmacen
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA REGISTRO DE LA TABLA alm_ingresodoc QUE ESTEN VINCULADOS CON LA COMPRA
'*                    ACTUAL, AQUI SE APLICA QUE LOS INGRESO QUE SE REGISTREN ATRAVEZ DE LA TABLA
'*                    alm_ingresodoc SE VINCULAN A UNA O VARIAS COMPRAS
'* Paranetros       : NOMBRE    |  TIPO         |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    IdCompra  |  INTEGER      |  ESPECIFICA EL ID DE LA COMPRA ACTUAL
'* Devuelve         :
'*****************************************************************************************************
Sub CargarIngresoAlmacen(IdCompra As Integer)
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    
    RST_Busq Rst, "SELECT alm_ingresodoc.iddoc, alm_ingreso.tipmov, alm_ingreso.fching, mae_documento.abrev, alm_ingreso.nombre, alm_ingreso.id, [alm_ingreso]![numser]+'-'+[alm_ingreso]![numdoc] AS numdoc " _
        & " FROM (alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id) RIGHT JOIN alm_ingresodoc ON alm_ingreso.id = alm_ingresodoc.id " _
        & " WHERE (((alm_ingresodoc.iddoc)=" & IdCompra & ") AND ((alm_ingreso.tipmov)=-1))", xCon

    Fg4.Rows = 1
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg4.Rows = Fg4.Rows + 1
            Fg4.TextMatrix(A, 1) = Rst("fching")
            Fg4.TextMatrix(A, 2) = NulosC(Rst("abrev"))
            Fg4.TextMatrix(A, 3) = NulosC(Rst("numdoc"))
            Fg4.TextMatrix(A, 4) = NulosC(Rst("nombre"))
            Fg4.TextMatrix(A, 5) = NulosN(Rst("id"))
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Bloquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES TextBox DEL FORMULARIO PARA EL INGRESO O MODIFI
'*                    CACION DE UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Bloquea()
    TxtTipCom.Locked = Not TxtTipCom.Locked
    TxtTipDoc.Locked = Not TxtTipDoc.Locked
    TxtNumRuc.Locked = Not TxtNumRuc.Locked
    TxtNumSer.Locked = Not TxtNumSer.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    TxtFchDoc.Locked = Not TxtFchDoc.Locked
    TxtFchVen.Locked = Not TxtFchVen.Locked
    TxtFchPago.Locked = Not TxtFchPago.Locked
    TxtConPag.Locked = Not TxtConPag.Locked
    TxtIdMon.Locked = Not TxtIdMon.Locked
    TxtIdAlmacen.Locked = Not TxtIdAlmacen.Locked
    TxtGlosa.Locked = Not TxtGlosa.Locked
    TxtTipDocRef.Locked = Not TxtTipDocRef.Locked
    TxtBruto.Locked = Not TxtBruto.Locked
    TxtBruto2.Locked = Not TxtBruto2.Locked
    TxtBruto3.Locked = Not TxtBruto3.Locked
    
    Frame9.Enabled = Not Frame9.Enabled
    CmdAddItem.Enabled = Not CmdAddItem.Enabled
    CmdDelItem.Enabled = Not CmdDelItem.Enabled
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LIMPIA LOS CONTROLES TextBox DEL FORMULARIO PARA EL INGRESO DE UN NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    lblReg.Caption = ""
    TxtTipCom.Text = ""
    TxtTipDoc.Text = ""
    TxtNumRuc.Text = ""
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    TxtFchDoc.Valor = ""
    TxtFchVen.Valor = ""
    TxtFchPago.Valor = ""
    TxtConPag.Text = ""
    TxtIdMon.Text = ""
    TxtIdAlmacen.Text = ""
    TxtGlosa.Text = ""
    TxtDocRef.Text = ""
    LblIdDocRef.Caption = ""
    
    LblIdCenCos.Caption = ""
    LblNomDoc.Caption = ""
    LblNomPro.Caption = ""
    LblCondPag.Caption = ""
    LblMoneda.Caption = ""
    LblIdProveedor.Caption = ""
    LblTipoCompra.Caption = ""
    LblDesAlmacen.Caption = ""
    
    Lbltipo.Caption = ""
    LblIdTipPer.Caption = ""
    TxtTipDocRef.Text = ""
    TxtDocRef2.Text = ""
    LblTipDocref.Caption = ""
    LblIdDocRef2.Caption = ""
    
    TxtBruto.Text = "0.00"
    TxtBruto2.Text = "0.00"
    TxtBruto3.Text = "0.00"
    TxtIGV.Text = "0.00"
    TxtIGV2.Text = "0.00"
    TxtIGV3.Text = "0.00"
    TxtTotal.Text = "0.00"
    TxtISC.Text = "0.00"
    TxtInafecto.Text = "0.00"
    TxtOtros.Text = "0.00"
    LblTipoCambio.Caption = "0.00"
    
    Label3(9).Visible = False
    TxtDocRef.Visible = False
    CmdBusDocRef.Visible = False
    Fg1.Rows = Fg1.FixedRows
    Fg4.Rows = 1 '--limpiar documentos adjuntos
    Fg5.Rows = 1 '--limpiar centro de costo
End Sub

Private Sub Check1_Click()
    ' CAMBIA EL ANCHO DEL FlexGrid Fg1
    If Check1.Value = 1 Then
        Fg1.ColWidth(1) = 4500 - 2000
        Fg1.ColWidth(15) = 1000
        Fg1.ColWidth(16) = 1000
    Else
        Fg1.ColWidth(1) = 4500
        Fg1.ColWidth(15) = 0
        Fg1.ColWidth(16) = 0
    End If
End Sub

Private Sub ChkAjusta_Click()
    ' ACTIVA O DESACTIVA LOS TOTALES PARA PODER HACER UNA MODIFICACION EN LOS VALORES
    If ChkAjusta.Value = 1 Then
        TxtBruto.Locked = False
        TxtInafecto.Locked = False
        TxtIGV.Locked = False
        TxtISC.Locked = False
        TxtTotal.Locked = False
    Else
        TxtBruto.Locked = True
        TxtInafecto.Locked = True
        TxtIGV.Locked = True
        TxtISC.Locked = True
        TxtTotal.Locked = True
    End If
End Sub

Private Sub ChkImpRen4_Click()
    BuscarImpuestos
End Sub

Private Sub CmdAcep_Click()
    TabOne1.Enabled = True
    Toolbar1.Enabled = True
    
    Frame11.Visible = False
End Sub

Private Sub CmdAceptar_Click()
    ' ACEPTA EL INGRESO DE LOS CENTROS DE COSTO ASIGNADOS A LA COMPRA
    If QueHace = 3 Then
        ActivarEntorno
        Frame6.Visible = False
        Exit Sub
    End If
    
    Dim xTot As Double
    If NulosN(TxtInafecto.Text) >= 0 Then
        xTot = NulosN(TxtBruto.Text) + NulosN(TxtBruto2.Text) + NulosN(TxtBruto3.Text) + NulosN(TxtInafecto.Text)
    Else
        If NulosN(TxtInafecto.Text) < 0 Then
            xTot = NulosN(TxtBruto.Text) + NulosN(TxtBruto2.Text) + NulosN(TxtBruto3.Text) + Val(TxtInafecto.Text)
        Else
            xTot = Val(TxtBruto.Text) + NulosN(TxtBruto2.Text) + NulosN(TxtBruto3.Text)
        End If
    End If
    
    ' SI LA DISTRIBUCION DEL CENTRO DE COSTO NO COINCIDE CON EL IMPORTE BRUTO DEL DOCUMENTO NO PERMITE SALIR DEL FRAME DE INGRESO
    If NulosN(Format(xTot, "0.00")) <> NulosN(Format(TxtTotImp.Text, "0.00")) Then
        MsgBox "la distribucion del centro de costo no coincide con el importe del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    LblIdCenCos.Caption = ""
    DetCenCos = True
    Frame6.Visible = False
    ActivarEntorno
End Sub

Private Sub CmdAddCenCos_Click()
    ' AGREGA UN CENTRO DE COSTO
    If QueHace = 3 Then Exit Sub
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    Dim xfrm As New SGI2_funciones.formularios
    Set Rst = xfrm.SeleCentroCosto(xCon)
    
    If Rst.State = 1 Then
        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst
            For A = 1 To Rst.RecordCount
                Fg5.Rows = Fg5.Rows + 1
                Fg5.TextMatrix(Fg5.Rows - 1, 1) = Rst("codigo")
                Fg5.TextMatrix(Fg5.Rows - 1, 2) = Rst("descripcion")
                Fg5.TextMatrix(Fg5.Rows - 1, 5) = Rst("idcencos")
                Rst.MoveNext
                If Rst.EOF = True Then Exit For
            Next A
        End If
    End If
    Set xfrm = Nothing
End Sub

Private Sub CmdAddItem_Click()
    ' AGREGA UNA FILA EN BLANCO AL FlexGrid Fg1 PARA PODER AGREGAR UN ITEM MAS A LA COMPRA, POR DEFECTO INVOCA AL EVENTO
    ' CellButtonClick DE fg1 PARA INICIAR LA BUSQUEDA DEL ITEM
    If QueHace = 3 Then Exit Sub
    
    If PermitirEdicion = False Then Exit Sub
    
    If NulosC(Fg1.TextMatrix(Fg1.Rows - 1, 1)) = "" Then Exit Sub
    Fg1.Rows = Fg1.Rows + 1
    
    '--agregando cantidad por defecto a 1 cuando es servcio
    If NulosN(TxtTipCom.Text) = 5 Then
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = 1
    End If
    Fg1.TextMatrix(Fg1.Rows - 1, 4) = 0
    Fg1.TextMatrix(Fg1.Rows - 1, 5) = 0
    '--
    With Fg1
        .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, 1
    End With
    
    fg1_CellButtonClick Fg1.Rows - 1, 1
    Fg1.SetFocus
End Sub

Private Sub CmdApertura_Click()
    AperturaDocumento xCon, xIdUsuario, 1, IdMenuActivo
    ' refrescar la consulta
    RstComp.Filter = ""
    TDB_FiltroLimpiar Dg1
    RstComp.Requery
End Sub

Private Sub CmdBusAlm_Click()
    ' BUSCA EL ALMACEN QUE SE LE ASIGNARA EL INGRESO DE LA COMPRA
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT alm_almacenes.* FROM alm_almacenes"
    
    xform.Titulo = "Buscando Almacenes"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        LblDesAlmacen.Caption = xRs("descripcion")
        TxtIdAlmacen.Text = xRs("id")
        TxtNumRuc.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusCondicion_Click()
    ' BUSCA LA CONDICION DE PAGO DEL DOCUMENTO
    If QueHace = 3 Then Exit Sub
    If PermitirEdicion = False Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_condpago ORDER BY descripcion"
    
    xform.Titulo = "Buscando Condicion de Pago"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtConPag.Text = xRs("id")
            LblCondPag.Caption = xRs("descripcion")
            
            TxtFchVen.Valor = CDate(TxtFchDoc.Valor) + xRs("numdia")
            TxtFchVen.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusDocRef_Click()
    ' SOLO SE APLICA CUANDO TxtTipDoc SEA IGUAL A 7
    ' BUSCA EL DOCUMENTO DE COMPRA AL QUE HACE REFERENCIA LA NOTA DE CREDITO
    If QueHace = 3 Then Exit Sub

    If NulosN(LblIdProveedor.Caption) = 0 Then
        MsgBox "No ha especificado el proveedor para referenciar este documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(6, 4) As String
    
    xCampos(0, 0) = "Tipo. Doc.":       xCampos(0, 1) = "abrev":                xCampos(0, 2) = "1000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Fch. Doc.":        xCampos(1, 1) = "fchdoc":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Nº Documento":     xCampos(2, 1) = "numdoc":               xCampos(2, 2) = "1500":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Fch. Ven.":        xCampos(3, 1) = "fchven":               xCampos(3, 2) = "1000":         xCampos(3, 3) = "C"
    xCampos(4, 0) = "Total":            xCampos(4, 1) = "imptot":               xCampos(4, 2) = "1000":         xCampos(4, 3) = "N"
    xCampos(5, 0) = "Condicion":        xCampos(5, 1) = "descripcion":          xCampos(5, 2) = "1000":         xCampos(5, 3) = "C"
    
    xform.SQLCad = "SELECT mae_documento.abrev, com_compras.fchdoc, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, com_compras.fchven," _
        & " mae_prov.nombre, mae_condpago.descripcion, com_compras.id, com_compras.imptot FROM mae_condpago LEFT JOIN (mae_documento RIGHT JOIN " _
        & " (mae_prov RIGHT JOIN com_compras ON mae_prov.id = com_compras.idpro) ON mae_documento.id = com_compras.tipdoc) ON mae_condpago.id = com_compras.idconpag " _
        & " WHERE (((com_compras.idpro)=" & NulosN(LblIdProveedor.Caption) & ") AND ((com_compras.tipdoc)<>7))"
    
    xform.Titulo = "Buscando Documentos del Proveedor"
    xform.FormaBusca = CualquierParte
    xform.Criterio = ""
    xform.Ordenado = "numdoc"
    xform.CampoBusca = "numdoc"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtDocRef.Text = xRs("numdoc")
            LblIdDocRef.Caption = xRs("id")
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusDocRef2_Click()
    ' BUSCAMOS EL DOCUMENTO DE REFERENCIA AL QUE HACE LA COMPRA, PREVIAMENTE SE DEBE DE HABER INGRESADO UN VALOR EN TxtTipDocRef
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(4, 4) As String
    
    If NulosN(TxtTipDocRef.Text) = 0 Then
        MsgBox "Falta especiticar el tipo de documento de referencia", vbInformation, xTitulo
        TxtTipDocRef.SetFocus
        Exit Sub
    End If
    
    PermitirEdicion = True
        
    If NulosN(TxtTipDocRef.Text) = 1 Then
        'Orden de Compra
        PermitirEdicion = False
        xCampos(0, 0) = "Documento":     xCampos(0, 1) = "descripcion":  xCampos(0, 2) = "1500":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Nº Documento":  xCampos(1, 1) = "numdoc2":      xCampos(1, 2) = "1400":         xCampos(1, 3) = "C"
        xCampos(2, 0) = "Fch. Emi.":     xCampos(2, 1) = "fchemi":       xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
        xCampos(3, 0) = "Proveedor":     xCampos(3, 1) = "nombre":       xCampos(3, 2) = "4000":         xCampos(3, 3) = "C"
        
        xform.SQLCad = "SELECT com_ordencompra.id, mae_documento.descripcion, [com_ordencompra]![numser] & '-' & [com_ordencompra]![numdoc] AS numdoc2, " _
            & " com_ordencompra.fchemi, com_ordencompra.fchent, mae_prov.nombre FROM (com_ordencompra LEFT JOIN mae_documento " _
            & " ON com_ordencompra.idtipdoc = mae_documento.id) LEFT JOIN mae_prov ON com_ordencompra.idpro = mae_prov.id Where (((com_ordencompra.idest) = 2)) " _
            & " ORDER BY [com_ordencompra]![numser] & '-' & [com_ordencompra]![numdoc]"
        
        xform.Titulo = "Orden de Compra"
        xform.Ordenado = "numdoc2"
        xform.CampoBusca = "numdoc2"
    End If
    If NulosN(TxtTipDocRef.Text) = 2 Then
        'Orden de Produccion
        MsgBox "Opcion no disponible"
        xform.Titulo = "Orden de Produccion"
        Exit Sub
    End If
    If NulosN(TxtTipDocRef.Text) = 3 Then
        'Orden de Mantenimiento
        xform.Titulo = "Orden de Matenimiento"
        MsgBox "Opcion no disponible"
        Exit Sub
    End If
    
    If NulosN(TxtTipDocRef.Text) = 4 Then
        xCampos(0, 0) = "Nº Documento":      xCampos(0, 1) = "numdoc":      xCampos(0, 2) = "1500":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Fch. Emi.":         xCampos(1, 1) = "fchemi":      xCampos(1, 2) = "1000":         xCampos(1, 3) = "C"
        xCampos(2, 0) = "Fch. Ven.":         xCampos(2, 1) = "fchven":      xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
        xCampos(3, 0) = "Proveedor":         xCampos(3, 1) = "nombre":      xCampos(3, 2) = "4000":         xCampos(3, 3) = "C"
        
        'Orden de Despacho
        xCampos(3, 0) = "Cliente"
        xform.SQLCad = "SELECT var_ordendespacho.id, var_ordendespacho.numerodoc AS numdoc,mae_cliente.nombre, var_ordendespacho.idcli, var_ordendespacho.fchemi, var_ordendespacho.fchven  " _
            & " FROM var_ordendespacho LEFT JOIN mae_cliente ON var_ordendespacho.idcli = mae_cliente.id "
        
        xform.Titulo = "Buscando Orden de Despacho"
        
        xform.FormaBusca = CualquierParte
        xform.Ordenado = "numdoc"
        xform.CampoBusca = "numdoc"
        
    End If
    
    If NulosN(TxtTipDocRef.Text) = 6 Then
        xCampos(0, 0) = "Nº Documento":      xCampos(0, 1) = "numdoc":      xCampos(0, 2) = "1500":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Fch. Emi.":         xCampos(1, 1) = "fchemi":      xCampos(1, 2) = "1000":         xCampos(1, 3) = "C"
        xCampos(2, 0) = "Fch. Ven.":         xCampos(2, 1) = "fchent":      xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
        xCampos(3, 0) = "Solicitante":       xCampos(3, 1) = "solicitante": xCampos(3, 2) = "4000":         xCampos(3, 3) = "C"
        
        'Orden de Requerimiento
        xform.SQLCad = "SELECT com_ordenreq.id, com_ordenreq.idtipdoc, [com_ordenreq]![numser] & '-' & [com_ordenreq]![numdoc] AS numdoc, " _
            & " com_ordenreq.fchemi, com_ordenreq.fchent, pla_empleados.nombre AS solicitante " _
            & " FROM (com_ordenreq LEFT JOIN com_usuario ON com_ordenreq.idsol = com_usuario.id) LEFT JOIN pla_empleados ON com_usuario.idper = pla_empleados.id " _
            & " ORDER BY [com_ordenreq]![numser] & '-' & [com_ordenreq]![numdoc]"

        xform.Titulo = "Orden de Requerimiento"
        xform.Ordenado = "numdoc"
        xform.CampoBusca = "numdoc"
    End If
    
    xform.FormaBusca = Principio
    xform.Criterio = ""
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            If NulosN(TxtTipDocRef.Text) = 1 Then
                ' ORDEN DE COMPRA
                TxtDocRef2.Text = NulosC(xRs("numdoc2"))
                LblIdDocRef2.Caption = xRs("id")
                CargarOrdenCompra xRs("id")
            End If
            If NulosN(TxtTipDocRef.Text) = 4 Then
                ' ORDEN DE DESPACHO
                TxtDocRef2.Text = NulosC(xRs("numdoc"))
                LblIdDocRef2.Caption = xRs("id")
            End If
            
            If NulosN(TxtTipDocRef.Text) = 6 Then
                ' ORDEN DE REQUERIMIENTO
                TxtDocRef2.Text = NulosC(xRs("numdoc"))
                LblIdDocRef2.Caption = xRs("id")
            End If
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Sub CargarOrdenCompra(IdOrdCompra As Integer)
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    
    RST_Busq xRs, "SELECT * FROM com_ordencompra WHERE id = " & IdOrdCompra & "", xCon
    If xRs.RecordCount <> 0 Then
        TxtNumRuc.Text = Busca_Codigo(xRs("idpro"), "id", "numruc", "mae_prov", "N", xCon)
        LblNomPro.Caption = Busca_Codigo(xRs("idpro"), "id", "nombre", "mae_prov", "N", xCon)
        LblIdProveedor.Caption = xRs("idpro")
        
        TxtConPag.Text = xRs("idconpag")
        TxtConPag_Validate True
        
        TxtIdMon.Text = xRs("idmon")
        TxtIdMon_Validate True
        
        TxtTipDoc.Text = "1"
        TxtTipDoc_Validate True
        
        Set xRs = Nothing
        
        RST_Busq xRs, "SELECT com_ordencompradet.*, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.idcuenta, alm_inventario.idtipcom, " _
            & " con_planctas.ctadesdeb, con_planctas.ctadeshab FROM ((com_ordencompradet LEFT JOIN alm_inventario ON com_ordencompradet.iditem = alm_inventario.id) " _
            & " LEFT JOIN mae_unidades ON com_ordencompradet.idunimed = mae_unidades.id) LEFT JOIN con_planctas ON alm_inventario.idcuenta = con_planctas.id " _
            & " WHERE (((com_ordencompradet.idcom)=1))", xCon

        If xRs.RecordCount <> 0 Then
            xRs.MoveFirst
            For A = 1 To xRs.RecordCount
                Fg1.Rows = Fg1.Rows + 1
                Fg1.TextMatrix(A, 1) = NulosC(xRs("descripcion"))
                Fg1.TextMatrix(A, 2) = NulosC(xRs("abrev"))
                Fg1.TextMatrix(A, 3) = Format(xRs("canpro"), "0.00")
                Fg1.TextMatrix(A, 4) = Format(xRs("preuni"), "0.00")
                Fg1.TextMatrix(A, 8) = Format(xRs("preuni") * xRs("canpro"), "0.00")
                Fg1.TextMatrix(A, 9) = xRs("iditem")
                Fg1.TextMatrix(A, 10) = xRs("idunimed")
                Fg1.TextMatrix(A, 11) = xRs("idcuenta")
                Fg1.TextMatrix(A, 12) = xRs("idtipcom")
                Fg1.TextMatrix(A, 13) = NulosC(xRs("ctadesdeb"))
                Fg1.TextMatrix(A, 14) = NulosC(xRs("ctadeshab"))
                xRs.MoveNext
                If xRs.EOF = True Then
                    Exit For
                End If
            Next A
            
            xRs.MoveFirst
            For A = 1 To xRs.RecordCount
                Fg5.Rows = Fg5.Rows + 1
                
                Fg5.TextMatrix(Fg5.Rows - 1, 1) = Busca_Codigo(xRs("idcencos"), "id", "codigo", "con_centrocosto", "N", xCon)
                Fg5.TextMatrix(Fg5.Rows - 1, 2) = Busca_Codigo(xRs("idcencos"), "id", "descripcion", "con_centrocosto", "N", xCon) 'NulosC(Rst("descripcion"))
                Fg5.TextMatrix(Fg5.Rows - 1, 4) = Format(xRs("impuni") * xRs("cantidad"), "0.00")
                Fg5.TextMatrix(Fg5.Rows - 1, 3) = "100.00"
                Fg5.TextMatrix(Fg5.Rows - 1, 5) = NulosN(xRs("idcencos"))
                xRs.MoveNext
                If xRs.EOF = True Then Exit For
            Next A
            
            BuscarImpuestos
            HallarTotal
        End If
    End If
End Sub

Private Sub CmdBusProv_Click()
    ' BUSCA AL PROVEEDOR QUE SE LE ESTA ASIGNADO LA COMPRA
    If QueHace = 3 Then Exit Sub
    If PermitirEdicion = False Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Proveedor":    xCampos(0, 1) = "nombre":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":    xCampos(1, 1) = "numruc":      xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_prov.nombre, mae_prov.numruc, mae_prov.id, mae_prov.idcondpag, mae_prov.tipper, mae_tipoempresa.descripcion" _
        & " FROM mae_tipoempresa RIGHT JOIN mae_prov ON mae_tipoempresa.id = mae_prov.tipper WHERE (((mae_prov.activo)=-1)) and mae_prov.id <> 0 "
    
    xform.Titulo = "Buscando Proveedor"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtNumRuc.Text = xRs("numruc")
            LblNomPro.Caption = xRs("nombre")
            LblIdProveedor.Caption = xRs("id")
            
            Lbltipo.Caption = xRs("descripcion")
            LblIdTipPer.Caption = xRs("tipper")
            
            If xRs("idcondpag") <> 0 Then
                TxtConPag.Text = xRs("idcondpag")
                TxtConPag_Validate True
            End If
            TxtFchDoc.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMon_Click()
    ' BUSCA LA MONEDA PARA LA COMPRA
    If QueHace = 3 Then Exit Sub
    If PermitirEdicion = False Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_moneda ORDER BY descripcion"
    
    xform.Titulo = "Buscando Moneda"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdMon.Text = xRs("id")
            LblMoneda.Caption = xRs("descripcion")
            'Fg1.SetFocus
            TxtTipDoc.SetFocus
            
            If Trim(TxtIdMon.Text) = "1" Then
                'LblTipCam.Visible = False
                'LblTipoCambio.Visible = False
            Else
                ' SI LA MONEDA ES DIFERENTE A 1 SE MOSTRARA EL TIPO DE CAMBIO DE LA MONEDA
                If TxtFchDoc.Valor = "" Then
                    MsgBox "No ha especificado la fecha del documento, no se puede determinar " & Chr(13) _
                        & "la fecha del tipo de cambio para este documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    
                    TxtIdMon.Text = ""
                    TxtFchDoc.SetFocus
                    Exit Sub
                End If
                'LblTipCam.Visible = True
                'LblTipoCambio.Visible = True
                LblTipoCambio.Caption = HallaTipoCambio(TxtFchDoc.Valor, 2, Venta, xCon)
                
                If NulosN(LblTipoCambio.Caption) = 0 Then
                    MsgBox "No se ha especificado el tipo de cambio para el dia " & NulosC(TxtFchDoc.Valor), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    TxtIdMon.Text = ""
                    LblMoneda.Caption = ""
                    Exit Sub
                End If
            End If
            'ACTUALIAMOS LA CUENTA CONTABLE DEL DOCUMENTO ASIGNADO A LA COMPRA EN FUNCION A LA MONEDA SELECCIONADA
            xCuentaDoc = HallaNumCuenta(TxtTipDoc.Text, TxtIdMon.Text)
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipDoc_Click()
    'BUSCA EL TIPO DE DOCUMENTO QUE SE LE ASIGNA A LA COMPRA
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_documento.*, mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, con_planctas.cuenta, " _
        & " mae_impuestos.Abrev AS abreimp, mae_impuestos.idcuen  as cuentaimp" _
        & " FROM mae_documento LEFT JOIN (mae_impuestos LEFT JOIN con_planctas ON mae_impuestos.idcuen = con_planctas.id) " _
        & " ON mae_documento.idimp = mae_impuestos.id"
    
    Dim xImpuesto As Double
    
    xform.Titulo = "Buscando Tipo de Documento"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            CodSunatDoc = xRs("codsun")
            TxtTipDoc.Text = xRs("id")
            LblNomDoc.Caption = NulosC(xRs("descripcion"))
            TasaImpuesto = NulosN(xRs("tasa"))
            xDescImp = NulosC(xRs("descripcion"))
            xIdCuenTasa = NulosN(xRs("cuentaimp"))
            LblRotulo = Trim(NulosC(xRs("abreimp"))) + " (       )"
            LblIgvTasa.Caption = Trim(Str(TasaImpuesto)) '+ "%"
            xPorIgv = (TasaImpuesto / 100)
            
            Frame3.Caption = "( Afecta : " + NulosC(xRs("descimp")) + ")"
            xCuentaDoc = HallaNumCuenta(TxtTipDoc.Text, TxtIdMon.Text)
            If xCuentaDoc = 0 Then
                TxtTipDoc.Text = ""
                LblNomDoc.Caption = ""
            End If
            
            If NulosN(TxtTipDoc.Text) = 2 Then
                Frame8.Visible = True
            Else
                Frame8.Visible = False
            End If
            
            If NulosN(TxtTipDoc.Text) = 7 Then
                Label3(9).Visible = True
                TxtDocRef.Visible = True
                CmdBusDocRef.Visible = True
            Else
                Label3(9).Visible = False
                TxtDocRef.Visible = False
                CmdBusDocRef.Visible = False
                
                LblIdDocRef.Caption = ""
                TxtDocRef.Text = ""
            End If
            
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipDocRef_Click()
    ' BUSCA EL TIPO DE DOCUMENTO DE REFERENCIA QUE SE LE ESTA ASIGNANDO A LA COMPRA
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_docreferencia ORDER BY descripcion"
    
    xform.Titulo = "Buscando Tipo de Documento de Referencia"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtTipDocRef.Text = xRs("id")
            LblTipDocref.Caption = xRs("descripcion")
            TxtDocRef2.Text = ""
            LblIdDocRef2.Caption = ""
            TxtDocRef2.SetFocus
            SeSeleTipDocRef = True
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipoCompra_Click()
    ' BUSCA EL TIPO DE PRODUCTO QUE SE ESTA COMPRANDO
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_tipoproducto.* FROM mae_tipoproducto"
    
    xform.Titulo = "Buscando Tipo de Producto"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtTipCom.Text = xRs("id")
            LblTipoCompra = xRs("descripcion")
            TxtIdMon.SetFocus
        End If
        
        If NulosN(TxtTipCom.Text) = 5 Then
            TxtIdAlmacen.Visible = False
            LblDesAlmacen.Visible = False
            CmdBusAlm.Visible = False
            Label3(11).Visible = False
            TxtIdAlmacen.Text = ""
            LblDesAlmacen.Caption = ""
        Else
            TxtIdAlmacen.Visible = True
            LblDesAlmacen.Visible = True
            CmdBusAlm.Visible = True
            Label3(11).Visible = True
        End If
        pGridConfigurar
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdCancelar_Click()
    ActivarEntorno
    DetCenCos = False
    Frame6.Visible = False
End Sub

Private Sub CmdCargaDoc_Click()
    If OptOpera3.Value = True Then
        PopupMenu Opciones
        
    End If
    If OptOpera2.Value = True Then
        'AdjuntarOrdenCompra
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : AdjuntarEntradas
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS DOCUMENTOS INGRESADOS POR EL METODO IngresoAlmacen, ESTO CON EL FIN DE
'*                    HACER UN AMARRE ENTRE LA FACTURA DE COMPRA Y LOS INGRESOS A ALMACEN
'* Paranetros       : NOMBRE    |  TIPO            |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Tipo      |  INTEGER         |  ESPECIFICA LOS DOCUENTOS DE INGRESO QUE MOSTRARA
'*                              |                  |  (TIPO = 1 MUESTRA LAS ENTRADAS NO PROCESADAS)
'*                              |                  |  (TIPO = 2 MUESTRA LAS ENTRADAS PROCESADAS)
'*                    Opcion    |  INTEGER         |  especifica si se mostrara los documentos del
'*                                                    proveedor o de otros proveedores
'*                                                    opcion = 1 muestra las entradas del proveedor
'*                                                    opcion = 2 muestra las entradas de otros proveedores
'* Devuelve         :
'*****************************************************************************************************
Sub AdjuntarEntradas(Tipo As Integer, Opcion As Integer)
    Dim xfrm As New eps_librerias.FormSeleccion
    Dim xCampos(5, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim xRs1 As New ADODB.Recordset
    Dim cSQL As String
    
    xCampos(0, 0) = "T.D.":            xCampos(0, 1) = "abrev":         xCampos(0, 2) = "700":   xCampos(0, 3) = "C":     xCampos(0, 4) = "N"
    xCampos(1, 0) = "Nº Documento":    xCampos(1, 1) = "numdoc":        xCampos(1, 2) = "1500":   xCampos(1, 3) = "C":     xCampos(1, 4) = "S"
    xCampos(2, 0) = "Fch. Giro":       xCampos(2, 1) = "fchdoc":        xCampos(2, 2) = "1000":   xCampos(2, 3) = "F":     xCampos(2, 4) = "F"
    xCampos(3, 0) = "Proveedor":       xCampos(3, 1) = "nombre":        xCampos(3, 2) = "2500":   xCampos(3, 3) = "C":     xCampos(3, 4) = "N"
    xCampos(4, 0) = "Doc. Ref.":       xCampos(4, 1) = "desdocref":     xCampos(4, 2) = "4300":   xCampos(4, 3) = "C":     xCampos(4, 4) = "N"

    If Tipo = 1 Then
        ' CARGAMOS LAS ENTRADAS NO PROCESADAS
        If Opcion = 1 Then
            cSQL = "SELECT 0 as xsel, alm_ingreso.fchdoc, mae_documento.abrev, alm_ingreso.nombre, alm_ingreso!numser+'-'+alm_ingreso!numdoc AS numdoc, alm_ingreso.id, (SELECT Count(1) AS numdocs From alm_ingresodoc WHERE (((alm_ingresodoc.id)=alm_ingreso.id))) AS numdocs, alm_ingreso.desdocref " _
                + vbCr + "FROM alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id " _
                + vbCr + "WHERE ((((SELECT Count(1) AS numdocs From alm_ingresodoc  WHERE (((alm_ingresodoc.id)=alm_ingreso.id))))=0) AND ((alm_ingreso.idpro)=" & NulosN(LblIdProveedor.Caption) & ") AND ((alm_ingreso.tipmov)=-1)) " _
                + vbCr + "ORDER BY alm_ingreso!numser+'-'+alm_ingreso!numdoc"
                
        Else
            cSQL = "SELECT 0 as xsel, alm_ingreso.fchdoc, mae_documento.abrev, alm_ingreso.nombre, alm_ingreso!numser+'-'+alm_ingreso!numdoc AS numdoc, alm_ingreso.id, (SELECT Count(1) AS numdocs From alm_ingresodoc WHERE (((alm_ingresodoc.id)=alm_ingreso.id))) AS numdocs, alm_ingreso.desdocref " _
                + vbCr + "FROM alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id " _
                + vbCr + "WHERE ((((SELECT Count(1) AS numdocs FROM alm_ingresodoc WHERE (((alm_ingresodoc.id)=alm_ingreso.id))))=0) AND ((alm_ingreso.idpro)<>" & NulosN(LblIdProveedor.Caption) & ") AND ((alm_ingreso.tipmov)=-1)) " _
                + vbCr + "ORDER BY alm_ingreso!numser+'-'+alm_ingreso!numdoc"
        End If
    Else
        ' CARGAMOS LAS ENTRADAS PROCESADAS
        If Opcion = 1 Then
            cSQL = "SELECT 0 as xsel, alm_ingreso.fchdoc, mae_documento.abrev, alm_ingreso.nombre, alm_ingreso!numser+'-'+alm_ingreso!numdoc AS numdoc, alm_ingreso.id, (SELECT Count(1) AS numdocs From alm_ingresodoc WHERE (((alm_ingresodoc.id)=alm_ingreso.id))) AS numdocs, alm_ingreso.desdocref " _
                + vbCr + "FROM alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id " _
                + vbCr + "WHERE ((((SELECT Count(1) AS numdocs From alm_ingresodoc  WHERE (((alm_ingresodoc.id)=alm_ingreso.id))))<>0) AND ((alm_ingreso.idpro)=" & NulosN(LblIdProveedor.Caption) & ")) " _
                + vbCr + "ORDER BY alm_ingreso.fchdoc"
        Else
            cSQL = "SELECT 0 as xsel, alm_ingreso.fchdoc, mae_documento.abrev, alm_ingreso.nombre, alm_ingreso!numser+'-'+alm_ingreso!numdoc AS numdoc, alm_ingreso.id, (SELECT Count(1) AS numdocs From alm_ingresodoc WHERE (((alm_ingresodoc.id)=alm_ingreso.id))) AS numdocs, alm_ingreso.desdocref " _
                + vbCr + "FROM alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id " _
                + vbCr + "WHERE ((((SELECT Count(1) AS numdocs From alm_ingresodoc WHERE (((alm_ingresodoc.id)=alm_ingreso.id))))<>0) AND ((alm_ingreso.idpro)<>" & NulosN(LblIdProveedor.Caption) & ")) " _
                + vbCr + "ORDER BY alm_ingreso.fchdoc"
        End If
    End If
        
    xfrm.SQLCad = cSQL
    xfrm.Titulo = "Buscando Entradas a Almacen"
    
    Set xfrm.Coneccion = xCon
    Set xRs = xfrm.Seleccionar(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount = 0 Then
            Set xRs = Nothing
            Exit Sub
        End If
        Dim xCadWHERE As String
        Dim A As Integer
        Dim Rst As New ADODB.Recordset
        
        xRs.MoveFirst
        ' CARGAMOS LOS DOCUMENTOS ADJUNTOS Y LO MOSTRAMOS EN LA LISTA DE "DOCUMENTOS ADJUNTOS"
        For A = 1 To xRs.RecordCount
            If GRID_BUSCAR_VALOR(Fg4, 5, xRs("id"), False) = -1 Then
                Fg4.Rows = Fg4.Rows + 1
                Fg4.TextMatrix(Fg4.Rows - 1, 1) = NulosC(xRs("fchdoc"))
                Fg4.TextMatrix(Fg4.Rows - 1, 2) = NulosC(xRs("abrev"))
                Fg4.TextMatrix(Fg4.Rows - 1, 3) = NulosC(xRs("numdoc"))
                Fg4.TextMatrix(Fg4.Rows - 1, 4) = NulosC(xRs("nombre"))
                Fg4.TextMatrix(Fg4.Rows - 1, 5) = xRs("id")
                xRs.MoveNext
            End If
            If xRs.EOF = True Then Exit For
        Next A
        ' CARGAMOS LOS ITEMS VINCULADOS A LOS DOCUMENTOS DE INGRESO
        CargarItems
    End If
    Set xfrm = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : CargarItems
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS ITEMS VINCULADOS A LOS DOCUMENTOS DE INGRESO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub CargarItems()
    Dim A As Integer
    Dim xCadWHERE As String
    Dim Rst As New ADODB.Recordset
    Dim cSQL As String
    
    ' ARMAMOS LA CADENA WHERE PARA BUSCAR LOS ITEMS VINCULADOS A LOS DOCUMENTOS DE INGRESO
    For A = 1 To Fg4.Rows - 1
        xCadWHERE = xCadWHERE + "(alm_ingresodet.id = " & Val(Fg4.TextMatrix(A, 5)) & ")"
        If A = Fg4.Rows - 1 Then
            Exit For
        End If
        xCadWHERE = xCadWHERE + " OR "
    Next A
    
    xCadWHERE = "(" + xCadWHERE + ")"
    
    ' CARGAMOS LOS ITEMS VINCULADOS A LOS INGRESOS
    cSQL = "SELECT alm_inventario.codpro, mae_unidades.abrev, alm_inventario.descripcion, Sum(alm_ingresodet.cantidad) AS cantidad, con_planctas.ctadesdeb, con_planctas.ctadeshab, alm_inventario.idcuenta, alm_inventario.iddet, alm_inventario.idtipcom, alm_inventario.id, alm_inventario.idunimed " _
        + vbCr + "FROM con_planctas RIGHT JOIN ((alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) ON con_planctas.id = alm_inventario.idcuenta " _
        + vbCr + "WHERE " + xCadWHERE _
        + vbCr + "GROUP BY alm_inventario.codpro, mae_unidades.abrev, alm_inventario.descripcion, con_planctas.ctadesdeb, con_planctas.ctadeshab, alm_inventario.idcuenta, alm_inventario.iddet, alm_inventario.idtipcom, alm_inventario.id, alm_inventario.idunimed"
        
    RST_Busq Rst, cSQL, xCon
    
    Fg1.Rows = 1
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        Mostrando = True

        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = NulosC(Rst("descripcion"))
            Fg1.TextMatrix(A, 2) = NulosC(Rst("abrev"))
            Fg1.TextMatrix(A, 3) = NulosN(Rst("cantidad"))
            Fg1.TextMatrix(A, 4) = 0
            Fg1.TextMatrix(A, 9) = NulosN(Rst("id"))
            Fg1.TextMatrix(A, 10) = NulosN(Rst("idunimed"))
            Fg1.TextMatrix(A, 11) = NulosN(Rst("idcuenta"))
            Fg1.TextMatrix(A, 12) = NulosN(Rst("idtipcom"))
            Fg1.TextMatrix(A, 13) = NulosN(Rst("ctadesdeb"))
            Fg1.TextMatrix(A, 14) = NulosN(Rst("ctadeshab"))
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
        Mostrando = False
    End If
End Sub

Private Sub CmdDelCenCos_Click()
    ' ELIMINA UN CENTRO DE COSTO
    If Fg5.Rows = 1 Then Exit Sub
    If Fg5.Row < 1 Then Exit Sub
    Fg5.RemoveItem Fg5.Row
    HallarTotCenCos
End Sub

Private Sub CmdDelItem_Click()
    ' ELIMINA UN ITEM DE LA LISTA
    If QueHace = 3 Then Exit Sub
    If PermitirEdicion = False Then Exit Sub
    If Fg1.Rows = 1 Then
        MsgBox "No hay items para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    If Fg1.Row < 1 Then
        MsgBox "Seleccione una fila correcta para eliminar", vbExclamation, xTitulo
        Exit Sub
    End If
    
    Fg1.RemoveItem Fg1.Row
    HallarTotal
    BuscarImpuestos
End Sub

'*****************************************************************************************************
'* Nombre           : ActivarEntorno
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES TabOne1 y ToolBar1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ActivarEntorno()
    TabOne1.Enabled = Not TabOne1.Enabled
    Toolbar1.Enabled = Not Toolbar1.Enabled
End Sub

Private Sub CmdDetCenCos_Click()
    ' MUESTRA LA DISTRIBUCION DE LOS CENTROS DE COSTO
    If ((NulosN(TxtBruto.Text) + NulosN(TxtInafecto.Text)) = 0) Then
        MsgBox "No ha especificado el importe del documento para distribuir el centro de costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    If QueHace = 3 Then
        CmdAddCenCos.Enabled = False
        CmdDelCenCos.Enabled = False
        CmdCancelar.Enabled = False
        Fg5.Editable = flexEDNone
        Fg5.SelectionMode = flexSelectionByRow
    Else
        CmdAddCenCos.Enabled = True
        CmdDelCenCos.Enabled = True
        CmdCancelar.Enabled = True
        Fg5.Editable = flexEDKbdMouse
        Fg5.SelectionMode = flexSelectionFree
    End If
    ActivarEntorno
    TxtTotPor.Text = ""
    TxtTotImp.Text = ""
    Frame6.Left = 1545
    Frame6.Top = 2190
    Frame6.Visible = True
    HallarTotCenCos
End Sub

Private Sub CmdSeleccionar_Click()
    If PermitirEdicion = False Then Exit Sub
    
    ' PERMITE AGREGAR ITEMS POR SELECCION, ES DECIR AGREGA UNO O MAS ITEMS EN UNA SOLA ACCION
    If Trim(CmdSeleccionar.Caption) = "Ver Documentos" Then
        TabOne1.Enabled = False
        Toolbar1.Enabled = False
        
        Frame11.Left = 2280
        Frame11.Top = 2550
        Frame11.Visible = True
        Exit Sub
    End If

    If QueHace = 3 Then Exit Sub
    
    If xOrigen = 0 Then
        If NulosC(TxtTipCom.Text) = "" Then
            MsgBox "No ha especificado el tipo de compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtTipCom.SetFocus
            Exit Sub
        End If
    End If
    
    Dim xCampos(3, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLId As String
    Dim A&
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "4000":         xCampos(0, 3) = "C":    xCampos(0, 4) = "N"
    xCampos(1, 0) = "Uni. Med":       xCampos(1, 1) = "abrev":         xCampos(1, 2) = "1000":         xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
    xCampos(2, 0) = "Codigo":         xCampos(2, 1) = "codpro":        xCampos(2, 2) = "1800":         xCampos(2, 3) = "C":    xCampos(2, 4) = "S"

    '*******************************************************************************************
    nSQLId = GRID_GENERAR_SQL_ID(Fg1, 9, "alm_inventario.id", " NOT IN ", True)
    '*******************************************************************************************
    If xOrigen = 0 Then
        If nSQLId <> "" Then nSQLId = " AND " & nSQLId
        nSQL = "SELECT CONSULTA1.*, CONSULTA2.precio FROM " _
            & " [SELECT 0 as xsel,alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.descripcion AS descuni, mae_unidades.abrev, " _
            & " con_planctas.ctadesdeb, con_planctas.ctadeshab,  alm_inventario.idunimed,  alm_inventario.idcuenta, alm_inventario.idtipcom " _
            & " FROM mae_unidades INNER JOIN (con_planctas RIGHT JOIN alm_inventario ON con_planctas.id = alm_inventario.idcuenta) ON mae_unidades.id = alm_inventario.idunimed " _
            & " Where (((alm_inventario.tippro) = " & NulosN(TxtTipCom.Text) & ")) ORDER BY alm_inventario.descripcion]. AS CONSULTA1 LEFT JOIN " _
            & " [SELECT com_comprasdet.iditem, Min(com_comprasdet.preuni) AS precio From com_comprasdet GROUP BY com_comprasdet.iditem]. AS CONSULTA2 ON CONSULTA1.id = CONSULTA2.iditem ORDER BY CONSULTA1.descripcion"
    Else
        If nSQLId <> "" Then nSQLId = " WHERE " & nSQLId
        nSQL = "SELECT 0 as xsel,alm_inventario.*, mae_unidades.descripcion AS descuni, mae_unidades.abrev, " _
                & " con_planctas.ctadesdeb, con_planctas.ctadeshab FROM con_planctas RIGHT JOIN (mae_unidades INNER JOIN " _
                & " alm_inventario ON mae_unidades.id = alm_inventario.idunimed) ON con_planctas.id = alm_inventario.idcuenta " _
                & " " & nSQLId & " ORDER BY alm_inventario.descripcion"
    End If
    
   '*******************************************************************************************
    CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), "Buscando Productos"
    '*******************************************************************************************
    
    If xRs.State = 1 Then
        Mostrando = True
        If xRs.RecordCount <> 0 Then xRs.MoveFirst
        ' MOSTRAMOS LOS ITEMS SELECCIONADOS
        Do While Not xRs.EOF
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(xRs("descripcion"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = xRs("abrev")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(NulosN(xRs("precio")), "0.0000")
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = xRs("id")
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = NulosN(xRs("idunimed"))
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosN(xRs("idcuenta"))
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = NulosN(xRs("idtipcom"))
            Fg1.TextMatrix(Fg1.Rows - 1, 13) = NulosN(xRs("ctadesdeb"))
            Fg1.TextMatrix(Fg1.Rows - 1, 14) = NulosN(xRs("ctadeshab"))
           
            If NulosN(TxtTipCom.Text) = 5 Then
                Fg1.TextMatrix(Fg1.Rows - 1, 3) = 1
                Fg1.TextMatrix(Fg1.Rows - 1, 4) = 0
            End If
            
            xRs.MoveNext
        Loop
    End If
    Mostrando = False
    Set xRs = Nothing
End Sub

Private Sub CmdPreHist_Click()
    ' MUESTRA EL PRECIO HISTORICO DE COMPRA DEL ITEM SELECCIONADO
    If Fg1.Rows < 1 Then Exit Sub
    If Fg1.Row < 1 Then
        MsgBox "Seleccione un Registro para ver el Histórico de Precios", vbInformation, xTitulo
        Exit Sub
    End If
    
    Dim xfrm As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    xfrm.PreciosHistoricos xCon, Fg1.TextMatrix(Fg1.Row, 9), True, NulosC(TxtNumRuc.Text)
    Set xfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub CmdVerAsiento_Click()
    VerAsiento
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA LAS COLUMNAS DEL CONTROL Dg1
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstComp.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_DblClick()
    ' MUESTRA INFORMACION EN LA PESTAÑA DETALLE
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstComp
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    ' MUESTRA LAS OPERACIONES EFECTUADAS SOBRE EL REGISTRO SELECCIONADO
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstComp("id")), xCon
    End If
End Sub

Private Sub fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    ' CARGA LA LISTA DE PRODUCCTOS PARA LA SELECCION
    If xOrigen = 0 Then
        If NulosN(TxtTipCom.Text) = 0 Then
            MsgBox "No ha especificado el tipo de compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtTipCom.SetFocus
            Exit Sub
        End If
    End If
    
    If PermitirEdicion = False Then Exit Sub
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim nSQLId As String
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(3, 4) As String
    
    xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5400":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Unid.":        xCampos(1, 1) = "abrev":          xCampos(1, 2) = "600":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Código":       xCampos(2, 1) = "codpro":         xCampos(2, 2) = "2000":    xCampos(2, 3) = "C"
    
    '*******************************************************************************************
    nSQLId = GRID_GENERAR_SQL_ID(Fg1, 9, "alm_inventario.id", " NOT IN ", True)
    '*******************************************************************************************
    
    If xOrigen = 0 Then
        If NulosN(TxtTipDocRef.Text) = 0 Or NulosN(TxtTipDocRef.Text) = 4 Then
            If nSQLId <> "" Then nSQLId = " and " & nSQLId
            xform.SQLCad = "SELECT CONSULTA1.*, CONSULTA2.precio FROM " _
                & " [SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.descripcion AS descuni, mae_unidades.abrev, " _
                & " con_planctas.ctadesdeb, con_planctas.ctadeshab,  alm_inventario.idunimed,  alm_inventario.idcuenta, alm_inventario.idtipcom, alm_inventario.idnetonodomi " _
                & " FROM mae_unidades INNER JOIN (con_planctas RIGHT JOIN alm_inventario ON con_planctas.id = alm_inventario.idcuenta) ON mae_unidades.id = alm_inventario.idunimed " _
                & " WHERE alm_inventario.activo = -1 and (((alm_inventario.tippro)=" & NulosN(TxtTipCom.Text) & ") AND ((alm_inventario.tipo)=1 Or (alm_inventario.tipo)=3)) " _
                & " ORDER BY alm_inventario.descripcion]. AS CONSULTA1 LEFT JOIN " _
                & " [SELECT com_comprasdet.iditem, Min(com_comprasdet.preuni) AS precio From com_comprasdet GROUP BY com_comprasdet.iditem]. AS CONSULTA2 ON CONSULTA1.id = CONSULTA2.iditem ORDER BY CONSULTA1.descripcion"
        Else
            If NulosN(TxtTipDocRef.Text) = 6 Then
                ' SI ES ORDEN DE REQUERIMIENTO
                xform.SQLCad = "SELECT com_ordenreqdet.idor, com_ordenreqdet.iditem AS id, alm_inventario.codpro, alm_inventario.descripcion, " _
                    & " mae_unidades.descripcion AS descuni, mae_unidades.abrev, con_planctas.ctadesdeb, con_planctas.ctadeshab, com_ordenreqdet.idunimed, " _
                    & " con_centrocosto.idctacon AS idcuenta, alm_inventario.idtipcom, alm_inventario.idnetonodomi, 0 AS precio " _
                    & " FROM ((mae_unidades RIGHT JOIN (com_ordenreqdet LEFT JOIN alm_inventario ON com_ordenreqdet.iditem = alm_inventario.id) " _
                    & " ON mae_unidades.id = alm_inventario.idunimed) LEFT JOIN con_centrocosto ON com_ordenreqdet.idcencos = con_centrocosto.id) " _
                    & " LEFT JOIN con_planctas ON con_centrocosto.idctacon = con_planctas.id WHERE (((com_ordenreqdet.idor)=" & LblIdDocRef2.Caption & "))"

            End If
            
        End If
    Else
        If nSQLId <> "" Then nSQLId = " where " & nSQLId
        xform.SQLCad = "SELECT alm_inventario.*, mae_unidades.descripcion AS descuni, mae_unidades.abrev, " _
            & " con_planctas.ctadesdeb, con_planctas.ctadeshab FROM con_planctas RIGHT JOIN (mae_unidades INNER JOIN " _
            & " alm_inventario ON mae_unidades.id = alm_inventario.idunimed) ON con_planctas.id = alm_inventario.idcuenta " _
            & " " & nSQLId & " AND alm_inventario.idcuenta <> 0 ORDER BY alm_inventario.descripcion "
    End If
    
    Dim RstCamBus As New ADODB.Recordset
    RST_Busq RstCamBus, "SELECT var_opcionesformulario.idform, var_opcionesformulario.campobus From var_opcionesformulario " _
        & " WHERE (((var_opcionesformulario.idform)=12))", xCon
    
    xform.Titulo = "Buscando Productos"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    
    If RstCamBus.RecordCount <> 0 Then
        xform.Ordenado = NulosC(RstCamBus("campobus"))
        xform.CampoBusca = NulosC(RstCamBus("campobus"))
    Else
        xform.Ordenado = "codpro"
        xform.CampoBusca = "codpro"
    End If
    Set RstCamBus = Nothing
    
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    Mostrando = True
    Dim A As Integer
    
    If xRs.State = 1 Then
        If Fg1.Rows <> 1 Then
            ' VERIFICAMOS QUE EL ITEM NO HAYA SIDO SELECCIONADO
            For A = 1 To Fg1.Rows - 1
                If Fg1.TextMatrix(A, 9) = xRs("id") Then
                    MsgBox "El item seleccionado ya fue agregado, seleccione otro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    A = Fg1.Rows - 1
                    Set xRs = Nothing
                    Exit Sub
                End If
            Next A
        End If
        
        ' MUESTRA LA INFORMACION DEL ITEM
        If xRs.RecordCount <> 0 Then
            Fg1.TextMatrix(Fg1.Row, 1) = NulosC(xRs("descripcion"))
            Fg1.TextMatrix(Fg1.Row, 2) = NulosC(xRs("abrev"))
            If NulosN(TxtTipCom.Text) <> 5 Then
                Fg1.TextMatrix(Fg1.Row, 4) = NulosN(xRs("precio"))
            Else
                Fg1.TextMatrix(Fg1.Row, 4) = 0
            End If
            Fg1.TextMatrix(Fg1.Row, 9) = NulosN(xRs("id"))
            Fg1.TextMatrix(Fg1.Row, 10) = NulosN(xRs("idunimed"))
            Fg1.TextMatrix(Fg1.Row, 11) = NulosN(xRs("idcuenta"))
            Fg1.TextMatrix(Fg1.Row, 12) = NulosN(xRs("idtipcom"))
            Fg1.TextMatrix(Fg1.Row, 13) = NulosN(xRs("ctadesdeb"))
            Fg1.TextMatrix(Fg1.Row, 14) = NulosN(xRs("ctadeshab"))
            Fg1.TextMatrix(Fg1.Row, 17) = NulosN(xRs("idnetonodomi"))
        End If
    End If
    Mostrando = False
    Set xform = Nothing
    
    If Fg1.Row >= 1 Then
        If NulosN(TxtTipCom.Text) = 5 Then
            Fg1.Col = 4
        Else
            Fg1.Col = 3
        End If
    End If

    Set xRs = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : AgregarCentroCosto2
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : AGREGA CENTRO DE COSTO DEL ITEM A LOS CENTROS DE COSTO YA CARGADOS
'* Paranetros       : NOMBRE        |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    CargarGrabado |  Boolean          |  especifica que se levantara un centro de
'*                                  |                   |  costos que haya sido grabado
'*                    IdCompra      |  Double           |  ESPECITICA EL ID DE LA COMPRA
'* Devuelve         :
'*****************************************************************************************************
Sub AgregarCentroCosto2(CargarGrabado As Boolean, Optional IdCompra As Double)
    'CargarGrabado = especifica que se levantara un centro de costos que haya sido grabado
    Dim Rst As New ADODB.Recordset
    Dim A, B, C, xFila As Integer
    Dim SeEncontro As Boolean
    
    
        
    If CargarGrabado = True Then
        RST_Busq Rst, "SELECT com_comprascosto.idcom, com_comprascosto.idcencos, con_centrocosto.codigo, con_centrocosto.descripcion, com_comprascosto.imppor, com_comprascosto.impcos, " _
            & " con_centrocosto.tipo FROM con_centrocosto INNER JOIN com_comprascosto ON con_centrocosto.id = com_comprascosto.idcencos " _
            & " WHERE (((com_comprascosto.idcom)=" & IdCompra & "))", xCon
            
        If Rst.RecordCount <> 0 Then
            Fg5.Rows = 1
            Rst.MoveFirst
            Mostrando = True
            For A = 1 To Rst.RecordCount
                Fg5.Rows = Fg5.Rows + 1
                Fg5.TextMatrix(Fg5.Rows - 1, 1) = NulosC(Rst("codigo"))
                Fg5.TextMatrix(Fg5.Rows - 1, 2) = NulosC(Rst("descripcion"))
                Fg5.TextMatrix(Fg5.Rows - 1, 3) = Format(NulosN(Rst("imppor")), "0.00")
                Fg5.TextMatrix(Fg5.Rows - 1, 4) = Format(NulosN(Rst("impcos")), "0.00")
                Fg5.TextMatrix(Fg5.Rows - 1, 5) = NulosN(Rst("idcencos"))
                
                Rst.MoveNext
                If Rst.EOF = True Then Exit For
            Next A
            Mostrando = False
        End If
    Else
        If NulosN(TxtTipDocRef.Text) = 1 Or NulosN(TxtTipDocRef.Text) = 6 Then
            
            If NulosN(TxtTipDocRef.Text) = 6 Then
                If Fg1.Rows <> 1 Then
                If NulosN(Fg1.TextMatrix(Fg1.Row, 9)) <> 0 Then
                    ' buscamos el centro de costo en el detalle del documento de referencia
                    Dim xRs As New ADODB.Recordset
                    
                    RST_Busq xRs, "SELECT com_ordenreqdet.idor, com_ordenreqdet.iditem, com_ordenreqdet.idcencos, con_centrocosto.codigo, con_centrocosto.descripcion" _
                        & " FROM com_ordenreqdet LEFT JOIN con_centrocosto ON com_ordenreqdet.idcencos = con_centrocosto.id WHERE (((com_ordenreqdet.idor)=" & NulosN(LblIdDocRef2.Caption) & ") " _
                        & " AND ((com_ordenreqdet.iditem)=" & NulosN(Fg1.TextMatrix(Fg1.Row, 9)) & "))", xCon
    
                    SeEncontro = False
                    xFila = 1
                    For C = 1 To Fg5.Rows - 1
                        If Fg5.TextMatrix(C, 5) = xRs("idcencos") Then
                            SeEncontro = True
                            xFila = C
                        End If
                    Next C
                    If SeEncontro = False Then Fg5.Rows = Fg5.Rows + 1
                    'Fg5.Rows = Fg5.Rows + 1
                    Fg5.TextMatrix(xFila, 1) = NulosC(xRs("codigo"))
                    Fg5.TextMatrix(xFila, 2) = NulosC(xRs("descripcion"))
                    Fg5.TextMatrix(xFila, 3) = "100.00"  'Format(NulosN(Rst("imppor")), "0.00")
                    Fg5.TextMatrix(xFila, 5) = NulosN(xRs("idcencos"))
                    Fg5.TextMatrix(xFila, 4) = Format(NulosN(Fg1.TextMatrix(Fg1.Row, 8)), "0.00")
                End If
                End If
            End If
        Else
            Fg5.Rows = 1
            For A = 1 To Fg1.Rows - 1
                ' buscamos si el item actual tiene centros de costo definido
                RST_Busq Rst, "SELECT alm_invencencos.idpro, alm_invencencos.idcencos, con_centrocosto.codigo, con_centrocosto.descripcion, " _
                & " alm_invencencos.imppor FROM alm_invencencos LEFT JOIN con_centrocosto ON alm_invencencos.idcencos = con_centrocosto.id " _
                & " WHERE (((alm_invencencos.idpro)=" & NulosN(Fg1.TextMatrix(A, 9)) & "))", xCon
                
                If Rst.RecordCount <> 0 Then
                    ' si tiene centro de costos agregamos a la cuadricula centro de costos
                    Rst.MoveFirst
                    For B = 1 To Rst.RecordCount
                        ' buscamos si el cetro de costo ya fue agregado a la cuadricula
                        SeEncontro = False
                        xFila = 0
                        For C = 1 To Fg5.Rows - 1
                            If Fg5.TextMatrix(C, 5) = Rst("idcencos") Then
                                SeEncontro = True
                                xFila = C
                            End If
                        Next C
                        
                        If SeEncontro = True Then
                            ' nos pocisionamos en la fila que contiene el centro de costos y sumamos el valor
                            If Rst("imppor") < 100 Then
                                Fg5.TextMatrix(Fg5.Rows - 1, 4) = NulosN(Fg5.TextMatrix(Fg5.Rows - 1, 4)) + (NulosN(Fg1.TextMatrix(A, 8)) * ((Rst("imppor") / 100) + 1))
                            Else
                                Fg5.TextMatrix(Fg5.Rows - 1, 4) = NulosN(Fg5.TextMatrix(Fg5.Rows - 1, 4)) + NulosN(Fg1.TextMatrix(A, 8))
                            End If
                        Else
                            ' agregamos una nueva fila a la cuadricula centro de costos
                            Fg5.Rows = Fg5.Rows + 1
                            Fg5.TextMatrix(Fg5.Rows - 1, 1) = NulosC(Rst("codigo"))
                            Fg5.TextMatrix(Fg5.Rows - 1, 2) = NulosC(Rst("descripcion"))
                            Fg5.TextMatrix(Fg5.Rows - 1, 3) = Format(NulosN(Rst("imppor")), "0.00")
                            Fg5.TextMatrix(Fg5.Rows - 1, 5) = NulosN(Rst("idcencos"))
                            If NulosN(Rst("imppor")) < 100 Then
                                Fg5.TextMatrix(Fg5.Rows - 1, 4) = NulosN(Fg1.TextMatrix(A, 8)) * ((Rst("imppor") / 100) + 1)
                            Else
                                Fg5.TextMatrix(Fg5.Rows - 1, 4) = NulosN(Fg1.TextMatrix(A, 8))
                            End If
                            Fg5.TextMatrix(Fg5.Rows - 1, 4) = Format(Fg5.TextMatrix(Fg5.Rows - 1, 4), "0.00")
                        End If
                        
                        Rst.MoveNext
                        If Rst.EOF = True Then Exit For
                    Next B
                Else
                    If NulosN(Fg1.TextMatrix(A, 9)) <> 0 Then
                        'MsgBox "El item " & NulosC(Fg1.TextMatrix(A, 1)) & ", no tiene especificado un centro de costos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    End If
                End If
            Next A
                
            If NulosN(TxtBruto.Text) <> 0 Or NulosN(TxtInafecto.Text) <> 0 Then
                For A = 1 To Fg5.Rows - 1
                    Fg5.TextMatrix(A, 3) = (NulosN(Fg5.TextMatrix(A, 4)) / (NulosN(TxtBruto.Text) + NulosN(TxtInafecto.Text))) * 100
                    Fg5.TextMatrix(A, 3) = Format(Fg5.TextMatrix(A, 3), "0.00")
                Next A
            End If
        End If
    End If
    Set Rst = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : AgregarCentroCosto
'* Tipo             : PROCEDIMIENTO
'* Descripcion      :
'* Paranetros       : NOMBRE       |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    xIdProducto  |  Integer          |  ESPECIFICA EL ID DEL ITEM AL QUE SE LE
'*                                 |                   |  AGREGARA EL CENTRO DE COSTOS
'*                    xImporte     |  Double           |  ESPECIFICA EL VALOR DEL ITEM QUE SE LE
'*                                 |                   |  AGREGARA EL CENTRO DE COSTOS
'* Devuelve         :
'*****************************************************************************************************
Sub AgregarCentroCosto(xIdProducto As Integer, xImporte As Double)
    Dim Rst As New ADODB.Recordset
    Dim A, B As Integer
    Dim SeEncontro As Boolean
    
    ' buscamos si el producto tiene centro de costo asignado
    RST_Busq Rst, "SELECT alm_invencencos.idpro, alm_invencencos.idcencos, con_centrocosto.codigo, con_centrocosto.descripcion, " _
        & " alm_invencencos.imppor FROM alm_invencencos LEFT JOIN con_centrocosto ON alm_invencencos.idcencos = con_centrocosto.id " _
        & " WHERE (((alm_invencencos.idpro)=" & xIdProducto & "))", xCon
    
    If Rst.RecordCount <> 0 Then
        For A = 1 To Rst.RecordCount
            For B = 1 To Fg5.Rows - 1
                SeEncontro = False
                If Fg5.TextMatrix(B, 5) = Rst("idcencos") Then
                    SeEncontro = True
                    Exit For
                End If
            Next B
            If SeEncontro = False Then
                ' si no lo encuentra lo debe de agregar a la lista de centro de costos
                Fg5.Rows = Fg5.Rows + 1
                Fg5.TextMatrix(Fg5.Rows - 1, 1) = Rst("codigo")
                Fg5.TextMatrix(Fg5.Rows - 1, 2) = Rst("descripcion")
                Fg5.TextMatrix(Fg5.Rows - 1, 3) = Format(Rst("imppor"), "0.00")
                If Rst("imppor") = 100 Then
                    Fg5.TextMatrix(Fg5.Rows - 1, 4) = xImporte * 1
                Else
                    Fg5.TextMatrix(Fg5.Rows - 1, 4) = xImporte * ((Rst("imppor") / 100) + 1)
                End If
                Fg5.TextMatrix(Fg5.Rows - 1, 4) = Format(Fg5.TextMatrix(Fg5.Rows - 1, 4), "0.00")
                Fg5.TextMatrix(Fg5.Rows - 1, 5) = Rst("idcencos")
            Else
                ' si el centro de costo ya existe, agregarlo al centro de costo ya existente
                MsgBox "Falta hacer esta opcion"
                'Fg5.TextMatrix(Fg5.Rows - 1, 1) = Rst("codigo")
                'Fg5.TextMatrix(Fg5.Rows - 1, 2) = Rst("descripcion")
                'Fg5.TextMatrix(Fg5.Rows - 1, 5) = Rst("idcencos")
            End If
        Next A
    End If
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Mostrando = True Then Exit Sub
    If Row = 0 Then Exit Sub
    If Agregando = True Then Exit Sub
    If Col = 3 Or Col = 4 Or Col = 5 Or Col = 6 Then
        ' verificamos si hay descuento
        ' chequeamo si es por porcentaje
        If OptDes1.Value = True Then
            ' Se esta aplicando descuento por porcentaje
            Dim xPorcen As Double
            If NulosN(Fg1.TextMatrix(Row, 6)) <> 0 Then
                xPorcen = (NulosN(Fg1.TextMatrix(Row, 6)) / 100)
                Fg1.TextMatrix(Row, 7) = (NulosN(Fg1.TextMatrix(Row, 4)) + NulosN(Fg1.TextMatrix(Row, 5))) * xPorcen
                Fg1.TextMatrix(Row, 7) = (NulosN(Fg1.TextMatrix(Row, 4)) + NulosN(Fg1.TextMatrix(Row, 5))) - NulosN(Fg1.TextMatrix(Row, 7))
            Else
                Fg1.TextMatrix(Row, 7) = (NulosN(Fg1.TextMatrix(Row, 4)) + NulosN(Fg1.TextMatrix(Row, 5)))
            End If
        End If
        If OptDes2.Value = True Then
            ' Se esta aplicando descuento por importe
            If NulosN(Fg1.TextMatrix(Row, 6)) <> 0 Then
                Fg1.TextMatrix(Row, 7) = (NulosN(Fg1.TextMatrix(Row, 4)) + NulosN(Fg1.TextMatrix(Row, 5))) - NulosN(Fg1.TextMatrix(Row, 6))
            Else
                Fg1.TextMatrix(Row, 7) = (NulosN(Fg1.TextMatrix(Row, 4)) + NulosN(Fg1.TextMatrix(Row, 5)))
            End If
        End If
        
        Fg1.TextMatrix(Row, 8) = NulosN(Fg1.TextMatrix(Row, 3)) * NulosN(Fg1.TextMatrix(Row, 7))
        
        HallarTotal
        BuscarImpuestos
    End If
    
    If Col = 15 Or Col = 16 Then
        Fg1.TextMatrix(Fg1.Row, 3) = Format(Fg1.TextMatrix(Fg1.Row, 16), "0.0000")
        
        Fg1.TextMatrix(Fg1.Row, 7) = "0.0000"
        Fg1.TextMatrix(Fg1.Row, 15) = Format(Fg1.TextMatrix(Fg1.Row, 15), "0.00")
        If NulosN(Fg1.TextMatrix(Fg1.Row, 16)) = 0 Then
            Fg1.TextMatrix(Fg1.Row, 4) = NulosN(Fg1.TextMatrix(Fg1.Row, 15)) / ((NulosN(LblIgvTasa.Caption) / 100) + 1)
        Else
            Fg1.TextMatrix(Fg1.Row, 4) = (NulosN(Fg1.TextMatrix(Fg1.Row, 15)) / ((NulosN(LblIgvTasa.Caption) / 100) + 1)) / NulosN(Fg1.TextMatrix(Fg1.Row, 16))
        End If
        Fg1.TextMatrix(Fg1.Row, 7) = Fg1.TextMatrix(Fg1.Row, 4)
        Fg1.TextMatrix(Fg1.Row, 8) = NulosN(Fg1.TextMatrix(Fg1.Row, 7)) * NulosN(Fg1.TextMatrix(Fg1.Row, 3))
        BuscarImpuestos
        HallarTotal
    End If
    
    
    Fg1.TextMatrix(Row, 3) = Format(Fg1.TextMatrix(Row, 3), "#,###,##0.0000")
    Fg1.TextMatrix(Row, 4) = Format(Fg1.TextMatrix(Row, 4), "0.000000")
    Fg1.TextMatrix(Row, 5) = Format(Fg1.TextMatrix(Row, 5), "0.000000")
    Fg1.TextMatrix(Row, 6) = Format(Fg1.TextMatrix(Row, 6), "0.000000")
    Fg1.TextMatrix(Row, 7) = Format(Fg1.TextMatrix(Row, 7), "0.000000")
    Fg1.TextMatrix(Row, 8) = Format(Fg1.TextMatrix(Row, 8), "#,###,##0.0000")
    
End Sub

'*****************************************************************************************************
'* Nombre           : BuscarImpuestos
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : BUSCA LOS IMPUESTO AL QUE ESTA AFECTO LOS ITEMS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub BuscarImpuestos()
    TxtIGV.Text = "0.00"
    TxtIGV2.Text = "0.00"
    TxtIGV3.Text = "0.00"
    TxtOtros.Text = "0.00"
    TxtTotal.Text = "0.00"
    If Fg1.Rows = 1 Then Exit Sub
    Dim A As Integer
    Dim xImpSEL, xImpIGV As Double
    
    Dim Rst As New ADODB.Recordset
    
    Set RstTempISC = Nothing
    PreparaRST_ISC
    xImpSEL = 0
    
    ' buscando selectivo
    For A = 1 To Fg1.Rows - 1
        If NulosC(Fg1.TextMatrix(A, 1)) <> "" Then
            RST_Busq Rst, "SELECT mae_impuestos.tasa, mae_impuestos.idcuen, con_planctas.cuenta " _
                & " FROM (alm_inventario LEFT JOIN mae_impuestos ON alm_inventario.idimpsel = mae_impuestos.id) " _
                & " LEFT JOIN con_planctas ON mae_impuestos.idcuen = con_planctas.id WHERE " _
                & " ((alm_inventario.id = " & Val(Fg1.TextMatrix(A, 9)) & " ))", xCon
            
            If Rst.RecordCount <> 0 Then
                If NulosN(Rst("idcuen")) <> 0 Then
                    xImpSEL = xImpSEL + NulosN(Fg1.TextMatrix(A, 5)) * (NulosN(Rst("tasa")) / 100)
                    
                    If RstTempISC.RecordCount = 0 Then
                        RstTempISC.AddNew
                        RstTempISC("idcuen") = NulosN(Rst("idcuen"))
                        RstTempISC("total") = NulosN(RstTempISC("total")) + NulosN(Fg1.TextMatrix(A, 5)) * (NulosN(Rst("tasa")) / 100)
                    Else
                        RstTempISC.MoveFirst
                        RstTempISC.Find "idcuen = " & Rst("idcuen") & ""
                        
                        If RstTempISC.EOF = False Then
                            RstTempISC("idcuen") = NulosN(Rst("idcuen"))
                            RstTempISC("total") = NulosN(RstTempISC("total")) + NulosN(Fg1.TextMatrix(A, 5)) * (NulosN(Rst("tasa")) / 100)
                        End If
                    End If
                End If
            End If
        End If
    Next A
    
    TxtISC.Text = Format(NulosN(xImpSEL), "0.00")
    
    ' buscando el impuesto a las ventas
    If NulosN(LblIdTipPer.Caption) <> 3 Then
        xImpIGV = 0
        
        xImpIGV = NulosN(TxtBruto.Text) * (NulosN(TasaImpuesto) / 100)
        
        If CodSunatDoc = "02" Then
            If ChkImpRen4.Value = 1 Then
                xImpIGV = NulosN(TxtBruto.Text) * (NulosN(TasaImpuesto) / 100)
            Else
                xImpIGV = 0
            End If
        End If
        If NulosN(TxtTipDoc.Text) <> 2 Then
            TxtIGV.Text = Format(xImpIGV, FORMAT_MONTO)
            TxtTotal.Text = NulosN(TxtBruto.Text) + NulosN(TxtInafecto.Text) + NulosN(TxtIGV.Text)
            TxtTotal.Text = Format(TxtTotal.Text, FORMAT_MONTO)
        Else
            TxtIGV.Text = Format(xImpIGV, FORMAT_MONTO)
            TxtTotal.Text = NulosN(TxtBruto.Text) + NulosN(TxtInafecto.Text) - NulosN(TxtIGV.Text)
            TxtTotal.Text = Format(TxtTotal.Text, FORMAT_MONTO)
        End If
    Else
        xImpIGV = 0
        Dim xNeto As Double
        Dim xNeto2 As Double
        
        For A = 1 To Fg1.Rows - 1
            xNeto = NulosN(Busca_Codigo(NulosN(Fg1.TextMatrix(A, 17)), "id", "neto", "mae_netonodomiciliado", "N", xCon))
            If xNeto <> 0 Then xNeto2 = Fg1.TextMatrix(A, 8) * (xNeto / 100)
            xImpIGV = xNeto2 * 0.3
        Next A
    
        TxtOtros.Text = Format(xImpIGV, FORMAT_MONTO)
        TxtTotal.Text = NulosN(TxtBruto.Text) + NulosN(TxtBruto2.Text) + NulosN(TxtBruto3.Text) + NulosN(TxtInafecto.Text) - NulosN(TxtOtros.Text)
        TxtTotal.Text = Format(TxtTotal.Text, FORMAT_MONTO)
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : PreparaRST_ISC
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA UN RECORDSET TEMPORAL PARA ALMACENAR LOS VALORS DEL IMPUESTO SELECTIVO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub PreparaRST_ISC()
    'Dim xFun As New Eps_DataAcces.FuncionesData
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(9, 3) As String

    xCampos(0, 0) = "idcuen":        xCampos(0, 1) = "N":      xCampos(0, 2) = "2"
    xCampos(1, 0) = "Total":         xCampos(1, 1) = "D":      xCampos(1, 2) = "2"
    Set RstTempISC = xFun.CrearRstTMP(xCampos)

    RstTempISC.Open
End Sub

'*****************************************************************************************************
'* Nombre           : HallarTotal
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : HALLA LA SUMA TOTAL DE LOS ITEMS CARGADOS EN EL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub HallarTotal()
    Dim A As Integer
    Dim Total, TotalIna As Double
    Dim xPorcen As Double
    Dim PreDes As Double
    Dim Valor As Double
    Total = 0
    TotalIna = 0
    For A = 1 To Fg1.Rows - 1
        If OptDes1.Value = True Then
            ' Se esta aplicando descuento por porcentaje
            If NulosN(Fg1.TextMatrix(A, 6)) <> 0 Then
                xPorcen = ((NulosN(Fg1.TextMatrix(A, 6)) / 100))
                PreDes = NulosN(Fg1.TextMatrix(A, 4)) - (NulosN(Fg1.TextMatrix(A, 4)) * xPorcen)
                Valor = PreDes * NulosN(Fg1.TextMatrix(A, 3))
                Total = Total + Valor
                
                Valor = (NulosN(Fg1.TextMatrix(A, 5)) / xPorcen) * NulosN(Fg1.TextMatrix(A, 3))
                TotalIna = TotalIna + Valor
            Else
                Valor = NulosN(Fg1.TextMatrix(A, 4)) * NulosN(Fg1.TextMatrix(A, 3))
                Total = Total + Valor
                
                Valor = NulosN(Fg1.TextMatrix(A, 5)) * NulosN(Fg1.TextMatrix(A, 3))
                TotalIna = TotalIna + Valor
            End If
        End If
        If OptDes2.Value = True Then
            ' Se esta aplicando descuento por importe
            If NulosN(Fg1.TextMatrix(A, 6)) <> 0 Then
                If NulosN(Fg1.TextMatrix(A, 4)) <> 0 Then
                    Valor = (NulosN(Fg1.TextMatrix(A, 4)) - NulosN(Fg1.TextMatrix(A, 6))) * NulosN(Fg1.TextMatrix(A, 3))
                    Total = Total + Valor
                End If
                If NulosN(Fg1.TextMatrix(A, 5)) <> 0 Then
                    Valor = (NulosN(Fg1.TextMatrix(A, 5)) - NulosN(Fg1.TextMatrix(A, 6))) * NulosN(Fg1.TextMatrix(A, 3))
                    TotalIna = TotalIna + Valor
                End If
            Else
                Total = Total + (NulosN(Fg1.TextMatrix(A, 4)) * NulosN(Fg1.TextMatrix(A, 3)))
                TotalIna = TotalIna + (NulosN(Fg1.TextMatrix(A, 5)) * NulosN(Fg1.TextMatrix(A, 3)))
            End If
        End If
    Next A

    TxtBruto.Text = Format(Total, FORMAT_MONTO)
    TxtInafecto.Text = Format(TotalIna, FORMAT_MONTO)
    AgregarCentroCosto2 False
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then Exit Sub
    If PermitirEdicion = False Then Exit Sub
    
    If Fg1.Col = 2 Or Fg1.Col = 7 Or Fg1.Col = 8 Then
        Fg1.Editable = flexEDNone
    Else
        Fg1.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = 3 Or Col = 4 Or Col = 5 Or Col = 6 Or Col = 15 Or Col = 16 Then
        If InStr(CaracteresNumericos2, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = 45 Then
        CmdAddItem_Click
    End If
    If KeyCode = 46 Then
        CmdDelItem_Click
    End If
    If KeyCode = 93 Then
        PopupMenu menu1
    End If
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If QueHace = 3 Then Exit Sub
    If Button = 2 Then PopupMenu menu1
End Sub

'*****************************************************************************************************
'* Nombre           : CargarRSTCom
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA EL REGISTRO DE COMPRAS DEL PERIODO ESPECIFICADO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub CargarRSTCom()
        
    RST_Busq RstComp, "SELECT DISTINCT com_compras.*, mae_prov.nombre,  IIf(IsNull([com_compras]![numser])=-1,[com_compras]![numdoc],[com_compras]![numser]+'-'+[com_compras]![numdoc]) AS numerodoc, mae_documento.descripcion AS nomdoc, " _
        & " mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_prov.numruc, mae_moneda.descripcion AS descmon, mae_moneda.simbolo, " _
        & " mae_tipoproducto.descripcion AS desctipcom, con_tc.impcom, Mid([com_compras].[numreg],1,2)+[mae_libros].[codsun]+Mid([com_compras].[numreg],3,4) AS numreg1, " _
        & " com_compras.fchdoc & '' as fchdoc1, com_compras.fchven & '' as fchven1, com_compras.impbru & '' as impbru1,com_compras.impigv & '' as impigv1,com_compras.imptot & '' as imptot1, com_compras.impsal & ''  as impsal1, " _
        & " IIf([com_compras].[tc]=0,[con_tc].[impven],[com_compras].[tc]) & '' AS impven1 ,0 as impdesc " _
        & " FROM (mae_condpago RIGHT JOIN (mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((com_compras LEFT JOIN mae_tipoproducto " _
        & " ON com_compras.idtipo = mae_tipoproducto.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) " _
        & " ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro) ON mae_condpago.id = com_compras.idconpag) LEFT JOIN mae_libros " _
        & " ON com_compras.idlib = mae_libros.id WHERE (((com_compras.numreg) Like '" & Format(mMesActivo, "00") & "%')) ORDER BY com_compras.numreg DESC", xCon
        
'    RST_Busq RstComp, "SELECT DISTINCT com_compras.*, mae_prov.nombre,  IIf(IsNull([com_compras]![numser])=-1,[com_compras]![numdoc],[com_compras]![numser]+'-'+[com_compras]![numdoc]) AS numerodoc, mae_documento.descripcion AS nomdoc, " _
        & " mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_prov.numruc, mae_moneda.descripcion AS descmon, mae_moneda.simbolo, " _
        & " mae_tipoproducto.descripcion AS desctipcom, con_tc.impcom, Mid([com_compras].[numreg],1,2)+[mae_libros].[codsun]+Mid([com_compras].[numreg],3,4) AS numreg1, " _
        & " com_compras.fchdoc & '' as fchdoc1, com_compras.fchven & '' as fchven1, com_compras.imptot & '' as imptot1, com_compras.impsal & ''  as impsal1 " _
        & " FROM (mae_condpago RIGHT JOIN (mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((com_compras LEFT JOIN mae_tipoproducto " _
        & " ON com_compras.idtipo = mae_tipoproducto.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) " _
        & " ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro) ON mae_condpago.id = com_compras.idconpag) LEFT JOIN mae_libros " _
        & " ON com_compras.idlib = mae_libros.id WHERE (((com_compras.numreg) Like '" & Format(mMesActivo, "00") & "%')) ORDER BY com_compras.numreg DESC", xCon
        
End Sub

Private Sub Fg5_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Mostrando = True Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    Dim xTot As Double
    xTot = NulosN(TxtBruto.Text) + NulosN(TxtInafecto.Text)
    
    If Col = 3 Then
        If NulosN(Fg5.TextMatrix(Fg5.Row, 3)) > 100 Then
            Fg5.TextMatrix(Fg5.Row, 3) = ""
            Fg5.TextMatrix(Fg5.Row, 4) = ""
            Exit Sub
        End If
        If NulosN(Fg5.TextMatrix(Fg5.Row, 3)) <> 0 Then
            Fg5.TextMatrix(Fg5.Row, 4) = xTot * NulosN(Fg5.TextMatrix(Fg5.Row, 3) / 100)
        End If
    End If
    
    If Col = 4 Then
        If Fg5.TextMatrix(Fg5.Row, 4) > xTot Then Exit Sub
        If NulosN(Fg5.TextMatrix(Fg5.Row, 4)) <> 0 And xTot <> 0 Then
            Fg5.TextMatrix(Fg5.Row, 3) = ((NulosN(Fg5.TextMatrix(Fg5.Row, 4)) / xTot) * 100)
            Fg5.TextMatrix(Fg5.Row, 3) = Format(Fg5.TextMatrix(Fg5.Row, 3), "0.00")
        End If
    End If
    
    HallarTotCenCos
End Sub

'*****************************************************************************************************
'* Nombre           : HallarTotCenCos
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : HALLA EL TOTAL DE LOS CENTROS DE COSTO DE LOS ITEMS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub HallarTotCenCos()
    Dim A As Integer
    Dim TotPor, TotImp As Double
    
    For A = 1 To Fg5.Rows - 1
        TotPor = TotPor + NulosN(Fg5.TextMatrix(A, 3))
        TotImp = TotImp + NulosN(Fg5.TextMatrix(A, 4))
    Next A
    
    TxtTotPor.Text = Format(TotPor, "0.00")
    TxtTotImp.Text = Format(TotImp, "0.00")
End Sub

Private Sub Fg5_EnterCell()
    If QueHace = 3 Then Exit Sub
    If Fg5.Col = 3 Or Fg5.Col = 4 Then
        Fg5.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        Dim Rpta As Integer
        Dim Rst As New ADODB.Recordset
        
        mMesActivo = xMes
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        If xOrigen = 1 Then
            LblPeriodo2.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
            Nuevo
        Else
            If CONTABILIZAR = True Then
                OpcionesPeriodo
            Else
                RST_Busq RstComp, "SELECT DISTINCT com_compras.*, mae_prov.nombre, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numerodoc, " _
                    & " mae_documento.descripcion AS nomdoc, mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_prov.numruc, " _
                    & " mae_moneda.descripcion AS descmon, mae_moneda.simbolo, mae_tipoproducto.descripcion AS desctipcom, " _
                    & " con_tc.impcom,com_compras.fchdoc & '' as fchdoc1, com_compras.fchven & '' as fchven1, com_compras.imptot & '' as imptot1, com_compras.impsal & '' as impsal1 " _
                    & " FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (mae_condpago RIGHT JOIN ((com_compras LEFT " _
                    & " JOIN mae_tipoproducto ON com_compras.idtipo = mae_tipoproducto.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) " _
                    & " ON mae_condpago.id = com_compras.idconpag) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) " _
                    & " ON mae_prov.id = com_compras.idpro", xCon
                    
            End If
            Set Rst = Nothing
            
            Set Dg1.DataSource = RstComp
'            If RstComp.RecordCount = 0 Then
'                Rpta = MsgBox("No se ha registrado ninguna compra, ¿Desea agregar una ahora?", vbYesNo + vbQuestion + vbDefaultButton1, xTitulo)
'                If Rpta = vbYes Then
'                    Nuevo
'                Else
'                    Dim xMesPro As Integer
'                    xMesPro = xMes
'                    xMes = SeleccionaMes(xCon)
'                    If xMes = 0 Then
'                        xMes = xMesPro
'                    End If
                    
                    'SelecionarPeriodo
'                End If
'            Else
'                OpcionesPeriodo
'                Dg1.SetFocus
'            End If
        End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then '--F3 Nuevo
        If fCierrePeriodo = False Then Exit Sub
        If QueHace <> 3 Then Exit Sub
        Nuevo
    End If
    
    If KeyCode = 115 Then '--F4 Modificar
        If fCierrePeriodo = False Then Exit Sub
        If QueHace <> 3 Then Exit Sub
        If RstComp.RecordCount = 0 Then Exit Sub
        Modificar
    End If
    
    If KeyCode = 113 Then '--F2 Grabar
        If fCierrePeriodo = False Then Exit Sub
        If QueHace = 3 Then Exit Sub
        If Grabar = True Then
            If xOrigen = 0 Then
                Cancelar
                RstComp.Requery
                Dg1.Refresh
            Else
                QueHace = 3
                Set RstComp = Nothing
                Unload Me
                Exit Sub
            End If
        End If
    End If
    
    If KeyCode = 116 Then '--F5 actualizar
    End If
    
    If KeyCode = 117 Then '--F6 '--cancelar
        If fCierrePeriodo = False Then Exit Sub
        If QueHace = 3 Then Exit Sub
        If xOrigen = 1 Then
            QueHace = 3
            IdCompraReg = 0
            Unload Me
            Exit Sub
        End If
        Cancelar
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE LE FORMULARIO
    QueHace = 3
    TabOne1.CurrTab = 0
    SeEjecuto = False
    PermitirEdicion = True
    Dg1.Columns("fchdoc1").NumberFormat = FORMAT_DATE
    Dg1.Columns("fchven1").NumberFormat = FORMAT_DATE
    Dg1.Columns("imptot1").NumberFormat = FORMAT_MONTO
    Dg1.Columns("impsal1").NumberFormat = FORMAT_MONTO
    
    CaracteresNumericos = "0123456789." & Chr(8)
    CaracteresNumericos2 = "0123456789.-" & Chr(8) & Chr(13)
    
    Fg4.ColWidth(5) = 0
    Fg5.ColWidth(5) = 0
    
    Fg1.ColWidth(9) = 0
    Fg1.ColWidth(10) = 0
    Fg1.ColWidth(11) = 0
    Fg1.ColWidth(12) = 0
    Fg1.ColWidth(13) = 0
    Fg1.ColWidth(14) = 0
    Fg1.ColWidth(15) = 0
    Fg1.ColWidth(16) = 0
    Fg1.ColWidth(17) = 0
    
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    
    ChkImpRen4.Value = 1
    
    LblIgvTasa.Caption = ""
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.ColComboList(1) = ""
    If CONTABILIZAR = True Then
        Toolbar1.Buttons(11).Visible = True
        LblPeriodo.Visible = True
        Frame5.Visible = True
    Else
        Toolbar1.Buttons(11).Visible = False
        LblPeriodo.Visible = False
        Frame5.Visible = False
    End If
    
    Fg4.SelectionMode = flexSelectionByRow
    Fg4.Editable = flexEDNone
    
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    
    '--dar formato a las columnas
'    Fg1.ColFormat(3) = "#,###,##0.0000" '--cantidad
'    Fg1.ColFormat(4) = "0.000000" '--afecto
'    Fg1.ColFormat(5) = "0.000000" '--inafecto
'    Fg1.ColFormat(6) = "0.000000" '--descuento
'    Fg1.ColFormat(7) = "0.000000" '--nvo precio
'    Fg1.ColFormat(8) = "#,###,##0.0000" '--total
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario miestras este agregando o editando una compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub Menu1_1_Click()
    CmdAddItem_Click
End Sub

Private Sub menu1_1_1_Click()
    CmdAddItem_Click
End Sub

Private Sub Menu1_3_Click()
    CmdDelItem_Click
End Sub

Private Sub Menu1_5_Click()
    CmdPreHist_Click
End Sub

Private Sub opciones_1_Click()
    AdjuntarEntradas 1, 1
End Sub

Private Sub opciones_2_Click()
    AdjuntarEntradas 2, 1
End Sub

Private Sub Opciones_4_Click()
    AdjuntarEntradas 1, 2
End Sub

Private Sub Opciones_5_Click()
    AdjuntarEntradas 2, 2
End Sub

Private Sub OptDes1_Click()
    ' INDICA QUE SE APLICARA EL DESCUENTO EN PORCENTAJE
    If OptDes1.Value = True Then
        Fg1.TextMatrix(0, 6) = " Dsct en %"
        
        Dim A As Integer
        For A = 1 To Fg1.Rows - 1
            Fg1_CellChanged A, 3
        Next A
    End If
End Sub

Private Sub OptDes2_Click()
    ' INDICA QUE SE APLICARA EL DESCUENTO EN IMPORTE
    If OptDes2.Value = True Then
        Fg1.TextMatrix(0, 6) = "Dsct en Imp."
        
        Dim A As Integer
        For A = 1 To Fg1.Rows - 1
            Fg1_CellChanged A, 3
        Next A
    End If
End Sub

Private Sub OptNo_Click()
    ' INDICA QUE LA COMPRA ESTA NO AFECTA AL IMPUESTO, REVISAR SI SE UTILIZA ESTA OPCION
    If OptNo.Value = True Then HallarTotal
End Sub

Private Sub OptOpera1_Click()
    ' ESPECIFICA QUE LA COMPRA SERA UNA COMPRA NORMAL
    If OptOpera1.Value = True Then
        Fg1.Rows = 1
        Fg4.Rows = 1
        CmdSeleccionar.Caption = "Seleccionar Item"
        CmdAddItem.Enabled = True
        CmdDelItem.Enabled = True
    End If
End Sub

Private Sub OptOpera2_Click()
    ' ESPECIFICA QUE LA COMPRA SE REALIZARA EN FUNCION A DOCUMENTOS DE ENTRADA REGISTRADOS EN EL EVENTO AlmacenIngreso
    If OptOpera2.Value = True Then
        CmdSeleccionar.Caption = "Ver Documentos"
        CmdAddItem.Enabled = False
        CmdDelItem.Enabled = False
    End If
End Sub

Private Sub OptOpera3_Click()
    ' ESPECIFICA QUE LA COMPRA SE REALIZARA EN FUNCIONA UNA O VARIAS ORDENES DE COMPRA
    If OptOpera3.Value = True Then
        CmdSeleccionar.Caption = "Ver Documentos"
        CmdAddItem.Enabled = False
        CmdDelItem.Enabled = False
    End If
End Sub

Private Sub OptSi_Click()
    ' ESPECIFICA QUE LA COMPRA ESTA AFECTA AL IMPUESTO
    If OptSi.Value = True Then HallarTotal
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If xOrigen = 0 Then
            If RstComp.State = 0 Then Exit Sub
            If RstComp.RecordCount = 0 And QueHace <> 1 Then
                Cancel = 1
                Exit Sub
            End If
            If QueHace = 3 Then MuestraSegundoTab
        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Filtrar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : FILTRA REGISTROS EN EL RECORDSET RstComp
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Filtrar()
    Dim xform As New eps_librerias.FormFiltrar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(7, 4) As String
   
    xCampos(0, 0) = "Tipo Documento":     xCampos(0, 1) = "abrev":         xCampos(0, 2) = "C":         xCampos(0, 3) = "1500"
    xCampos(1, 0) = "Moneda":             xCampos(1, 1) = "simbolo":       xCampos(1, 2) = "C":         xCampos(1, 3) = "1500"
    xCampos(2, 0) = "Fch. Emi.":          xCampos(2, 1) = "fchdoc":        xCampos(2, 2) = "F":         xCampos(2, 3) = "1500"
    xCampos(3, 0) = "Proveedor":          xCampos(3, 1) = "nombre":      xCampos(3, 2) = "C":         xCampos(3, 3) = "1500"
    xCampos(4, 0) = "Forma Pago":         xCampos(4, 1) = "desccond":      xCampos(4, 2) = "C":         xCampos(4, 3) = "1500"
    xCampos(5, 0) = "Fch. Vencimiento":   xCampos(5, 1) = "fchven":        xCampos(5, 2) = "C":         xCampos(5, 3) = "1500"
    xCampos(6, 0) = "Importe":            xCampos(6, 1) = "imptot":        xCampos(6, 2) = "C":         xCampos(6, 3) = "1500"
    xCampos(7, 0) = "Saldo":              xCampos(7, 1) = "impsal":        xCampos(7, 2) = "C":         xCampos(7, 3) = "1500"
    
    Set xform.Coneccion = xCon
    Set xform.Rst = RstComp   'recorset que llena el grid
    Set RstComp = xform.FiltrarReg(xCampos)
    Set Dg1.DataSource = RstComp
    Dg1.Refresh
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then
        If RstComp.State = 0 Then Exit Sub
        If RstComp.RecordCount = 0 Then Exit Sub
        ' preguntamos si la compra esta vinculada a una orden de compra
        If RstComp("idordcom") <> 0 Then
            ' no se puede modificar una compra que tenga un orden de compra asignada
            MsgBox "La compra no se puede modificar por tener una Orden de Compra asignada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        Modificar
    End If
    
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        ' GRABAMOS LA COMPRA
        If Grabar = True Then
            If xOrigen = 0 Then
                Cancelar
                RstComp.Requery
                Dg1.Refresh
                
                If RstComp.RecordCount <> 0 Then
                    RstComp.MoveFirst
                    RstComp.Find "id=" & mIdRegistro
                    If RstComp.EOF = True Then RstComp.MoveFirst
                End If
                
            Else
                QueHace = 3
                Set RstComp = Nothing
                Unload Me
                Exit Sub
            End If
        End If
    End If
    
    If Button.Index = 6 Then
        If xOrigen = 1 Then
            QueHace = 3
            IdCompraReg = 0
            Unload Me
            Exit Sub
        End If
        
        Cancelar
    End If
    
    If Button.Index = 8 Then
        Filtrar
    End If
    
    If Button.Index = 9 Then
'        TabOne1.CurrTab = 0
'        TDB_FiltroLimpiar Dg1
'        RstComp.Filter = ""
        TDB_Actualizar Me, TabOne1, Dg1, RstComp
    End If
    
    If Button.Index = 10 Then Buscar
    
    If Button.Index = 11 Then
        mMesActivo = SeleccionaMes(xCon)
        OpcionesPeriodo
    End If
    
    If Button.Index = 13 Then
        pExportar
        'ExportarExcel
'        If RstComp("tipdoc") = 4 Then
'            Imprimir
'        Else
'            MsgBox "No puede imprimir este documento, seleccione una liquidación de compras para efectuar esta operación", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        End If
    End If
    
    If Button.Index = 16 Then
        Set RstComp = Nothing
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : OpcionesPeriodo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LAS OPCIONES DEL PERIODO ESPECIFICADO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub OpcionesPeriodo()
'Modificado 12/01/11 Johan Castro
'           Agregar envío de parametro xIdUsuario a procidimiento CierrePeriodo

     Dim NomMes As String
     Dim Cerrado As Boolean
     Dim Rpta  As Integer
     Dim xFechaMes As String
     
    ' mostrar el boton para agregar apertura
    If mMesActivo = 0 Then CmdApertura.Visible = True Else CmdApertura.Visible = False
     
     LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    '------------------------------------------------------------------------------------------
    ' bloqueamos los botones del toolbar
    CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
    '------------------------------------------------------------------------------------------
    TDB_FiltroLimpiar Dg1
    Set RstComp = Nothing
    '------------------------------------------
    
    LblPeriodo.Caption = LblMes.Caption
    LblPeriodo2.Caption = LblPeriodo.Caption
    
    If mMesActivo <> 0 And mMesActivo <> 13 Then
        xFechaMes = "01/" + Trim(Format(mMesActivo, "00")) + "/" + Trim(Format(AnoTra, "0000"))
        xFchIni = xFechaMes
        xFchFin = Format(HallaDiasMes(CDate(xFechaMes)), "00") + "/" + Mid(xFechaMes, 4, 7)
    End If
    '------------------------------------------
    CargarRSTCom
   
    Set Dg1.DataSource = RstComp

    TabOne1.CurrTab = 0
    
    Dg1.SetFocus

End Sub

Private Sub TxtBruto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtBruto_Validate(Cancel As Boolean)
    If NulosN(TxtBruto.Text) <> 0 Then
        TxtBruto.Text = Format(TxtBruto.Text, FORMAT_MONTO)
        TxtIGV.Text = NulosN(TxtBruto.Text) * xPorIgv
        TxtIGV.Text = Format(TxtIGV.Text, FORMAT_MONTO)
    Else
        TxtIGV.Text = "0.00"
    End If
    BuscarImpuestos
End Sub

Private Sub TxtBruto2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtBruto2_Validate(Cancel As Boolean)
    If NulosN(TxtBruto2.Text) <> 0 Then
        TxtBruto2.Text = Format(TxtBruto2.Text, FORMAT_MONTO)
        TxtIGV2.Text = NulosN(TxtBruto2.Text) * xPorIgv
        TxtIGV2.Text = Format(TxtIGV2.Text, FORMAT_MONTO)
    Else
        TxtIGV2.Text = "0.00"
    End If
    BuscarImpuestos
End Sub

Private Sub TxtBruto3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtBruto3_Validate(Cancel As Boolean)
    If NulosN(TxtBruto3.Text) <> 0 Then
        TxtBruto3.Text = Format(TxtBruto3.Text, FORMAT_MONTO)
        TxtIGV3.Text = NulosN(TxtBruto3.Text) * xPorIgv
        TxtIGV3.Text = Format(TxtIGV3.Text, FORMAT_MONTO)
    Else
        TxtIGV3.Text = "0.00"
    End If
    BuscarImpuestos
End Sub

Private Sub TxtConPag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtConPag_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCondicion_Click
    End If
End Sub

Private Sub TxtConPag_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtConPag.Text) = "" Then Exit Sub
    Dim xRs1 As New ADODB.Recordset
    
    RST_Busq xRs1, "SELECT * FROM mae_condpago WHERE id = " & Val(TxtConPag.Text) & "", xCon
    
    If xRs1.RecordCount = 0 Then
        TxtConPag.Text = ""
        LblCondPag.Caption = ""
    Else
        LblCondPag.Caption = Trim(xRs1("descripcion"))
        If NulosC(TxtFchDoc.Valor) <> "" Then
            TxtFchVen.Valor = CDate(TxtFchDoc.Valor) + xRs1("numdia")
        End If
    End If
    Set xRs1 = Nothing
End Sub

Private Sub TxtDocRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtDocRef_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusDocRef_Click
    End If
End Sub

Private Sub TxtDocRef2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Fg1.Rows = 1 Then
            CmdAddItem.SetFocus
        Else
            Fg1.Row = 1
            Fg1.Col = 1
            Fg1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDocRef2_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = 116 Then
        CmdBusDocRef2_Click
    End If
    If KeyCode = 46 Then
        TxtDocRef2.Text = ""
        LblIdDocRef2.Caption = ""
        EliminarDatosCargados
        PermitirEdicion = True
    End If
End Sub

Sub EliminarDatosCargados()
    Fg1.Rows = 1  ' ELIMINAMOS LOS ITEMS CARGADOS
    Fg5.Rows = 1  ' ELIMINAMOS LOS CENTROS DE COSTOS CARGADOS
    
    ' ELIMINAMOS LOS DATOS DE LA COMPRA CARGADOS
    TxtNumRuc.Text = ""
    LblNomPro.Caption = ""
    LblIdProveedor.Caption = ""
    TxtConPag.Text = ""
    LblCondPag.Caption = ""
    TxtIdMon.Text = ""
    LblMoneda.Caption = ""
End Sub

Private Sub TxtFchDoc_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtFchDoc.Valor) <> "" Then
        Dim xRs1 As New ADODB.Recordset
        
        LblTipoCambio.Caption = Format(HallaTipoCambio(TxtFchDoc.Valor, 2, Venta, xCon))
        RST_Busq xRs1, "SELECT * FROM mae_condpago WHERE id = " & Val(TxtConPag.Text) & "", xCon
        
        If xRs1.RecordCount = 0 Then
            TxtConPag.Text = ""
            LblCondPag.Caption = ""
        Else
            If NulosC(TxtFchDoc.Valor) <> "" Then
                TxtFchVen.Valor = CDate(TxtFchDoc.Valor) + xRs1("numdia")
            End If
        End If
        Set xRs1 = Nothing
    End If
End Sub

Private Sub TxtGlosa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And QueHace <> 3 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdAlmacen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdAlmacen_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusAlm_Click
    End If
End Sub

Private Sub TxtIdAlmacen_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtIdAlmacen.Text) <> "" Then
        Set RstTmp = BuscaConCriterio("SELECT * FROM alm_almacenes WHERE id = " & Val(TxtIdAlmacen.Text) & "", xCon)
        If RstTmp.RecordCount <> 0 Then
            LblDesAlmacen.Caption = RstTmp("descripcion")
        Else
            TxtIdAlmacen.Text = ""
            LblDesAlmacen.Caption = ""
        End If
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
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtIdMon.Text) = "" Then Exit Sub
    Dim xRs1 As New ADODB.Recordset
    
    'buscamos el codigo de la moneda         digitada
    RST_Busq xRs1, "SELECT * FROM mae_moneda WHERE id = " & Val(TxtIdMon.Text) & "", xCon
    
    If xRs1.RecordCount = 0 Then
        TxtIdMon.Text = ""
        LblMoneda.Caption = ""
    Else
        LblMoneda.Caption = Trim(xRs1("descripcion"))
        
        If Trim(TxtIdMon.Text) = "1" Then
            'LblTipCam.Visible = False
            'LblTipoCambio.Visible = False
        Else
            If TxtFchDoc.Valor = "" Then
                MsgBox "No ha especificado la fecha del documento, no se puede determinar " & Chr(13) _
                    & "la fecha del tipo de cambio para este documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                
                TxtIdMon.Text = ""
                LblMoneda.Caption = ""
                TxtFchDoc.SetFocus
                Exit Sub
            End If
            'LblTipCam.Visible = True
            'LblTipoCambio.Visible = True
            LblTipoCambio.Caption = HallaTipoCambio(TxtFchDoc.Valor, 2, Venta, xCon)
            If NulosN(LblTipoCambio.Caption) = 0 Then
                MsgBox "No se ha especificado el tipo de cambio para el dia " & NulosC(TxtFchDoc.Valor), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                TxtIdMon.Text = ""
                LblMoneda.Caption = ""
                Exit Sub
            End If
        End If
    End If
    Set xRs1 = Nothing
    xCuentaDoc = HallaNumCuenta(TxtTipDoc.Text, TxtIdMon.Text)
End Sub

Private Sub TxtIGV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIGV2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIGV3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtInafecto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtISC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumDoc_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtNumDoc.Text) <> "" Then
        If IsNumeric(TxtNumDoc.Text) = True Then
            TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000000")
        End If
        If NulosC(TxtNumDoc.Text) <> "" And NulosC(TxtNumSer.Text) <> "" Then
            If ExisteNumDocCompra = True Then
                Exit Sub
            End If
        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : ExisteNumDocCompra
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : BUSCA EL NUMERO DEL DOCUMENTO DE COMPRA EN LA TABLA Com_compras, ESTA FUNCION
'*                    DEVUELVE VERDADERO EN CASO DE ENCONTRAR EL NUMERO DE DOCUMENTO
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Function ExisteNumDocCompra() As Boolean
    Dim Rst As New ADODB.Recordset
    Dim nSQL As String
    If QueHace <> 1 Then nSQL = " and com_compras.id <> " & NulosN(RstComp("id"))
    
    ' BUSCAMOS EN NUMERO DE DOCUMENTO
    RST_Busq Rst, "SELECT com_compras.fchdoc, Left([com_compras].[numreg],2) & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & Right([com_compras].[numreg],4) AS registro FROM com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id WHERE numser = '" & NulosC(TxtNumSer.Text) & "' and numdoc = '" & NulosC(TxtNumDoc.Text) & "' AND idpro = " & NulosN(LblIdProveedor.Caption) & nSQL, xCon
    If Rst.RecordCount = 0 Then
        ' SI NO EXISTE ESTA BIEN
        ExisteNumDocCompra = False
    Else
        ' SI EXISTE ESTA MAL
        MsgBox "El número de documento ingresado ya fue registrado" & vbCr & "Nº Registro: " & NulosC(Rst("registro")) & vbCr & "Fecha Doc.   " & NulosC(Rst("fchdoc")) & vbCr & "Ingrese Otro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDoc.Text = ""

        ExisteNumDocCompra = True
    End If
    Set Rst = Nothing
End Function

Private Sub TxtNumRuc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumRuc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusProv_Click
    End If
End Sub

Private Sub TxtNumRuc_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If TxtNumRuc.Text = "" Then Exit Sub
    Dim xRs1 As New ADODB.Recordset
    
    RST_Busq xRs1, "SELECT mae_prov.id, mae_prov.numruc, mae_prov.nombre, mae_tipoempresa.descripcion, mae_prov.tipper, mae_prov.idcondpag FROM mae_tipoempresa RIGHT JOIN mae_prov " _
        & " ON mae_tipoempresa.id = mae_prov.tipper WHERE (((mae_prov.numruc) Like '" & TxtNumRuc.Text & "%'))", xCon
    
    If xRs1.RecordCount <> 0 Then
        TxtNumRuc.Text = xRs1("numruc")
        LblNomPro.Caption = xRs1("nombre")
        LblIdProveedor.Caption = xRs1("id")
        
        Lbltipo.Caption = xRs1("descripcion")
        LblIdTipPer.Caption = xRs1("tipper")

        If xRs1("idcondpag") <> 0 Then
            TxtConPag.Text = xRs1("idcondpag")
            TxtConPag_Validate True
        End If
    Else
        TxtNumRuc.Text = ""
        LblNomPro.Caption = ""
        LblIdProveedor.Caption = ""
    End If
    Set xRs1 = Nothing
End Sub

Private Sub TxtNumSer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumSer_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtNumSer.Text) <> "" Then
        TxtNumSer.Text = Format(TxtNumSer.Text, "0000")
        
        If NulosC(TxtNumDoc.Text) <> "" And NulosC(TxtNumSer.Text) <> "" Then
            If ExisteNumDocCompra = True Then
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub TxtTipCom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipCom_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipoCompra_Click
    End If
End Sub

Private Sub TxtTipCom_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtTipCom.Text) <> "" Then
        Set RstTmp = BuscaConCriterio("SELECT * FROM mae_tipoproducto WHERE id = " & Val(TxtTipCom.Text) & "", xCon)
        If RstTmp.RecordCount <> 0 Then
            LblTipoCompra.Caption = RstTmp("descripcion")
        Else
            TxtTipCom.Text = ""
            LblTipoCompra.Caption = ""
        End If
        
        If NulosN(TxtTipCom.Text) = 5 Then
            TxtIdAlmacen.Visible = False
            LblDesAlmacen.Visible = False
            CmdBusAlm.Visible = False
            Label3(11).Visible = False
            TxtIdAlmacen.Text = ""
            LblDesAlmacen.Caption = ""
        Else
            TxtIdAlmacen.Visible = True
            LblDesAlmacen.Visible = True
            CmdBusAlm.Visible = True
            Label3(11).Visible = True
        End If
    End If
    pGridConfigurar
End Sub

Private Sub TxtTipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipDoc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipDoc_Click
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : GRABA LA NUEVA COMPRA, ESTA FUNCION DEVUELVE VERDADERO SI TIENE EXITO AL GRABAR
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Function Grabar() As Boolean
    Dim A, B, Rpta As Integer
    
    ' VALIDA QUE LOS DATOS SEAN LOS CORRECTOS
    
    If NulosN(TxtTipDoc.Text) <> 0 Then
        If xCuentaDoc = 0 Then
            MsgBox "No se ha asignado una cuenta contable al documento " + LblNomDoc.Caption & Chr(13) _
                & "Asignele una cuenta en el menu Contabilidad opcion Asignar Ctas. Contables a documentos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
    
        If xIdCuenTasa = 0 Then
            MsgBox "El impuesto asignado al documento " + LblNomDoc.Caption & Chr(13) & " no tiene cuenta contable" & Chr(13) _
                & "Asignele una cuenta en el menu Contabilidad opcion Maestro de Impuestos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
    Else
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc.SetFocus
        Exit Function
    End If
    
    Dim Rst As New ADODB.Recordset
    
    For A = 1 To Fg1.Rows - 1
        '--validamos cuando sea diferente a servicios
        If NulosN(TxtTipCom.Text) <> 5 Then
            
            ' validamos que el precio ingresado este en un rango de precios especificado
            RST_Busq Rst, "SELECT * FROM com_precios WHERE idpro = " & NulosN(Fg1.TextMatrix(A, 9)) & "", xCon
            If Rst.RecordCount <> 0 Then
                If NulosN(Fg1.TextMatrix(A, 4)) > NulosN(Rst("pretop")) Then
                    Set Rst = Nothing
                    ' buscamos una autorizacion de ingreso para el precio del proveedor
                    RST_Busq Rst, "SELECT com_preciosdet.idpro, com_preciosdet.fecreg, com_preciosdet.idprov, com_preciosdet.precio" _
                        & " From com_preciosdet  " _
                        & " WHERE (((com_preciosdet.idpro)=" & NulosN(Fg1.TextMatrix(A, 9)) & ") AND ((com_preciosdet.fecreg)=CDate('" & Format(TxtFchDoc.Valor, "dd/mm/yyyy") & " ')) " _
                        & " AND ((com_preciosdet.idprov)=" & NulosN(LblIdProveedor.Caption) & "))", xCon
                    
                    If Rst.RecordCount = 0 Then
                        ' si no encontramos una autorizacion de precio para el proveedor en el dia de la operacion se rechaza
                        MsgBox "El precio ingresado para el item " + NulosC(Fg1.TextMatrix(A, 1)) & Chr(13) _
                            & "excede el precio fijado por el administrador de precios, verifique el precio fijado" & Chr(13) _
                            & "en el modulo de Compras opcion  Fijar Precios de Compra a Item", vbCritical + vbOKOnly + vbDefaultButton1, xTitulo
                        Set Rst = Nothing
                        Exit Function
                    Else
                        If NulosN(Fg1.TextMatrix(A, 4)) > NulosN(Rst("precio")) Then
                            ' si el precio ingresado es aun mayor que el precio autorizado se rechaza la compra
                            MsgBox "El precio ingresado para el item " + NulosC(Fg1.TextMatrix(A, 1)) & Chr(13) _
                                & "excede el precio fijado por el administrador de precios, verifique el precio fijado" & Chr(13) _
                                & "en el modulo de Compras opcion  Fijar Precios de Compra a Item", vbCritical + vbOKOnly + vbDefaultButton1, xTitulo
                            Set Rst = Nothing
                            Exit Function
                        End If
                    End If
                End If
            End If
            
            Set Rst = Nothing
            
            ' validamos que el ingreso de items no exceda el stock maximo
            If (OptOpera1.Value = True) Or (OptOpera2.Value = True) Then
                RST_Busq Rst, "SELECT * FROM alm_inventario WHERE id = " & NulosN(Fg1.TextMatrix(A, 8)) & "", xCon
                
                If Rst.RecordCount <> 0 Then
                    If (NulosN(Rst("stckact")) + NulosN(Fg1.TextMatrix(A, 3))) > NulosN(Rst("stckmax")) Then
                        Rpta = MsgBox("La cantidad sumada al stock actual del item " & NulosC(Fg1.TextMatrix(A, 1)) & Chr(13) _
                            & "sobrepasa el Stock Maximo asignado ¿Esta seguro de agregar la cantidad especificada?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
                        If Rpta = vbNo Then
                            Set Rst = Nothing
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        
        ' validamos la cuenta contable del item
        If NulosN(Fg1.TextMatrix(A, 10)) = 0 Then
            MsgBox "No se le ha asignado una Cuenta Contable al item : " & Chr(13) _
                & Fg1.TextMatrix(A, 1) & Chr(13) _
                & "Asígnele una cuenta en el menu Almacén opción Mantenimiento Items de Compra y Venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
    Next A
        
    If TxtTipCom.Text = "" Then
        MsgBox "No ha especificado el Tipo de Compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipCom.SetFocus
        Exit Function
    End If
    
    If TxtNumRuc.Text = "" Then
        MsgBox "No ha especificado proveedor de la compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumRuc.SetFocus
        Exit Function
    End If
    
    If TxtNumSer.Text = "" Or TxtNumDoc.Text = "" Then
        MsgBox "No ha especificado el numero del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumSer.SetFocus
        Exit Function
    End If
    
    If TxtFchDoc.Valor = "" Then
        MsgBox "No ha especificado la fecha de emision del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchDoc.SetFocus
        Exit Function
    End If
    
    If NulosN(LblTipoCambio.Caption) = 0 Then
        MsgBox "No se ha especificado tipo de cambio para el dia " & NulosC(TxtFchDoc.Valor), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Function
    End If
    
    If TxtFchVen.Valor = "" Then
        MsgBox "No ha especificado la fecha de vencimiento del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchVen.SetFocus
        Exit Function
    End If
    
    If TxtConPag.Text = "" Then
        MsgBox "No ha especificado la condicion de pago del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtConPag.SetFocus
        Exit Function
    End If
    
    If TxtIdMon.Text = "" Then
        MsgBox "No ha especificado la moneda del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMon.SetFocus
        Exit Function
    End If

    
    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado items para la compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        '------------
        Fg1.Col = 1
        '------------
        Fg1.SetFocus
        Exit Function
    End If
    
    ' verificamos que la fecha de vencimiento no sea menor a la fecha de vencimiento
    If CDate(TxtFchDoc.Valor) > CDate(xFchFin) Then
        MsgBox "La fecha de Emisión del documento no puede ser mayor a la fecha de cierre del periodo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchDoc.SetFocus
        Exit Function
    End If
    
    If NulosN(TxtTipCom.Text) <> 5 Then
        If NulosC(TxtIdAlmacen.Text) = "" Then
            MsgBox "No ha especificado el almacen de destino de la compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtIdAlmacen.SetFocus
            Exit Function
        End If
    End If
    
    ' verificamos que la fecha de vencimiento no sea mayor al periodo contable
    If NulosN(TxtTipDoc.Text) = 14 Then
        If CDate(TxtFchVen.Valor) > (CDate(xFchFin)) Then
            If NulosC(TxtFchPago.Valor) = "" Then
                MsgBox "No puede registrar este documento en el mes de " + Trim(LblPeriodo.Caption) + ", la fecha de " & Chr(13) _
                    & "vencimiento es mayor a la fecha del periodo, para registrar este documento en el periodo" & Chr(13) _
                    & "actual ingrese la fecha de pago menor o igual a la fecha de cierre del periodo ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Function
            End If
        End If
    End If
    
    'VERIFICAMOS QUE LOS ITEMS IGRESADOS SON LOS CORRECTOS
    'VERIFICAMOS QUE NO EXISTAS FILAS SIN ITEMS
    For A = 1 To Fg1.Rows - 1
        If NulosC(Fg1.TextMatrix(A, 1)) = "" Then
            Fg1.RemoveItem A
        End If
    Next A
    
    If Fg1.Rows <> 1 Then
        For A = 1 To Fg1.Rows - 1
            If NulosN(Fg1.TextMatrix(A, 3)) = 0 Then
                MsgBox "No ha especificado la cantidad para el item : " + Trim(Fg1.TextMatrix(A, 1)), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Fg1.Col = 3: Fg1.Row = A
                Fg1.SetFocus
                Exit Function
            End If
        Next A
    Else
        MsgBox "No se ha especificado ningún item para esta compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
    
    'verificamos que el total de items sea igual al total de los totales
    A = NulosN(Format(GRID_SUMAR_COL(Fg1, 8), "0.00"))
    
    B = NulosN(TxtBruto.Text) + NulosN(TxtBruto2.Text) + NulosN(TxtBruto3.Text) + NulosN(TxtInafecto.Text)
    
    If Round(A, 1) <> Round(B, 1) Then
        MsgBox "El monto del detalle del documento no coincide con la sumatoria de los totales", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtBruto.SetFocus
        Exit Function
    End If
    
    ' VERIFICAMOS QUE LA FACTURA NO TENGA IMPORTE 0
    If NulosN(TxtTotal.Text) = 0 Then
        MsgBox "El importe total de la factura no puede ser 0", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.SetFocus
        Exit Function
    End If
    
    Dim RstDeta2 As New ADODB.Recordset
    Dim RstActPro As New ADODB.Recordset
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstDia As New ADODB.Recordset
    Dim RstCosto As New ADODB.Recordset
    
    Dim xIdCuen As Integer
    Dim xId As Double
    Dim xTotal As Double
    Dim xNumAsiento As String
    Dim xSaldo As Double '--indica el saldo actual del documento
    
    Dim nSQL As String
    
'    On Error GoTo LaCague
    
    xCon.BeginTrans
    
    If QueHace = 1 Then
        ' SI ESTA AGREGANDO UN NUEVO REGISTRO
        xId = HallaCodigoTabla("com_compras", xCon, "id")
        xNumAsiento = NuevoNumAsiento(1, mMesActivo, xCon)
        RST_Busq RstCab, "SELECT TOP 1 * FROM com_compras", xCon
        
        RstCab.AddNew
        RstCab("id") = xId
        IdCompraReg = xId
        
'        If NulosN(TxtTipDoc.Text) = 7 Then
'            xSaldo = 0
'        Else
            xSaldo = NulosN(TxtTotal.Text)
'        End If
    Else
        ' SI ESTA MODIFICANDO UN REGISTRO
        xId = RstComp("id")
        RST_Busq RstCab, "SELECT * FROM com_compras WHERE id = " & xId & "", xCon
        
        '------------------------------------------
        'eliminamos el sotck agregado con la compra
        RST_Busq RstDeta2, "SELECT com_comprasdet.* From com_comprasdet WHERE (((com_comprasdet.idcom)=" & xId & "))", xCon

        If RstDeta2.RecordCount <> 0 Then
            RstDeta2.MoveFirst
            For A = 1 To RstDeta2.RecordCount
                RST_Busq RstActPro, "SELECT alm_inventario.id, alm_inventario.stckact  From alm_inventario WHERE ((alm_inventario.id=" & RstDeta2("iditem") & "))", xCon
                If RstActPro.RecordCount = 1 Then
                    RstActPro("stckact") = NulosN(RstActPro("stckact")) - NulosN(RstDeta2("canpro"))
                    RstActPro.Update
                End If
                Set RstActPro = Nothing
            Next A
        End If
        Set RstDeta2 = Nothing
        
        '----------------------------------
        'eliminamos el detalle de la compra
        xCon.Execute "DELETE * FROM com_comprasdet WHERE idcom = " & xId & ""
                
        'eliminamos el asiento contable
        If mMesActivo = 0 Then
            xCon.Execute "DELETE * FROM con_diario WHERE idmes = " & mMesActivo & " AND idlib = 36 AND idmov = " & xId & ""
        Else
            xCon.Execute "DELETE * FROM con_diario WHERE idmes = " & mMesActivo & " AND idlib = 1 AND idmov = " & xId & ""
        End If
        
        'eliminamos el centro de costos
        xCon.Execute "DELETE * FROM com_comprascosto WHERE idcom = " & xId & ""
        
        xNumAsiento = Mid(RstComp("numreg"), 3, 4)
        
        'Borramos los flag de las tablas alm_ingreso y com_ordencompra
        xCon.Execute "DELETE * FROM alm_ingresodoc WHERE iddoc = " & RstComp("id") & " "

        'actualizamos campo idfac en la tabla com_ordencompra a 0 para que se vuelva a procesar
        xCon.Execute "UPDATE com_ordencompra SET com_ordencompra.idfac = 0 WHERE (((com_ordencompra.idfac)=" & RstComp("id") & "))"
        
        '--obtener el ultimo saldo

        nSQL = " SELECT tes_cajadestinodet.iddoc, tes_caja.idmon, tes_cajadestino.tc AS tipcam, tes_cajadestinodet.acuenta AS imptotal, IIf(tes_caja!idmon=1,tes_cajadestinodet!acuenta,tes_cajadestinodet!acuenta*tipcam) AS imptotsol, IIf(tes_caja!idmon=2,tes_cajadestinodet!acuenta,tes_cajadestinodet!acuenta/tipcam) AS imptotdol " _
            + vbCr + " FROM tes_caja INNER JOIN (tes_cajadestino INNER JOIN tes_cajadestinodet ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON tes_caja.id = tes_cajadestino.idtes " _
            + vbCr + " Where (((tes_cajadestinodet.iddoc) = " & xId & ") And ((tes_cajadestinodet.idmod) = 1) And ((tes_caja.tipmov) = 2)) " _
            + vbCr + " Union " _
            + vbCr + " SELECT con_canjesdet.iddoc, con_canjes.idmon, con_tc.impven AS tipcam, con_canjesdet.impdoc AS imptotal, IIf(con_canjes.idmon=1,con_canjesdet.impdoc,IIf(con_tc.impven Is Null Or con_tc.impven=0,0,con_canjesdet.impdoc*con_tc.impven)) AS imptotsol, IIf(con_canjes.idmon=2,con_canjesdet.impdoc,IIf(con_tc.impven Is Null Or con_tc.impven=0,0,con_canjesdet.impdoc/con_tc.impven)) AS imptotdol " _
            + vbCr + " FROM (con_canjes LEFT JOIN con_tc ON con_canjes.fchemi = con_tc.fecha) INNER JOIN con_canjesdet ON con_canjes.id = con_canjesdet.idcan " _
            + vbCr + " Where (((con_canjesdet.iddoc) = " & xId & ") And ((con_canjesdet.Tipo) = 2)) " _
            + vbCr + " Union " _
            + vbCr + " SELECT com_compras.iddocref, com_compras.idmon, IIf(com_compras.tc=0,IIf(con_tc.impven Is Null,0,con_tc.impven),com_compras.tc) AS tipcam, com_compras.imptot, IIf(com_compras!idmon=1,com_compras!imptot,com_compras!imptot*tipcam) AS imptotsol, IIf(com_compras!idmon=2,com_compras!imptot,com_compras!imptot/tipcam) AS imptotdol " _
            + vbCr + " FROM com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha " _
            + vbCr + " Where (((com_compras.iddocref) = " & xId & ")) "

        RST_Busq Rst, nSQL, xCon
        
        
        If Rst.RecordCount <> 0 Then
            If NulosN(TxtIdMon.Text) = 1 Then
                xSaldo = NulosN(TxtTotal.Text) - NulosN(RstRegistroSumar(Rst, "imptotsol"))
            Else
                xSaldo = NulosN(TxtTotal.Text) - NulosN(RstRegistroSumar(Rst, "imptotdol"))
            End If
        Else
            xSaldo = NulosN(TxtTotal.Text)
            
        End If
        Set Rst = Nothing
        '************************************************************************************************************************
    End If
    
    ' ESCRIBIMOS LOS DATOS DE LA COMPRA
    RST_Busq RstDet, "SELECT TOP 1 * FROM com_comprasdet", xCon
    RST_Busq RstDia, "SELECT TOP 1 * FROM con_diario", xCon
    RST_Busq RstCosto, "SELECT TOP 1 * FROM com_comprascosto", xCon
    
    mIdRegistro = xId
    
    RstCab("idlib") = 1
    RstCab("idtipo") = NulosN(TxtTipCom.Text)
    RstCab("tipdoc") = NulosN(TxtTipDoc.Text)
    RstCab("idpro") = NulosN(LblIdProveedor.Caption)
    RstCab("numser") = TxtNumSer.Text
    RstCab("numdoc") = TxtNumDoc.Text
    RstCab("fchdoc") = TxtFchDoc.Valor
    RstCab("fchven") = TxtFchVen.Valor
    If IsDate(TxtFchPago.Valor) = True Then RstCab("fchpag") = TxtFchPago.Valor
    RstCab("idconpag") = NulosN(TxtConPag.Text)
    RstCab("idmon") = NulosN(TxtIdMon.Text)
    RstCab("impbru") = NulosN(TxtBruto.Text)
    RstCab("impbru2") = NulosN(TxtBruto2.Text)
    RstCab("impbru3") = NulosN(TxtBruto3.Text)
    RstCab("impina") = NulosN(TxtInafecto.Text)
    RstCab("impigv") = NulosN(TxtIGV.Text)
    RstCab("impigv2") = NulosN(TxtIGV2.Text)
    RstCab("impigv3") = NulosN(TxtIGV3.Text)
    RstCab("otroscargos") = NulosN(TxtOtros.Text)
    RstCab("imptot") = NulosN(TxtTotal.Text)
    RstCab("glosa") = NulosC(TxtGlosa.Text)
    
    If NulosN(TxtTipDocRef.Text) <> 0 Then
        RstCab("idtipdocref") = NulosN(TxtTipDocRef.Text)
        RstCab("iddocref2") = NulosN(LblIdDocRef2.Caption)
    Else
        RstCab("idtipdocref") = 0
        RstCab("iddocref2") = 0
    End If
    '--uso temporal
    RstCab("numerodocref") = NulosC(TxtDocRef2.Text)
    
    RstCab("impsal") = xSaldo
    
    RstCab("impisc") = NulosN(TxtISC.Text)
    
    If NulosN(TxtTipCom.Text) <> 5 Then
        RstCab("idalm") = NulosN(TxtIdAlmacen.Text)
    End If
    
    ' documento al que hace referencia en caso de ser nota de credito
    RstCab("iddocref") = NulosN(LblIdDocRef.Caption)
    
    ' Actualizamos el saldo del documento
'    ActualizaSaldoDoc NulosN(LblIdDocRef.Caption), 1, NulosN(TxtTotal.Text)
    
    RstCab("fchreg") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
    
    ' SI EL SISTEMA ESTA EN MODO CONTABILIZAR GRABAMOS EL NUMERO DE REGISTRO DE LA OPERACION
    If CONTABILIZAR = True Then
        RstCab("numreg") = Format(Trim(Str(mMesActivo)), "00") + xNumAsiento
    End If
    
    'grabamos el tipo de descuento
    If OptDes1.Value = True Then
        RstCab("tipdes") = 1
    End If
    If OptDes2.Value = True Then
        RstCab("tipdes") = 2
    End If
    
    If OptSi.Value = True Then
        RstCab("afecto") = -1
    Else
        RstCab("afecto") = 0
    End If
    
    ' especificamos como en que contexto se esta haciendo la compra
    If OptOpera1.Value = True Then RstCab("tipcom") = 1  'Compra normal
    If OptOpera3.Value = True Then RstCab("tipcom") = 2  'Compra vinculada con documentos de entrada
    If OptOpera2.Value = True Then RstCab("tipcom") = 3  'Compra vinculada con Orden de Compra
    
    '--grabar la tasa del igv aplicada; solo si hay impuesto
    If NulosN(TxtIGV.Text) + NulosN(TxtIGV2.Text) + NulosN(TxtIGV3.Text) <> 0 Then
        RstCab("tasaigv") = TasaImpuesto
    Else
        RstCab("tasaigv") = 0
    End If
    
    RstCab.Update
    
    'Grabamos los items de la compra
    For A = 1 To Fg1.Rows - 1
        RstDet.AddNew
        RstDet("idcom") = xId
        RstDet("iditem") = NulosN(Fg1.TextMatrix(A, 9))
        RstDet("idunimed") = NulosN(Fg1.TextMatrix(A, 10))
        RstDet("canpro") = NulosN(Fg1.TextMatrix(A, 3))
        RstDet("preunibru") = NulosN(Fg1.TextMatrix(A, 4)) 'precio bruto afecto
        RstDet("preunibruina") = NulosN(Fg1.TextMatrix(A, 5)) 'precio bruto inafecto
        RstDet("valdes") = NulosN(Fg1.TextMatrix(A, 6))
        RstDet("preuni") = NulosN(Fg1.TextMatrix(A, 7))
        RstDet("imptot") = NulosN(Fg1.TextMatrix(A, 8))
        RstDet("idcue") = NulosN(Fg1.TextMatrix(A, 11))
        RstDet.Update
        
        If NulosN(TxtTipCom.Text) = 1 Or NulosN(TxtTipCom.Text) = 4 Or NulosN(TxtTipCom.Text) = 2 Then
            RST_Busq RstActPro, "SELECT alm_inventario.id, alm_inventario.stckact  From alm_inventario WHERE (((alm_inventario.id)=" & Val(Val(Fg1.TextMatrix(A, 8))) & "))", xCon

            If RstActPro.RecordCount = 1 Then
                RstActPro("stckact") = NulosN(RstActPro("stckact")) + NulosN(Fg1.TextMatrix(A, 3))
                RstActPro.Update
            End If
            Set RstActPro = Nothing
        End If
    Next A
    
    ' Actualizamos los documentos relacionados con la factura
    If OptOpera3.Value = True Then
        If Fg4.Rows <> 1 Then
            For A = 1 To Fg4.Rows - 1
                ' actualizamos el flag de los partes de entrada para saber con que documento de compra se valorizaran
                xCon.Execute "INSERT INTO alm_ingresodoc (id, iddoc) values (" & NulosN(Fg4.TextMatrix(A, 5)) & "," & xId & ")"
            Next A
        End If
    End If
    
    If OptOpera2.Value = True Then
        If Fg4.Rows <> 3 Then
            For A = 1 To Fg4.Rows - 1
                'actualizamos el flag de las ordenes de compra para saber con que documento ingresaron las ordenes de compra
                xCon.Execute "UPDATE com_ordencompra SET com_ordencompra.idfac = " & xId & " WHERE (((com_ordencompra.id)=" & NulosN(Fg4.TextMatrix(A, 5)) & "))"
            Next A
        End If
    End If
    
    'grabamos el centro de costos
    If Fg5.Rows > 1 Then
        For A = 1 To Fg5.Rows - 1
            RstCosto.AddNew
            RstCosto("idcom") = xId
            RstCosto("idcencos") = NulosN(Fg5.TextMatrix(A, 5))
            RstCosto("imppor") = NulosN(Fg5.TextMatrix(A, 3))
            RstCosto("impcos") = NulosN(Fg5.TextMatrix(A, 4))
            RstCosto.Update
        Next A
    End If
    
    If CONTABILIZAR = True Then
        'Grabamos el libro diario del movimiento
        'grabamos a facturas por pagar Plan de cuentas 42.1 o dependiendo del caso
'        RstDia.AddNew
'        RstDia("año") = AnoTra
'        RstDia("idmes") = mMesActivo
'        RstDia("idlib") = 1
'        RstDia("idmov") = xId
'        RstDia("numasi") = xNumAsiento
'        RstDia("tc") = ValTipCam
'        RstDia("idcue") = xCuentaDoc
'        RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'        RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'        If NulosN(TxtTipDoc.Text) <> 0 Then
'            If NulosN(TxtTipDoc.Text) <> 7 Then
'                ' cuando se factura u otro comprabante excepto nota de credito hace su asiento norma
'                If TxtIdMon.Text = "1" Then
'                    RstDia("imphabsol") = Format(NulosN(TxtTotal.Text), "0.000000")
'                    RstDia("imphabdol") = 0
'                Else
'                    RstDia("imphabsol") = Format(NulosN(TxtTotal.Text), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.000000")
'                    RstDia("imphabdol") = Format(NulosN(TxtTotal.Text), "0.000000")
'                End If
'            Else
'                ' cuando sea nota de credito hace el asiento inverso al de una venta
'                If TxtIdMon.Text = "1" Then
'                    RstDia("impdebsol") = Format(NulosN(TxtTotal.Text), "0.000000")
'                    RstDia("impdebdol") = 0
'                Else
'                    RstDia("impdebsol") = Format(NulosN(TxtTotal.Text), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.000000")
'                    RstDia("impdebdol") = Format(NulosN(TxtTotal.Text), "0.000000")
'                End If
'            End If
'        End If
'        RstDia.Update
        
        ' grabamos el impuesto si la operacion esta afecta a el
'        If NulosN(TxtIGV.Text) <> 0 Then
'                RstDia.AddNew
'                RstDia("año") = AnoTra
'                RstDia("idmes") = mMesActivo
'                If mMesActivo = 0 Then
'                    RstDia("idlib") = 36
'                Else
'                    RstDia("idlib") = 1
'                End If
'                RstDia("idmov") = xId
'                RstDia("numasi") = xNumAsiento
'                RstDia("tc") = ValTipCam
'                RstDia("idcue") = xIdCuenTasa
'                RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'                RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'
'                ' si el tipo de l proveedor es diferente a no domiciliado
'                If NulosN(LblIdTipPer.Caption) <> 3 Then
'                    If NulosN(TxtTipDoc.Text) <> 0 Then
'                        If NulosN(TxtTipDoc.Text) <> 7 And NulosN(TxtTipDoc.Text) <> 2 Then
'                            If TxtIdMon.Text = "1" Then
'                                RstDia("impdebsol") = Format(NulosN(TxtIGV.Text), "0.000000")
'                                RstDia("impdebdol") = 0
'                            Else
'                                RstDia("impdebsol") = Format(NulosN(TxtIGV.Text), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.000000")
'                                RstDia("impdebdol") = Format(NulosN(TxtIGV.Text), "0.000000")
'                            End If
'                        Else
'                            If TxtIdMon.Text = "1" Then
'                                RstDia("imphabsol") = Format(NulosN(TxtIGV.Text), "0.000000")
'                                RstDia("imphabdol") = 0
'                            Else
'                                RstDia("imphabsol") = Format(NulosN(TxtIGV.Text), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.000000")
'                                RstDia("imphabdol") = Format(NulosN(TxtIGV.Text), "0.000000")
'                            End If
'                        End If
'                    End If
'                Else
'                    If TxtIdMon.Text = "1" Then
'                        RstDia("imphabsol") = Format(NulosN(TxtIGV.Text), "0.000000")
'                        RstDia("imphabdol") = 0
'                    Else
'                        RstDia("imphabsol") = Format(NulosN(TxtIGV.Text), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.000000")
'                        RstDia("imphabdol") = Format(NulosN(TxtIGV.Text), "0.000000")
'                    End If
'                End If
'                RstDia.Update
'            Else
'            End If
'        End If
'
        ' grabamos el impuesto si la operacion esta no afecta a el
'        If NulosN(TxtIGV2.Text) <> 0 Then
'            RstDia.AddNew
'            RstDia("año") = AnoTra
'            RstDia("idmes") = mMesActivo
'            If mMesActivo = 0 Then
'                RstDia("idlib") = 36
'            Else
'                RstDia("idlib") = 1
'            End If
'            RstDia("idmov") = xId
'            RstDia("numasi") = xNumAsiento
'            RstDia("tc") = ValTipCam
'            RstDia("idcue") = xIdCuenTasa
'            RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'            RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'
'            ' si el tipo de l proveedor es diferente a no domiciliado
'            If NulosN(LblIdTipPer.Caption) <> 3 Then
'                If NulosN(TxtTipDoc.Text) <> 0 Then
'                    If NulosN(TxtTipDoc.Text) <> 7 And NulosN(TxtTipDoc.Text) <> 2 Then
'                        If TxtIdMon.Text = "1" Then
'                            RstDia("impdebsol") = Format(NulosN(TxtIGV2.Text), "0.000000")
'                            RstDia("impdebdol") = 0
'                        Else
'                            RstDia("impdebsol") = Format(NulosN(TxtIGV2.Text), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.000000")
'                            RstDia("impdebdol") = Format(NulosN(TxtIGV2.Text), "0.000000")
'                        End If
'                    Else
'                        If TxtIdMon.Text = "1" Then
'                            RstDia("imphabsol") = Format(NulosN(TxtIGV2.Text), "0.000000")
'                            RstDia("imphabdol") = 0
'                        Else
'                            RstDia("imphabsol") = Format(NulosN(TxtIGV2.Text), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.000000")
'                            RstDia("imphabdol") = Format(NulosN(TxtIGV2.Text), "0.000000")
'                        End If
'                    End If
'                End If
'            Else
'                If TxtIdMon.Text = "1" Then
'                    RstDia("imphabsol") = Format(NulosN(TxtIGV2.Text), "0.000000")
'                    RstDia("imphabdol") = 0
'                Else
'                    RstDia("imphabsol") = Format(NulosN(TxtIGV2.Text), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.000000")
'                    RstDia("imphabdol") = Format(NulosN(TxtIGV2.Text), "0.000000")
'                End If
'            End If
'            RstDia.Update
'        End If
        
        ' grabamos el impuesto si la operacion sin derecho a credito fiscal
'        If NulosN(TxtIGV3.Text) <> 0 Then
'            RstDia.AddNew
'            RstDia("año") = AnoTra
'            RstDia("idmes") = mMesActivo
'            If mMesActivo = 0 Then
'                RstDia("idlib") = 36
'            Else
'                RstDia("idlib") = 1
'            End If
'            RstDia("idmov") = xId
'            RstDia("numasi") = xNumAsiento
'            RstDia("tc") = ValTipCam
'            RstDia("idcue") = xIdCuenTasa
'            RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'            RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'
'            ' si el tipo de l proveedor es diferente a no domiciliado
'            If NulosN(LblIdTipPer.Caption) <> 3 Then
'                If NulosN(TxtTipDoc.Text) <> 0 Then
'                    If NulosN(TxtTipDoc.Text) <> 7 And NulosN(TxtTipDoc.Text) <> 2 Then
'                        If TxtIdMon.Text = "1" Then
'                            RstDia("impdebsol") = Format(NulosN(TxtIGV3.Text), "0.000000")
'                            RstDia("impdebdol") = 0
'                        Else
'                            RstDia("impdebsol") = Format(NulosN(TxtIGV3.Text), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.000000")
'                            RstDia("impdebdol") = Format(NulosN(TxtIGV3.Text), "0.000000")
'                        End If
'                    Else
'                        If TxtIdMon.Text = "1" Then
'                            RstDia("imphabsol") = Format(NulosN(TxtIGV3.Text), "0.000000")
'                            RstDia("imphabdol") = 0
'                        Else
'                            RstDia("imphabsol") = Format(NulosN(TxtIGV3.Text), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.000000")
'                            RstDia("imphabdol") = Format(NulosN(TxtIGV3.Text), "0.000000")
'                        End If
'                    End If
'                End If
'            Else
'                If TxtIdMon.Text = "1" Then
'                    RstDia("imphabsol") = Format(NulosN(TxtIGV3.Text), "0.000000")
'                    RstDia("imphabdol") = 0
'                Else
'                    RstDia("imphabsol") = Format(NulosN(TxtIGV3.Text), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.000000")
'                    RstDia("imphabdol") = Format(NulosN(TxtIGV3.Text), "0.000000")
'                End If
'            End If
'            RstDia.Update
'        End If
        
        
        
'        ' grabamos el impuesto si la operacion a sujeto a no domiciliado
'        If NulosN(TxtOtros.Text) <> 0 And NulosN(TxtTipDoc.Text) = 107 Then
'            RstDia.AddNew
'            RstDia("año") = AnoTra
'            RstDia("idmes") = mMesActivo
'            RstDia("idlib") = 1
'            RstDia("idmov") = xId
'            RstDia("numasi") = xNumAsiento
'            RstDia("tc") = ValTipCam
'            RstDia("idcue") = xIdCuenTasa
'            RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'            RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'            If NulosN(TxtTipDoc.Text) <> 0 Then
'                If NulosN(TxtTipDoc.Text) <> 7 Then
'                    ' cuando se factura u otro comprabante excepto nota de credito hace su asiento norma
'                    If TxtIdMon.Text = "1" Then
'                        RstDia("imphabsol") = Format(NulosN(TxtOtros.Text), "0.000000")
'                        RstDia("imphabdol") = 0
'                    Else
'                        RstDia("imphabsol") = Format(NulosN(TxtOtros.Text), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.000000")
'                        RstDia("imphabdol") = Format(NulosN(TxtOtros.Text), "0.000000")
'                    End If
'                Else
'                    ' cuando sea nota de credito hace el asiento inverso al de una venta
'                    If TxtIdMon.Text = "1" Then
'                        RstDia("impdebsol") = Format(NulosN(TxtOtros.Text), "0.000000")
'                        RstDia("impdebdol") = 0
'                    Else
'                        RstDia("impdebsol") = Format(NulosN(TxtOtros.Text), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.000000")
'                        RstDia("impdebdol") = Format(NulosN(TxtOtros.Text), "0.000000")
'                    End If
'                End If
'            End If
'            RstDia.Update
        End If
        '***********************************************
        ' ESCRIBIMOS EL ASIENTO CONTABLE DE LA OPERACION
        Dim xFchAsi As String
        
        xFchAsi = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
        
        ' grabamos el haber de la compra
        EscribirAsiento NulosN(AnoTra), mMesActivo, NulosN(TxtIdMon.Text), 1, xId, xNumAsiento, NulosN(LblTipoCambio.Caption), xCuentaDoc, TxtFchDoc.Valor, xFchAsi, NulosN(TxtTipDoc.Text), NulosN(TxtTotal.Text), 2, xCon
        
        ' GRABAMOS EL IMPUESTO DE LA OPERACION
        
        If NulosN(TxtIGV.Text) <> 0 Then        ' SI ES EL PRIMER IGV DE LA COMPRA
            EscribirAsiento NulosN(AnoTra), mMesActivo, TxtIdMon.Text, 1, xId, xNumAsiento, LblTipoCambio.Caption, xIdCuenTasa, TxtFchDoc.Valor, xFchAsi, TxtTipDoc.Text, NulosN(TxtIGV.Text), 1, xCon
        End If
        
        'grabamos el impuesto si la operacion esta no afecta a el
        If NulosN(TxtIGV2.Text) <> 0 Then       ' SI ES EL SEGUNDO IGV DE LA COMPRA
            EscribirAsiento NulosN(AnoTra), mMesActivo, TxtIdMon.Text, 1, xId, xNumAsiento, LblTipoCambio.Caption, xIdCuenTasa, TxtFchDoc.Valor, xFchAsi, TxtTipDoc.Text, NulosN(TxtIGV2.Text), 1, xCon
        End If
        
        'grabamos el impuesto si la operacion sin derecho a credito fiscal
        If NulosN(TxtIGV3.Text) <> 0 Then       ' SI ES EL TERCER IGV DE LA COMPRA
            EscribirAsiento NulosN(AnoTra), mMesActivo, TxtIdMon.Text, 1, xId, xNumAsiento, LblTipoCambio.Caption, xIdCuenTasa, TxtFchDoc.Valor, xFchAsi, TxtTipDoc.Text, NulosN(TxtIGV3.Text), 1, xCon
        End If
        
                
                
        ' GRABAMOS EL IMPUESTO PARA LAS OPERACIONES CON PROVEEDORES NO DOMICILIADOS
        If NulosN(TxtOtros.Text) <> 0 And NulosN(TxtTipDoc.Text) = 107 Then
            EscribirAsiento NulosN(AnoTra), mMesActivo, TxtIdMon.Text, 1, xId, xNumAsiento, LblTipoCambio.Caption, xIdCuenTasa, TxtFchDoc.Valor, xFchAsi, TxtTipDoc.Text, TxtOtros.Text, 2, xCon
        End If
        
        ' GRABAMOS EL IMPONIBLE DEL DOCUMENTO EN FUNCION A LOS ITEMS DEL DOCUMENTO
        'Dim A As Integer
        Dim RstItem As New ADODB.Recordset
        Dim xFun As New eps_librerias.FuncionesData
        Dim xCampos(3, 3) As String
        Dim xIdCue As Integer
        'Dim xTotal As Double
        
        xCampos(0, 0) = "iditem":        xCampos(0, 1) = "N":      xCampos(0, 2) = "2"
        xCampos(1, 0) = "idcuen":        xCampos(1, 1) = "N":      xCampos(1, 2) = "2"
        xCampos(2, 0) = "importe":       xCampos(2, 1) = "D":      xCampos(2, 2) = "2"
        Set RstItem = xFun.CrearRstTMP(xCampos)
        RstItem.Open
        
        For A = 1 To Fg1.Rows - 1
            RstItem.AddNew
            RstItem("iditem") = NulosN(Fg1.TextMatrix(A, 9))
            RstItem("idcuen") = NulosN(Fg1.TextMatrix(A, 11))
            RstItem("importe") = NulosN(Fg1.TextMatrix(A, 8))
        Next A
        RstItem.MoveFirst
        RstItem.Sort = "idcuen"
        
        xIdCue = RstItem("idcuen")
        xTotal = 0
        For A = 1 To RstItem.RecordCount
            'ASGINAR DATOS A VARIABLES
            xIdCue = RstItem("idcuen")
            xTotal = NulosN(RstItem("importe"))
            
            EscribirAsiento NulosN(AnoTra), mMesActivo, TxtIdMon.Text, 1, xId, xNumAsiento, LblTipoCambio.Caption, xIdCue, TxtFchDoc.Valor, xFchAsi, TxtTipDoc.Text, xTotal, 1, xCon
            
            RstItem.MoveNext
            
        Next A
        
        ' GRABAMOS EL MOVIMIENTO EN LA TABLA var_ctacte PARA ANALIZAR POR DOCUMENTO DE REFERENCIA
        xNumAsiento = Format(mMesActivo, "00") & Format(Busca_Codigo(1, "id", "codsun", "mae_libros", "N", xCon), "00") & xNumAsiento
        
        If NulosN(TxtTipDoc.Text) <> 7 Then
            ' SI ES DIFERENTE A NOTA DE CREDITO
            If NulosN(TxtIdMon.Text) = 1 Then
                GrabarOperacionCtaCteDocRef 1, xId, NulosC(TxtDocRef2.Text), NulosN(LblIdProveedor.Caption), NulosN(TxtTipDoc.Text), TxtNumSer.Text & "-" & TxtNumDoc.Text, _
                    TxtFchDoc.Valor, NulosN(TxtIdMon.Text), LblTipoCambio.Caption, 0, NulosN(TxtTotal.Text), 0, 0, xNumAsiento, xCon, , , NulosN(TxtTipDocRef), NulosN(LblIdDocRef2.Caption), 2
            Else
                GrabarOperacionCtaCteDocRef 1, xId, NulosC(TxtDocRef2.Text), NulosN(LblIdProveedor.Caption), NulosN(TxtTipDoc.Text), TxtNumSer.Text & "- " & TxtNumDoc.Text, _
                    TxtFchDoc.Valor, NulosN(TxtIdMon.Text), LblTipoCambio.Caption, 0, 0, 0, NulosN(TxtTotal.Text), xNumAsiento, xCon, , , NulosN(TxtTipDocRef), NulosN(LblIdDocRef2.Caption), 2
            End If
        Else
            ' SI ES IGUAL A NOTA DE CREDITO
            If NulosN(TxtIdMon.Text) = 1 Then
                GrabarOperacionCtaCteDocRef 1, xId, NulosC(TxtDocRef2.Text), NulosN(LblIdProveedor.Caption), NulosN(TxtTipDoc.Text), TxtNumSer.Text & "- " & TxtNumDoc.Text, _
                    TxtFchDoc.Valor, NulosN(TxtIdMon.Text), LblTipoCambio.Caption, NulosN(TxtTotal.Text), 0, 0, 0, xNumAsiento, xCon, , , NulosN(TxtTipDocRef), NulosN(LblIdDocRef2.Caption), 2
            Else
                GrabarOperacionCtaCteDocRef 1, xId, NulosC(TxtDocRef2.Text), NulosN(LblIdProveedor.Caption), NulosN(TxtTipDoc.Text), TxtNumSer.Text & "- " & TxtNumDoc.Text, _
                    TxtFchDoc.Valor, NulosN(TxtIdMon.Text), LblTipoCambio.Caption, 0, 0, NulosN(TxtTotal.Text), 0, xNumAsiento, xCon, , , NulosN(TxtTipDocRef), NulosN(LblIdDocRef2.Caption), 2
            
            End If
        End If
        
'        'grabamos el imponible en function a los items de la factura
'        Set Rst = Nothing
'        RST_Busq Rst, "SELECT com_comprasdet.idcom, alm_inventario.idcuenta, Sum(com_comprasdet.imptot) AS SumaDeimptot FROM alm_inventario INNER JOIN com_comprasdet " _
'            & " ON alm_inventario.id = com_comprasdet.iditem GROUP BY com_comprasdet.idcom, alm_inventario.idcuenta HAVING (((com_comprasdet.idcom)=" & xId & "))", xCon
'
'        If Rst.RecordCount <> 0 Then
'            Rst.MoveFirst
'            For A = 1 To Rst.RecordCount
'                RstDia.AddNew
'                RstDia("año") = AnoTra
'                RstDia("idmes") = mMesActivo             'LLAVE - CODIGO DEL MES
'                If mMesActivo = 0 Then
'                    RstDia("idlib") = 36                 'LLAVE - CODIGO DEL LIBRO
'                Else
'                    RstDia("idlib") = 1
'                End If
'                RstDia("idmov") = xId                    'LLAVE - CODIGO DEL MOVIMIENTO
'                RstDia("numasi") = xNumAsiento           'LLAVE - NUMERO DE ASIENTO
'                RstDia("tc") = ValTipCam
'                RstDia("idcue") = Rst("idcuenta")
'                RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'                RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'                If NulosN(TxtTipDoc.Text) <> 0 Then
'                    If NulosN(TxtTipDoc.Text) <> 7 Then
'                        If TxtIdMon.Text = "1" Then
'                            RstDia("impdebsol") = Format(Rst("SumaDeimptot"), "0.000000")
'                            RstDia("impdebdol") = 0
'                        Else
'                            RstDia("impdebsol") = Format(Rst("SumaDeimptot"), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.0000000")
'                            RstDia("impdebdol") = Format(Rst("SumaDeimptot"), "0.000000")
'                        End If
'                    Else
'                        If TxtIdMon.Text = "1" Then
'                            RstDia("imphabsol") = Format(Rst("SumaDeimptot"), "0.000000")
'                            RstDia("imphabdol") = 0
'                        Else
'                            RstDia("imphabsol") = Format(Rst("SumaDeimptot"), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.0000000")
'                            RstDia("imphabdol") = Format(Rst("SumaDeimptot"), "0.000000")
'                        End If
'                    End If
'                End If
'                RstDia.Update
'
'                Rst.MoveNext
'                If Rst.EOF = True Then Exit For
'            Next A
'        End If
        
        ' GRABAMOS LOS ASIENTOS AUTOMATICOS
        'grabamos la cuenta de destino debe
'        Set Rst = Nothing
'
'        RST_Busq Rst, "SELECT com_comprasdet.idcom, con_planctas.ctadesdeb, Sum(com_comprasdet.imptot) AS SumaDeimptot FROM con_planctas RIGHT JOIN (alm_inventario " _
'            & " INNER JOIN com_comprasdet ON alm_inventario.id = com_comprasdet.iditem) ON con_planctas.id = alm_inventario.idcuenta GROUP BY com_comprasdet.idcom, " _
'            & " con_planctas.ctadesdeb HAVING (((com_comprasdet.idcom)=" & xId & "))", xCon
'
'        If Rst.RecordCount <> 0 Then
'            Rst.MoveFirst
'            For A = 1 To Rst.RecordCount
'                If Rst("ctadesdeb") <> 0 Then
'                    RstDia.AddNew
'                    RstDia("año") = AnoTra
'                    RstDia("idmes") = mMesActivo         'LLAVE - CODIGO DEL MES
'                    RstDia("idlib") = 1                  'LLAVE - CODIGO DEL LIBRO
'                    RstDia("idmov") = xId                'LLAVE - CODIGO DEL MOVIMIENTO
'                    RstDia("numasi") = xNumAsiento       'LLAVE - NUMERO DE ASIENTO
'                    RstDia("tc") = ValTipCam
'                    RstDia("idcue") = Rst("ctadesdeb") 'xIdCuen
'                    RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'                    RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'                    If NulosN(TxtTipDoc.Text) <> 0 Then
'                        If NulosN(TxtTipDoc.Text) <> 7 Then
'                            If TxtIdMon.Text = "1" Then
'                                RstDia("impdebsol") = Format(Rst("SumaDeimptot"), "0.000000")
'                                RstDia("impdebdol") = 0
'                            Else
'                                RstDia("impdebsol") = Format(Rst("SumaDeimptot"), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.0000000")
'                                RstDia("impdebdol") = Format(Rst("SumaDeimptot"), "0.000000")
'                            End If
'                        Else
'                            If TxtIdMon.Text = "1" Then
'                                RstDia("imphabsol") = Format(Rst("SumaDeimptot"), "0.000000")
'                                RstDia("imphabdol") = 0
'                            Else
'                                RstDia("imphabsol") = Format(Rst("SumaDeimptot"), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.0000000")
'                                RstDia("imphabdol") = Format(Rst("SumaDeimptot"), "0.000000")
'                            End If
'                        End If
'                    End If
'                    RstDia.Update
'                End If
'
'                Rst.MoveNext
'                If Rst.EOF = True Then Exit For
'            Next A
'        End If
'
'        'grabamos la cuenta de destino haber
'        Set Rst = Nothing
'
'        RST_Busq Rst, "SELECT com_comprasdet.idcom, con_planctas.ctadeshab, Sum(com_comprasdet.imptot) AS SumaDeimptot FROM con_planctas RIGHT JOIN (alm_inventario " _
'            & " INNER JOIN com_comprasdet ON alm_inventario.id = com_comprasdet.iditem) ON con_planctas.id = alm_inventario.idcuenta GROUP BY com_comprasdet.idcom, " _
'            & " con_planctas.ctadeshab HAVING (((com_comprasdet.idcom)=" & xId & "))", xCon
'
'        If Rst.RecordCount <> 0 Then
'            Rst.MoveFirst
'            For A = 1 To Rst.RecordCount
'                If Rst("ctadeshab") <> 0 Then
'                    RstDia.AddNew
'                    RstDia("año") = AnoTra
'                    RstDia("idmes") = mMesActivo         'LLAVE - CODIGO DEL MES
'                    RstDia("idlib") = 1                  'LLAVE - CODIGO DEL LIBRO
'                    RstDia("idmov") = xId                'LLAVE - CODIGO DEL MOVIMIENTO
'                    RstDia("numasi") = xNumAsiento       'LLAVE - NUMERO DE ASIENTO
'                    RstDia("tc") = ValTipCam
'                    RstDia("idcue") = Rst("ctadeshab") 'xIdCuen
'                    RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'                    RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'                    If NulosN(TxtTipDoc.Text) <> 0 Then
'                        If NulosN(TxtTipDoc.Text) <> 7 Then
'                            If TxtIdMon.Text = "1" Then
'                                RstDia("imphabsol") = Format(Rst("SumaDeimptot"), "0.000000")
'                                RstDia("imphabdol") = 0
'                            Else
'                                RstDia("imphabsol") = Format(Rst("SumaDeimptot"), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.0000000")
'                                RstDia("imphabdol") = Format(Rst("SumaDeimptot"), "0.000000")
'                            End If
'                        Else
'                            If TxtIdMon.Text = "1" Then
'                                RstDia("impdebsol") = Format(Rst("SumaDeimptot"), "0.000000")
'                                RstDia("impdebdol") = 0
'                            Else
'                                RstDia("impdebsol") = Format(Rst("SumaDeimptot"), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.0000000")
'                                RstDia("impdebdol") = Format(Rst("SumaDeimptot"), "0.000000")
'                            End If
'                        End If
'                    End If
'                    RstDia.Update
'                End If
'
'                Rst.MoveNext
'                If Rst.EOF = True Then Exit For
'            Next A
'        End If
    
'    '----------------------------------------------------------
'    'grabamos el selectivo en funcion a los items de la factura
'    If RstTempISC.RecordCount <> 0 Then
'        RstTempISC.MoveFirst
'
'        For A = 1 To RstTempISC.RecordCount
'            RstDia.AddNew
'            RstDia("año") = AnoTra
'            RstDia("idmes") = mMesActivo         'LLAVE - CODIGO DEL MES
'            RstDia("idlib") = 1                  'LLAVE - CODIGO DEL LIBRO
'            RstDia("idmov") = xId                'LLAVE - CODIGO DEL MOVIMIENTO
'            RstDia("numasi") = xNumAsiento       'LLAVE - NUMERO DE ASIENTO
'            RstDia("tc") = ValTipCam
'            RstDia("idcue") = RstTempISC("idcuen")
'            RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'            RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'            If NulosN(TxtTipDoc.Text) <> 0 Then
'                If NulosN(TxtTipDoc.Text) <> 7 Then
'                    If TxtIdMon.Text = "1" Then
'                        RstDia("impdebsol") = Format(RstTempISC("total"), "0.000000")
'                        RstDia("impdebdol") = 0
'                    Else
'                        RstDia("impdebsol") = Format(RstTempISC("total"), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.000000")
'                        RstDia("impdebdol") = Format(RstTempISC("total"), "0.000000")
'                    End If
'                Else
'                    If TxtIdMon.Text = "1" Then
'                        RstDia("imphabsol") = Format(RstTempISC("total"), "0.000000")
'                        RstDia("imphabdol") = 0
'                    Else
'                        RstDia("imphabsol") = Format(RstTempISC("total"), "0.000000") * Format(NulosN(LblTipoCambio.Caption), "0.000000")
'                        RstDia("imphabdol") = Format(RstTempISC("total"), "0.000000")
'                    End If
'                End If
'            End If
'            RstDia.Update
'
'            RstTempISC.MoveNext
'
'            If RstTempISC.EOF = True Then
'                Exit For
'            End If
'        Next A
'    End If
    
    ' AVERIGUAMOS SI EL ITEM ESTA AFECTO A LA DETRACCION
    Set Rst = Nothing
''    RST_Busq Rst, "SELECT mae_detraccion.id, mae_detraccion.descripcion, mae_detraccion.tasa, alm_inventario.iddet " _
''        & " FROM alm_inventario LEFT JOIN mae_detraccion ON alm_inventario.iddet = mae_detraccion.id " _
''        & " WHERE ((alm_inventario.id= " & NulosN(Fg1.TextMatrix(Fg1.Row, 9)) & "))", xCon

    RST_Busq Rst, "SELECT mae_detraccion.id, mae_detraccion.descripcion, mae_detraccion.tasa, alm_inventario.iddet " _
        & " FROM com_comprasdet INNER JOIN (alm_inventario INNER JOIN mae_detraccion ON alm_inventario.iddet = mae_detraccion.id) ON com_comprasdet.iditem = alm_inventario.id " _
        & " WHERE (((com_comprasdet.idcom)=" & xId & "));", xCon

    If Rst.RecordCount <> 0 Then
        If Rst("iddet") <> 0 Then
            MsgBox "Se ha detectado que la compra registrada esta afecta al regimen de la Detraccion " + Chr(13) _
                & "Decripcion : " + Rst("descripcion") + Chr(13) _
                & "tasa : " + Format(Rst("tasa"), "0.00") + "%", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                
            Dim RstDeta As New ADODB.Recordset
            Dim xId2 As Integer
            
            If QueHace = 1 Then
                xId2 = HallaCodigoTabla("con_detraccion", xCon, "id")
                RST_Busq RstDeta, "SELECT * FROM con_detraccion", xCon
                RstDeta.AddNew
                RstDeta("id") = xId2
            Else
                RST_Busq RstDeta, "SELECT con_detraccion.* From con_detraccion " _
                    & " WHERE (((con_detraccion.iddoc)=" & xId & ")) and con_detraccion.tipo =1 ", xCon
            End If
            
            If RstDeta.RecordCount = 0 Then
                'este procedimiento es solo para cuando se este modificando una compra afecta a la detraccion y no se le haya hecho la detraccion
                'a la hora de ingresar la compra
                xId2 = HallaCodigoTabla("con_detraccion", xCon, "id")
                RstDeta.AddNew
                RstDeta("id") = xId2
            End If
            
            RstDeta("iddet") = Rst("iddet")
            RstDeta("por") = Rst("tasa")
            RstDeta("iddoc") = xId
            RstDeta("idmon") = NulosN(TxtIdMon.Text)
            RstDeta("tipo") = 1
            RstDeta("fchmov") = Date
            RstDeta("Glosa") = ""
            RstDeta("imp") = Format((NulosN(TxtTotal.Text) * (Rst("tasa") / 100)), "0.00")
            RstDeta("numdet") = "SIN NUMERO"
            RstDeta.Update
        End If
    End If
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    '-----------------------------------------------------------------------------------------------------------
    '--grabar datos adicionales en el diario
    nSQL = "UPDATE ((com_compras INNER JOIN con_diario ON com_compras.id = con_diario.idmov) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha  " _
        + vbCr + " SET con_diario.tc = [con_tc].[impven] , con_diario.fchdoc=com_compras.fchdoc, con_diario.idmon=com_compras.idmon, con_diario.ridlib = 1, con_diario.ridtipper = 1, con_diario.ridper = [com_compras].[idpro], con_diario.rtipdoc = [com_compras].[tipdoc], con_diario.rfchope = [com_compras].[fchdoc], con_diario.rnumerodoc = IIf([com_compras].[numser] Is Null Or [com_compras].[numser]='','',[com_compras].[numser] & '-') & [com_compras].[numdoc], con_diario.rglosaope = [com_compras].[glosa] & '', con_diario.rregistro = Left([com_compras].[numreg],2) & [mae_libros].[codsun] & Right([com_compras].[numreg],4) " _
        + vbCr + " WHERE (((con_diario.idlib)=1) AND ((con_diario.idmov)=" & xId & ")); "
        
    xCon.Execute nSQL
    '-----------------------------------------------------------------------------------------------------------
    
    xCon.CommitTrans
    MsgBox "La compra se registró con éxito" & vbCr & "Registro Nº: " & xNumAsiento, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    Set RstDeta = Nothing
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDia = Nothing
    Set RstCosto = Nothing
    Grabar = True
    Exit Function
    
LaCague:
    'Resume
    xCon.RollbackTrans
    Set RstDeta = Nothing
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDia = Nothing
    Set RstCosto = Nothing
    
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" & vbCr & Trim(Err.Description)
End Function

'*****************************************************************************************************
'* Nombre           : HallaNumAsiento
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : HALLA EL ULTIMO NUMERO DE ASIENTO DEL PERIODO ESPECIFICADO, DEVUELVE UNA CADENA
'* Paranetros       : NOMBRE    |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Mes       |  INTEGER          |  ESPECIFICA EL NUMERO DEL MES ACTUAL
'* Devuelve         : STRING
'*****************************************************************************************************
Function HallaNumAsiento(Mes As Integer) As String
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT con_diario.idmes, con_diario.idlib, con_diario.numasi From con_diario " _
        & " WHERE (((con_diario.idmes)=" & Mes & ") AND ((con_diario.idlib)=1)) ORDER BY numasi", xCon
    
    If Rst.RecordCount = 0 Then
        HallaNumAsiento = "0001"
    Else
        Rst.MoveLast
        HallaNumAsiento = Format(NulosN(Rst("numasi")) + 1, "0000")
    End If
    Exit Function
End Function

Private Sub TxtTipDoc_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtTipDoc.Text) = "" Then Exit Sub
    Dim xRs As New ADODB.Recordset
    
    RST_Busq xRs, "SELECT mae_documento.*, mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, con_planctas.cuenta, " _
        & " mae_impuestos.Abrev AS abreimp, mae_impuestos.idcuen as cuentaimp " _
        & " FROM mae_documento LEFT JOIN (mae_impuestos LEFT JOIN con_planctas ON mae_impuestos.idcuen = con_planctas.id) " _
        & " ON mae_documento.idimp = mae_impuestos.id WHERE mae_documento.id  = " & NulosN(TxtTipDoc.Text) & "", xCon
    
    If NulosN(TxtTipDoc.Text) = 2 Then
        Frame8.Visible = True
    Else
        Frame8.Visible = False
    End If
    
    If TxtTipDoc.Text = 7 Then
        Label3(9).Visible = True
        TxtDocRef.Visible = True
        CmdBusDocRef.Visible = True
    Else
        Label3(9).Visible = False
        TxtDocRef.Visible = False
        CmdBusDocRef.Visible = False
    End If
    
    If xRs.RecordCount = 0 Then
        TxtTipDoc.Text = ""
        LblNomDoc.Caption = ""
    Else
        CodSunatDoc = xRs("codsun")
        TxtTipDoc.Text = xRs("id")
        LblNomDoc.Caption = xRs("descripcion")
        TasaImpuesto = NulosN(xRs("tasa"))
        xDescImp = xRs("descripcion")
        xIdCuenTasa = NulosN(xRs("cuentaimp"))
        
        LblRotulo = Trim(NulosC(xRs("abreimp"))) + " (       )"
        LblIgvTasa.Caption = Trim(Str(TasaImpuesto)) ' + "%"
        xPorIgv = (TasaImpuesto / 100)
        
        Frame3.Caption = "( Afecta : " + NulosC(xRs("descimp")) + ")"
    End If
    
    Set xRs = Nothing
    xCuentaDoc = HallaNumCuenta(TxtTipDoc.Text, TxtIdMon.Text)
    If xCuentaDoc = 0 Then
        TxtTipDoc.Text = ""
        LblNomDoc.Caption = ""
    End If
    
End Sub



Sub Imprimir()
    Dim RsPDoc As New ADODB.Recordset
    Dim RsPCab As New ADODB.Recordset
    Dim RsPDet As New ADODB.Recordset
    Dim xRsDoc As New ADODB.Recordset
    Dim xRsDet As New ADODB.Recordset
    Dim RstGui As New ADODB.Recordset
    Dim A As Integer
    Dim xCadGuias As String

    RST_Busq xRsDoc, "SELECT com_compras.fchdoc, mae_prov.nombre, mae_prov.numdoc, com_compras.imptot, com_compras.tipdoc, com_compras.idmon, " _
        & " mae_prov.dir FROM mae_prov RIGHT JOIN com_compras " _
        & " ON mae_prov.id = com_compras.idpro Where (((com_compras.id) = " & RstComp("id") & "))", xCon
    
    RST_Busq xRsDet, "SELECT com_comprasdet.idcom, alm_inventario.descripcion, mae_unidades.abrev, com_comprasdet.canpro, com_comprasdet.preuni, " _
        & " com_comprasdet.imptot FROM mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN com_comprasdet ON alm_inventario.id = com_comprasdet.iditem) " _
        & " ON mae_unidades.id = com_comprasdet.idunimed WHERE (((com_comprasdet.idcom)=" & RstComp("id") & "))", xCon

    RST_Busq RsPDoc, "SELECT * FROM var_plantilladoc WHERE tipdoc = " & xRsDoc("tipdoc") & " ", xCon

    If RsPDoc.RecordCount = 0 Then
        MsgBox "No se ha definido la plantilla de impresion para este tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Set xRsDoc = Nothing
        Set xRsDet = Nothing
        Set RsPDoc = Nothing
        Exit Sub
    End If
    RST_Busq RsPCab, "SELECT * FROM var_plantilladoc WHERE tipdoc = " & RsPDoc("tipdoc") & " ", xCon
    If RsPCab.RecordCount <> 0 Then
        A = RsPCab("id")
        RST_Busq RsPCab, "SELECT * FROM var_plantillacab WHERE idplan = " & A & " ORDER BY item", xCon
        RST_Busq RsPDet, "SELECT * FROM var_plantilladet WHERE idplan = " & A & " ORDER BY item", xCon
    End If

    Printer.Font = "Super Draft 15cpi"
    Printer.FontBold = True
    Printer.FontSize = 11
    Printer.ScaleMode = 6

    Dim xCam, xFor As String

    'imprime cabezera
    Do While RsPCab.EOF = False
        xCam = RsPCab("campo")
        xFor = NulosC(RsPCab("formato"))

        Printer.CurrentX = RsPCab("posx")
        Printer.CurrentY = RsPCab("posy")

        If NulosC(UCase(xCam)) <> UCase("x-numeletra") And NulosC(UCase(xCam)) <> UCase("x-numguia") And NulosC(UCase(xCam)) <> UCase("x-docref") Then
            Printer.Print Format((NulosC(xRsDoc(xCam))), xFor)
        Else
            If NulosC(UCase(xCam)) = UCase("x-numeletra") Then
                Printer.Print "Son : "; NumeroLetra(xRsDoc("imptot"), xRsDoc("idmon"))
            End If
            If NulosC(UCase(xCam)) = UCase("x-numguia") Then
                Printer.Print xCadGuias
            End If
            If NulosC(UCase(xCam)) = UCase("x-docref") Then
                Printer.Print "Referente a Factura(s) : "; xRsDoc("docref")
            End If
        End If

        RsPCab.MoveNext
    Loop

    'imprime detalle
    Dim Fila As Integer

    Fila = RsPDet("posy")
    xRsDet.MoveFirst
    Do While xRsDet.EOF = False
        RsPDet.MoveFirst
        Do While RsPDet.EOF = False
            xCam = RsPDet("campo")
            xFor = NulosC(RsPDet("formato"))
            Printer.CurrentX = RsPDet("posx")
            Printer.CurrentY = Fila
            If xFor = "" Then
                Printer.Print NulosC(xRsDet(xCam))
            Else
                Printer.Print Format((NulosC(xRsDet(xCam))), xFor)
            End If
            RsPDet.MoveNext
        Loop
        Fila = Fila + 4

        xRsDet.MoveNext
    Loop

    Printer.EndDoc
End Sub

'*****************************************************************************************************
'* Nombre           : Buscar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PERMITE EFECTUAR UNA BUSQUEDA EN EL RECORDSET RstComp
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Buscar()
    TabOne1.CurrTab = 0
    Dim xRs As New ADODB.Recordset
    
    Dim nSQL As String
    Dim xCampos(8, 4) As String
    
    xCampos(0, 0) = "N°Reg":         xCampos(0, 1) = "numreg":     xCampos(0, 2) = "820":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "T.D.":          xCampos(1, 1) = "abrev":      xCampos(1, 2) = "400":   xCampos(1, 3) = "C"
    xCampos(2, 0) = "N°. Documento": xCampos(2, 1) = "numerodoc":  xCampos(2, 2) = "1400":  xCampos(2, 3) = "C"
    xCampos(3, 0) = "FchEmi":        xCampos(3, 1) = "fchdoc":     xCampos(3, 2) = "830":   xCampos(3, 3) = "C"
    xCampos(4, 0) = "FchVenc":       xCampos(4, 1) = "fchven":     xCampos(4, 2) = "830":   xCampos(4, 3) = "C"
    xCampos(5, 0) = "Proveedor":     xCampos(5, 1) = "nombre":     xCampos(5, 2) = "2600":  xCampos(5, 3) = "C"
    xCampos(6, 0) = "M":             xCampos(6, 1) = "simbolo":    xCampos(6, 2) = "450":    xCampos(6, 3) = "C"
    xCampos(7, 0) = "Importe":       xCampos(7, 1) = "imptot":     xCampos(7, 2) = "850":    xCampos(7, 3) = "N"
    
    nSQL = "SELECT com_compras.id,Mid([com_compras].[numreg],1,2)+[mae_libros].[codsun]+Mid([com_compras].[numreg],3,4) AS numreg, mae_prov.nombre, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numerodoc, mae_documento.abrev, format(com_compras.fchdoc,'dd/mm/yy') as fchdoc, format(com_compras.fchven,'dd/mm/yy') as fchven, mae_prov.numruc, mae_moneda.simbolo, com_compras.imptot, com_compras.impsal " _
        + vbCr + " FROM (mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN com_compras ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id " _
        + vbCr + " WHERE (((com_compras.numreg) Like '" & Format(mMesActivo, "00") & "%')) " _
        + vbCr + " ORDER BY com_compras.numreg DESC;"

    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Compras", "nombre", "nombre", Principio

    If xRs.State = 1 Then
        RstComp.MoveFirst
        RstComp.Find "id = " & xRs("id") & ""
    End If
    Set xRs = Nothing
End Sub

Private Sub TxtTipDocRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipDocRef_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipDocRef_Click
    End If
End Sub

Private Sub TxtTipDocRef_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    
    If NulosN(TxtTipDocRef.Text) = 0 Then
        TxtTipDocRef.Text = ""
        LblTipDocref.Caption = ""
        TxtDocRef2.Text = ""
        LblIdDocRef2.Caption = ""
        
        If SeSeleTipDocRef = False Then Exit Sub
        EliminarDatosCargados
        PermitirEdicion = True
        SeSeleTipDocRef = False
        Exit Sub
    End If
    
    Dim xRs1 As New ADODB.Recordset
    SeSeleTipDocRef = True
    RST_Busq xRs1, "SELECT * FROM mae_docreferencia WHERE id = " & NulosN(TxtTipDocRef.Text) & "", xCon
    
    If xRs1.RecordCount = 0 Then
        TxtTipDocRef.Text = ""
        LblTipDocref.Caption = ""
        TxtDocRef2.Text = ""
        LblIdDocRef2.Caption = ""
        EliminarDatosCargados
        PermitirEdicion = True
    Else
        LblTipDocref.Caption = Trim(xRs1("descripcion"))
        TxtDocRef2.Text = ""
        LblIdDocRef2.Caption = ""
    End If
    Set xRs1 = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : ActualizaSaldoDoc
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTUALIZA EL SALDO DEL DOCUMENTO, PARA ELLO HACE UNA CONSULTA DE ABONOS EN LA
'*                    Tes_caja
'* Paranetros       : NOMBRE        |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    idDocumento   |  Integer          |  ID DEL DOCUMENTO QUE SE VA ACTUALIZAR
'*                    Tabla         |  Integer          |  ID DE LA TABLA QUE SE ACTUALIZARA
'*                                  |                   |  1 = COMPRAS; 2 = VENTAS; 3 HONORARIOS
'*                    ImporteRestar |  Double           |
'* Devuelve         :
'* Observaciones    : ESTA FUNCION DEBERIA DE ESTAR EN LA CLASE Sgi2_Funciones
'*****************************************************************************************************
Sub ActualizaSaldoDoc(idDocumento As Double, Tabla As Integer, ImporteRestar As Double)
    '1 = compras
    '2 = Ventas
    '3 = honorarios
    
    Dim Rst As New ADODB.Recordset
    Dim Total As Double
    
    If Tabla = 1 Then
        ' ACTUALIZAMOS EL SALDO DE COMPRAS
        RST_Busq Rst, "SELECT Sum(tes_cajadestinodet.acuenta) AS total FROM tes_caja LEFT JOIN tes_cajadestinodet ON tes_caja.id = tes_cajadestinodet.idtes " _
            & " GROUP BY tes_cajadestinodet.iddoc, tes_caja.tipmov HAVING (((tes_cajadestinodet.iddoc)=" & idDocumento & ") AND ((tes_caja.tipmov)=2))", xCon
            
        Total = BuscaImporteDocumento(idDocumento, 1)
    End If
    
    If Rst.RecordCount <> 0 Then
        Total = ((Total - Rst("total")) - ImporteRestar)
    Else
        Total = (Total - ImporteRestar)
    End If
    
    xCon.Execute "UPDATE com_compras SET com_compras.impsal = " & Total & " WHERE (((com_compras.id)=" & idDocumento & "))"
    Set Rst = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : BuscaImporteDocumento
'* Tipo             : FUNCION
'* Descripcion      : DEVUELVE EL SALDO ACTUAL DE UN DOCUMENTO ESPECIFICADO, DEVUELVE UN VALOR NUMERICO
'*                    DE TIPO DOBLE
'* Paranetros       : NOMBRE      |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    idDocumento |  Integer          |  ESPECIFICA EL ID DEL DOCUMENTO
'*                    Tabla       |  Integer          |  ESPECIFICA EL ID DE LA TABLA EN QUE SE BUSCARA
'*                                |                   |  1 = COMPRAS; 2 = VENTAS; 3 = HONORARIOS
'* Devuelve         : DOUBLE
'*****************************************************************************************************
Function BuscaImporteDocumento(idDocumento As Double, Tabla As Integer) As Double
    '1 = compras
    '2 = Ventas
    '3 = honorarios
    Dim Rst As New ADODB.Recordset
    
    'compras
    If Tabla = 1 Then RST_Busq Rst, "SELECT * FROM com_compras WHERE id = " & idDocumento & "", xCon
    
    If Rst.RecordCount <> 0 Then
        BuscaImporteDocumento = Rst("imptot")
    Else
        BuscaImporteDocumento = 0
    End If
    
    Set Rst = Nothing
End Function

'*****************************************************************************************************
'* Nombre           : pGridConfigurar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CONFIGURA EL GRID DE INGRESO DE ITEMS DEL FOMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pGridConfigurar()
    If NulosN(TxtTipCom.Text) = 5 Then
        Fg1.ColWidth(2) = 0
        Fg1.ColWidth(3) = 0
        Fg1.ColWidth(4) = 1100
        Fg1.ColWidth(5) = 1100
        Fg1.ColWidth(8) = 1300
        If Fg1.Rows > 1 Then Fg1.TextMatrix(Fg1.Rows - 1, 3) = 1
        
    Else
        Fg1.ColWidth(2) = 420
        Fg1.ColWidth(3) = 1200
        Fg1.ColWidth(4) = 1005
        Fg1.ColWidth(5) = 1005
        Fg1.ColWidth(8) = 1095
    End If
End Sub

Private Sub VerAsiento()
    '===================================================================================================
    'Creado : 17/11/09 Por: Johan Castro
    'Propósito: Mostrar el asiento
    '
    'Entradas:  Tomara como base la informacion degistrada para que a partir de allí genere el asiento
    '
    'Resultados:Asiento en pantalla
    
    'Modificado: 15/10/10 Por Johan Castro
    '            Cuando el documento es nota de credito; las cuentas contables automaticas no se invertian
    '            como si lo hacia la cta 12, cta igv 40
    '===================================================================================================
    
    
    Dim RstAsi As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    Dim ValTipCam As Double
    Dim mFila As Long
    Dim mIdCta As Long '--codigo de la cuenta
    
    '--validar datos
    If IsDate(TxtFchDoc.Valor) = False Then
        MsgBox "Falta especificar la Fecha de Emision", vbInformation, xTitulo
        TxtFchDoc.SetFocus
        Exit Sub
    End If
    
    '---------------------------------------------------------------------------------------------------------------------------
    '--definir la estructura del rst
    RST_Busq RstTmp, "SELECT TOP 1 con_diario.idcue, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, con_diario.tc AS tipcam, con_diario.impdebsol AS impdebmn, con_diario.imphabsol AS imphabmn, con_diario.imphabdol AS impdebme, con_diario.imphabdol AS imphabme " _
                   & " FROM con_diario INNER JOIN con_planctas ON con_diario.idcue = con_planctas.id; ", xCon
    
    DEFINIR_RST_TMP RstAsi, RstTmp
   
    '---------------------------------------------------------------------------------------------------------------------------
    'ValTipCam = NulosN(LblTipoCambio.Caption)
    '--almacenar el tipo de cambio para hacer las conversiones mas adelante
    ValTipCam = HallaTipoCambio(TxtFchDoc.Valor, 2, Venta, xCon)
    If ValTipCam = 0 Then
        MsgBox "No hay tipo de Cambio" & vbCr & "Indique el tipo de cambio para continuar", vbInformation, xTitulo
        Exit Sub
    End If
        
    '-------------------------------------------------------------------------
    'grabamos a facturas por pagar Plan de cuentas 42.1 o dependiendo del caso
    RstAsi.AddNew
    RstAsi("idcue") = xCuentaDoc
    RstAsi("ctanum") = NulosC(Busca_Codigo(xCuentaDoc, "id", "cuenta", "con_planctas", "N", xCon))
    RstAsi("ctadesc") = NulosC(Busca_Codigo(xCuentaDoc, "id", "descripcion", "con_planctas", "N", xCon))
    RstAsi("tipcam") = ValTipCam

    If NulosN(TxtTipDoc.Text) <> 0 Then
        If NulosN(TxtTipDoc.Text) <> 7 Then
            'cuando se factura u otro comprabante excepto nota de credito hace su asiento norma
            If TxtIdMon.Text = "1" Then
                RstAsi("imphabmn") = NulosN(TxtTotal.Text)
                RstAsi("imphabme") = NulosN(TxtTotal.Text) / ValTipCam
            Else
                RstAsi("imphabmn") = NulosN(TxtTotal.Text) * ValTipCam
                RstAsi("imphabme") = NulosN(TxtTotal.Text)
            End If
        Else
            'cuando sea nota de credito hace el asiento inverso al de una venta
            If TxtIdMon.Text = "1" Then
                RstAsi("impdebmn") = NulosN(TxtTotal.Text)
                RstAsi("impdebme") = NulosN(TxtTotal.Text) / ValTipCam
            Else
                RstAsi("impdebmn") = NulosN(TxtTotal.Text) * ValTipCam
                RstAsi("impdebme") = NulosN(TxtTotal.Text)
            End If
        End If
    End If
    RstAsi.Update
        
    '-----------------------------------------------------
    'grabamos el impuesto si la operacion esta afecta a el
    If NulosN(TxtIGV.Text) <> 0 Then
        RstAsi.AddNew
        RstAsi("idcue") = xIdCuenTasa
        RstAsi("ctanum") = NulosC(Busca_Codigo(xIdCuenTasa, "id", "cuenta", "con_planctas", "N", xCon))
        RstAsi("ctadesc") = NulosC(Busca_Codigo(xIdCuenTasa, "id", "descripcion", "con_planctas", "N", xCon))
        RstAsi("tipcam") = ValTipCam
            
        'si el tipo del proveedor es diferente a no domiciliado
        If NulosN(LblIdTipPer.Caption) <> 3 Then
            If NulosN(TxtTipDoc.Text) <> 0 Then
                If NulosN(TxtTipDoc.Text) <> 7 And NulosN(TxtTipDoc.Text) <> 2 Then
                    If TxtIdMon.Text = "1" Then
                        RstAsi("impdebmn") = NulosN(TxtIGV.Text)
                        RstAsi("impdebme") = NulosN(TxtIGV.Text) / ValTipCam
                    Else
                        RstAsi("impdebmn") = NulosN(TxtIGV.Text) * ValTipCam
                        RstAsi("impdebme") = NulosN(TxtIGV.Text)
                    End If
                Else
                    If TxtIdMon.Text = "1" Then
                        RstAsi("imphabmn") = NulosN(TxtIGV.Text)
                        RstAsi("imphabme") = NulosN(TxtIGV.Text) / ValTipCam
                    Else
                        RstAsi("imphabmn") = NulosN(TxtIGV.Text) * ValTipCam
                        RstAsi("imphabme") = NulosN(TxtIGV.Text)
                    End If
                End If
            End If
        Else
            If TxtIdMon.Text = "1" Then
                RstAsi("imphabmn") = NulosN(TxtIGV.Text)
                RstAsi("imphabme") = NulosN(TxtIGV.Text) / ValTipCam
            Else
                RstAsi("imphabmn") = NulosN(TxtIGV.Text) * ValTipCam
                RstAsi("imphabme") = NulosN(TxtIGV.Text)
            End If
        End If
        RstAsi.Update
    End If
    '***********************************************************************
        
    'grabamos el impuesto si la operacion esta no afecta a el
    If NulosN(TxtIGV2.Text) <> 0 Then
        RstAsi.AddNew
        RstAsi("idcue") = xIdCuenTasa
        RstAsi("ctanum") = NulosC(Busca_Codigo(xIdCuenTasa, "id", "cuenta", "con_planctas", "N", xCon))
        RstAsi("ctadesc") = NulosC(Busca_Codigo(xIdCuenTasa, "id", "descripcion", "con_planctas", "N", xCon))
        RstAsi("tipcam") = ValTipCam
        'si el tipo de l proveedor es diferente a no domiciliado
        If NulosN(LblIdTipPer.Caption) <> 3 Then
            If NulosN(TxtTipDoc.Text) <> 0 Then
                If NulosN(TxtTipDoc.Text) <> 7 And NulosN(TxtTipDoc.Text) <> 2 Then
                    If TxtIdMon.Text = "1" Then
                        RstAsi("impdebmn") = NulosN(TxtIGV2.Text)
                        RstAsi("impdebme") = NulosN(TxtIGV2.Text) / ValTipCam
                    Else
                        RstAsi("impdebmn") = NulosN(TxtIGV2.Text) * ValTipCam
                        RstAsi("impdebme") = NulosN(TxtIGV2.Text)
                    End If
                Else
                    If TxtIdMon.Text = "1" Then
                        RstAsi("imphabmn") = NulosN(TxtIGV2.Text)
                        RstAsi("imphabme") = NulosN(TxtIGV2.Text) / ValTipCam
                    Else
                        RstAsi("imphabmn") = NulosN(TxtIGV2.Text) * ValTipCam
                        RstAsi("imphabme") = NulosN(TxtIGV2.Text)
                    End If
                End If
            End If
        Else
            If TxtIdMon.Text = "1" Then
                RstAsi("imphabmn") = NulosN(TxtIGV2.Text)
                RstAsi("imphabme") = NulosN(TxtIGV2.Text) / ValTipCam
            Else
                RstAsi("imphabmn") = NulosN(TxtIGV2.Text) * ValTipCam
                RstAsi("imphabme") = NulosN(TxtIGV2.Text)
            End If
        End If
        RstAsi.Update
    End If
        
    'grabamos el impuesto si la operacion sin derecho a credito fiscal
    If NulosN(TxtIGV3.Text) <> 0 Then
        RstAsi.AddNew
        RstAsi("idcue") = xIdCuenTasa
        RstAsi("ctanum") = NulosC(Busca_Codigo(xIdCuenTasa, "id", "cuenta", "con_planctas", "N", xCon))
        RstAsi("ctadesc") = NulosC(Busca_Codigo(xIdCuenTasa, "id", "descripcion", "con_planctas", "N", xCon))
        RstAsi("tipcam") = ValTipCam
        'si el tipo de l proveedor es diferente a no domiciliado
        If NulosN(LblIdTipPer.Caption) <> 3 Then
            If NulosN(TxtTipDoc.Text) <> 0 Then
                If NulosN(TxtTipDoc.Text) <> 7 And NulosN(TxtTipDoc.Text) <> 2 Then
                    If TxtIdMon.Text = "1" Then
                        RstAsi("impdebmn") = NulosN(TxtIGV3.Text)
                        RstAsi("impdebme") = NulosN(TxtIGV3.Text) / ValTipCam
                    Else
                        RstAsi("impdebmn") = NulosN(TxtIGV3.Text) * ValTipCam
                        RstAsi("impdebme") = NulosN(TxtIGV3.Text)
                    End If
                Else
                    If TxtIdMon.Text = "1" Then
                        RstAsi("imphabmn") = NulosN(TxtIGV3.Text)
                        RstAsi("imphabme") = NulosN(TxtIGV3.Text) / ValTipCam
                    Else
                        RstAsi("imphabmn") = NulosN(TxtIGV3.Text) * ValTipCam
                        RstAsi("imphabme") = NulosN(TxtIGV3.Text)
                    End If
                End If
            End If
        Else
            If TxtIdMon.Text = "1" Then
                RstAsi("imphabmn") = NulosN(TxtIGV3.Text)
                RstAsi("imphabme") = NulosN(TxtIGV3.Text) / ValTipCam
            Else
                RstAsi("imphabmn") = NulosN(TxtIGV3.Text) * ValTipCam
                RstAsi("imphabme") = NulosN(TxtIGV3.Text)
            End If
        End If
        RstAsi.Update
    
    End If
        
    '***********************************************************************
    
    'grabamos el impuesto si la operacion a sujeto a no domiciliado
    If NulosN(TxtOtros.Text) <> 0 And NulosN(TxtTipDoc.Text) = 107 Then
        RstAsi.AddNew
        RstAsi("idcue") = xIdCuenTasa
        RstAsi("ctanum") = NulosC(Busca_Codigo(xIdCuenTasa, "id", "cuenta", "con_planctas", "N", xCon))
        RstAsi("ctadesc") = NulosC(Busca_Codigo(xIdCuenTasa, "id", "descripcion", "con_planctas", "N", xCon))
        RstAsi("tipcam") = ValTipCam
        If NulosN(TxtTipDoc.Text) <> 0 Then
            If NulosN(TxtTipDoc.Text) <> 7 Then
                'cuando se factura u otro comprabante excepto nota de credito hace su asiento norma
                If TxtIdMon.Text = "1" Then
                    RstAsi("imphabmn") = NulosN(TxtOtros.Text)
                    RstAsi("imphabme") = NulosN(TxtOtros.Text) / ValTipCam
                Else
                    RstAsi("imphabmn") = NulosN(TxtOtros.Text) * ValTipCam
                    RstAsi("imphabme") = NulosN(TxtOtros.Text)
                End If
            Else
                'cuando sea nota de credito hace el asiento inverso al de una venta
                If TxtIdMon.Text = "1" Then
                    RstAsi("impdebmn") = NulosN(TxtOtros.Text)
                    RstAsi("impdebme") = NulosN(TxtOtros.Text) / ValTipCam
                Else
                    RstAsi("impdebmn") = NulosN(TxtOtros.Text) * ValTipCam
                    RstAsi("impdebme") = NulosN(TxtOtros.Text)
                End If
            End If
        End If
        RstAsi.Update
    End If
                
    '***********************************************************************
    '--grabamos el imponible en function a los items de la factura
    For mFila = 1 To Fg1.Rows - 1
        mIdCta = NulosN(Fg1.TextMatrix(mFila, 11))
        If mIdCta <> 0 Then
            RstAsi.AddNew
            RstAsi("idcue") = mIdCta
            RstAsi("ctanum") = NulosC(Busca_Codigo(mIdCta, "id", "cuenta", "con_planctas", "N", xCon))
            RstAsi("ctadesc") = NulosC(Busca_Codigo(mIdCta, "id", "descripcion", "con_planctas", "N", xCon))
            RstAsi("tipcam") = ValTipCam
                
            If NulosN(TxtTipDoc.Text) <> 0 Then
                If NulosN(TxtTipDoc.Text) <> 7 Then
                    If TxtIdMon.Text = "1" Then
                        RstAsi("impdebmn") = NulosN(Fg1.TextMatrix(mFila, 8))
                        RstAsi("impdebme") = NulosN(Fg1.TextMatrix(mFila, 8)) / ValTipCam
                    Else
                        RstAsi("impdebmn") = NulosN(Fg1.TextMatrix(mFila, 8)) * ValTipCam
                        RstAsi("impdebme") = NulosN(Fg1.TextMatrix(mFila, 8))
                    End If
                Else
                    If TxtIdMon.Text = "1" Then
                        RstAsi("imphabmn") = NulosN(Fg1.TextMatrix(mFila, 8))
                        RstAsi("imphabme") = NulosN(Fg1.TextMatrix(mFila, 8)) / ValTipCam
                    Else
                        RstAsi("imphabmn") = NulosN(Fg1.TextMatrix(mFila, 8)) * ValTipCam
                        RstAsi("imphabme") = NulosN(Fg1.TextMatrix(mFila, 8))
                    End If
                End If
            End If
            RstAsi.Update
        End If
        
    Next mFila
        
    '***********************************************************************
    
    '--mostramos los asientos automaticos
    For mFila = 1 To Fg1.Rows - 1
        '--cta debe
        mIdCta = NulosN(Busca_Codigo(NulosN(Fg1.TextMatrix(mFila, 11)), "id", "ctadesdeb", "con_planctas", "N", xCon))
        If mIdCta <> 0 Then
            RstAsi.AddNew
            RstAsi("idcue") = mIdCta
            RstAsi("ctanum") = Busca_Codigo(mIdCta, "id", "cuenta", "con_planctas", "N", xCon)
            RstAsi("ctadesc") = Busca_Codigo(mIdCta, "id", "descripcion", "con_planctas", "N", xCon)
            RstAsi("tipcam") = ValTipCam
            
            If NulosN(TxtTipDoc.Text) <> 7 Then
                If TxtIdMon.Text = "1" Then
                    RstAsi("impdebmn") = NulosN(Fg1.TextMatrix(mFila, 8))
                    RstAsi("impdebme") = NulosN(Fg1.TextMatrix(mFila, 8)) / ValTipCam
                Else
                    RstAsi("impdebmn") = NulosN(Fg1.TextMatrix(mFila, 8)) * ValTipCam
                    RstAsi("impdebme") = NulosN(Fg1.TextMatrix(mFila, 8))
                End If
                RstAsi("imphabmn") = 0
                RstAsi("imphabme") = 0
            Else
                If TxtIdMon.Text = "1" Then
                    RstAsi("imphabmn") = NulosN(Fg1.TextMatrix(mFila, 8))
                    RstAsi("imphabme") = NulosN(Fg1.TextMatrix(mFila, 8)) / ValTipCam
                Else
                    RstAsi("imphabmn") = NulosN(Fg1.TextMatrix(mFila, 8)) * ValTipCam
                    RstAsi("imphabme") = NulosN(Fg1.TextMatrix(mFila, 8))
                End If
                RstAsi("impdebmn") = 0
                RstAsi("impdebme") = 0

            End If
            
            
                
            RstAsi.Update
        End If
        
        '-----------------
        '--cta haber
        mIdCta = NulosN(Busca_Codigo(NulosN(Fg1.TextMatrix(mFila, 11)), "id", "ctadeshab", "con_planctas", "N", xCon))
        If mIdCta <> 0 Then
            RstAsi.AddNew
            RstAsi("idcue") = mIdCta
            RstAsi("ctanum") = Busca_Codigo(mIdCta, "id", "cuenta", "con_planctas", "N", xCon)
            RstAsi("ctadesc") = Busca_Codigo(mIdCta, "id", "descripcion", "con_planctas", "N", xCon)
            RstAsi("tipcam") = ValTipCam
            
            If NulosN(TxtTipDoc.Text) <> 7 Then
                If TxtIdMon.Text = "1" Then
                    RstAsi("imphabmn") = NulosN(Fg1.TextMatrix(mFila, 8))
                    RstAsi("imphabme") = NulosN(Fg1.TextMatrix(mFila, 8)) / ValTipCam
                Else
                    RstAsi("imphabmn") = NulosN(Fg1.TextMatrix(mFila, 8)) * ValTipCam
                    RstAsi("imphabme") = NulosN(Fg1.TextMatrix(mFila, 8))
                End If
                RstAsi("impdebmn") = 0
                RstAsi("impdebme") = 0
            Else
                If TxtIdMon.Text = "1" Then
                    RstAsi("impdebmn") = NulosN(Fg1.TextMatrix(mFila, 8))
                    RstAsi("impdebme") = NulosN(Fg1.TextMatrix(mFila, 8)) / ValTipCam
                Else
                    RstAsi("impdebmn") = NulosN(Fg1.TextMatrix(mFila, 8)) * ValTipCam
                    RstAsi("impdebme") = NulosN(Fg1.TextMatrix(mFila, 8))
                End If
                RstAsi("imphabmn") = 0
                RstAsi("imphabme") = 0
            End If
            RstAsi.Update
        End If
        
    Next mFila
    '***********************************************************************
    'grabamos el selectivo en funcion a los items de la factura
    '--pendiente de implementar 17/11/09
'    If RstTempISC.RecordCount <> 0 Then
'        RstTempISC.MoveFirst
'
'        dowhile Not RstTempISC.EOF
'            mIdCta = RstTempISC("idcuen")
'            RstAsi.AddNew
'            RstAsi("idcue") = mIdCta
'            RstAsi("ctanum") = Busca_Codigo(mIdCta, "id", "cuenta", "con_planctas", "N", xCon)
'            RstAsi("ctadesc") = Busca_Codigo(mIdCta, "id", "descripcion", "con_planctas", "N", xCon)
'            RstAsi("tipcam") = ValTipCam
'            If NulosN(TxtTipDoc.Text) <> 0 Then
'                If NulosN(TxtTipDoc.Text) <> 7 Then
'                    If TxtIdMon.Text = "1" Then
'                        RstAsi("impdebmn") = RstTempISC("total")
'                        RstAsi("impdebme") = RstTempISC("total") / ValTipCam
'                    Else
'                        RstAsi("impdebmn") = RstTempISC("total") * ValTipCam
'                        RstAsi("impdebme") = RstTempISC("total")
'                    End If
'                Else
'                    If TxtIdMon.Text = "1" Then
'                        RstAsi("imphabmn") = RstTempISC("total")
'                        RstAsi("imphabme") = RstTempISC("total") / ValTipCam
'                    Else
'                        RstAsi("imphabmn") = RstTempISC("total") * ValTipCam
'                        RstAsi("imphabme") = RstTempISC("total")
'                    End If
'                End If
'            End If
'            RstAsi.Update
'
'            RstTempISC.MoveNext
'
'        Loop
'    End If

    '--mostrar el asiento
    Dim xfrm As New SGI2_funciones.formularios
    Dim xId As Double
    
    '--verificar que accion se esta haciendo
    If QueHace = 1 Then
        xId = 0
    Else
        xId = RstComp("id")
    End If
    
    xfrm.AsientoVerTmp xCon, RstAsi, 1, xId
    Set xfrm = Nothing

End Sub

Sub ExportarExcel()
    '-- solo exporta datos de la grilla
    '--mo se usa, cambiado el 14/10/10 por pExportar()
    Dim xFun As New eps_librerias.FuncionesDGrid
    xFun.xNomEmp = NomEmp
    xFun.xNumRuc = NumRUC
    xFun.ExportarDGExcel RstComp, Dg1, "DOCUMENTOS DE COMPRA DEL MES DE " & UCase(LblMes.Caption)
    Set xFun = Nothing
End Sub

Private Sub pExportar()
    
    TabOne1.CurrTab = 0

    Dim nSQL As String
    Dim oExport As New SGI2_funciones.formularios
    Dim RstTmp  As New ADODB.Recordset

    Dim xCampos(22, 3) As String
    
    '0::Nombre a Mostrar;
    '1::nombre de Campo del Rst;
    '2::alineacion(0::derecha, 1::centro, 2::izquierda);
    '3::ancho de columna
    '--obs: el rst puede tener mas columnas solo se consideran los campos del array
    xCampos(0, 0) = "Id":           xCampos(0, 1) = "id":           xCampos(0, 2) = 2:   xCampos(0, 3) = "500"
    xCampos(1, 0) = "Nº Reg":       xCampos(1, 1) = "numreg1":      xCampos(1, 2) = 0:   xCampos(1, 3) = "900"
    xCampos(2, 0) = "R.U.C.":       xCampos(2, 1) = "numruc":       xCampos(2, 2) = 0:   xCampos(2, 3) = "1200"
    xCampos(3, 0) = "Proveedor":    xCampos(3, 1) = "nombre":       xCampos(3, 2) = 0:   xCampos(3, 3) = "3290"
    xCampos(4, 0) = "T.D.":         xCampos(4, 1) = "abrev":        xCampos(4, 2) = 0:   xCampos(4, 3) = "350"
    xCampos(5, 0) = "Num. Doc":     xCampos(5, 1) = "numerodoc":    xCampos(5, 2) = 0:   xCampos(5, 3) = "1600"
    xCampos(6, 0) = "Fch.Emi":      xCampos(6, 1) = "fchdoc1":      xCampos(6, 2) = 1:   xCampos(6, 3) = "900"
    xCampos(7, 0) = "Fch. Venc":    xCampos(7, 1) = "fchven1":      xCampos(7, 2) = 1:   xCampos(7, 3) = "900"
    xCampos(8, 0) = "Glosa":        xCampos(8, 1) = "glosa":        xCampos(8, 2) = 0:   xCampos(8, 3) = "2000"
    xCampos(9, 0) = "M":            xCampos(9, 1) = "simbolo":      xCampos(9, 2) = 1:   xCampos(9, 3) = "500"
    xCampos(10, 0) = "T.C.":        xCampos(10, 1) = "impven1":     xCampos(10, 2) = 2:  xCampos(10, 3) = "700"
    xCampos(11, 0) = "Imp Bru1":    xCampos(11, 1) = "impbru":      xCampos(11, 2) = 2:  xCampos(11, 3) = "900"
    xCampos(12, 0) = "Imp Bru2":    xCampos(12, 1) = "impbru2":     xCampos(12, 2) = 2:  xCampos(12, 3) = "900"
    xCampos(13, 0) = "Imp Bru3":    xCampos(13, 1) = "impbru3":     xCampos(13, 2) = 2:  xCampos(13, 3) = "900"
    xCampos(14, 0) = "Imp Inaf":    xCampos(14, 1) = "impina":      xCampos(14, 2) = 2:  xCampos(14, 3) = "900"
    xCampos(15, 0) = "Descuento":   xCampos(15, 1) = "impdesc":     xCampos(15, 2) = 2:  xCampos(15, 3) = "900"
    xCampos(16, 0) = "Imp ISC":     xCampos(16, 1) = "impisc":      xCampos(16, 2) = 2:  xCampos(16, 3) = "900"
    xCampos(17, 0) = "Imp Igv1":    xCampos(17, 1) = "impigv":      xCampos(17, 2) = 2:  xCampos(17, 3) = "900"
    xCampos(18, 0) = "Imp Igv2":    xCampos(18, 1) = "impigv2":     xCampos(18, 2) = 2:  xCampos(18, 3) = "900"
    xCampos(19, 0) = "Imp Igv3":    xCampos(19, 1) = "impigv3":     xCampos(19, 2) = 2:  xCampos(19, 3) = "900"
    xCampos(20, 0) = "Imp Otros":   xCampos(20, 1) = "otroscargos": xCampos(20, 2) = 2:  xCampos(20, 3) = "900"
    xCampos(21, 0) = "Imp Total":   xCampos(21, 1) = "imptot":      xCampos(21, 2) = 2:  xCampos(21, 3) = "1000"
    xCampos(22, 0) = "Imp Saldo":   xCampos(22, 1) = "impsal":      xCampos(22, 2) = 2:  xCampos(22, 3) = "1000"
    
    Set RstTmp = RstComp.Clone
    
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "LISTADO DE COMPRAS", "Periodo " & LblMes.Caption, "", "Listado de Compras - " & LblMes.Caption, RstTmp, xCampos
    Set oExport = Nothing
    Set RstTmp = Nothing
    
End Sub



Private Sub Frame11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OrigFX = X
    OrigFY = Y
    Frame11.ZOrder 0
End Sub

Private Sub Frame11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        With Frame11
            .Move .Left + X - OrigFX, .Top + Y - OrigFY
        End With
    End If
End Sub

Private Sub Frame6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OrigFX = X
    OrigFY = Y
    Frame6.ZOrder 0
End Sub

Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        With Frame6
            .Move .Left + X - OrigFX, .Top + Y - OrigFY
        End With
    End If
End Sub

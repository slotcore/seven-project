VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmManReceta2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Produccion - Mantenimiento de Recetas"
   ClientHeight    =   7260
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   11205
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmDuplicaReceta 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4215
      Left            =   11520
      TabIndex        =   35
      Top             =   690
      Visible         =   0   'False
      Width           =   9330
      Begin VB.Frame Frame10 
         Height          =   675
         Left            =   120
         TabIndex        =   40
         Top             =   3480
         Width           =   9090
         Begin VB.CommandButton CmdGenerar 
            Caption         =   "&Generar"
            Height          =   390
            Left            =   3165
            TabIndex        =   42
            Top             =   195
            Width           =   1350
         End
         Begin VB.CommandButton CmdCancelar 
            Caption         =   "&Cancelar"
            Height          =   390
            Left            =   4545
            TabIndex        =   41
            Top             =   195
            Width           =   1350
         End
      End
      Begin VB.Frame Frame7 
         Height          =   2550
         Left            =   7770
         TabIndex        =   37
         Top             =   975
         Width           =   1440
         Begin VB.CommandButton CmdAgregar 
            Caption         =   "&Agregar"
            Height          =   435
            Left            =   180
            TabIndex        =   39
            Top             =   840
            Width           =   1095
         End
         Begin VB.CommandButton CmdEliminar 
            Caption         =   "Eliminar"
            Height          =   435
            Left            =   180
            TabIndex        =   38
            Top             =   1290
            Width           =   1095
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg4 
         Height          =   2445
         Left            =   120
         TabIndex        =   36
         Top             =   1065
         Width           =   7590
         _cx             =   13388
         _cy             =   4313
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
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmManReceta2.frx":0000
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
         Editable        =   2
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
      Begin VB.Label LblCodPro 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblCodPro"
         Height          =   300
         Left            =   1185
         TabIndex        =   49
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "Codigo"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   48
         Top             =   390
         Width           =   795
      End
      Begin VB.Label LblDescPro 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblDescPro"
         Height          =   300
         Left            =   1185
         TabIndex        =   47
         Top             =   675
         Width           =   8010
      End
      Begin VB.Label Label8 
         Caption         =   "Descripcion"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   46
         Top             =   705
         Width           =   855
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -15
         X2              =   9315
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
         Y2              =   4185
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   9315
         X2              =   9300
         Y1              =   15
         Y2              =   4260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   15
         X2              =   9345
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label7 
         BackColor       =   &H00800000&
         Height          =   270
         Left            =   45
         TabIndex        =   45
         Top             =   45
         Width           =   9240
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Duplicar Recetas"
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
         Left            =   165
         TabIndex        =   44
         Top             =   75
         Width           =   1485
      End
      Begin VB.Label LblIdProd 
         AutoSize        =   -1  'True
         Caption         =   "LblIdProd"
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
         Height          =   195
         Left            =   3645
         TabIndex        =   43
         Top             =   390
         Width           =   825
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7470
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6930
      Left            =   15
      TabIndex        =   0
      Top             =   375
      Width           =   11190
      _cx             =   19738
      _cy             =   12224
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
      Appearance      =   2
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
         BorderStyle     =   0  'None
         Height          =   6510
         Left            =   11835
         TabIndex        =   4
         Top             =   375
         Width           =   11100
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   3060
            Left            =   15
            TabIndex        =   21
            Top             =   3450
            Width           =   11085
            _cx             =   19553
            _cy             =   5397
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
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
            Caption         =   "     Insumos     |       Tareas       |   Centro de Costo    "
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   0
            Position        =   1
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
            Begin VB.Frame Frame3 
               BorderStyle     =   0  'None
               Caption         =   "Frame6"
               Height          =   2640
               Left            =   12030
               TabIndex        =   50
               Top             =   45
               Width           =   10995
               Begin VB.Frame Frame11 
                  Height          =   2685
                  Left            =   9735
                  TabIndex        =   51
                  Top             =   -60
                  Width           =   1275
                  Begin VB.CommandButton cmdCos 
                     Caption         =   "Eliminar Todos"
                     Enabled         =   0   'False
                     Height          =   450
                     Index           =   3
                     Left            =   60
                     TabIndex        =   56
                     TabStop         =   0   'False
                     ToolTipText     =   "Eliminar Tarea"
                     Top             =   1710
                     Width           =   1110
                  End
                  Begin VB.CommandButton cmdCos 
                     Caption         =   "Eliminar"
                     Enabled         =   0   'False
                     Height          =   450
                     Index           =   2
                     Left            =   60
                     TabIndex        =   55
                     TabStop         =   0   'False
                     ToolTipText     =   "Eliminar Tarea"
                     Top             =   1200
                     Width           =   1110
                  End
                  Begin VB.CommandButton cmdCos 
                     Caption         =   "Agregar"
                     Enabled         =   0   'False
                     Height          =   450
                     Index           =   0
                     Left            =   60
                     TabIndex        =   53
                     ToolTipText     =   "Agregar Tarea"
                     Top             =   180
                     Width           =   1110
                  End
                  Begin VB.CommandButton cmdCos 
                     Caption         =   "Seleccionar"
                     Enabled         =   0   'False
                     Height          =   450
                     Index           =   1
                     Left            =   60
                     TabIndex        =   52
                     TabStop         =   0   'False
                     ToolTipText     =   "Eliminar Tarea"
                     Top             =   690
                     Width           =   1110
                  End
               End
               Begin VSFlex7Ctl.VSFlexGrid fg5 
                  Height          =   2580
                  Left            =   -30
                  TabIndex        =   54
                  Top             =   0
                  Width           =   9690
                  _cx             =   17092
                  _cy             =   4551
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
                  AllowUserResizing=   1
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   2
                  Cols            =   4
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmManReceta2.frx":003D
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
            End
            Begin VB.Frame Frame6 
               BorderStyle     =   0  'None
               Caption         =   "Frame6"
               Height          =   2640
               Left            =   11730
               TabIndex        =   23
               Top             =   45
               Width           =   10995
               Begin VB.Frame Frame9 
                  Height          =   2685
                  Left            =   9735
                  TabIndex        =   29
                  Top             =   -60
                  Width           =   1275
                  Begin VB.CommandButton cmdTar 
                     Caption         =   "Exportar"
                     Enabled         =   0   'False
                     Height          =   450
                     Index           =   2
                     Left            =   60
                     TabIndex        =   32
                     TabStop         =   0   'False
                     ToolTipText     =   "Eliminar Tarea"
                     Top             =   1680
                     Width           =   1110
                  End
                  Begin VB.CommandButton cmdTar 
                     Caption         =   "Eliminar"
                     Enabled         =   0   'False
                     Height          =   450
                     Index           =   1
                     Left            =   60
                     TabIndex        =   31
                     TabStop         =   0   'False
                     ToolTipText     =   "Eliminar Tarea"
                     Top             =   1140
                     Width           =   1110
                  End
                  Begin VB.CommandButton cmdTar 
                     Caption         =   "Agregar"
                     Enabled         =   0   'False
                     Height          =   450
                     Index           =   0
                     Left            =   60
                     TabIndex        =   30
                     ToolTipText     =   "Agregar Tarea"
                     Top             =   600
                     Width           =   1110
                  End
               End
               Begin VSFlex7Ctl.VSFlexGrid Fg3 
                  Height          =   2580
                  Left            =   0
                  TabIndex        =   25
                  Top             =   30
                  Width           =   9690
                  _cx             =   17092
                  _cy             =   4551
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
                  AllowUserResizing=   1
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
                  FormatString    =   $"FrmManReceta2.frx":00BF
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
            End
            Begin VB.Frame Frame5 
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   2640
               Left            =   45
               TabIndex        =   22
               Top             =   45
               Width           =   10995
               Begin VSFlex7Ctl.VSFlexGrid Fg1 
                  Height          =   2595
                  Left            =   0
                  TabIndex        =   24
                  Top             =   15
                  Width           =   9690
                  _cx             =   17092
                  _cy             =   4577
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
                  AllowUserResizing=   1
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   1
                  Cols            =   10
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmManReceta2.frx":0276
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
               Begin VB.Frame Frame8 
                  Height          =   2685
                  Left            =   9720
                  TabIndex        =   26
                  Top             =   -75
                  Width           =   1275
                  Begin VB.CommandButton cmdIngr 
                     Caption         =   "Exportar"
                     Enabled         =   0   'False
                     Height          =   450
                     Index           =   2
                     Left            =   75
                     TabIndex        =   33
                     TabStop         =   0   'False
                     ToolTipText     =   "Eliminar Ingrediente"
                     Top             =   1740
                     Width           =   1125
                  End
                  Begin VB.CommandButton cmdIngr 
                     Caption         =   "Agregar"
                     Enabled         =   0   'False
                     Height          =   450
                     Index           =   0
                     Left            =   75
                     TabIndex        =   28
                     ToolTipText     =   "Agregar Ingrediente"
                     Top             =   720
                     Width           =   1125
                  End
                  Begin VB.CommandButton cmdIngr 
                     Caption         =   "&Eliminar"
                     Enabled         =   0   'False
                     Height          =   450
                     Index           =   1
                     Left            =   75
                     TabIndex        =   27
                     TabStop         =   0   'False
                     ToolTipText     =   "Eliminar Ingrediente"
                     Top             =   1230
                     Width           =   1125
                  End
               End
            End
         End
         Begin VB.Frame Frame4 
            Height          =   1590
            Left            =   9810
            TabIndex        =   18
            Top             =   1110
            Width           =   1275
            Begin VB.CommandButton CmdVerProc 
               Caption         =   "Modificar Procedimiento"
               Enabled         =   0   'False
               Height          =   450
               Left            =   75
               TabIndex        =   34
               ToolTipText     =   "Agregar Receta"
               Top             =   150
               Width           =   1125
            End
            Begin VB.CommandButton CmdAddReceta 
               Caption         =   "Agregar Receta"
               Enabled         =   0   'False
               Height          =   450
               Left            =   75
               TabIndex        =   20
               ToolTipText     =   "Agregar Receta"
               Top             =   615
               Width           =   1125
            End
            Begin VB.CommandButton CmdDelReceta 
               Caption         =   "Eliminar Receta"
               Enabled         =   0   'False
               Height          =   450
               Left            =   75
               TabIndex        =   19
               ToolTipText     =   "Eliminar Receta"
               Top             =   1095
               Width           =   1125
            End
         End
         Begin VB.TextBox TxtNotas 
            Height          =   525
            Left            =   75
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Text            =   "FrmManReceta2.frx":03A7
            Top             =   2895
            Width           =   11010
         End
         Begin VB.TextBox TxtUnidad 
            Height          =   300
            Left            =   9810
            Locked          =   -1  'True
            TabIndex        =   14
            Text            =   "TxtUnidad"
            Top             =   615
            Width           =   1065
         End
         Begin VB.CommandButton CmdUsuario 
            Height          =   240
            Left            =   3600
            Picture         =   "FrmManReceta2.frx":03B0
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   330
            Width           =   240
         End
         Begin VB.TextBox TxtCodigo 
            Height          =   300
            Left            =   1815
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "TxtCodigo"
            Top             =   300
            Width           =   2055
         End
         Begin VB.TextBox TxtDescripcion 
            Height          =   300
            Left            =   1815
            Locked          =   -1  'True
            TabIndex        =   6
            Text            =   "TxtDescripcion"
            Top             =   615
            Width           =   6945
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg2 
            Height          =   1485
            Left            =   75
            TabIndex        =   15
            Top             =   1200
            Width           =   9690
            _cx             =   17092
            _cy             =   2619
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
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmManReceta2.frx":04E2
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
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Observaciones de la Receta"
            Height          =   195
            Left            =   105
            TabIndex        =   17
            Top             =   2700
            Width           =   2025
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            Height          =   195
            Left            =   105
            TabIndex        =   13
            Top             =   345
            Width           =   495
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Recetas Disponible"
            Height          =   195
            Left            =   105
            TabIndex        =   10
            Top             =   975
            Width           =   1440
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unidad"
            Height          =   195
            Left            =   9090
            TabIndex        =   9
            Top             =   660
            Width           =   510
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción Producto"
            Height          =   195
            Left            =   105
            TabIndex        =   8
            Top             =   660
            Width           =   1530
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Recetas del Producto"
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
            Left            =   15
            TabIndex        =   7
            Top             =   30
            Width           =   11070
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6510
         Left            =   45
         TabIndex        =   1
         Top             =   375
         Width           =   11100
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6195
            Left            =   45
            TabIndex        =   2
            Top             =   315
            Width           =   11040
            _ExtentX        =   19473
            _ExtentY        =   10927
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "IdItem"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Código"
            Columns(1).DataField=   "codpro"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Descripción"
            Columns(2).DataField=   "descripcion"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Unidad"
            Columns(3).DataField=   "abrev"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Nº Recetas"
            Columns(4).DataField=   "numrec"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Receta Princ."
            Columns(5).DataField=   "recpri"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=3387"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=3307"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=9525"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=9446"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=1296"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1217"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=1984"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1905"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=2223"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=2143"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HE0FEFE&,.bold=0,.fontsize=825"
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
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H80&"
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
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H80&"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80&"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
            _StyleDefs(60)  =   "Named:id=33:Normal"
            _StyleDefs(61)  =   ":id=33,.parent=0"
            _StyleDefs(62)  =   "Named:id=34:Heading"
            _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(64)  =   ":id=34,.wraptext=-1"
            _StyleDefs(65)  =   "Named:id=35:Footing"
            _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(67)  =   "Named:id=36:Selected"
            _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(69)  =   "Named:id=37:Caption"
            _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(71)  =   "Named:id=38:HighlightRow"
            _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=39:EvenRow"
            _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(75)  =   "Named:id=40:OddRow"
            _StyleDefs(76)  =   ":id=40,.parent=33"
            _StyleDefs(77)  =   "Named:id=41:RecordSelector"
            _StyleDefs(78)  =   ":id=41,.parent=34"
            _StyleDefs(79)  =   "Named:id=42:FilterBar"
            _StyleDefs(80)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Productos Disponibles"
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
            Left            =   15
            TabIndex        =   3
            Top             =   30
            Width           =   11070
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
            Picture         =   "FrmManReceta2.frx":0615
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManReceta2.frx":0B59
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManReceta2.frx":0CDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManReceta2.frx":1131
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManReceta2.frx":1249
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManReceta2.frx":178D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManReceta2.frx":1CD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManReceta2.frx":1DE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManReceta2.frx":1EF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManReceta2.frx":234D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManReceta2.frx":24B9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11205
      _ExtentX        =   19764
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar Receta"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Duplicar Receta"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
   Begin VB.Menu rece_1 
      Caption         =   "menureceta"
      Visible         =   0   'False
      Begin VB.Menu rece_1_1 
         Caption         =   "Agregar Receta                     "
      End
      Begin VB.Menu rece_1_2 
         Caption         =   "-"
      End
      Begin VB.Menu rece_1_3 
         Caption         =   "Eliminar Receta"
      End
   End
   Begin VB.Menu ingre_1 
      Caption         =   "menuIngre"
      Visible         =   0   'False
      Begin VB.Menu ingre_1_1 
         Caption         =   "Agregar "
      End
      Begin VB.Menu ingre_1_2 
         Caption         =   "-"
      End
      Begin VB.Menu ingre_1_3 
         Caption         =   "Eliminar"
      End
   End
   Begin VB.Menu menu_2 
      Caption         =   "menuTarea"
      Visible         =   0   'False
      Begin VB.Menu menu_2_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu menu_2_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu_2_3 
         Caption         =   "Eliminar"
      End
   End
End
Attribute VB_Name = "FrmManReceta2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMMANRECETA.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : PERMITE REGISTRAR Y MODIFICAR LAS RECETAS QUE SE UTILIZARAN EN EL SISTEMA
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 06/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim RstPro As New ADODB.Recordset              ' RECORDSET QUE ALAMCENARA LOS PRODUCTOS QUE MANEJA EL SISTEMA
Dim QueHace As Integer                         ' INDICA EN QUE MODO SE ENCUENTRA EL FORMULARUI
Dim SeEjecuto As Boolean                       ' VERIFICA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim RstDetRec As New ADODB.Recordset
Dim RstDetTar As New ADODB.Recordset
'*******************************************
Dim RstDetCos As New ADODB.Recordset
'*******************************************
Dim TipIng As Integer                          ' 1 = Insumo         2 = Producto
Dim ListandoInsumos As Boolean
Dim xNumRec As Double                          ' PARA CONTROLAR EL NUMERO DE RECETA solo temporal
Dim fOrdenLista As Boolean                     ' especfica el orden de la lista de la consulta
Dim Agregando As Boolean
Dim mIdRegistro&                               ' identificador del registro
Dim IdMenuActivo As Integer                    ' INDICA EL CODIGO DEL MENU ACTIVO
Dim xHorIni As Date                            ' ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim xSeleccionoArchivo As Boolean              ' Indica si se selecciono un archivo de procedimientos para grabarlo como nuevo
Dim cSQL As String

'*****************************************************************************************************
'* Nombre           : VerIngredienteReceta
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA LOS INGREDIENTES DE LA RECETA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub VerIngredienteReceta()
    Fg1.Rows = 1
    Agregando = True
    If RstDetRec.RecordCount <> 0 Then
        RstDetRec.MoveFirst
        RstDetRec.Sort = "idtipo DESC"
        
        Do While Not RstDetRec.EOF
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(RstDetRec("idtipo"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstDetRec("codpro"))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(RstDetRec("descripcion"))
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(RstDetRec("abrev"))
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosN(RstDetRec("canpro"))
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosN(RstDetRec("iditem"))
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosN(RstDetRec("idunimed"))
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosN(RstDetRec("idrec"))
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosN(RstDetRec("canpropra"))
            RstDetRec.MoveNext
        Loop
    End If
    Agregando = False
End Sub

'*****************************************************************************************************
'* Nombre           : VerTareasReceta
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA LAS TAREAS DE LA RECETA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub VerTareasReceta()
    Fg3.Rows = 1
    Agregando = True
    If RstDetTar.RecordCount <> 0 Then
        RstDetTar.MoveFirst
        Do While Not RstDetTar.EOF
            Fg3.Rows = Fg3.Rows + 1
            Fg3.TextMatrix(Fg3.Rows - 1, 1) = NulosC(RstDetTar("codigo"))
            Fg3.TextMatrix(Fg3.Rows - 1, 2) = NulosC(RstDetTar("descripcion"))
            Fg3.TextMatrix(Fg3.Rows - 1, 3) = NulosN(RstDetTar("idrec"))
            Fg3.TextMatrix(Fg3.Rows - 1, 4) = NulosN(RstDetTar("idtar"))
            Fg3.TextMatrix(Fg3.Rows - 1, 5) = NulosN(RstDetTar("idunimed"))
            Fg3.TextMatrix(Fg3.Rows - 1, 6) = NulosC(RstDetTar("abrev"))
            
            If IsNull(RstDetTar("horarr")) = False Then
                Fg3.TextMatrix(Fg3.Rows - 1, 7) = Format(RstDetTar("horarr"), "hh:mm")
            Else
                Fg3.TextMatrix(Fg3.Rows - 1, 7) = "00:00"
            End If
            Fg3.TextMatrix(Fg3.Rows - 1, 8) = Format(NulosN(RstDetTar("factor")), "0.000000")
            Fg3.TextMatrix(Fg3.Rows - 1, 9) = Format(NulosN(RstDetTar("numper")), "00")
            Fg3.TextMatrix(Fg3.Rows - 1, 10) = Format(NulosN(RstDetTar("costokg")), "0.000000")
            Fg3.TextMatrix(Fg3.Rows - 1, 11) = Format(NulosN(RstDetTar("aplpor")), "0.00")
            Fg3.TextMatrix(Fg3.Rows - 1, 12) = NulosN(RstDetTar("orden"))
            
            Fg3.TextMatrix(Fg3.Rows - 1, 13) = NulosC(RstDetTar("area"))
            Fg3.TextMatrix(Fg3.Rows - 1, 14) = NulosC(RstDetTar("destiptrab"))
            Fg3.TextMatrix(Fg3.Rows - 1, 15) = NulosC(RstDetTar("desformapag"))
            Fg3.TextMatrix(Fg3.Rows - 1, 16) = NulosN(RstDetTar("idarea"))
            Fg3.TextMatrix(Fg3.Rows - 1, 17) = NulosN(RstDetTar("idtiptrab"))
            Fg3.TextMatrix(Fg3.Rows - 1, 18) = NulosN(RstDetTar("idformapag"))
            
            RstDetTar.MoveNext
        Loop
    End If
    Agregando = False
End Sub

Private Sub VerCentroCostoReceta()
    fg5.Rows = 1
    Agregando = True
    If RstDetCos.RecordCount <> 0 Then
        RstDetCos.MoveFirst
        Do While Not RstDetCos.EOF
            fg5.Rows = fg5.Rows + 1
            fg5.TextMatrix(fg5.Rows - 1, 1) = NulosC(RstDetCos("cuenta"))
            fg5.TextMatrix(fg5.Rows - 1, 2) = NulosC(RstDetCos("descripcion"))
            fg5.TextMatrix(fg5.Rows - 1, 3) = NulosN(RstDetCos("idcuenta"))
            
            RstDetCos.MoveNext
        Loop
    End If
    Agregando = False
End Sub

Private Sub CmdAddReceta_Click()
    Dim Rst As New ADODB.Recordset
    Dim xFam As String
    Dim xNum As Integer
    
    RST_Busq Rst, "SELECT alm_inventario.id, mae_familia.descripcion AS desfam, alm_inventario.tippro " _
        & " FROM alm_inventario LEFT JOIN mae_familia ON alm_inventario.idfam = mae_familia.id WHERE (((alm_inventario.id)=" & RstPro("id") & ") " _
        & " AND ((alm_inventario.tippro) IN (3, 8)))", xCon

    xFam = Mid(Rst("desfam"), 1, 5)
    xNumRec = xNumRec + 1
    
    RST_Busq Rst, "SELECT pro_receta.codrec From pro_receta Where (((pro_receta.codrec) Like '" & xFam & "%')) " _
        & " ORDER BY pro_receta.codrec", xCon
    
    Agregando = True
    Fg2.Rows = Fg2.Rows + 1
    
    If Rst.RecordCount = 0 Then
        Fg2.TextMatrix(Fg2.Rows - 1, 1) = Mid(xFam, 1, 5) + "001"
        Fg2.TextMatrix(Fg2.Rows - 1, 3) = Trim(TxtUnidad.Text)
        Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(1, "0.00")
        Fg2.TextMatrix(Fg2.Rows - 1, 7) = xNumRec
    Else
        Rst.MoveLast
        xNum = NulosN(Mid(Rst("codrec"), 6, 3)) + 1
        
        Fg2.TextMatrix(Fg2.Rows - 1, 1) = Mid(xFam, 1, 5) + Format(xNum, "000")
        Fg2.TextMatrix(Fg2.Rows - 1, 3) = Trim(TxtUnidad.Text)
        Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(1, "0.00")
        Fg2.TextMatrix(Fg2.Rows - 1, 7) = xNumRec
    End If
    
    Agregando = False
    
    If Fg2.Rows > 1 Then
        habilitar cmdIngr, True
        habilitar cmdTar, True
    Else
        habilitar cmdIngr, False
        habilitar cmdTar, False
    End If
    
    If Fg2.Rows >= 1 Then
        Fg2.Col = 2
        Fg2.Row = Fg2.Rows - 1
        Fg2.SetFocus
    Else
        CmdAddReceta.SetFocus
    End If
    Set Rst = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : pIngredienteDel
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UNA FILA DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pIngredienteDel()
    If Fg1.Row < 1 Then Exit Sub
    
    If Fg1.Rows = 1 Then
        MsgBox "No hay registros para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Seguro desea eliminar el registro seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo) = vbNo Then Exit Sub
    
    If RstDetRec.RecordCount <> 0 Then
        RstDetRec.MoveFirst
        RstDetRec.Find "iditem = " & NulosN(Fg1.TextMatrix(Fg1.Row, 6)) & ""
        If RstDetRec.EOF = False Then
            RstDetRec.Delete
        End If
    End If
    
    Fg1.RemoveItem Fg1.Row
    If Fg1.Rows > 1 Then
        Fg1.Row = 1
        Fg1.Col = 1
        Fg1.SetFocus
    Else
        cmdIngr(0).SetFocus
    End If
End Sub

Private Sub CmdAgregar_Click()
    If NulosC(Fg4.TextMatrix(Fg4.Rows - 1, 1)) = "" Then
        Fg4.Rows = Fg4.Rows + 1
    End If
End Sub

Private Sub CmdCancelar_Click()
    TabOne1.Enabled = True
    Toolbar1.Enabled = True
    FrmDuplicaReceta.Visible = False
End Sub

Private Sub cmdCos_Click(Index As Integer)
    Dim Rpta As Integer
    Dim nTitulo As String
    Dim xRs As New ADODB.Recordset
    Dim xCampos() As String
    Dim xform As New eps_librerias.FormSeleccion
    Dim nSQLId As String
        
    If QueHace = 3 Then Exit Sub
            
    Select Case Index
        Case 0 ' AGREGAR CUENTA
            ReDim xCampos(2, 4) As String

            xCampos(0, 0) = "Cuenta":           xCampos(0, 1) = "cuenta":           xCampos(0, 2) = "1000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            xCampos(1, 0) = "Descripción":      xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "4500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
            
            nSQLId = " AND (Left(con_planctas.cuenta, 1) In ('9'))"
            nSQLId = nSQLId & GENERAR_SQL_ID(fg5, 3, " AND con_planctas.id", "NOT IN", True)
            
            cSQL = "SELECT con_planctas.id, con_planctas.cuenta, con_planctas.descripcion " _
                + vbCr + "FROM con_planctas " _
                + vbCr + "WHERE (((con_planctas.tipo)=0)) " & nSQLId
                
            nTitulo = "Buscando Cuentas"
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "cuenta", "cuenta", Principio
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub

            Agregando = True
            With fg5
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = NulosC(xRs("cuenta"))
                .TextMatrix(.Rows - 1, 2) = NulosC(xRs("descripcion"))
                .TextMatrix(.Rows - 1, 3) = NulosC(xRs("id"))
                
                RstDetCos.AddNew
                RstDetCos("idrec") = NulosN(Fg2.TextMatrix(Fg2.Row, 7))
                RstDetCos("cuenta") = NulosC(xRs("cuenta"))
                RstDetCos("descripcion") = NulosC(xRs("descripcion"))
                RstDetCos("idcuenta") = NulosC(xRs("id"))
                RstDetCos.Update
            End With
            Agregando = False
            
        Case 1 ' SELECCIONAR CUENTA
            ReDim xCampos(3, 4) As String

            xCampos(0, 0) = "Cuenta":           xCampos(0, 1) = "cuenta":           xCampos(0, 2) = "1000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            xCampos(1, 0) = "Descripción":      xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "4500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
            xCampos(2, 0) = "Importe":          xCampos(2, 1) = "importe":          xCampos(2, 2) = "1200":     xCampos(2, 3) = "N":    xCampos(2, 4) = "N"
            
            nSQLId = " AND (Left(con_planctas.cuenta, 1) In ('9'))"
            nSQLId = nSQLId & GENERAR_SQL_ID(fg5, 3, " AND con_planctas.id", "NOT IN", True)
            
            cSQL = "SELECT 0 AS xsel, con_planctas.id, con_planctas.cuenta, con_planctas.descripcion " _
                + vbCr + "FROM con_planctas " _
                + vbCr + "WHERE (((con_planctas.tipo)=0)) " & nSQLId
                
            xform.SQLCad = cSQL
            xform.titulo = "Seleccionando Cuentas"
            Set xform.Coneccion = xCon
            Set xRs = Nothing
            Set xRs = xform.seleccionar(xCampos)
            Set xform = Nothing
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub

            Agregando = True
            With fg5
                xRs.MoveFirst
                While Not xRs.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = NulosC(xRs("cuenta"))
                    .TextMatrix(.Rows - 1, 2) = NulosC(xRs("descripcion"))
                    .TextMatrix(.Rows - 1, 3) = NulosC(xRs("id"))
                
                    RstDetCos.AddNew
                    RstDetCos("idrec") = NulosN(Fg2.TextMatrix(Fg2.Row, 7))
                    RstDetCos("cuenta") = NulosC(xRs("cuenta"))
                    RstDetCos("descripcion") = NulosC(xRs("descripcion"))
                    RstDetCos("idcuenta") = NulosC(xRs("id"))
                    RstDetCos.Update
                
                    xRs.MoveNext
                Wend
            End With
            Agregando = False
            
        Case 2 ' ELIMINAR CUENTA
            If fg5.Rows <= fg5.FixedRows Then Exit Sub
            Rpta = MsgBox("¿Está seguro de eliminar el registro actual?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
            If Rpta = vbYes Then
                RstDetCos.Filter = "idrec=" & NulosN(Fg2.TextMatrix(Fg2.Row, 7)) & " AND idcuenta=" & NulosN(fg5.TextMatrix(fg5.Row, 3))
                limpiarRST RstDetCos, False
                fg5.RemoveItem fg5.Row
            End If
            
        Case 3 ' ELIMINAR TODO
            If fg5.Rows <= fg5.FixedRows Then Exit Sub
            Rpta = MsgBox("¿Está seguro de eliminar todos los registros?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
            If Rpta = vbYes Then
                RstDetCos.Filter = "idrec=" & NulosN(Fg2.TextMatrix(Fg2.Row, 7))
                limpiarRST RstDetCos, False
                fg5.Rows = fg5.FixedRows
            End If
            
    End Select
End Sub

Private Sub CmdDelReceta_Click()
    ' ELIMINA UNA FILA DEL CONTROL Fg2
    Dim Rpta As Integer
    If Fg2.Row < 1 Then Exit Sub
    Rpta = MsgBox("Esta seguro de eliminar la receta seleccionada", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        ' eliminar los registros del recordset temporal
        RstRegistroEliminar RstDetRec, "idrec", NulosN(Fg2.TextMatrix(Fg2.Row, 7)), True
        RstRegistroEliminar RstDetTar, "idrec", NulosN(Fg2.TextMatrix(Fg2.Row, 7)), True
        Fg2.RemoveItem (Fg2.Row)
    End If
    
    If Fg2.Rows > 1 Then
        habilitar cmdIngr, True
        habilitar cmdTar, True
    Else
        habilitar cmdIngr, False
        habilitar cmdTar, False
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : pTareaDel
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UNA FILA DEL CONTROL Fg3
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pTareaDel()
    If Fg3.Row < 1 Then
        MsgBox "Elija el registro correcto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    If Fg3.Rows = 1 Then
        MsgBox "No hay tareas para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg3.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    
    RstDetTar.MoveFirst
    RstDetTar.Find "idtar = " & NulosN(Fg3.TextMatrix(Fg3.Row, 4)) & ""
    If RstDetTar.EOF = False Then
        RstDetTar.Delete
    End If
    
    Fg3.RemoveItem Fg3.Row
    If Fg3.Rows > 1 Then
        Fg3.Row = 1
        Fg3.Col = 1
        Fg3.SetFocus
    Else
        cmdTar(0).SetFocus
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : pTareaAdd
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : AGREGA UN FILA AL CONTROL Fg3
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pTareaAdd()
    Dim mCol%
    Dim fInsertar As Boolean
    
    If QueHace = 3 Then Exit Sub
    Agregando = True
    
    If Fg3.Rows > Fg3.FixedRows Then
        If NulosN(Fg3.TextMatrix(Fg3.Rows - 1, 4)) = 0 Then
            MsgBox "Seleccione una tarea", vbExclamation, xTitulo
        Else
            fInsertar = True
        End If
    Else
        fInsertar = True
    End If
    
    If fInsertar = True Then Fg3.AddItem ""
    
    Fg3.Row = Fg3.Rows - 1
    Fg3.Col = 1
    
    If fInsertar = True Then Fg3_CellButtonClick Fg1.Rows - 1, 1

    Fg3.SetFocus
    Agregando = False
End Sub

Private Sub CmdEliminar_Click()
    If Fg4.Rows = 1 Then Exit Sub
    Fg4.Rows = Fg4.Rows - 1
End Sub

Private Sub CmdGenerar_Click()
    Dim xSQL As String
    Dim xRstCab As New ADODB.Recordset
    Dim xRstIns As New ADODB.Recordset
    Dim xRstTar As New ADODB.Recordset
    Dim xIdTabla As Double
    Dim A, B, C, X As Integer
    Dim xRst As New ADODB.Recordset

On Error GoTo HORROR_

    xCon.BeginTrans
    
    xSQL = "SELECT pro_receta.* From pro_receta WHERE (((pro_receta.iditem)=" & NulosN(LblIdProd.Caption) & "))"
    RST_Busq xRstCab, xSQL, xCon
    
    RST_Busq xRstCab, xSQL, xCon
    
    xIdTabla = HallaCodigoTabla("pro_receta", xCon, "id")
    
    If xRstCab.RecordCount <> 0 Then
        xRstCab.MoveFirst
        For X = 1 To Fg4.Rows - 1
            While Not xRstCab.EOF
                'hallamos los insumos de la receta
                xSQL = "SELECT pro_recetains.* From pro_recetains WHERE (((pro_recetains.idrec)=" & xRstCab("id") & "))"
                RST_Busq xRstIns, xSQL, xCon
                
                xSQL = "SELECT pro_recetatar.* From pro_recetatar WHERE (((pro_recetatar.idrec)=" & xRstCab("id") & "))"
                RST_Busq xRstTar, xSQL, xCon
                
                
                'Grabamos la cabecera
                RST_Busq xRst, "SELECT * FROM pro_receta", xCon
                xRst.AddNew
                xRst("id") = xIdTabla
                xRst("codrec") = xRstCab("codrec")
                xRst("iditem") = xRstCab("iditem")
                xRst("descripcion") = NulosC(Fg4.TextMatrix(X, 1)) 'xRstCab("descripcion")
                xRst("idunimed") = xRstCab("idunimed")
                xRst("cantidad") = xRstCab("cantidad")
                xRst("prirec") = xRstCab("prirec")
                xRst("observaciones") = xRstCab("observaciones")
                xRst("archpro") = xRstCab("archpro")
                xRst("idtiptrab") = xRstCab("idtiptrab")
                xRst("idformapag") = xRstCab("idformapag")
                xRst.Update
                
                ' GRABAMOS LOS INSUMOS
                xRstIns.MoveFirst
                While Not xRstIns.EOF
                    RST_Busq xRst, "SELECT * FROM pro_recetains", xCon
                    xRst.AddNew
                    xRst("idrec") = xIdTabla
                    xRst("iditem") = xRstIns("iditem")
                    xRst("idunimed") = xRstIns("idunimed")
                    xRst("canpro") = xRstIns("canpro")
                    xRst("canpropra") = xRstIns("canpropra")
                    xRst.Update
                    xRstIns.MoveNext
                Wend
    
                ' GRABAMOS LAS TAREAS
                xRstTar.MoveFirst
                While Not xRstTar.EOF
                    RST_Busq xRst, "SELECT * FROM pro_recetatar", xCon
                    xRst.AddNew
                    xRst("idrec") = xIdTabla
                    xRst("idtar") = xRstTar("idtar")
                    xRst("idunimed") = xRstTar("idunimed")
                    xRst("cantidad") = xRstTar("cantidad")
                    xRst("orden") = xRstTar("orden")
                    xRst("factor") = xRstTar("factor")
                    xRst("costokg") = xRstTar("costokg")
                    xRst("costohr") = xRstTar("costohr")
                    xRst("jornalkg") = xRstTar("jornalkg")
                    xRst("descpro") = xRstTar("descpro")
                    xRst("numper") = xRstTar("numper")
                    xRst("horarr") = xRstTar("horarr")
                    xRst("aplpor") = xRstTar("aplpor")
                    xRst("idarea") = xRstTar("idarea")
                    xRst("idtiptrab") = xRstTar("idtiptrab")
                    xRst("idformapag") = xRstTar("idformapag")
                    xRst.Update
                    xRstTar.MoveNext
                Wend
                
                xRstCab.MoveNext
                xIdTabla = xIdTabla + 1
            Wend
        Next X
    End If
    
    xCon.CommitTrans
    Set xRstCab = Nothing
    Set xRstIns = Nothing
    Set xRstTar = Nothing
    Set xRst = Nothing
    
    MsgBox "El/las recetas se duplicaron con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    TabOne1.Enabled = True
    Toolbar1.Enabled = True
    FrmDuplicaReceta.Visible = False
    Exit Sub
HORROR_:
    'Resume
    xCon.RollbackTrans
    Set xRstCab = Nothing
    Set xRstIns = Nothing
    Set xRstTar = Nothing
    Set xRst = Nothing
    
    MsgBox "No se pudo copiar el/las Recetas por el siguiente motivo :" + Trim(Err.Description)
End Sub

Private Sub CmdVerProc_Click()
    If NulosC(Fg2.TextMatrix(Fg2.Row, 8)) = "" Then
        Exit Sub
    End If
    
    Dim xArch As String
    Dim xWord As Object
    Set xWord = CreateObject("Word.application")
    
    xWord.Visible = True
    xWord.Application.Activate
    '...massimizzandolo...
    xWord.WindowState = 1 'wdWindowStateMaximize
    
    xArch = LeerLineaINI(App.Path & "\seven.ini", "RUTAAR", "RUTAS") & "procedimientos\" & NulosC(Fg2.TextMatrix(Fg2.Row, 8))
    xWord.Documents.Open(xArch).Activate
    Exit Sub
        
'            Dim xF As New eps_librerias.browser
'            Dim xRuta As String
'            xRuta = LeerLineaINI(App.Path & "\seven.ini", "RUTAAR", "RUTAS") & "procedimientos\" & NulosC(Fg2.TextMatrix(Fg2.Row, 8))
'
'            'xRuta = NulosC(Fg2.TextMatrix(Fg2.Row, 8))
'            xF.Navegador (xRuta)
'            Set xF = Nothing
    
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstPro
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA LAS COLUMNAS DEL CONTROL Dg1
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstPro.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        'RstPro("id")=iditem de almacen
        VerMovimientos1 IdMenuActivo, NulosN(RstPro("id")), xCon
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim xRs As New ADODB.Recordset
    Dim xCodItem As Double
    ReDim xCampos(3, 4) As String
    Dim nTitulo As String
    Dim nWhereIn As String
    Dim nSQL As String
    Dim nSQLNotId As String
    Set xRs = Nothing
    
    ' verificar si eligio el tipo de ingrediente
    If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 1)) = 0 Then
        MsgBox "Seleccione el Tipo de Ingrediente ", vbExclamation, xTitulo
        Fg1.Col = 1
        Fg1.SetFocus
        Exit Sub
    End If
    
    Select Case NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 1))
        Case 1
            nTitulo = "Materia Prima"
            nWhereIn = NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 1))
        
        Case 3
            nTitulo = "Producto"
            nWhereIn = "" & NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 1))
        
        Case 4
            nTitulo = "Insumo"
            nWhereIn = NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 1))
        
        Case 8
            nTitulo = "Productos Intermedios"
            nWhereIn = NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 1))
        
        Case Else
            Exit Sub
    End Select
    
    If Col = 2 Then
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":    xCampos(0, 3) = "C"
        xCampos(1, 0) = "Und":          xCampos(1, 1) = "Abrev":         xCampos(1, 2) = "1000":    xCampos(1, 3) = "C"
        xCampos(2, 0) = "Código":       xCampos(2, 1) = "codpro":        xCampos(2, 2) = "2000":    xCampos(2, 3) = "C"
        
        ' si hay registros seleccionados anteriormente no considerarlos de nuevo
        nSQLNotId = GRID_GENERAR_SQL_ID(Fg1, 6, " AND alm_inventario.id", "NOT IN", True)

        nSQL = "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, " _
           & " alm_inventario.tippro, mae_tipoproducto.descripcion AS destippro, alm_inventario.idunimed " _
            & vbCr & " FROM mae_tipoproducto RIGHT JOIN (mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed) ON mae_tipoproducto.id = alm_inventario.tippro " _
            & vbCr & " WHERE alm_inventario.activo =-1 and (((alm_inventario.tippro) In (" & nWhereIn & " ))) " & nSQLNotId _
            & vbCr & " ORDER BY alm_inventario.descripcion "
            
        Agregando = True
        
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando " & nTitulo, "descripcion", "descripcion", Principio
        
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                ' si ya tiene item asociado al registro
                xCodItem = NulosN(Fg1.TextMatrix(Fg1.Row, 6))
                
                Fg1.TextMatrix(Fg1.Row, 2) = NulosC(xRs("codpro"))
                Fg1.TextMatrix(Fg1.Row, 3) = NulosC(xRs("descripcion"))
                Fg1.TextMatrix(Fg1.Row, 4) = NulosC(xRs("abrev"))
                Fg1.TextMatrix(Fg1.Row, 5) = 0
                Fg1.TextMatrix(Fg1.Row, 6) = NulosN(xRs("id"))
                Fg1.TextMatrix(Fg1.Row, 7) = NulosN(xRs("idunimed"))
                Fg1.TextMatrix(Fg1.Row, 8) = Fg2.TextMatrix(Fg2.Row, 7) '--id de receta
                
                If xCodItem = 0 Then
                    RstDetRec.AddNew
                Else
                    RstDetRec.MoveFirst
                    RstDetRec.Find "iditem = " & xCodItem
                End If
                
                RstDetRec("tipoproducto") = NulosC(xRs("destippro"))
                RstDetRec("codpro") = NulosC(xRs("codpro"))
                RstDetRec("descripcion") = NulosC(xRs("descripcion"))
                RstDetRec("abrev") = NulosC(xRs("abrev"))
                RstDetRec("canpro") = 0
                RstDetRec("idrec") = NulosN(Fg2.TextMatrix(Fg2.Row, 7))
                RstDetRec("idtipo") = NulosN(xRs("tippro"))
                RstDetRec("iditem") = NulosN(xRs("id"))
                RstDetRec("idunimed") = NulosC(xRs("idunimed"))
                Fg1.Col = 5
            End If
        End If
        
    ElseIf Col = 4 Then
        If NulosN(Fg1.TextMatrix(Fg1.Row, 6)) = 0 Then
            MsgBox "Seleccione primero " & nTitulo, vbExclamation, xTitulo
            Fg1.Col = 2
            Fg1.SetFocus
            Exit Sub
        End If
        
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "3500":    xCampos(0, 3) = "C"
        xCampos(1, 0) = "Abrev":        xCampos(1, 1) = "Abrev":         xCampos(1, 2) = "600":    xCampos(1, 3) = "C"
        xCampos(2, 0) = "Id":       xCampos(2, 1) = "id":        xCampos(2, 2) = "450":    xCampos(2, 3) = "N"
        
        CARGAR_DLL_EPSBUSCAR xCon, xRs, "SELECT * FROM mae_unidades", xCampos(), "Buscando Unidades", "descripcion", "descripcion", Principio
        
        Agregando = True
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                xCodItem = NulosN(Fg1.TextMatrix(Fg1.Row, 6))
                Fg1.TextMatrix(Fg1.Row, 4) = NulosC(xRs("abrev"))
                Fg1.TextMatrix(Fg1.Row, 7) = NulosN(xRs("id"))
                RstDetRec.MoveFirst
                RstDetRec.Find "iditem = " & xCodItem
                
                If RstDetRec.EOF = False Then
                    RstDetRec("idunimed") = NulosC(xRs("id"))
                    RstDetRec("abrev") = NulosC(xRs("abrev"))
                End If
            End If
        End If
        Set xRs = Nothing
    End If
    
    Agregando = False
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    
    If Col = 5 Or Col = 9 Then
        RstDetRec.MoveFirst
        RstDetRec.Find "iditem = " & NulosN(Fg1.TextMatrix(Fg1.Row, 6))
        
        If IsNumeric(Fg1.TextMatrix(Fg1.Row, Col)) = False Then
            MsgBox "El valor ingresado no es incorrecto", vbExclamation, xTitulo
            Fg1.TextMatrix(Fg1.Row, Col) = 0
        End If
        
        If RstDetRec.EOF = False Then
            If Col = 5 Then RstDetRec("canpro") = NulosN(Fg1.TextMatrix(Fg1.Row, Col))
            If Col = 9 Then RstDetRec("canpropra") = NulosN(Fg1.TextMatrix(Fg1.Row, Col))
        End If
    End If
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
        Exit Sub
    End If
    
    If Fg1.Col = 1 Or Fg1.Col = 2 Or Fg1.Col = 4 Or Fg1.Col = 5 Or Fg1.Col = 9 Then
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case Col
        Case 1, 2, 4
            KeyAscii = 0
        
        Case 5
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End Select
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    
    If KeyCode = 114 Or KeyCode = 45 Then cmdIngr_click 0   ' F3 = Agrega Item
    
    If KeyCode = 115 Or KeyCode = 46 Then cmdIngr_click 1   ' F4 = Eliminar Item
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then If QueHace <> 3 Then If Fg2.Rows > 1 Then PopupMenu ingre_1
End Sub

Private Sub Fg2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim cSQL As String
    Dim xCampos(2, 4) As String
    Dim xRs As New ADODB.Recordset
    Dim nTitulo As String
    
    If Col = 8 Then
        If QueHace = 3 Then
            If NulosC(Fg2.TextMatrix(Fg2.Row, 8)) = "" Then
                Exit Sub
            End If
            
            Dim xF As New eps_librerias.browser
            Dim xRuta As String
            xRuta = LeerLineaINI(App.Path & "\seven.ini", "RUTAAR", "RUTAS") & "procedimientos\" & NulosC(Fg2.TextMatrix(Fg2.Row, 8))
            
            'xRuta = NulosC(Fg2.TextMatrix(Fg2.Row, 8))
            xF.Navegador (xRuta)
            Set xF = Nothing
        Else
            With CommonDialog1
                .DefaultExt = "htm"
                .Filter = "Todos los htm (*.htm)|*.htm|Archivos html (*.html)|*.html"
                .ShowOpen
            End With
            
            If NulosC(CommonDialog1.FileName) <> "" Then
                'xSeleccionoArchivo = True
                Fg2.TextMatrix(Fg2.Row, 8) = CommonDialog1.FileName
                Fg2.TextMatrix(Fg2.Row, 9) = 1
            End If
        End If
    End If
End Sub

Private Sub Fg2_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    
    If KeyCode = 114 Then CmdAddReceta_Click  ' F3 = Agrega Item
    
    If KeyCode = 115 Then CmdDelReceta_Click  ' F4 = Eliminar Item
End Sub

Private Sub Fg2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace <> 3 Then PopupMenu rece_1
    End If
End Sub

Private Sub Fg2_RowColChange()
    If Agregando = True Then Exit Sub
    
    If Fg2.Row < 1 Then
        habilitar cmdIngr, False
        habilitar cmdTar, False
        Exit Sub
    End If
    
    If QueHace <> 3 Then
        habilitar cmdIngr, True
        habilitar cmdTar, True
    End If
    
    TxtNotas.Text = Fg2.TextMatrix(Fg2.Row, 6)
    ' Mostramos los insumos de la receta
    RstDetRec.Filter = adFilterNone
    RstDetRec.Filter = "idrec= " & NulosN(Fg2.TextMatrix(Fg2.Row, 7))
    If RstDetRec.RecordCount <> 0 Then
        VerIngredienteReceta
    Else
        Fg1.Rows = 1
    End If
    
    ' Mostramos las tareas de la receta
    RstDetTar.Filter = adFilterNone
    RstDetTar.Filter = "idrec = " & NulosN(Fg2.TextMatrix(Fg2.Row, 7))
    If RstDetTar.RecordCount <> 0 Then
        VerTareasReceta
    Else
        Fg3.Rows = 1
    End If
    
    ' Mostramos el cento de costos de la receta
    RstDetCos.Filter = adFilterNone
    RstDetCos.Filter = "idrec = " & NulosN(Fg2.TextMatrix(Fg2.Row, 7))
    If RstDetCos.RecordCount <> 0 Then
        VerCentroCostoReceta
    Else
        fg5.Rows = 1
    End If
End Sub

Private Sub Fg3_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim xRs As New ADODB.Recordset
    Dim xCodTar As Double
    Dim nSQL As String
    Dim nSQLId As String
    Dim nTitulo As String
    Dim cSQL As String
    Dim xCampos() As String
    
    If Col = 1 Then
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        ReDim xCampos(3, 4) As String
        
        xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":    xCampos(0, 3) = "C"
        xCampos(1, 0) = "Uni. Medida":  xCampos(1, 1) = "undabrev":      xCampos(1, 2) = "1500":    xCampos(1, 3) = "C"
        xCampos(2, 0) = "Codigo":       xCampos(2, 1) = "codigo":        xCampos(2, 2) = "2000":    xCampos(2, 3) = "C"
        
        ' si hay tareas anteriomente seleccionadas
        nSQLId = GRID_GENERAR_SQL_ID(Fg3, 4, " and pro_tareas.id", "NOT IN")
        nSQL = "SELECT pro_tareas.*, mae_unidades.abrev as undabrev " _
            & vbCr & " FROM pro_tareas LEFT JOIN mae_unidades ON pro_tareas.idunimed = mae_unidades.id " _
            & vbCr & " WHERE pro_tareas.diverso = 0 " _
            & vbCr & " ORDER BY pro_tareas.descripcion "
        
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Tareas", "descripcion", "descripcion", Principio
        Agregando = True
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                ' si ya es una tarea anteriomente seleccionado
                xCodTar = NulosN(Fg3.TextMatrix(Fg3.Row, 4))
                Fg3.TextMatrix(Fg3.Row, 1) = NulosC(xRs("codigo"))
                Fg3.TextMatrix(Fg3.Row, 2) = NulosC(xRs("descripcion"))
                Fg3.TextMatrix(Fg3.Row, 3) = Fg2.TextMatrix(Fg2.Row, 7) '--cod receta
                Fg3.TextMatrix(Fg3.Row, 4) = NulosN(xRs("id")) '--cod tarea
                Fg3.TextMatrix(Fg3.Row, 5) = NulosN(xRs("idunimed"))
                Fg3.TextMatrix(Fg3.Row, 6) = NulosC(xRs("undabrev"))
                
                If xCodTar = 0 Then
                    RstDetTar.AddNew
                Else
                    RstDetTar.MoveFirst
                    RstDetTar.Find "idtar = " & xCodTar
                End If
                
                RstDetTar("idrec") = NulosN(Fg2.TextMatrix(Fg2.Row, 7))
                RstDetTar("idtar") = NulosN(xRs("id"))
                RstDetTar("codigo") = NulosC(xRs("codigo"))
                RstDetTar("descripcion") = NulosC(xRs("descripcion"))
                RstDetTar("abrev") = NulosC(xRs("abrev"))
                RstDetTar("idunimed") = NulosN(xRs("idunimed"))
            End If
        End If
        Set xRs = Nothing
        
    ElseIf Col = 13 Then ' Area
        If QueHace = 3 Then Exit Sub
        
        ReDim xCampos(2, 4) As String
        xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
        xCampos(1, 0) = "Id":           xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"
        
        nTitulo = "Buscando Area"
        
        cSQL = "SELECT mae_area.id, mae_area.descripcion AS nombre, mae_area.id AS cod, mae_area.id AS idarea, pro_emp.id AS idper, pla_empleados.id AS idemp, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS encargado, pla_empleados.numdoc " _
            + vbCr + " FROM ((pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) RIGHT JOIN pro_area ON pro_emp.id = pro_area.idper) INNER JOIN mae_area ON pro_area.idarea = mae_area.id "
            
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
                                                      
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        Fg3.TextMatrix(Fg3.Row, Col) = NulosC(xRs("nombre"))
        Fg3.TextMatrix(Fg3.Row, 16) = NulosN(xRs("id"))
        
        ' Se busca el codigo de Tarea
        xCodTar = NulosN(Fg3.TextMatrix(Fg3.Row, 4))
        RstDetTar.MoveFirst
        RstDetTar.Find "idtar = " & xCodTar
                
        RstDetTar("idarea") = NulosN(xRs("id"))
        RstDetTar("area") = NulosC(xRs("nombre"))
            
    ElseIf Col = 14 Then ' Tipo de Trabajo
        If QueHace = 3 Then Exit Sub
        
        ReDim xCampos(2, 4) As String
        xCampos(0, 0) = "Id":           xCampos(0, 1) = "id":           xCampos(0, 2) = "500":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
        xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "2500":    xCampos(1, 3) = "C":   xCampos(1, 4) = "C"
        
        nTitulo = "Buscando Tipos de Trabajo"
        
        cSQL = "SELECT * FROM pro_tiptrab"
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                        "descripcion", "descripcion", Principio, ""
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        Fg3.TextMatrix(Fg3.Row, Col) = NulosC(xRs("descripcion"))
        Fg3.TextMatrix(Fg3.Row, 17) = NulosN(xRs("id"))
        
        ' Se busca el codigo de Tarea
        xCodTar = NulosN(Fg3.TextMatrix(Fg3.Row, 4))
        RstDetTar.MoveFirst
        RstDetTar.Find "idtar = " & xCodTar
                
        RstDetTar("idtiptrab") = NulosN(xRs("id"))
        RstDetTar("destiptrab") = NulosC(xRs("descripcion"))
    
    ElseIf Col = 15 Then ' Forma de Pago
        If QueHace = 3 Then Exit Sub
        
        ReDim xCampos(2, 4) As String
        xCampos(0, 0) = "Id":           xCampos(0, 1) = "id":           xCampos(0, 2) = "500":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
        xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "2500":    xCampos(1, 3) = "C":   xCampos(1, 4) = "C"
        
        nTitulo = "Buscando Formas de Pago"
        
        cSQL = "SELECT * FROM pro_formapag"
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                        "descripcion", "descripcion", Principio, ""
                                                      
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        Fg3.TextMatrix(Fg3.Row, Col) = NulosC(xRs("descripcion"))
        Fg3.TextMatrix(Fg3.Row, 18) = NulosN(xRs("id"))
        
        ' Se busca el codigo de Tarea
        xCodTar = NulosN(Fg3.TextMatrix(Fg3.Row, 4))
        RstDetTar.MoveFirst
        RstDetTar.Find "idtar = " & xCodTar
                
        RstDetTar("idformapag") = NulosN(xRs("id"))
        RstDetTar("desformapag") = NulosC(xRs("descripcion"))
    End If
    
    Agregando = False
    Fg3.Col = 1
    Fg3.SetFocus
End Sub

Private Sub Fg3_CellChanged(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    If Agregando = True Then Exit Sub
    If Row <= 0 Then Exit Sub
    
    Select Case Col
        Case 7
            RstDetTar.MoveFirst
            RstDetTar.Find "idtar = " & NulosN(Fg3.TextMatrix(Row, 4))
            
            If RstDetTar.EOF = False Then
                If NulosC(Fg3.TextMatrix(Row, 7)) <> "" Then
                    RstDetTar("horarr") = Format((Fg3.TextMatrix(Row, 7)), "hh:mm")
                End If
            End If
        
        Case 8, 9, 10, 11
            If IsNumeric(Fg3.TextMatrix(Row, Col)) = False Then
                MsgBox "El valor ingresado no es incorrecto", vbExclamation, xTitulo
                Fg3.TextMatrix(Row, Col) = 0
            Else
                Fg3.TextMatrix(Row, Col) = NulosN(Fg3.TextMatrix(Row, Col))
                RstDetTar.MoveFirst
                RstDetTar.Find "idtar = " & NulosN(Fg3.TextMatrix(Row, 4))
                
                If RstDetTar.EOF = False Then
                    RstDetTar("factor") = NulosN(Fg3.TextMatrix(Row, 8))
                    RstDetTar("numper") = NulosN(Fg3.TextMatrix(Row, 9))
                    RstDetTar("costokg") = NulosN(Fg3.TextMatrix(Row, 10))
                    RstDetTar("aplpor") = NulosN(Fg3.TextMatrix(Row, 11))
                End If
            End If
        
        Case 12
            ' buscar si existe el valor en la lista de tareas deberan de ser numeros únicos
            RstDetTar.MoveFirst
            RstDetTar.Find "idtar = " & NulosN(Fg3.TextMatrix(Row, 4))
            If RstDetTar.EOF = False Then
                RstDetTar("orden") = NulosN(Fg3.TextMatrix(Row, Col))
            End If
    End Select

    Exit Sub

error:
    Resume
    SHOW_ERROR Me.Name, "Fg3_CellChanged"
End Sub

Private Sub Fg3_EnterCell()
    If QueHace = 3 Then
        Fg3.Editable = flexEDNone
        Exit Sub
    End If
    If Fg3.Col >= 2 And Fg3.Col <= 6 Then
        Fg3.Editable = flexEDNone
    Else
        Fg3.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub Fg3_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    Select Case Col
        Case 7, 8, 9, 10, 11, 12
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Fg3_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = 114 Or KeyCode = 45 Then cmdTar_click 0   'F3 = Agrega Item
    If KeyCode = 115 Or KeyCode = 46 Then cmdTar_click 1   'F4 = Eliminar Item
End Sub

Private Sub Fg3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then If QueHace <> 3 Then If Fg2.Rows > 1 Then PopupMenu menu_2
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        '--ocultar el boton a agregar
        Toolbar1.Buttons(1).Visible = False
        
        RST_Busq RstPro, "SELECT item.* ,rec.recpri  FROM " _
            & " (SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.tippro, (SELECT Count([iditem]) AS numrec From pro_receta WHERE (pro_receta.iditem=alm_inventario.id)) AS numrec " _
            & " FROM mae_unidades RIGHT JOIN  alm_inventario ON mae_unidades.id = alm_inventario.idunimed Where (((alm_inventario.tippro) In (3, 8)) ) " _
            & " ) as item LEFT JOIN ( SELECT pro_receta.iditem, pro_receta.codrec AS recpri, pro_receta.prirec FROM pro_receta WHERE (((pro_receta.prirec)=1)) ) as rec ON item.id = rec.iditem ORDER BY item.descripcion", xCon
        
        Set Dg1.DataSource = RstPro
        xNumRec = -999     ' Temporal si es que se agrega mas recetas, al momento de grabar se creara el verdadero numero
        Agregando = True
        pConfigurarGrilla
        Agregando = False
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Muestradatos
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA LOS DATOS DE LA RECETA EN FORMA DETALLADA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Muestradatos()
    Dim RstRecetas As New ADODB.Recordset
    Dim cSQL As String
   
    Agregando = True
   
    pConfigurarGrilla
   
    TabOne2.CurrTab = 0
    
    TxtCodigo.Text = NulosC(RstPro("codpro"))
    TxtDescripcion.Text = NulosC(RstPro("descripcion"))
    TxtUnidad.Text = NulosC(RstPro("abrev"))
    
'    cSQL = "SELECT pro_receta.id, pro_receta.codrec, pro_receta.descripcion, pro_receta.archpro, mae_unidades.abrev, pro_receta.cantidad, pro_receta.prirec, pro_receta.iditem, pro_receta.observaciones " _
'        + vbCr + "FROM pro_receta LEFT JOIN mae_unidades ON pro_receta.idunimed = mae_unidades.id " _
'        + vbCr + "Where (((pro_receta.iditem) = " & NulosN(RstPro("id")) & ")) " _
'        + vbCr + "ORDER BY pro_receta.prirec"
    
    cSQL = "SELECT pro_receta.id, pro_receta.codrec, pro_receta.descripcion, pro_receta.archpro, mae_unidades.abrev, pro_receta.cantidad, pro_receta.prirec, pro_receta.iditem, pro_receta.observaciones, pro_tiptrab.id AS idtiptrab, pro_tiptrab.descripcion AS destiptrab, pro_formapag.id AS idformapag, pro_formapag.descripcion AS desformapag " _
        + vbCr + "FROM ((pro_receta LEFT JOIN mae_unidades ON pro_receta.idunimed = mae_unidades.id) LEFT JOIN pro_tiptrab ON pro_receta.idtiptrab = pro_tiptrab.id) LEFT JOIN pro_formapag ON pro_receta.idformapag = pro_formapag.id " _
        + vbCr + "Where (((pro_receta.iditem) = " & NulosN(RstPro("id")) & ")) " _
        + vbCr + "ORDER BY pro_receta.prirec"
    
    RST_Busq RstRecetas, cSQL, xCon
    
    Fg2.Rows = 1
    
    ' limpiando los rst termporales
    Set RstDetRec = Nothing
    Set RstDetTar = Nothing
    Set RstDetCos = Nothing
        
    If RstRecetas.RecordCount <> 0 Then
        RstRecetas.MoveFirst
        
        Do While Not RstRecetas.EOF
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosC(RstRecetas("codrec"))
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = NulosC(RstRecetas("descripcion"))
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = NulosC(RstRecetas("abrev"))
            Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(NulosN(RstRecetas("cantidad")), "0.00")
            Fg2.TextMatrix(Fg2.Rows - 1, 5) = NulosN(RstRecetas("prirec"))
            Fg2.TextMatrix(Fg2.Rows - 1, 6) = NulosC(RstRecetas("observaciones"))
            Fg2.TextMatrix(Fg2.Rows - 1, 7) = NulosN(RstRecetas("id"))
            Fg2.TextMatrix(Fg2.Rows - 1, 8) = NulosC(RstRecetas("archpro"))
            
            ' cargar datos de ingredientes al rst tmp
            pDefinirRstTmp E_INSUMO, RstDetRec, NulosN(RstRecetas("id"))
            
            ' cargar datos de tareas al rst tmp
            pDefinirRstTmp e_TAREA, RstDetTar, NulosN(RstRecetas("id"))
        
            ' definir el rst tmp de costos
            pDefinirRstTmp E_COSTO, RstDetCos, NulosN(RstRecetas("id"))
        
            RstRecetas.MoveNext
        Loop
    Else
        ' definir el rst tmp de ingredientes
        pDefinirRstTmp E_INSUMO, RstDetRec, -1111
        ' definir el rst tmp de tareas
        pDefinirRstTmp e_TAREA, RstDetTar, -1111
        
        ' definir el rst tmp de costos
        pDefinirRstTmp E_COSTO, RstDetCos, -1111
    End If
    
    Agregando = False
    Fg2_RowColChange
    
    Set RstRecetas = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : CargarTareasReceta
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGAS LAS TAREAS DE LA RECETA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub CargarTareasReceta()
    PreparaRSTTareas
    
    Set RstDetRec.ActiveConnection = Nothing
    Dim Rst As New ADODB.Recordset
    Dim A, B As Integer
    
    For B = 1 To Fg2.Rows - 1
        RST_Busq Rst, "SELECT pro_recetatar.idrec, pro_recetatar.idtar, pro_tareas.codigo, pro_tareas.descripcion, mae_unidades.abrev, pro_recetatar.cantidad, pro_recetatar.idunimed,pro_recetatar.orden " _
            & " FROM (pro_recetatar LEFT JOIN pro_tareas ON pro_recetatar.idtar = pro_tareas.id) LEFT JOIN mae_unidades ON pro_recetatar.idunimed = mae_unidades.id " _
            & " Where (((pro_recetatar.idrec) = " & NulosN(Fg2.TextMatrix(B, 7)) & ")) ORDER BY pro_recetatar.orden asc ", xCon

        Fg1.Rows = 1
        If Rst.RecordCount <> 0 Then Rst.MoveFirst
        Do While Not Rst.EOF
            RstDetTar.AddNew
            RstDetTar("idrec") = NulosN(Rst("idrec"))
            RstDetTar("idtar") = NulosN(Rst("idtar"))
            RstDetTar("codigo") = NulosC(Rst("codigo"))
            RstDetTar("descripcion") = NulosC(Rst("descripcion"))
            RstDetTar("abrev") = NulosC(Rst("abrev"))
            RstDetTar("cantidad") = NulosN(Rst("cantidad"))
            RstDetTar("idunimed") = NulosN(Rst("idunimed"))
            RstDetTar("orden") = NulosN(Rst("orden"))

            Rst.MoveNext
        Loop
        Set Rst = Nothing
    Next B
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO QUE SE EJECUTARA CUANDO SE CARGUE EL FORMULARIO
    SeEjecuto = False
    QueHace = 3
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Frame6.BackColor = &H8000000F
    Frame5.BackColor = &H8000000F
    TabOne1.CurrTab = 0
    Fg1.ColWidth(6) = 0
    Fg1.ColWidth(7) = 0
    Fg1.ColWidth(8) = 0

    Fg2.ColWidth(6) = 0
    Fg2.ColWidth(7) = 0
    Fg2.ColWidth(9) = 0
    
    Fg3.ColWidth(5) = 0
    Fg3.ColWidth(6) = 0
    Fg3.ColWidth(7) = 0
    
    ListandoInsumos = False
    Fg1.SelectionMode = flexSelectionByRow
    Fg2.SelectionMode = flexSelectionByRow
    Fg3.SelectionMode = flexSelectionByRow
    TabOne2.CurrTab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 1
        Exit Sub
    End If
End Sub

Private Sub ingre_1_1_Click()
    pIngredienteAdd
End Sub

Private Sub ingre_1_3_Click()
    pIngredienteDel
End Sub

Private Sub menu_2_1_Click()
    pTareaAdd
End Sub

Private Sub menu_2_3_Click()
    pTareaDel
End Sub

Private Sub rece_1_1_Click()
    CmdAddReceta_Click
End Sub

Private Sub rece_1_3_Click()
    CmdDelReceta_Click
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If NulosN(RstPro("numrec")) = 0 Then
            If QueHace = 3 Then
                MsgBox "El producto seleccionado no tiene una receta asignada," & Chr(13) _
                    & "Para agregar una Receta haga clic en el boton Modificar de la barra de herramientas", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    Cancel = True
            Else
                Muestradatos
            End If
        Else
            If QueHace <> 1 Then Muestradatos
        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : PreparaRST
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : DEFINE LA ESTRUCTURA DE DATOS DEL RECORDSET TEMPORAL RstDetRec
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub preparaRST()
    Dim xFun As New eps_librerias.FuncionesData
    
    Dim xCampos(9, 3) As String

    xCampos(0, 0) = "cod_receta":   xCampos(0, 1) = "C":      xCampos(0, 2) = "8"
    xCampos(1, 0) = "cod_item":     xCampos(1, 1) = "C":      xCampos(1, 2) = "16"
    xCampos(2, 0) = "cantidad":     xCampos(2, 1) = "D":      xCampos(2, 2) = "2"
    xCampos(3, 0) = "cod_unidad":   xCampos(3, 1) = "N":      xCampos(3, 2) = "5"
    xCampos(4, 0) = "tipo":         xCampos(4, 1) = "C":      xCampos(4, 2) = "15"
    xCampos(5, 0) = "descripcion":  xCampos(5, 1) = "C":      xCampos(5, 2) = "100"
    xCampos(6, 0) = "descabrevia":  xCampos(6, 1) = "C":      xCampos(6, 2) = "5"
    xCampos(7, 0) = "idrec":        xCampos(7, 1) = "N":      xCampos(7, 2) = "5"
    xCampos(8, 0) = "iditem":       xCampos(8, 1) = "N":      xCampos(8, 2) = "5"
    
    Set RstDetRec = xFun.CrearRstTMP(xCampos)
    RstDetRec.Open
End Sub

'*****************************************************************************************************
'* Nombre           : PreparaRST
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : DEFINE LA ESTRUCTURA DE DATOS DEL RECORDSET TEMPORAL RstDetTar
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub PreparaRSTTareas()
    Dim xFun As New eps_librerias.FuncionesData
    
    Dim xCampos(11, 3) As String

    xCampos(0, 0) = "idrec":       xCampos(0, 1) = "N":      xCampos(0, 2) = "8"
    xCampos(1, 0) = "idtar":       xCampos(1, 1) = "N":      xCampos(1, 2) = "8"
    xCampos(2, 0) = "codigo":      xCampos(2, 1) = "C":      xCampos(2, 2) = "9"
    xCampos(3, 0) = "descripcion": xCampos(3, 1) = "C":      xCampos(3, 2) = "100"
    xCampos(4, 0) = "abrev":       xCampos(4, 1) = "C":      xCampos(4, 2) = "5"
    xCampos(5, 0) = "cantidad":    xCampos(5, 1) = "D":      xCampos(5, 2) = "8"
    xCampos(6, 0) = "idunimed":    xCampos(6, 1) = "N":      xCampos(6, 2) = "8"
    xCampos(7, 0) = "orden":       xCampos(7, 1) = "N":      xCampos(7, 2) = "8"
    xCampos(8, 0) = "numper":      xCampos(8, 1) = "N":      xCampos(8, 2) = "8"
    xCampos(9, 0) = "horarr":      xCampos(9, 1) = "N":      xCampos(9, 2) = "8"
    xCampos(10, 0) = "aplpor":     xCampos(10, 1) = "N":     xCampos(10, 2) = "8"
    
    Set RstDetTar = xFun.CrearRstTMP(xCampos)
    RstDetTar.Open
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
    xHorIni = Time
    ActivarTool
    TabOne1.TabEnabled(0) = Not TabOne1.TabEnabled(0)
    Label1.Caption = "Modificado Recetas del Producto"
    xSeleccionoArchivo = False
    TabOne1.CurrTab = 1
    
    ActivarControles True
    TxtNotas.Locked = Not TxtNotas.Locked
    
    Fg2.Editable = flexEDKbd
    Fg3.Editable = flexEDKbd
    fg5.Editable = flexEDKbd
    Fg2.SetFocus
    
    Fg1.SelectionMode = flexSelectionFree
    Fg2.SelectionMode = flexSelectionFree
    Fg3.SelectionMode = flexSelectionFree
    fg5.SelectionMode = flexSelectionFree
    
    If Fg2.Rows = 1 Then
        habilitar cmdIngr, False
        habilitar cmdTar, False
        habilitar cmdCos, False
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL PROCESO DE INGRESAR O MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    ActivarTool
    ActivarControles False
    TabOne1.TabEnabled(0) = Not TabOne1.TabEnabled(0)
    TabOne1.CurrTab = 0
    Label1.Caption = "Recetas del Producto"
    TxtNotas.Locked = Not TxtNotas.Locked
    
    Fg1.SelectionMode = flexSelectionByRow
    Fg2.SelectionMode = flexSelectionByRow
    Fg3.SelectionMode = flexSelectionByRow
End Sub

'*****************************************************************************************************
'* Nombre           : ActivarTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS BOTONES DE LA BARRA DE HERRAMIENTAS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ActivarTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

'*****************************************************************************************************
'* Nombre           : ActivarControles
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES DEL FORMULARIO
'* Paranetros       : NOMBRE    |  TIPO         |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    band      |               |
'* Devuelve         :
'*****************************************************************************************************
Sub ActivarControles(band As Boolean)
    CmdAddReceta.Enabled = band
    CmdDelReceta.Enabled = band
    CmdVerProc.Enabled = Not CmdVerProc.Enabled
        
    habilitar cmdIngr, band
    habilitar cmdTar, band
    habilitar cmdCos, band
End Sub

'*****************************************************************************************************
'* Nombre           : CargarIngredientesReceta
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS INGREDIENTES DE UNA RECETA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub CargarIngredientesReceta()
    preparaRST
    
    Set RstDetRec.ActiveConnection = Nothing
    Dim Rst As New ADODB.Recordset
    Dim A, B As Integer
    
    For B = 1 To Fg2.Rows - 1
        RST_Busq Rst, "SELECT pro_recetains.idrec, pro_receta.codrec, pro_recetains.iditem, pro_recetains.idunimed, alm_inventario.codpro, " _
            & " alm_inventario.descripcion, mae_tipoproducto.descripcion AS desctippro, alm_inventario.tippro, mae_unidades.abrev, pro_recetains.canpro " _
            & " FROM mae_tipoproducto RIGHT JOIN (pro_receta RIGHT JOIN ((alm_inventario RIGHT JOIN pro_recetains ON alm_inventario.id = pro_recetains.iditem) " _
            & " LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id) ON pro_receta.id = pro_recetains.idrec) " _
            & " ON mae_tipoproducto.id = alm_inventario.tippro Where (((pro_recetains.idrec) = " & NulosN(Fg2.TextMatrix(B, 7)) & ")) ORDER BY alm_inventario.descripcion", xCon
        
        Fg1.Rows = 1
        If Rst.State = 1 Then
            If Rst.RecordCount <> 0 Then Rst.MoveFirst
            Do While Not Rst.EOF
                RstDetRec.AddNew
                RstDetRec("cod_item") = NulosC(Rst("codpro"))
                RstDetRec("cod_receta") = NulosC(Rst("codrec"))
                RstDetRec("cantidad") = NulosN(Rst("canpro"))
                RstDetRec("cod_unidad") = NulosC(Rst("idunimed"))
                RstDetRec("tipo") = NulosC(Rst("desctippro"))
                RstDetRec("descripcion") = NulosC(Rst("descripcion"))
                RstDetRec("descabrevia") = NulosC(Rst("abrev"))
                RstDetRec("idrec") = NulosN(Rst("idrec"))
                RstDetRec("iditem") = NulosN(Rst("iditem"))
                Rst.MoveNext
            Loop
        End If
        Set Rst = Nothing
    Next B
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 2 Then
        Modificar
    End If
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstPro.Requery
            RstPro.MoveFirst
            Dg1.Refresh
            
            RstPro.Find "id=" & mIdRegistro
            If RstPro.EOF = True Then RstPro.MoveFirst
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 8 Then Filtrar
    
    If Button.Index = 9 Then
        TabOne1.CurrTab = 0
        RstPro.Filter = adFilterNone
        TDB_FiltroLimpiar Dg1
        RstPro.Requery
        Dg1.Refresh
    End If
    
    If Button.Index = 10 Then Buscar
    
    If Button.Index = 14 Then
        Set RstPro = Nothing
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : FUNCCION
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA pro_receta, ESTA FUNCION DEVUELVE VERDADERO CUANDO
'*                    TIENE EXITO
'* Paranetros       :
'* Devuelve         : Boolean
'*****************************************************************************************************
Function Grabar() As Boolean
    Dim A&, B&, xNum&
    
    Dim fs As New Scripting.FileSystemObject
    Dim d As Scripting.Folder
    
    If Fg2.Rows = 1 Then
        MsgBox "No ha especificado recetas para el producto " & Trim(TxtDescripcion.Text), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Grabar = False
        Exit Function
    End If
    
    If Fg2.Rows - 1 > 1 Then
        xNum = -1
        For A = 1 To Fg2.Rows - 1
            If xNum = NulosN(Fg2.TextMatrix(A, 5)) Then
                MsgBox "2 recetas no pueden tener la misma prioridad", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Fg2.SetFocus
                Exit Function
            End If
            
            If NulosC(Fg2.TextMatrix(A, 2)) = "" Then
                MsgBox "No ha especificado en nombre de la receta " + Fg2.TextMatrix(Fg2.Row, 1), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Fg2.SetFocus
                Exit Function
            End If
            
            If NulosC(Fg2.TextMatrix(A, 5)) = "" Then
                MsgBox "La receta " & Fg2.TextMatrix(Fg2.Row, 1) & vbCr & "No tiene prioridad", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Fg2.SetFocus
                Exit Function
            End If
            
            If A = Fg2.Rows - 1 Then
                Exit For
            End If
            
            xNum = NulosN(Fg2.TextMatrix(A, 5))
        Next A
    End If
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstTar As New ADODB.Recordset
    Dim RstCos As New ADODB.Recordset
    Dim xId As Double
    
    On Error GoTo LaCague

    xCon.BeginTrans
    
    ' eliminando recetas que no se usan
    Dim nSQLIdRec As String
    Dim RstTemp As New ADODB.Recordset
    nSQLIdRec = GRID_GENERAR_SQL_ID(Fg2, 7, "and id ", "not in", True)
    If nSQLIdRec <> "" Then
       RST_Busq RstTemp, "Select * from pro_receta where iditem= " & RstPro("id") & " " & nSQLIdRec, xCon
       If RstTemp.State = 1 Then
            If RstTemp.RecordCount <> 0 Then
                Do While Not RstTemp.EOF
                    ' eliminamos sus insumos y tareas
                    xCon.Execute "DELETE * FROM  pro_recetains WHERE idrec = " & NulosN(RstTemp("id")) & ""
                    xCon.Execute "DELETE * FROM  pro_recetatar WHERE idrec = " & NulosN(RstTemp("id")) & ""
                    xCon.Execute "DELETE * FROM  pro_recetacos WHERE idrec = " & NulosN(RstTemp("id")) & ""
                    ' eliminanos el registro de receta
                    xCon.Execute "DELETE * FROM  pro_receta WHERE id = " & NulosN(RstTemp("id")) & ""
                    RstTemp.MoveNext
                Loop
            End If
       End If
    End If
    Set RstTemp = Nothing
    
    RST_Busq RstDet, "SELECT TOP 1 * FROM pro_recetains", xCon
    RST_Busq RstTar, "SELECT TOP 1 * FROM pro_recetatar", xCon
    RST_Busq RstCos, "SELECT TOP 1 * FROM pro_recetacos", xCon
    
    For A = 1 To Fg2.Rows - 1
        ' buscamos la receta
        RST_Busq RstCab, "SELECT * FROM pro_receta WHERE id = " & NulosN(Fg2.TextMatrix(A, 7)) & "", xCon
        
        ' eliminamos sus insumos y tareas
        xCon.Execute "DELETE * FROM  pro_recetains WHERE idrec = " & NulosN(Fg2.TextMatrix(A, 7)) & ""
        xCon.Execute "DELETE * FROM  pro_recetatar WHERE idrec = " & NulosN(Fg2.TextMatrix(A, 7)) & ""
        xCon.Execute "DELETE * FROM  pro_recetacos WHERE idrec = " & NulosN(Fg2.TextMatrix(A, 7)) & ""
    
        If RstCab.RecordCount = 0 Then
            xId = HallaCodigoTabla("pro_receta", xCon, "id")
            RstCab.AddNew
            RstCab("id") = xId
        Else
            xId = NulosN(Fg2.TextMatrix(A, 7))
        End If
        
        
        mIdRegistro = RstPro("id")
        
        RstCab("iditem") = RstPro("id")
        RstCab("codrec") = Fg2.TextMatrix(A, 1)
        RstCab("descripcion") = Fg2.TextMatrix(A, 2)
        RstCab("idunimed") = Busca_Codigo(Fg2.TextMatrix(A, 3), "abrev", "id", "mae_unidades", "C", xCon)
        RstCab("cantidad") = NulosN(Fg2.TextMatrix(A, 4))
        RstCab("prirec") = NulosN(Fg2.TextMatrix(A, 5))
        RstCab("observaciones") = NulosC(Fg2.TextMatrix(A, 6))
        
        '*************************************************************************
'        RstCab("idtiptrab") = NulosN(Fg2.TextMatrix(A, 12)) ' tipo de trabajo
'        RstCab("idformapag") = NulosN(Fg2.TextMatrix(A, 13)) ' Forma de Pago
        '*************************************************************************
        
        Dim Ruta As String
        Ruta = LeerLineaINI(App.Path & "\seven.ini", "RUTAAR", "RUTAS") & "procedimientos\"
        If NulosC(Fg2.TextMatrix(A, 8)) <> "" Then
        
            If NulosN(Fg2.TextMatrix(Fg2.Row, 9)) = 1 Then
                If fs.FolderExists(Trim(Ruta)) = True Then
                    If fs.FileExists(Trim(Fg2.TextMatrix(A, 8))) = True Then
                        fs.CopyFile Trim(Fg2.TextMatrix(A, 8)), Ruta + "\" & "PROC-" & Fg2.TextMatrix(A, 1) + ".htm"
                        RstCab("archpro") = "PROC-" & Fg2.TextMatrix(A, 1) + ".htm"
                    Else
                        MsgBox "No se ha encontrado el archivo especificado", vbInformation + vbOKOnly + vbDefaultButton1, "Normas legales"
                        xCon.RollbackTrans
                        Exit Function
                    End If
                Else
                    MsgBox "No se ha encontrado la carpeta especificada", vbInformation + vbOKOnly + vbDefaultButton1, "Normas legales"
                    xCon.RollbackTrans
                    Exit Function
                End If
            
            End If
            
        End If
        
        
'    If NulosC(TxtArch.Text) <> "" Then
'        If Mid(TxtArch.Text, 1, 1) <> "0" Then
'            Ruta = AP_RUTATX + Format(Val(TxtIdBL.Text), "000000") '+ "\"
'
'            If fs.FolderExists(Trim(Ruta)) = True Then
'                If fs.FileExists(Trim(TxtArch.Text)) = True Then
'                    fs.CopyFile Trim(TxtArch.Text), Ruta + "\" + Format(xRstGraba("id"), "0000000") + ".htm"
'                Else
'                    MsgBox "No se ha encontrado el archivo especificado", vbInformation + vbOKOnly + vbDefaultButton1, "Normas legales"
'                    xConeccion.RollbackTrans
'                    Exit Sub
'                End If
'            Else
'                MsgBox "No se ha encontrado la carpeta especificada", vbInformation + vbOKOnly + vbDefaultButton1, "Normas legales"
'                xConeccion.RollbackTrans
'                Exit Sub
'            End If
'        End If
'    Else
'        If TipNorAct <> Val(TxtIdBL.Text) Then
'            'Se ha modificado el tipo de norma legal
'            Dim RutaAnt As String
'            Dim RutaNew As String
'            Dim NomArch As String
'
'            RutaAnt = AP_RUTATX + Format(Val(TipNorAct), "000000") + "\"
'            RutaNew = AP_RUTATX + Format(Val(TxtIdBL.Text), "000000") + "\"
'
'            NomArch = RutaAnt + Trim(xNomArhc)
'            If fs.FolderExists(Trim(RutaAnt)) = True Then
'                If fs.FileExists(Trim(NomArch)) = True Then
'                    fs.CopyFile NomArch, RutaNew
'                    fs.DeleteFile NomArch
'                Else
'                    MsgBox "No se ha encontrado el archivo especificado", vbInformation + vbOKOnly + vbDefaultButton1, "Normas legales"
'                    xConeccion.RollbackTrans
'                    Exit Sub
'                End If
'            Else
'                MsgBox "No se ha encontrado la carpeta especificada", vbInformation + vbOKOnly + vbDefaultButton1, "Normas legales"
'                xConeccion.RollbackTrans
'                Exit Sub
'            End If
'        Else
'            'se modifico el archivo de texto
'            TxtArch.Text = xNomArhc
'            If Mid(TxtArch.Text, 1, 1) <> "0" Then
'                Ruta = AP_RUTATX + Format(Val(TxtIdBL.Text), "000000") '+ "\"
'
'                If fs.FolderExists(Trim(Ruta)) = True Then
'                    If fs.FileExists(Trim(TxtArch.Text)) = True Then
'                        fs.CopyFile Trim(TxtArch.Text), Ruta + "\" + Format(xRstGraba("id"), "0000000") + ".htm"
'                    Else
'                        MsgBox "No se ha encontrado el archivo especificado", vbInformation + vbOKOnly + vbDefaultButton1, "Normas legales"
'                        xConeccion.RollbackTrans
'                        Exit Sub
'                    End If
'                Else
'                    MsgBox "No se ha encontrado la carpeta especificada", vbInformation + vbOKOnly + vbDefaultButton1, "Normas legales"
'                    xConeccion.RollbackTrans
'                    Exit Sub
'                End If
'            End If
'        End If
'    End If
        
        RstCab.Update
        
        ' grabamos los insumos de la receta
        RstDetRec.Filter = adFilterNone
        RstDetRec.Filter = "idrec = " & NulosN(Fg2.TextMatrix(A, 7))
        
        If RstDetRec.RecordCount <> 0 Then
            RstDetRec.MoveFirst
            Do While Not RstDetRec.EOF
                RstDet.AddNew
                RstDet("idrec") = xId
                RstDet("iditem") = NulosN(RstDetRec("iditem"))
                RstDet("idunimed") = NulosN(RstDetRec("idunimed"))
                RstDet("canpro") = NulosN(RstDetRec("canpro"))
                RstDet("canpropra") = NulosN(RstDetRec("canpropra"))
                RstDet.Update
                RstDetRec.MoveNext
            Loop
        End If
        
        ' grabamos las tareas de la receta
        RstDetTar.Filter = adFilterNone
        RstDetTar.Filter = "idrec = " & NulosN(Fg2.TextMatrix(A, 7))
        
        If RstDetTar.RecordCount <> 0 Then
            RstDetTar.MoveFirst
            Do While Not RstDetTar.EOF
                RstTar.AddNew
                RstTar("idrec") = xId
                RstTar("idtar") = NulosN(RstDetTar("idtar"))
                RstTar("idunimed") = NulosN(RstDetTar("idunimed"))
                RstTar("cantidad") = NulosN(RstDetTar("cantidad"))
                RstTar("orden") = NulosN(RstDetTar("orden"))
                
                RstTar("factor") = NulosN(RstDetTar("factor"))
                RstTar("jornalkg") = NulosN(RstDetTar("jornalkg"))
                RstTar("costokg") = NulosN(RstDetTar("costokg"))
                RstTar("costohr") = NulosN(RstDetTar("costohr"))
                
                If IsNull(RstDetTar("horarr")) = False Then
                    RstTar("horarr") = RstDetTar("horarr")
                Else
                    RstTar("horarr") = "00:00"
                End If
                RstTar("numper") = NulosN(RstDetTar("numper"))
                RstTar("aplpor") = NulosN(RstDetTar("aplpor"))
                
                '*************************************************************
                RstTar("idarea") = NulosN(RstDetTar("idarea"))
                RstTar("idtiptrab") = NulosN(RstDetTar("idtiptrab"))
                RstTar("idformapag") = NulosN(RstDetTar("idformapag"))
                '*************************************************************
                
                RstTar.Update
                
                RstDetTar.MoveNext
            Loop
        End If
        
        
        ' grabamos los centros de costo de la receta
        RstDetCos.Filter = adFilterNone
        RstDetCos.Filter = "idrec = " & NulosN(Fg2.TextMatrix(A, 7))
        
        If RstDetCos.RecordCount <> 0 Then
            RstDetCos.MoveFirst
            Do While Not RstDetCos.EOF
                RstCos.AddNew
                RstCos("idrec") = xId
                RstCos("idcuenta") = NulosN(RstDetCos("idcuenta"))
                RstCos.Update
                RstDetCos.MoveNext
            Loop
        End If
    Next A
    
    'grabamos el movimiento en la tabla var_edicion, el codigo del movimiento es iditem
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, NulosN(RstPro("id"))
        
    xCon.CommitTrans
    Set RstDet = Nothing
    Set RstTar = Nothing
    Set RstCos = Nothing
    Set RstCab = Nothing
    MsgBox "Las recetas se guardaron con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Exit Function
    
LaCague:
    xCon.RollbackTrans
    Set RstDet = Nothing
    Set RstTar = Nothing
    Set RstCos = Nothing
    Set RstCab = Nothing
    
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
End Function

'*****************************************************************************************************
'* Nombre           : Buscar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA UNA BUSQUEDA EN EL RECORDSET RstPro
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Buscar()
    TabOne1.CurrTab = 0
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCodItem As String
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":       xCampos(1, 1) = "codpro":        xCampos(1, 2) = "2000":    xCampos(1, 3) = "C"

    xform.SQLCad = "SELECT item.* ,rec.recpri  FROM " _
            & " (SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.tippro, (SELECT Count([iditem]) AS numrec From pro_receta WHERE (pro_receta.iditem=alm_inventario.id)) AS numrec " _
            & " FROM mae_unidades RIGHT JOIN  alm_inventario ON mae_unidades.id = alm_inventario.idunimed Where (((alm_inventario.tippro) = 3) ) " _
            & " ) as item LEFT JOIN ( SELECT pro_receta.iditem, pro_receta.codrec AS recpri, pro_receta.prirec FROM pro_receta WHERE (((pro_receta.prirec)=1)) ) as rec ON item.id = rec.iditem "
    
    xform.titulo = "Buscando Productos"
    
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.BOF = True Or xRs.EOF = True Then Exit Sub
        RstPro.MoveFirst
        RstPro.Find "codpro = '" & xRs("codpro") & "'"
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : Filtrar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA UN FILTRO EN EL RECORDSET RstPro
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Filtrar()
    Dim xform As New eps_librerias.FormFiltrar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":  xCampos(0, 2) = "C":         xCampos(0, 3) = "6200"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "codpro":       xCampos(1, 2) = "C":         xCampos(1, 3) = "2000"
    
    Set xform.Coneccion = xCon
    Set xform.Rst = RstPro
    Set RstPro = xform.FiltrarReg(xCampos)
    Set Dg1.DataSource = RstPro
    Dg1.Refresh
End Sub

Sub MostrarFrameDuplica()
    LblIdProd.Caption = RstPro("id")
    LblCodPro.Caption = RstPro("codpro")
    LblDescPro.Caption = RstPro("descripcion")
    TabOne1.Enabled = False
    Toolbar1.Enabled = False
    FrmDuplicaReceta.Left = 1050
    FrmDuplicaReceta.Top = 1830
    FrmDuplicaReceta.Visible = True
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index = 1 Then
        Modificar
    End If
    
    If ButtonMenu.Index = 2 Then
        MostrarFrameDuplica
    End If
End Sub

Private Sub TxtNotas_Validate(Cancel As Boolean)
    If QueHace <> 3 Then
        Fg2.TextMatrix(Fg2.Row, 6) = NulosC(TxtNotas.Text)
    End If
End Sub

Private Sub cmdIngr_click(Index As Integer)
    Select Case Index
        Case 0 ' agregar ingrediente
            pIngredienteAdd
        
        Case 1 ' eliminar ingrediente
            pIngredienteDel
        
        Case 2 ' exportar excel
            pExportarExcel 1
    End Select
End Sub

Private Sub cmdTar_click(Index As Integer)
    Select Case Index
        Case 0 ' agregar tarea
            pTareaAdd
        
        Case 1 ' eliminar tarea
            pTareaDel
        
        Case 2 ' exportar excel
            pExportarExcel 2
    End Select
End Sub

'*****************************************************************************************************
'* Nombre           : pIngredienteAdd
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : AGREGA UNA FILA AL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pIngredienteAdd()
    Dim mCol%
    Dim fInsertar As Boolean
    If QueHace = 3 Then Exit Sub
    Agregando = True
    
    If Fg1.Rows > Fg1.FixedRows Then
        If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 1)) = 0 Then
            MsgBox "Seleccione el Tipo de Ingrediente ", vbExclamation, xTitulo
        Else
            fInsertar = True
        End If
    Else
        fInsertar = True
    End If
    
    If fInsertar = True Then Fg1.AddItem ""
    
    Fg1.Row = Fg1.Rows - 1
    Fg1.Col = 1
    
    Fg1.SetFocus
    Agregando = False
End Sub

'*****************************************************************************************************
'* Nombre           : pConfigurarGrilla
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA LAS COLUMNAS Y CABECERAS DE LOS CONTROLES Fg1, Fg3
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pConfigurarGrilla()
    With Fg2      ' Producto
        .Rows = 1
    End With
    
    With Fg1      ' de los ingredientes
        .Rows = 1
        .Cols = 10
        .FixedRows = 1
        .RowHeight(0) = 300
        .TextMatrix(0, 1) = "Tipo Ingrediente": .ColWidth(1) = 1440:  .ColAlignment(1) = flexAlignLeftCenter:  .Row = 0: .Col = 1: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 2) = "Código":           .ColWidth(2) = 1200:  .ColAlignment(2) = flexAlignLeftCenter:  .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 3) = "Descripción":      .ColWidth(3) = 3800:  .ColAlignment(3) = flexAlignLeftCenter:  .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 4) = "Und":              .ColWidth(4) = 700:  .ColAlignment(4) = flexAlignLeftCenter:  .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 5) = "Can. Teorica":     .ColWidth(5) = 1100:  .ColAlignment(5) = flexAlignRightCenter: .Row = 0: .Col = 5: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 6) = "IdItem":           .ColWidth(6) = 0:  '.ColAlignment(6) = flexAlignLeftCenter:  .Row = 0: .Col = 6: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 7) = "IdUnd":            .ColWidth(7) = 0:  '.ColAlignment(7) = flexAlignRightCenter: .Row = 0: .Col = 7: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 8) = "IdReceta":         .ColWidth(8) = 0:  '.ColAlignment(8) = flexAlignLeftCenter:  .Row = 0: .Col = 8: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 9) = "Can. Practica":    .ColWidth(9) = 1100:  .ColAlignment(9) = flexAlignLeftCenter:  .Row = 0: .Col = 9: .CellAlignment = flexAlignLeftCenter
        .ColFormat(5) = "0.000000"
        .ColFormat(9) = "0.000000"
        .SelectionMode = flexSelectionByRow
    End With
    
    With Fg3 '--de las tareas
        .Rows = 1
        .Cols = 19
        .FixedRows = 1
        '.RowHeight(0) = 300
        .TextMatrix(0, 1) = "Código":           .ColWidth(1) = 1440:    .ColAlignment(1) = flexAlignLeftCenter:   .Row = 0: .Col = 1: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 2) = "Descripción":      .ColWidth(2) = 3250:    .ColAlignment(2) = flexAlignLeftCenter:   .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 3) = "IdReceta":         .ColWidth(3) = 0:      '.ColAlignment(3) = flexAlignLeftCenter:   .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 4) = "IdTar":            .ColWidth(4) = 0:      '.ColAlignment(4) = flexAlignRightCenter:  .Row = 0: .Col = 4: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 5) = "IdUnd":            .ColWidth(5) = 0:      '.ColAlignment(5) = flexAlignLeftCenter:   .Row = 0: .Col = 5: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 6) = "Und":              .ColWidth(6) = 0:       .ColAlignment(6) = flexAlignLeftCenter:   .Row = 0: .Col = 6: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 7) = "Arranca":          .ColWidth(7) = 0:       .ColAlignment(7) = flexAlignRightCenter:  .Row = 0: .Col = 7: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 8) = "Factor ":          .ColWidth(8) = 0:       .ColAlignment(8) = flexAlignRightCenter:  .Row = 0: .Col = 8: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 9) = "Nº Oper.":         .ColWidth(9) = 0:       .ColAlignment(9) = flexAlignRightCenter:  .Row = 0: .Col = 9: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 10) = "Costo/Kg":        .ColWidth(10) = 0:      .ColAlignment(10) = flexAlignRightCenter: .Row = 0: .Col = 10: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 11) = "Rendto":          .ColWidth(11) = 0:      .ColAlignment(11) = flexAlignRightCenter: .Row = 0: .Col = 11: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 12) = "Orden":           .ColWidth(12) = 550:    .ColAlignment(12) = flexAlignRightCenter: .Row = 0: .Col = 12: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 13) = "Area":            .ColWidth(13) = 1500:    .ColAlignment(13) = flexAlignRightCenter: .Row = 0: .Col = 13: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 14) = "Tip. Trab.":      .ColWidth(14) = 1250:    .ColAlignment(14) = flexAlignRightCenter: .Row = 0: .Col = 14: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 15) = "Form. Pag":       .ColWidth(15) = 1250:    .ColAlignment(15) = flexAlignRightCenter: .Row = 0: .Col = 15: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 16) = "idarea":          .ColWidth(16) = 0:    .ColAlignment(16) = flexAlignRightCenter: .Row = 0: .Col = 16: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 17) = "idtiptrab":       .ColWidth(17) = 0:    .ColAlignment(17) = flexAlignRightCenter: .Row = 0: .Col = 17: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 18) = "idformapag":      .ColWidth(18) = 0:    .ColAlignment(18) = flexAlignRightCenter: .Row = 0: .Col = 18: .CellAlignment = flexAlignRightCenter
        
        '.ColFormat(7) = "##:##"
        .ColFormat(8) = "0.000000"
        .ColFormat(9) = "00"
        .ColFormat(10) = "0.000000"
        .ColFormat(11) = "0.00"
        .SelectionMode = flexSelectionByRow
    End With
    
    With fg5
        .Rows = .FixedRows
    End With
    
    ' Producto
    GRID_COMBOLIST Fg2, 8     ' unidad
    
    ' Insumos
    GRID_COMBOLIST Fg1, 2    ' descripcion
    GRID_COMBOLIST Fg1, 4    ' unidad
    
    
    ' Tareas
    GRID_COMBOLIST Fg3, 1    ' tarea
    GRID_COMBOLIST Fg3, 13    ' Area
    GRID_COMBOLIST Fg3, 14    ' Tipo de Trabajo
    GRID_COMBOLIST Fg3, 15    ' Forma de PAgo
    
    ' Tipo de Ingrediente
    Dim RstTmp As New ADODB.Recordset
    Dim tFormat$
    RST_Busq RstTmp, "SELECT mae_tipoproducto.id, mae_tipoproducto.descripcion FROM mae_tipoproducto WHERE (((mae_tipoproducto.id) In (1,3,4,8))) ORDER BY mae_tipoproducto.descripcion; ", xCon
    tFormat = Fg1.BuildComboList(RstTmp, "descripcion", "id", vbYellow)
    Fg1.ColComboList(1) = tFormat
    Set RstTmp = Nothing
    DoEvents
End Sub

'*****************************************************************************************************
'* Nombre           : pDefinirRstTmp
'* Tipo             : PROCEDIMIENTO
'* Descripcion      :
'* Paranetros       : NOMBRE    |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    eTipo     |  e_PROGRAMA       |
'*                    rst       |  ADODB.Recordset  |
'*                    IdReceta  |                   |
'* Devuelve         :
'*****************************************************************************************************
Private Sub pDefinirRstTmp(eTipo As e_PROGRAMA, Rst As ADODB.Recordset, Optional IdReceta = -1)
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Dim mIdReceta
    
    If IdReceta = -1 Then mIdReceta = -999    ' comodin
    
    If eTipo = E_INSUMO Then
        nSQL = "SELECT mae_tipoproducto.descripcion AS tipoproducto, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro, pro_recetains.idrec, alm_inventario.tippro AS idtipo, pro_recetains.iditem, alm_inventario.idunimed, pro_recetains.canpropra " _
            + vbCr + " FROM ((mae_tipoproducto RIGHT JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) RIGHT JOIN pro_recetains ON alm_inventario.id = pro_recetains.iditem) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id " _
            + vbCr + " WHERE pro_recetains.idrec = " & IdReceta & " ORDER BY mae_tipoproducto.descripcion ASC "
            
    ElseIf eTipo = e_TAREA Then
        nSQL = "SELECT pro_recetatar.idrec, pro_recetatar.idtar, pro_recetatar.idunimed, pro_tareas.codigo, pro_tareas.descripcion, mae_unidades.abrev, pro_recetatar.cantidad, pro_recetatar.orden, pro_recetatar.factor, pro_recetatar.jornalkg, pro_recetatar.costokg, pro_recetatar.costohr, pro_recetatar.numper, pro_recetatar.horarr, pro_recetatar.aplpor, pro_recetatar.idtiptrab, pro_tiptrab.descripcion AS destiptrab, pro_recetatar.idformapag, pro_formapag.descripcion AS desformapag, pro_recetatar.idarea, mae_area.descripcion AS area " _
            + vbCr + "FROM ((((pro_recetatar LEFT JOIN pro_tareas ON pro_recetatar.idtar = pro_tareas.id) LEFT JOIN mae_unidades ON pro_recetatar.idunimed = mae_unidades.id) LEFT JOIN pro_tiptrab ON pro_recetatar.idtiptrab = pro_tiptrab.id) LEFT JOIN pro_formapag ON pro_recetatar.idformapag = pro_formapag.id) LEFT JOIN mae_area ON pro_recetatar.idarea = mae_area.id " _
            + vbCr + "Where pro_recetatar.idrec = " & IdReceta & " " _
            + vbCr + "ORDER BY pro_recetatar.orden; "
            
    ElseIf eTipo = E_COSTO Then
        nSQL = "SELECT pro_recetacos.idrec, con_planctas.cuenta, con_planctas.descripcion, pro_recetacos.idcuenta " _
            + vbCr + "FROM pro_recetacos INNER JOIN con_planctas ON pro_recetacos.idcuenta = con_planctas.id " _
            + vbCr + "Where (((pro_recetacos.IDREC) = " & IdReceta & ")) " _
            + vbCr + "ORDER BY con_planctas.cuenta;"
    Else
    
    End If
    
    RST_Busq RstTmp, nSQL, xCon
    ' definir la estructura del recordset
    If Rst.State = 0 Then DEFINIR_RST_TMP Rst, RstTmp
    ' cargar los datos al recordset temporal
    If RstTmp.RecordCount <> 0 Then CARGAR_RST_TMP Rst, RstTmp
End Sub

'*****************************************************************************************************
'* Nombre           : pExportarExcel
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : Exportar Excel tanto lista de materiales, tareas
'* Paranetros       : NOMBRE    |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Tipo      |  Integer    | tipo = indica que se va exporar  1=Lista de materiales
'*                                               2 = Lista de Tareas
'* Devuelve         :
'*****************************************************************************************************
Private Sub pExportarExcel(Tipo As Integer)
    On Error GoTo error
    Dim ObjExport As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    ObjExport.VSFlexGrid_Exportar_MSExcel xCon, IIf(Tipo = 1, Fg1, Fg3), "Receta de Producción - " & IIf(Tipo = 1, "Lista de Materiales", "Lista de Tareas"), Fg2.TextMatrix(Fg2.Row, 2), , "Receta de Producción"
    Set ObjExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub

error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pExportarExcel"
End Sub


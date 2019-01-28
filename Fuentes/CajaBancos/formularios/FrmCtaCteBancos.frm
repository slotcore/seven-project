VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmCtaCteBancos 
   Caption         =   "Contabilidad - Libro Bancos"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   3600
      Left            =   10410
      TabIndex        =   27
      Top             =   -1980
      Visible         =   0   'False
      Width           =   8325
      Begin VB.CommandButton CmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   350
         Left            =   3480
         TabIndex        =   28
         Top             =   3165
         Width           =   1365
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg4 
         Height          =   2670
         Left            =   60
         TabIndex        =   29
         Top             =   420
         Width           =   8190
         _cx             =   14446
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
         BackColorSel    =   64
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmCtaCteBancos.frx":0000
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
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   8310
         Y1              =   3585
         Y2              =   3585
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   8310
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   8310
         X2              =   8310
         Y1              =   15
         Y2              =   3585
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   3555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resumen de la Conciliacion"
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
         Left            =   240
         TabIndex        =   30
         Top             =   90
         Width           =   2370
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000D&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000D&
         Height          =   300
         Left            =   45
         Top             =   45
         Width           =   8235
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6270
      Left            =   -15
      TabIndex        =   15
      Top             =   1440
      Width           =   11895
      _cx             =   20981
      _cy             =   11060
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
      Caption         =   "  Libro Bancos  |  Conciliación  "
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
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   5850
         Left            =   12540
         TabIndex        =   17
         Top             =   375
         Width           =   11805
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   2580
            Left            =   30
            TabIndex        =   32
            Top             =   3255
            Width           =   11760
            _cx             =   20743
            _cy             =   4551
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
            Caption         =   " Cheques Emitidos y no Cobrados (Mes Anterior)| Otros Movimientos no Considerados "
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
            Begin VB.Frame Frame8 
               BackColor       =   &H00C0E0FF&
               BorderStyle     =   0  'None
               Height          =   2160
               Left            =   12405
               TabIndex        =   34
               Top             =   45
               Width           =   11670
               Begin VB.Frame Frame2 
                  Height          =   2115
                  Left            =   10050
                  TabIndex        =   37
                  Top             =   -15
                  Width           =   1590
                  Begin VB.CommandButton CmdAgregaMov 
                     Caption         =   "Agregar Movimiento"
                     Enabled         =   0   'False
                     Height          =   495
                     Left            =   165
                     Style           =   1  'Graphical
                     TabIndex        =   39
                     Top             =   555
                     Width           =   1260
                  End
                  Begin VB.CommandButton CmdDelMov 
                     Caption         =   "Eliminar Movimiento"
                     Enabled         =   0   'False
                     Height          =   495
                     Left            =   165
                     Style           =   1  'Graphical
                     TabIndex        =   38
                     Top             =   1080
                     Width           =   1260
                  End
               End
               Begin VB.TextBox TxtTotHab 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Height          =   300
                  Left            =   7395
                  Locked          =   -1  'True
                  TabIndex        =   36
                  Text            =   "TxtTotHab"
                  Top             =   1815
                  Width           =   1095
               End
               Begin VB.TextBox TxtTotDeb 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Height          =   300
                  Left            =   6300
                  Locked          =   -1  'True
                  TabIndex        =   35
                  Text            =   "TxtTotDeb"
                  Top             =   1815
                  Width           =   1095
               End
               Begin VSFlex7Ctl.VSFlexGrid Fg3 
                  Height          =   1725
                  Left            =   45
                  TabIndex        =   40
                  Top             =   75
                  Width           =   9945
                  _cx             =   17542
                  _cy             =   3043
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
                  BackColorSel    =   64
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
                  AllowUserResizing=   0
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   20
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmCtaCteBancos.frx":00B8
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
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  Caption         =   "Total ==>"
                  Height          =   195
                  Left            =   5325
                  TabIndex        =   41
                  Top             =   1860
                  Width           =   675
               End
            End
            Begin VB.Frame Frame7 
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   0  'None
               Height          =   2160
               Left            =   45
               TabIndex        =   33
               Top             =   45
               Width           =   11670
               Begin VB.TextBox TxtTotHab5 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Height          =   300
                  Left            =   9465
                  Locked          =   -1  'True
                  TabIndex        =   44
                  Text            =   "TxtTotHab5"
                  Top             =   1815
                  Width           =   1095
               End
               Begin VB.TextBox TxtTotDeb5 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Height          =   300
                  Left            =   8370
                  Locked          =   -1  'True
                  TabIndex        =   43
                  Text            =   "TxtTotDeb5"
                  Top             =   1815
                  Width           =   1095
               End
               Begin VSFlex7Ctl.VSFlexGrid Fg5 
                  Height          =   1725
                  Left            =   45
                  TabIndex        =   42
                  Top             =   75
                  Width           =   11595
                  _cx             =   20452
                  _cy             =   3043
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
                  BackColorSel    =   64
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
                  AllowUserResizing=   0
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   20
                  Cols            =   9
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmCtaCteBancos.frx":01A9
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
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  Caption         =   "Total ==>"
                  Height          =   195
                  Left            =   7440
                  TabIndex        =   45
                  Top             =   1860
                  Width           =   675
               End
            End
         End
         Begin VB.TextBox TxtTot2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   9450
            Locked          =   -1  'True
            TabIndex        =   25
            Text            =   "TxtTot2"
            Top             =   2910
            Width           =   1095
         End
         Begin VB.TextBox TxtTot1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   8355
            Locked          =   -1  'True
            TabIndex        =   24
            Text            =   "TxtTot1"
            Top             =   2910
            Width           =   1095
         End
         Begin VB.TextBox TxtImpEstado 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   21
            Text            =   "TxtImpEstado"
            Top             =   105
            Width           =   1200
         End
         Begin VB.CommandButton CmdProcesar 
            Caption         =   "&Ver Conciliacion"
            Height          =   630
            Left            =   9510
            TabIndex        =   20
            Top             =   15
            Width           =   2280
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg2 
            Height          =   2205
            Left            =   15
            TabIndex        =   19
            Top             =   690
            Width           =   11775
            _cx             =   20770
            _cy             =   3889
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
            BackColorSel    =   64
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
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   20
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmCtaCteBancos.frx":02F5
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
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Total ==>"
            Height          =   195
            Left            =   7380
            TabIndex        =   26
            Top             =   2955
            Width           =   675
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Movimientos en Libros y no en Estado de Cuenta"
            Height          =   195
            Left            =   60
            TabIndex        =   23
            Top             =   465
            Width           =   3465
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Importe Estado Cuenta"
            Height          =   195
            Left            =   60
            TabIndex        =   22
            Top             =   150
            Width           =   1620
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   5850
         Left            =   45
         TabIndex        =   16
         Top             =   375
         Width           =   11805
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   5565
            Left            =   15
            TabIndex        =   18
            Top             =   120
            Width           =   11775
            _cx             =   20770
            _cy             =   9816
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
            BackColorSel    =   64
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
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmCtaCteBancos.frx":040C
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
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6360
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCtaCteBancos.frx":05A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCtaCteBancos.frx":0AE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCtaCteBancos.frx":0E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCtaCteBancos.frx":0FD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCtaCteBancos.frx":1362
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCtaCteBancos.frx":14E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCtaCteBancos.frx":193A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCtaCteBancos.frx":1A52
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCtaCteBancos.frx":1F96
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCtaCteBancos.frx":24DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCtaCteBancos.frx":25EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCtaCteBancos.frx":2702
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCtaCteBancos.frx":2B56
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCtaCteBancos.frx":2CC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCtaCteBancos.frx":320A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCtaCteBancos.frx":359C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   3
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
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Conciliar "
            ImageIndex      =   16
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar documentos"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a Excel"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   1125
      Left            =   15
      TabIndex        =   4
      Top             =   285
      Width           =   11865
      Begin VB.CommandButton CmdBusCliPro 
         Height          =   240
         Left            =   3195
         Picture         =   "FrmCtaCteBancos.frx":392E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   345
         Width           =   240
      End
      Begin VB.TextBox TxtCuenta 
         Height          =   300
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "TxtCuenta"
         Top             =   315
         Width           =   2430
      End
      Begin VB.Frame Frame3 
         Caption         =   "[  Ordenado Por  ]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   855
         Left            =   9930
         TabIndex        =   5
         Top             =   165
         Width           =   1890
         Begin VB.OptionButton OptSel2 
            Caption         =   "Fch. Emision"
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
            Left            =   195
            TabIndex        =   7
            Top             =   525
            Width           =   1560
         End
         Begin VB.OptionButton OptSel1 
            Caption         =   "Nº Registro"
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
            Left            =   195
            TabIndex        =   6
            Top             =   270
            Width           =   1560
         End
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   315
         Left            =   1035
         TabIndex        =   1
         Top             =   645
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
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
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
         Height          =   315
         Left            =   3510
         TabIndex        =   2
         Top             =   645
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
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
      Begin VB.Label LblIdCuentaContable 
         AutoSize        =   -1  'True
         Caption         =   "LblIdCuentaContable"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2520
         TabIndex        =   31
         Top             =   120
         Visible         =   0   'False
         Width           =   1485
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
         Left            =   6390
         TabIndex        =   14
         Top             =   315
         Width           =   1365
      End
      Begin VB.Label LblBanco 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblBanco"
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
         Left            =   3510
         TabIndex        =   13
         Top             =   315
         Width           =   2835
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Final"
         Height          =   195
         Index           =   2
         Left            =   2535
         TabIndex        =   12
         Top             =   690
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Cuenta"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   11
         Top             =   345
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Inicio"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   10
         Top             =   690
         Width           =   735
      End
      Begin VB.Label LblIdCuenta 
         AutoSize        =   -1  'True
         Caption         =   "LblIdCuenta"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   4575
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmCtaCteBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim xIdMon As Integer      'almacena el codigo de la moneda que se carge en funcion a la cuenta de banco
Dim xFchIni As String

Sub Cargar()
    If NulosC(TxtCuenta.Text) = "" Then
        MsgBox "No ha seleccionado el numero de la cuenta de bancos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtCuenta.SetFocus
        Exit Sub
    End If
    
    If NulosC(TxtFchIni.Valor) = "" Then
        MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    If NulosC(TxtFchIni.Valor) = "" Then
        MsgBox "No ha especificado la fecha de final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchFin.SetFocus
        Exit Sub
    End If
    
    If CDate(TxtFchIni.Valor) > TxtFchFin.Valor Then
        MsgBox "El rango de fechas comprendido no es valido", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    Dim rst As New ADODB.Recordset
    Dim Rst2 As New ADODB.Recordset
    Dim A, B As Integer
    
    Fg1.Rows = 1
    
    RST_Busq rst, "SELECT DISTINCT con_cajabanco.id, con_cajabanco.conciliado, Format([con_diario]![idmes],'00')+Format([mae_libros]![codsun],'00')+[numasi] AS numasi2, " _
        & " con_cajabanco.fchope, mae_doccajaban.descripcion, con_cajabanco.numdoc, 0 AS debe, IIf([con_bancocuenta]![idmon]=1,IIf([con_cajabanco]![idmon]=1, " _
        & " [con_diario]![imphabsol],[con_diario]![imphabdol]*[con_tc]![impven]),IIf([con_cajabanco]![idmon]=2,[con_diario]![imphabdol],[con_diario]![imphabsol]/[con_tc]![impven])) AS haber, " _
        & " con_cajabanco.tipmov FROM mae_libros INNER JOIN ((((con_cajabanco LEFT JOIN mae_doccajaban ON con_cajabanco.iddoc = mae_doccajaban.id) RIGHT JOIN con_diario " _
        & " ON con_cajabanco.id = con_diario.idmov) LEFT JOIN con_tc ON (con_cajabanco.idmon = con_tc.idmon) AND (con_cajabanco.fchope = con_tc.fecha)) LEFT JOIN con_bancocuenta " _
        & " ON con_diario.idcue = con_bancocuenta.idcuen) ON mae_libros.id = con_diario.idlib WHERE (((IIf([con_bancocuenta]![idmon]=1," _
        & " IIf([con_cajabanco]![idmon]=1,[con_diario]![imphabsol],[con_diario]![imphabdol]*[con_tc]![impven]),IIf([con_cajabanco]![idmon]=2,[con_diario]![imphabdol]," _
        & " [con_diario]![imphabsol]/[con_tc]![impven])))<>0) AND ((con_diario.idlib)=6) AND ((con_cajabanco.fchreg)>=CDate('" & TxtFchIni.Valor & "') " _
        & " And (con_cajabanco.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((con_diario.idcue)=" & NulosN(LblIdCuentaContable.Caption) & ")) " _
        & " Union " _
        & " SELECT DISTINCT con_cajabanco.id, con_cajabanco.conciliado, Format([con_diario]![idmes],'00')+Format([mae_libros]![codsun],'00')+[numasi] AS numasi2, " _
        & " con_cajabanco.fchope, mae_doccajaban.descripcion, con_cajabanco.numdoc, IIf([con_bancocuenta]![idmon]=1,IIf([con_cajabanco]![idmon]=1,[con_diario]![impdebsol]," _
        & " [con_diario]![impdebdol]*[con_tc]![impven]),IIf([con_cajabanco]![idmon]=2,[con_diario]![impdebdol],[con_diario]![impdebsol]/[con_tc]![impven])) AS debe, " _
        & " 0 AS haber, con_cajabanco.tipmov FROM ((((con_cajabanco LEFT JOIN mae_doccajaban ON con_cajabanco.iddoc = mae_doccajaban.id) RIGHT JOIN con_diario " _
        & " ON con_cajabanco.id = con_diario.idmov) LEFT JOIN con_bancocuenta ON con_diario.idcue = con_bancocuenta.idcuen) LEFT JOIN con_tc " _
        & " ON (con_cajabanco.idmon = con_tc.idmon) AND (con_cajabanco.fchope = con_tc.fecha)) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id " _
        & " WHERE (((IIf([con_bancocuenta]![idmon]=1,IIf([con_cajabanco]![idmon]=1,[con_diario]![impdebsol],[con_diario]![impdebdol]*[con_tc]![impven]), " _
        & " IIf([con_cajabanco]![idmon]=2,[con_diario]![impdebdol],[con_diario]![impdebsol]/[con_tc]![impven])))<>0) AND ((con_diario.idlib)=6) AND " _
        & " ((con_cajabanco.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (con_cajabanco.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((con_diario.idcue)=" & NulosN(LblIdCuentaContable.Caption) & "))", xCon

    If OptSel1.Value = True Then
        rst.Sort = "numasi2"
    Else
        rst.Sort = "fchope"
    End If
    
    Dim Saldo, xTotDeb, xTotHab As Double
    Dim xTotalDeb, xTotalHab As Double
    Dim xFila As Integer
    
    If CDate(TxtFchIni.Valor) <= CDate("01/01/" + Trim(AnoTra)) Then
        RST_Busq Rst2, "SELECT con_diario.idlib, con_diario.idcue, IIf([con_diario]![impdebdol]<>0,[con_diario]![impdebdol]*[con_tc]![impven],[impdebsol]) AS impdebsol1," _
            & " IIf([con_diario]![imphabdol]<>0,[con_diario]![imphabdol]*[con_tc]![impven],[imphabsol]) AS imphabsol1, con_diario.impdebdol, con_diario.imphabdol, " _
            & " con_diario.idmes FROM con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha WHERE (((con_diario.idlib)=3) " _
            & " AND ((con_diario.idcue)=" & Val(LblIdCuentaContable.Caption) & ") AND ((con_diario.idmes)=0))", xCon
        
        If Rst2.RecordCount <> 0 Then
            If NulosN(Rst2("impdebsol1")) <> 0 Then
                If xIdMon = 1 Then
                    Saldo = NulosN(Rst2("impdebsol1"))
                    xTotalDeb = NulosN(Rst2("impdebsol1"))
                Else
                    Saldo = NulosN(Rst2("impdebdol"))
                    xTotalDeb = NulosN(Rst2("impdebdol"))
                End If
            Else
                If xIdMon = 1 Then
                    Saldo = NulosN(Rst2("imphabsol1"))
                    xTotalHab = NulosN(Rst2("imphabsol1"))
                Else
                    Saldo = NulosN(Rst2("imphabdol"))
                    xTotalHab = NulosN(Rst2("imphabdol"))
                End If
            End If
        End If
    Else
        RST_Busq Rst2, "SELECT con_diario.idlib, con_diario.idcue, IIf([con_diario]![impdebdol]<>0,[con_diario]![impdebdol]*[con_tc]![impven],[impdebsol]) AS impdebsol1, " _
            & " IIf([con_diario]![imphabdol]<>0,[con_diario]![imphabdol]*[con_tc]![impven],[imphabsol]) AS imphabsol1, con_diario.impdebdol, con_diario.imphabdol, " _
            & " con_diario.fchasi, con_diario.idmov FROM con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha WHERE (((con_diario.idlib)=3) " _
            & " AND ((con_diario.idcue)=" & NulosN(LblIdCuentaContable.Caption) & ") AND ((con_diario.fchasi) Is Null))" _
            & " Union " _
            & " SELECT con_diario.idlib, con_diario.idcue, IIf([con_diario]![impdebdol]<>0,[con_diario]![impdebdol]*[con_tc]![impven],[impdebsol]) AS impdebsol1, " _
            & " IIf([con_diario]![imphabdol]<>0,[con_diario]![imphabdol]*[con_tc]![impven],[imphabsol]) AS imphabsol1, con_diario.impdebdol, con_diario.imphabdol, " _
            & " con_diario.fchasi, con_diario.idmov FROM con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha WHERE (((con_diario.idlib)=6) " _
            & " AND ((con_diario.idcue)=" & NulosN(LblIdCuentaContable.Caption) & ") AND ((con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "')))", xCon

        Saldo = 0
        
        If Rst2.RecordCount <> 0 Then
            Rst2.MoveFirst
            For A = 1 To Rst2.RecordCount
            
                If xIdMon = 1 Then
                    xTotalDeb = xTotalDeb + NulosN(Rst2("impdebsol1"))
                    xTotalHab = xTotalHab + NulosN(Rst2("imphabsol1"))
                Else
                    xTotalDeb = xTotalDeb + NulosN(Rst2("impdebdol"))
                    xTotalHab = xTotalHab + NulosN(Rst2("imphabdol"))
                End If
                Rst2.MoveNext
                If Rst2.EOF = True Then Exit For
            Next A
        End If
    End If
    
    Saldo = xTotalDeb - xTotalHab
    Fg1.Rows = Fg1.Rows + 1
    Fg1.TextMatrix(Fg1.Rows - 1, 6) = "Saldo Anterior ==>"
    Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(Saldo, "0.00")
    Fg1.Rows = Fg1.Rows + 1
    
    If rst.RecordCount = 0 Then
        MsgBox "La cuenta de banco seleccionada no tiene movimientos", vbInformation + vbOKCancel + vbDefaultButton1, xTitulo
        Set rst = Nothing
        Exit Sub
    Else
        rst.MoveFirst
        xFila = 3
        For A = 1 To rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(xFila, 1) = rst("numasi2") & ""
            Fg1.TextMatrix(xFila, 2) = Format(rst("fchope"), "dd/mm/yy")
            Fg1.TextMatrix(xFila, 3) = rst("descripcion") & ""
            Fg1.TextMatrix(xFila, 4) = rst("numdoc") & ""
            Fg1.TextMatrix(xFila, 7) = Format(NulosN(rst("debe")), FORMAT_MONTO)
            Fg1.TextMatrix(xFila, 8) = Format(NulosN(rst("haber")), FORMAT_MONTO)
            xTotDeb = xTotDeb + NulosN(rst("debe"))
            xTotHab = xTotHab + NulosN(rst("haber"))
            If rst("debe") <> 0 Then
                Saldo = Saldo + NulosN(rst("debe"))
            Else
                Saldo = Saldo - NulosN(rst("haber"))
            End If
            
            Fg1.TextMatrix(xFila, 9) = Format(Saldo, FORMAT_MONTO)
            If rst("conciliado") = True Then
                Fg1.TextMatrix(xFila, 10) = -1
            Else
                Fg1.TextMatrix(xFila, 10) = 0
            End If
            Fg1.TextMatrix(xFila, 11) = rst("id")
            rst.MoveNext
            If rst.EOF = True Then
                Exit For
            End If
            xFila = xFila + 1
        Next A
        
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 6) = "Total del Periodo ==>"
        Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(xTotDeb, FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(xTotHab, FORMAT_MONTO)
    End If
    
    'CARGAMOS LOS CHEQUES GIRADOS PERO NO COBRADOS
    Set rst = Nothing
    RST_Busq rst, "SELECT DISTINCT con_cajabanco.id, con_cajabanco.conciliado, con_diario.numasi, con_cajabanco.fchope, mae_doccajaban.descripcion, con_cajabanco.numdoc, 0 AS debe, " _
        & " con_cajabanco.importe AS haber, con_cajabanco.tipmov FROM (con_cajabanco LEFT JOIN mae_doccajaban ON con_cajabanco.iddoc = mae_doccajaban.id) " _
        & " LEFT JOIN con_diario ON con_cajabanco.id = con_diario.idmov WHERE (((con_cajabanco.tipmov)=2) AND ((con_diario.idlib)=6) AND ((con_cajabanco.tipope)=2) " _
        & " AND ((con_cajabanco.idcueban)=" & NulosN(LblIdCuenta.Caption) & ") AND ((con_cajabanco.fchreg)<CDate('" & TxtFchIni.Valor & "')) AND ((con_cajabanco.conciliado)=0))", xCon

    Fg5.Rows = 1
    
    xTotDeb = 0:    xTotHab = 0
    For A = 1 To rst.RecordCount
        Fg5.Rows = Fg5.Rows + 1
        Fg5.TextMatrix(A, 1) = rst("numasi") & ""
        Fg5.TextMatrix(A, 2) = rst("descripcion") & ""
        Fg5.TextMatrix(A, 3) = Format(rst("fchope"), "dd/mm/yy")
        Fg5.TextMatrix(A, 4) = rst("numdoc") & ""
        Fg5.TextMatrix(A, 5) = Format(NulosN(rst("debe")), FORMAT_MONTO)
        Fg5.TextMatrix(A, 6) = Format(NulosN(rst("haber")), FORMAT_MONTO)
        xTotDeb = xTotDeb + NulosN(rst("debe"))
        xTotHab = xTotHab + NulosN(rst("haber"))
        If rst("debe") <> 0 Then
            Saldo = Saldo + NulosN(rst("debe"))
        Else
            Saldo = Saldo - NulosN(rst("haber"))
        End If
        
        Fg5.TextMatrix(A, 8) = rst("id")
        rst.MoveNext
        If rst.EOF = True Then
            Exit For
        End If
    Next A
    TxtTotDeb5.Text = Format(xTotDeb, FORMAT_MONTO)
    TxtTotHab5.Text = Format(xTotHab, FORMAT_MONTO)
    
    TabOne1.CurrTab = 0
    TabOne2.CurrTab = 0
    
    Set rst = Nothing
    'actualizamos los cheques girados y no cobrados
    xFchIni = "01/" + Format(CDate(TxtFchIni.Valor), "MM") + "/" + AnoTra
    RST_Busq rst, "SELECT * FROM con_cajabancocobrados WHERE fchreg = cdate('" & xFchIni & "') and idcueban = " & NulosN(LblIdCuenta.Caption) & "", xCon
    
    If rst.RecordCount <> 0 Then
        TxtImpEstado.Text = Format(rst("salban"), "0.00")
        rst.MoveFirst
        For A = 1 To rst.RecordCount
            For B = 1 To Fg5.Rows - 1
                If rst("idcajban") = NulosN(Fg5.TextMatrix(B, 8)) Then
                    Fg5.TextMatrix(B, 7) = -1
                    Exit For
                End If
            Next B
            rst.MoveNext
            If rst.EOF = True Then Exit For
        Next A
        HallarTotalNoCobrados
    End If
    
    'borramos los cheques que fueron cobrados en otros mes
    Set rst = Nothing
    RST_Busq rst, "SELECT * From con_cajabancocobrados WHERE (((con_cajabancocobrados.fchreg)<CDate('" & xFchIni & "')) AND ((con_cajabancocobrados.idcueban)=" & NulosN(LblIdCuenta.Caption) & "))", xCon
    If rst.RecordCount <> 0 Then
        rst.MoveFirst
        For A = 1 To rst.RecordCount
            For B = 1 To Fg5.Rows - 1
                If rst("idcajban") = NulosN(Fg5.TextMatrix(B, 8)) Then
                    Fg5.RemoveItem B
                    Exit For
                End If
            Next B
            
            rst.MoveNext
            If rst.EOF = True Then Exit For
        Next A
        HallarTotalNoCobrados
    End If
End Sub



Private Sub CmdAgregaMov_Click()
    If NulosC(Fg3.TextMatrix(Fg3.Rows - 1, 1)) = "" Then
        Exit Sub
    End If
    Fg3.Rows = Fg3.Rows + 1
End Sub

Private Sub CmdBusCliPro_Click()
    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Banco":           xCampos(0, 1) = "desban":        xCampos(0, 2) = "3600":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº Cuenta":       xCampos(1, 1) = "numcue":        xCampos(1, 2) = "1600":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Moneda":          xCampos(2, 1) = "desmon":        xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Nº Cta Contable": xCampos(3, 1) = "cuenta":        xCampos(3, 2) = "1500":         xCampos(3, 3) = "C"
    
    xForm.SQLCad = "SELECT mae_bancos.descripcion AS desban, con_bancocuenta.*, mae_moneda.descripcion AS desmon, " _
        & " con_planctas.cuenta FROM mae_bancos INNER JOIN (con_planctas RIGHT JOIN (con_bancocuenta LEFT JOIN mae_moneda " _
        & " ON con_bancocuenta.idmon = mae_moneda.id) ON con_planctas.id = con_bancocuenta.idcuen) ON " _
        & " mae_bancos.id = con_bancocuenta.idban " _
        & " ORDER BY mae_bancos.descripcion"
    
    xForm.Titulo = "Buscando Cuentas de Banco"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "desban"
    xForm.CampoBusca = "desban"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtCuenta.Text = xRs("numcue")
        LblIdCuenta.Caption = xRs("id")
        LblBanco.Caption = Trim(xRs("desban")) '"   Cuenta Nº " & xRs("numcue")
        xIdMon = xRs("idmon")
        LblMoneda.Caption = xRs("desmon")
        LblIdCuentaContable.Caption = xRs("idcuen")
        TxtFchIni.SetFocus
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Sub SumarSaldos()
'    Fg2.TextMatrix(3, 6) = Val(Fg1.TextMatrix(Fg1.Rows - 2, 9)) - Val(TxtImpEstado.Text)
'    Fg2.TextMatrix(3, 6) = Format(Fg2.TextMatrix(3, 6), "0.00")
End Sub


'Sub aaaa()
'    Fg2.Rows = 1
'
'    Fg2.Rows = Fg2.Rows + 1
'    Fg2.TextMatrix(Fg2.Rows - 1, 2) = "Saldo segun Libro Al "
'    Fg2.TextMatrix(Fg2.Rows - 1, 3) = "31/01/01"
'    Fg2.TextMatrix(Fg2.Rows - 1, 6) = Fg1.TextMatrix(Fg1.Rows - 2, 9)
'
'    Fg2.Rows = Fg2.Rows + 1
'    Fg2.TextMatrix(Fg2.Rows - 1, 2) = "Saldo segun Estado de Cuenta Al"
'    Fg1.TextMatrix(Fg2.Rows - 1, 3) = "31/01/01"
'
'    If NulosN(TxtImpEstado.Text) = 0 Then
'        Fg2.TextMatrix(Fg2.Rows - 1, 6) = "0.00"
'    Else
'        Fg2.TextMatrix(Fg2.Rows - 1, 6) = TxtImpEstado.Text
'    End If
'
'    Fg2.Rows = Fg2.Rows + 1
'    Fg2.TextMatrix(Fg2.Rows - 1, 2) = "Importe a Conciliar"
'    Fg2.TextMatrix(Fg2.Rows - 1, 6) = Val(Fg1.TextMatrix(Fg1.Rows - 2, 9)) - Val(TxtImpEstado.Text)
'    Fg2.TextMatrix(Fg2.Rows - 1, 6) = Format(Fg2.TextMatrix(Fg2.Rows - 1, 6), "0.00")
'
'    Fg2.Rows = Fg2.Rows + 2
'    Fg2.TextMatrix(Fg2.Rows - 1, 2) = "ANALISIS DE LA DIFERENCIA"
'
'    Fg2.Rows = Fg2.Rows + 1
'    Fg2.TextMatrix(Fg2.Rows - 1, 2) = "Documentos Pendientes (Cargos)"
'    Fg2.TextMatrix(Fg2.Rows - 1, 6) = ""
'    xFilaActual = Fg2.Rows - 1
'End Sub

Sub Procesar()
    Dim xFilaActual As Integer
    Dim A As Integer
    Dim xTotal As Double
    
    Fg2.Rows = 1
    For A = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(A, 10)) <> -1 And NulosC(Fg1.TextMatrix(A, 1)) <> "" Then
            Fg2.Rows = Fg2.Rows + 1
            
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = Fg1.TextMatrix(A, 1)
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = ""
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = Fg1.TextMatrix(A, 2)
            Fg2.TextMatrix(Fg2.Rows - 1, 4) = Fg1.TextMatrix(A, 4)
            Fg2.TextMatrix(Fg2.Rows - 1, 5) = Fg1.TextMatrix(A, 7)
            Fg2.TextMatrix(Fg2.Rows - 1, 6) = Fg1.TextMatrix(A, 8)
        End If
    Next A
    
    Fg2.Visible = True
    HallarTotalFg2
    
    Dim rst As New ADODB.Recordset
    
    Fg3.Rows = 1
    RST_Busq rst, "SELECT con_cajabancocon.*, mae_bancoscarabo.descripcion AS descarabo " _
        & " FROM con_cajabancocon LEFT JOIN mae_bancoscarabo ON con_cajabancocon.idcon = mae_bancoscarabo.id ORDER BY fchope", xCon
    
    If rst.RecordCount <> 0 Then
        rst.MoveFirst
        TxtImpEstado.Text = Format(rst("impestcue"), "0.00")
        For A = 1 To rst.RecordCount
            Fg3.Rows = Fg3.Rows + 1
            Fg3.TextMatrix(A, 1) = rst("descarabo")
            Fg3.TextMatrix(A, 2) = rst("fchope")
            Fg3.TextMatrix(A, 3) = Format(rst("impdeb"), "0.00")
            Fg3.TextMatrix(A, 4) = Format(rst("imphab"), "0.00")
            Fg3.TextMatrix(A, 5) = rst("idcon")
            rst.MoveNext
            If rst.EOF = True Then Exit For
        Next A
    End If
    Set rst = Nothing
End Sub

Private Sub CmdCerrar_Click()
    Toolbar1.Enabled = True
    Frame1.Enabled = True
    TabOne1.Enabled = True
    
    Frame6.Visible = False
End Sub

Private Sub CmdDelMov_Click()
    If Fg3.Rows = 1 Then Exit Sub
    Fg3.RemoveItem Fg3.Row
End Sub

Private Sub CmdProcesar_Click()
    Frame6.Left = 1778
    Frame6.Top = 2085
    
    Toolbar1.Enabled = False
    Frame1.Enabled = False
    TabOne1.Enabled = False
    Frame6.Visible = True
    ConciliarConciliacion
        
End Sub

Sub ConciliarConciliacion()
    Fg4.Rows = 1
    Dim xDiferencia As Double

    Fg4.Rows = Fg4.Rows + 1
    Fg4.TextMatrix(Fg4.Rows - 1, 1) = "Saldo segun Libro Al " + Trim(TxtFchFin.Valor)
    Fg4.TextMatrix(Fg4.Rows - 1, 2) = NulosN(Fg1.TextMatrix(Fg1.Rows - 2, 9))

    Fg4.Rows = Fg4.Rows + 1
    Fg4.TextMatrix(Fg4.Rows - 1, 1) = "Movimientos en Libros y no en Estados de Cuenta"
    'If NulosN(TxtTot1.Text) > NulosN(TxtTot2.Text) Then
    
    '    Fg4.TextMatrix(Fg4.Rows - 1, 2) = NulosN(TxtTot1.Text) - NulosN(TxtTot2.Text)
    'Else
        Fg4.TextMatrix(Fg4.Rows - 1, 2) = NulosN(TxtTot2.Text) - NulosN(TxtTot1.Text)
    'End If
    Fg4.TextMatrix(Fg4.Rows - 1, 2) = Format(Fg4.TextMatrix(Fg4.Rows - 1, 2), "0.00")

    Fg4.Rows = Fg4.Rows + 1
    Fg4.TextMatrix(Fg4.Rows - 1, 1) = "Cheques Girados y No Cobrados"
    Fg4.TextMatrix(Fg4.Rows - 1, 2) = Format(TxtTotHab5.Text, "0.00")
    
    
    Fg4.Rows = Fg4.Rows + 1
    Fg4.TextMatrix(Fg4.Rows - 1, 1) = "Otros Gastos No Cosiderados"
    Fg4.TextMatrix(Fg4.Rows - 1, 2) = Format(TxtTotHab.Text, "0.00")
    
    Fg4.Rows = Fg4.Rows + 1
    Fg4.TextMatrix(Fg4.Rows - 1, 1) = "Diferencia a Conciliar"
    Fg4.TextMatrix(Fg4.Rows - 1, 2) = ((NulosN(Fg4.TextMatrix(Fg4.Rows - 4, 2)) + NulosN(Fg4.TextMatrix(Fg4.Rows - 3, 2))) + NulosN(TxtTotHab5.Text)) + NulosN(TxtTotHab.Text)
    Fg4.TextMatrix(Fg4.Rows - 1, 2) = Format(Fg4.TextMatrix(Fg4.Rows - 1, 2), "0.00")
    xDiferencia = NulosN(Fg4.TextMatrix(Fg4.Rows - 1, 2))

    Fg4.Rows = Fg4.Rows + 2
    Fg4.TextMatrix(Fg4.Rows - 1, 1) = "Saldos Segun Banco"
    Fg4.TextMatrix(Fg4.Rows - 1, 2) = TxtImpEstado.Text


    Fg4.Rows = Fg4.Rows + 1
    Fg4.TextMatrix(Fg4.Rows - 1, 1) = "Diferencia a Ubicar"
    Fg4.TextMatrix(Fg4.Rows - 1, 2) = NulosN(TxtImpEstado.Text) - xDiferencia
    Fg4.TextMatrix(Fg4.Rows - 1, 2) = Format(Fg4.TextMatrix(Fg4.Rows - 1, 2), "0.00")
   


'    Fg4.Rows = 1
'
'    Fg4.Rows = Fg4.Rows + 1
'    Fg4.TextMatrix(Fg4.Rows - 1, 1) = "Saldo segun Libro Al " + Trim(TxtFchFin.Valor)
'    Fg4.TextMatrix(Fg4.Rows - 1, 2) = Fg1.TextMatrix(Fg1.Rows - 2, 9)
'
'    Fg4.Rows = Fg4.Rows + 1
'    Fg4.TextMatrix(Fg4.Rows - 1, 1) = "Saldo segun Estado de Cuenta Al " + Trim(TxtFchFin.Valor)
'    If NulosN(TxtImpEstado.Text) = 0 Then
'        Fg4.TextMatrix(Fg4.Rows - 1, 2) = "0.00"
'    Else
'        Fg4.TextMatrix(Fg4.Rows - 1, 2) = TxtImpEstado.Text
'    End If
'
'    Fg4.Rows = Fg4.Rows + 1
'    Fg4.TextMatrix(Fg4.Rows - 1, 1) = "Importe a Conciliar"
'    Fg4.TextMatrix(Fg4.Rows - 1, 2) = Val(Fg1.TextMatrix(Fg1.Rows - 2, 9)) - Val(TxtImpEstado.Text)
'    Fg4.TextMatrix(Fg4.Rows - 1, 2) = Format(Fg4.TextMatrix(Fg4.Rows - 1, 2), "0.00")
'
'    'agregamos los cargos y abonos del libro que no estan en el estado de cuenta
'    Dim xTot As Double
'    Dim A As Integer
'
'    Fg4.Rows = Fg4.Rows + 1
'
'    '---------------------------------------------------------
'    'Hallamos los Cargos de libros y no en bancos
'    xTot = 0
'    For A = 1 To Fg2.Rows - 1
'        If Val(Fg2.TextMatrix(A, 5)) <> 0 Then
'            xTot = xTot + Val(Fg2.TextMatrix(A, 5))
'        End If
'    Next A
'    Fg4.Rows = Fg4.Rows + 1
'    Fg4.TextMatrix(Fg4.Rows - 1, 1) = "Cargos en Libros y no en Bancos (-)"
'    Fg4.TextMatrix(Fg4.Rows - 1, 2) = Format(xTot, "0.00")
'
'    'Hallamos los Abonos de libros y no en bancos
'    xTot = 0
'    For A = 1 To Fg2.Rows - 1
'        If Val(Fg2.TextMatrix(A, 6)) <> 0 Then
'            xTot = xTot + Val(Fg2.TextMatrix(A, 6))
'        End If
'    Next A
'    Fg4.Rows = Fg4.Rows + 1
'    Fg4.TextMatrix(Fg4.Rows - 1, 1) = "Abonos en Libros y no en Bancos (+)"
'    Fg4.TextMatrix(Fg4.Rows - 1, 2) = Format(xTot, "0.00")
'
'
'    '---------------------------------------------------------
'    'Hallamos los Cargos de libros y no en bancos
'    Fg4.Rows = Fg4.Rows + 1
'    xTot = 0
'    For A = 1 To Fg3.Rows - 1
'        If Val(Fg3.TextMatrix(A, 3)) <> 0 Then
'            xTot = xTot + Val(Fg3.TextMatrix(A, 3))
'        End If
'    Next A
'    Fg4.Rows = Fg4.Rows + 1
'    Fg4.TextMatrix(Fg4.Rows - 1, 1) = "Cargos en Bancos y no en Libros (-)"
'    Fg4.TextMatrix(Fg4.Rows - 1, 2) = Format(xTot, "0.00")
'
'    'Hallamos los Abonos de libros y no en bancos
'    xTot = 0
'    For A = 1 To Fg3.Rows - 1
'        If Val(Fg3.TextMatrix(A, 4)) <> 0 Then
'            xTot = xTot + Val(Fg3.TextMatrix(A, 4))
'        End If
'    Next A
'    Fg4.Rows = Fg4.Rows + 1
'    Fg4.TextMatrix(Fg4.Rows - 1, 1) = "Abonos en Bancos y no en Libros (+)"
'    Fg4.TextMatrix(Fg4.Rows - 1, 2) = Format(xTot, "0.00")

End Sub

Private Sub Fg1_EnterCell()
    If QueHace <> 3 Then
        If Fg1.Col = 10 Then
            If NulosC(Fg1.TextMatrix(Fg1.Row, 1)) <> "" Then
                Fg1.Editable = flexEDKbdMouse
            Else
                Fg1.TextMatrix(Fg1.Row, 10) = 0
                Fg1.Editable = flexEDNone
            End If
        Else
            Fg1.Editable = flexEDNone
        End If
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Fg3_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset

    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":       xCampos(1, 1) = "id":             xCampos(1, 2) = "1000":    xCampos(1, 3) = "C"
    
    xForm.SQLCad = "SELECT * FROM mae_bancoscarabo ORDER BY descripcion"
    
    xForm.Titulo = "Buscando Cargos y Abonos de Bancos"
    xForm.FormaBusca = CualquierParte
    xForm.Criterio = ""
    xForm.Ordenado = "descripcion"
    xForm.CampoBusca = "descripcion"
    
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            Fg3.TextMatrix(Fg3.Row, 1) = xRs("descripcion")
            Fg3.TextMatrix(Fg3.Row, 5) = xRs("id")
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub Fg3_CellChanged(ByVal Row As Long, ByVal Col As Long)
    HallarTotalFg3
End Sub

Sub HallarTotalFg3()
    Dim A As Integer
    Dim TotDeb As Double
    Dim TotHab As Double
    
    For A = 1 To Fg3.Rows - 1
        TotDeb = TotDeb + Val(Fg3.TextMatrix(A, 3))
        TotHab = TotHab + Val(Fg3.TextMatrix(A, 4))
    Next A
    
    TxtTotDeb.Text = Format(TotDeb, "0.00")
    TxtTotHab.Text = Format(TotHab, "0.00")
End Sub

Sub HallarTotalFg2()
    Dim A As Integer
    Dim TotDeb As Double
    Dim TotHab As Double
    
    For A = 1 To Fg2.Rows - 1
        TotDeb = TotDeb + NulosN(Fg2.TextMatrix(A, 5))
        TotHab = TotHab + NulosN(Fg2.TextMatrix(A, 6))
    Next A
    
    TxtTot1.Text = Format(TotDeb, FORMAT_MONTO)
    TxtTot2.Text = Format(TotHab, FORMAT_MONTO)
End Sub

Private Sub Fg3_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = 45 Then
        CmdAgregaMov_Click
    End If
    If KeyCode = 46 Then
        CmdDelMov_Click
    End If
End Sub

Private Sub Fg5_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Fg5.Col = 7 Then
        HallarTotalNoCobrados
    End If
End Sub

Sub HallarTotalNoCobrados()
    Dim A As Integer
    Dim xTotal As Double
    For A = 1 To Fg5.Rows - 1
        If NulosN(Fg5.TextMatrix(A, 7)) = 0 Then
            xTotal = xTotal + NulosN(Fg5.TextMatrix(A, 6))
        End If
    Next A
    TxtTotDeb5.Text = "0.00"
    TxtTotHab5.Text = Format(xTotal, FORMAT_MONTO)
    TxtTotHab5.Refresh
    TxtTotDeb5.Refresh
End Sub

Private Sub Fg5_EnterCell()
    If QueHace = 3 Then
        Fg5.Editable = flexEDNone
        Exit Sub
    End If
    If Fg5.Col = 7 Then
        Fg5.Editable = flexEDKbdMouse
    Else
        Fg5.Editable = flexEDNone
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        TabOne1.CurrTab = 0
        OptSel2.Value = True
        TxtCuenta.SetFocus
    End If
End Sub

Private Sub Form_Load()
    QueHace = 3
    SeEjecuto = False
    Blanquea
    Frame4.BackColor = &H8000000F
    Frame5.BackColor = &H8000000F
    Frame7.BackColor = &H8000000F
    Frame8.BackColor = &H8000000F
    Fg5.ColWidth(8) = 0
    Fg1.ColWidth(11) = 0
    Fg3.ColWidth(5) = 0
    Fg5.Rows = 1
    Fg3.Rows = 1
    Fg2.Rows = 1
    TxtImpEstado.Text = ""
'    TxtFchIni.Valor = "01/01/07"
'    TxtFchFin.Valor = "31/01/07"
    'LblIdCuenta.Caption = "1"
    'TxtCuenta.Text = "solo haz clic"
End Sub

Sub ActivarTool()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    
    Toolbar1.Buttons(3).Enabled = Not Toolbar1.Buttons(3).Enabled
    Toolbar1.Buttons(4).Enabled = Not Toolbar1.Buttons(4).Enabled
    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
    
    Toolbar1.Buttons(7).Enabled = Not Toolbar1.Buttons(7).Enabled
    Toolbar1.Buttons(8).Enabled = Not Toolbar1.Buttons(8).Enabled
    
    Toolbar1.Buttons(10).Enabled = Not Toolbar1.Buttons(10).Enabled
End Sub

Sub Conciliar()
    If TxtCuenta.Text = "" Then
        MsgBox "No ha especificado la cuenta a conciliar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtCuenta.SetFocus
        Exit Sub
    End If
    
    If TxtFchIni.Valor = "" Then
        MsgBox "No ha especificado la fecha de inicio de la consulta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    If TxtFchFin.Valor = "" Then
        MsgBox "No ha especificado la fecha final de la consulta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchFin.SetFocus
        Exit Sub
    End If
    
    If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
        MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    If Fg1.Rows = 1 Then
        MsgBox "No se ha mostrado el libro bancos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtCuenta.SetFocus
        Exit Sub
    End If

    QueHace = 2
    
    Frame1.Enabled = False
    
    Fg3.ColComboList(1) = "|..."
    Fg3.Rows = 1
    Fg3.Editable = flexEDKbdMouse
    Fg3.SelectionMode = flexSelectionFree
    
    Fg5.Editable = flexEDKbdMouse
    Fg5.SelectionMode = flexSelectionFree
    ActivarTool
    
    CmdAgregaMov.Enabled = Not CmdAgregaMov.Enabled
    CmdDelMov.Enabled = Not CmdDelMov.Enabled

    Bloquea
End Sub

Sub Bloquea()
    CmdBusCliPro.Enabled = Not CmdBusCliPro.Enabled
    TxtFchIni.Locked = Not TxtFchIni.Locked
    TxtFchFin.Locked = Not TxtFchFin.Locked
    TxtImpEstado.Locked = Not TxtImpEstado.Locked
End Sub

Sub Cancelar()
    Bloquea
    ActivarTool
    'Bloquea
    QueHace = 3
    Frame1.Enabled = True
    Fg3.Editable = flexEDNone
    Fg3.SelectionMode = flexSelectionByRow
    TabOne1.CurrTab = 0
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If Fg1.Rows = 1 Then
            MsgBox "No se ha mostrado el libro bancos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Cancel = 1: Exit Sub
        End If
        
        'TxtImpEstado.Text = ""
        Procesar
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Conciliar
        
    If Button.Index = 3 Then Cargar
    
    If Button.Index = 4 Then Imprimir
    
    If Button.Index = 5 Then EXPORTAR
    
    If Button.Index = 7 Then Cancelar
        
    If Button.Index = 8 Then
        If Grabar = True Then
            Cancelar
            Cargar1
        End If
    End If
    
    If Button.Index = 10 Then
        Unload Me
    End If
End Sub

Function Grabar() As Boolean
    If Fg1.Rows = 1 Then
        MsgBox "No ha habido movimientos para esta cuenta en el periodo especificado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Grabar = False
        Exit Function
    End If
   
    TabOne1.CurrTab = 1
    
    If NulosN(TxtImpEstado.Text) = 0 Then
        MsgBox "No ha especificado el importe del estado de cuenta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtImpEstado.SetFocus
        Grabar = False
        Exit Function
    End If
    
    Dim A As Integer
    Dim RstCon As New ADODB.Recordset
    Dim RstConCheCob As New ADODB.Recordset
    
    For A = 1 To Fg1.Rows - 1
        If NulosC(Fg1.TextMatrix(A, 1)) <> "" Then
            If NulosN(Fg1.TextMatrix(A, 10)) = -1 Then
                xCon.Execute "UPDATE con_cajabanco SET con_cajabanco.conciliado = -1 WHERE (((con_cajabanco.id)=" & Val(Fg1.TextMatrix(A, 11)) & "))"
            Else
                xCon.Execute "UPDATE con_cajabanco SET con_cajabanco.conciliado = 0 WHERE (((con_cajabanco.id)=" & Val(Fg1.TextMatrix(A, 11)) & "))"
            End If
        End If
    Next A
    
    xCon.Execute "DELETE * FROM con_cajabancocon WHERE idmes = 1"
    RST_Busq RstCon, "SELECT * FROM con_cajabancocon", xCon
    
    If Fg3.Rows <> 1 Then
        For A = 1 To Fg3.Rows - 1
            RstCon.AddNew
            RstCon("idmes") = Val(Format(CDate(TxtFchIni.Valor), "MM"))
            RstCon("idcon") = NulosN(Fg3.TextMatrix(A, 5))
            RstCon("fchope") = CDate(Fg3.TextMatrix(A, 2))
            RstCon("impdeb") = NulosN(Fg3.TextMatrix(A, 3))
            RstCon("imphab") = NulosN(Fg3.TextMatrix(A, 4))
            RstCon("impestcue") = NulosN(TxtImpEstado.Text)
            RstCon.Update
        Next A
    End If
    
    'grabamos los cheques girados en el periodo anterior y recien estennsiendo cobrados
    If Fg5.Rows <> 1 Then
        xFchIni = "01/" + Format(CDate(TxtFchIni.Valor), "MM") + "/" + AnoTra
        xCon.Execute "DELETE * FROM con_cajabancocobrados WHERE fchreg =cdate('" & xFchIni & "') and idcueban = " & NulosN(LblIdCuenta.Caption) & ""
        
        RST_Busq RstConCheCob, "SELECT * FROM con_cajabancocobrados", xCon
        For A = 1 To Fg5.Rows - 1
            If NulosN(Fg5.TextMatrix(A, 7)) = -1 Then
                RstConCheCob.AddNew
                RstConCheCob("idcajban") = Fg5.TextMatrix(A, 8)
                RstConCheCob("fchreg") = xFchIni
                RstConCheCob("idcueban") = Val(LblIdCuenta.Caption)
                RstConCheCob("salban") = NulosN(TxtImpEstado.Text)
                RstConCheCob.Update
            End If
        Next A
    End If
    
    MsgBox "La conciliacion se guardo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Set RstCon = Nothing
    Exit Function
End Function

Sub Blanquea()
    TxtCuenta.Text = ""
    LblIdCuenta.Caption = ""
    LblMoneda.Caption = ""
    LblBanco.Caption = ""
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    
    TxtTotDeb.Text = ""
    TxtTotHab.Text = ""
    TxtTot1.Text = ""
    TxtTot2.Text = ""
End Sub

Private Sub TxtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = vbKeyF5 Then CmdBusCliPro_Click
End Sub

Private Sub TxtImpEstado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtImpEstado_Validate(Cancel As Boolean)
    If NulosN(TxtImpEstado.Text) <> 0 Then
        TxtImpEstado.Text = Format(TxtImpEstado.Text, "0.00")
    End If
End Sub

Private Sub VSFlexGrid1_Click()

End Sub

'--------
Private Sub EXPORTAR()
    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios
    Dim nTitulo1 As String
    Dim nTitulo As String
    Dim nPeriodo As String
    
    
    nTitulo = "Libro Bancos"
    
    nTitulo1 = "Nº Cuenta: " + TxtCuenta.Text + " " + LblBanco.Caption + " " + LblMoneda.Caption
    
    If CDate(TxtFchIni.Valor) <> CDate(TxtFchFin.Valor) Then
        nPeriodo = "Del" + Format(TxtFchIni.Valor, "dd/mm/yy") + " Al " + Format(TxtFchIni.Valor, "dd/mm/yy")
    Else
        nPeriodo = "Al " + Format(TxtFchIni.Valor, "dd/mm/yy")
    End If

    Me.MousePointer = vbHourglass
    
    X_PRINT.VSFlexGrid_Exportar_MSExcel xCon, Fg1, nTitulo, nPeriodo, nTitulo1, nTitulo
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Imprimir"
End Sub

Private Sub Imprimir()
    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios
    Dim nTitulo1 As String
    Dim nTitulo As String
    Dim nPeriodo As String
    
    Me.MousePointer = vbHourglass
    
    nTitulo = "Libro Bancos"
    
    nTitulo1 = "Nº Cuenta: " + TxtCuenta.Text + " " + LblBanco.Caption + " " + LblMoneda.Caption
    
    If CDate(TxtFchIni.Valor) <> CDate(TxtFchFin.Valor) Then
        nPeriodo = "Del" + Format(TxtFchIni.Valor, "dd/mm/yy") + " Al " + Format(TxtFchIni.Valor, "dd/mm/yy")
    Else
        nPeriodo = "Al " + Format(TxtFchIni.Valor, "dd/mm/yy")
    End If
    
    X_PRINT.Imprimir_x_VSFlexGrid Fg1, nTitulo, nTitulo1, nPeriodo, True, True
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR
End Sub



Sub Cargar1()
    '--corregido por Johan Castro
    '-- fecha 27/02/09
    If NulosC(TxtCuenta.Text) = "" Then
        MsgBox "No ha seleccionado el numero de la cuenta de bancos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtCuenta.SetFocus
        Exit Sub
    End If
    
    If NulosC(TxtFchIni.Valor) = "" Then
        MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    If NulosC(TxtFchIni.Valor) = "" Then
        MsgBox "No ha especificado la fecha de final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchFin.SetFocus
        Exit Sub
    End If
    
    If CDate(TxtFchIni.Valor) > TxtFchFin.Valor Then
        MsgBox "El rango de fechas comprendido no es valido", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    Dim rst As New ADODB.Recordset
    Dim Rst2 As New ADODB.Recordset
    Dim A&, B&
    Dim nSQL As String
    
    Fg1.Rows = 1
    
    DoEvents
    nSQL = "SELECT DISTINCT tes_caja.id, tes_caja.conciliado, Format(con_diario.idmes,'00')+Format(mae_libros.codsun,'00')+numasi AS numasi2, tes_caja.fchope, tes_documentos.descripcion, tes_documentos.abrev, tes_cajaorigendet.numdoc, " _
        + vbCr + " IIf(con_bancocuenta.idmon=1,IIf(tes_caja.idmon=1,con_diario.impdebsol,con_diario.impdebdol*con_tc.impven),IIf(tes_caja.idmon=2,con_diario.impdebdol,con_diario.impdebsol/con_tc.impven)) AS debe, " _
        + vbCr + " IIf(con_bancocuenta.idmon=1,IIf(tes_caja.idmon=1,con_diario.imphabsol,con_diario.imphabdol*con_tc.impven),IIf(tes_caja.idmon=2,con_diario.imphabdol,con_diario.imphabsol/con_tc.impven)) AS haber, tes_caja.tipmov, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc " _
        + vbCr + " FROM ((((mae_libros RIGHT JOIN (con_diario INNER JOIN con_bancocuenta ON con_diario.idcue = con_bancocuenta.idcuen) ON mae_libros.id = con_diario.idlib) LEFT JOIN (tes_cajaorigendet LEFT JOIN tes_caja ON tes_cajaorigendet.idtes = tes_caja.id) ON (con_diario.idmov = tes_cajaorigendet.idtes) AND (con_diario.idorides = tes_cajaorigendet.idori)) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id) LEFT JOIN con_planctas ON con_diario.idcue = con_planctas.id " _
        + vbCr + " WHERE (((con_diario.idlib)=6) AND ((tes_caja.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (tes_caja.fchreg)<=CDate('" & TxtFchFin.Valor & "'))) and con_diario.idcue = " & NulosN(LblIdCuentaContable.Caption)

    RST_Busq rst, nSQL, xCon
    
    If OptSel1.Value = True Then
        rst.Sort = "numasi2"
    Else
        rst.Sort = "fchope"
    End If
    
    
    '*****************************************************************************************************************
    Dim Saldo, xTotDeb, xTotHab As Double
    Dim xTotalDeb, xTotalHab As Double
    Dim xFila As Integer
    
    '--cargando los saldos iniciales
    If CDate(TxtFchIni.Valor) <= CDate("01/01/" + Trim(AnoTra)) Then
        '--cargando las aperturas
        RST_Busq Rst2, "SELECT con_diario.idlib, con_diario.idcue, IIf([con_diario]![impdebdol]<>0,[con_diario]![impdebdol]*[con_tc]![impven],[impdebsol]) AS impdebsol1," _
            & " IIf([con_diario]![imphabdol]<>0,[con_diario]![imphabdol]*[con_tc]![impven],[imphabsol]) AS imphabsol1, con_diario.impdebdol, con_diario.imphabdol, " _
            & " con_diario.idmes FROM con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha WHERE (((con_diario.idlib)=3) " _
            & " AND ((con_diario.idcue)=" & NulosN(LblIdCuentaContable.Caption) & ") AND ((con_diario.idmes)=0))", xCon
        
         
        If Rst2.RecordCount <> 0 Then
            If NulosN(Rst2("impdebsol1")) <> 0 Then
                If xIdMon = 1 Then
                    Saldo = NulosN(Rst2("impdebsol1"))
                    xTotalDeb = NulosN(Rst2("impdebsol1"))
                Else
                    Saldo = NulosN(Rst2("impdebdol"))
                    xTotalDeb = NulosN(Rst2("impdebdol"))
                End If
            Else
                If xIdMon = 1 Then
                    Saldo = NulosN(Rst2("imphabsol1"))
                    xTotalHab = NulosN(Rst2("imphabsol1"))
                Else
                    Saldo = NulosN(Rst2("imphabdol"))
                    xTotalHab = NulosN(Rst2("imphabdol"))
                End If
            End If
        End If
    
    Else
        
        '--si la fecha de consulta es superior al inicio del año
        
        nSQL = "SELECT con_diario.idcue, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
            + vbCr + " sum( IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol)) AS impdebesol, " _
            + vbCr + " sum( IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol)) AS imphabersol, " _
            + vbCr + " sum( IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(con_tc.impven Is Null Or con_diario.impdebsol=0,0,(con_diario.impdebsol/con_tc.impven)))) AS impdebedol, " _
            + vbCr + " sum(IIf(con_diario.idmon = 2, con_diario.imphabdol, IIf(con_tc.impven Is Null Or con_diario.imphabsol = 0, 0, (con_diario.imphabsol / con_tc.impven)))) As imphaberdol " _
            + vbCr + " FROM (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN con_planctas ON con_diario.idcue = con_planctas.id " _
            + vbCr + " WHERE (((con_diario.idlib)=3) AND ((con_diario.idcue)=" & NulosN(LblIdCuentaContable.Caption) & ") AND ((con_diario.fchasi) Is Null)) OR (((con_diario.idlib)=6) AND ((con_diario.idcue)=" & NulosN(LblIdCuentaContable.Caption) & ") AND ((con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "'))) " _
            + vbCr + " GROUP BY con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion;"
        
        RST_Busq Rst2, nSQL, xCon
        
        Saldo = 0
               
        If Rst2.RecordCount <> 0 Then
            If xIdMon = 1 Then
                xTotalDeb = xTotalDeb + NulosN(Rst2("impdebesol"))
                xTotalHab = xTotalHab + NulosN(Rst2("imphabersol"))
            Else
                xTotalDeb = xTotalDeb + NulosN(Rst2("impdebedol"))
                xTotalHab = xTotalHab + NulosN(Rst2("imphaberdol"))
            
            End If
        
        End If
        
    End If
    
    Saldo = xTotalDeb - xTotalHab
    Fg1.Rows = Fg1.Rows + 1
    Fg1.TextMatrix(Fg1.Rows - 1, 6) = "Saldo Anterior ==>"
    Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(Saldo, "0.00")
    Fg1.Rows = Fg1.Rows + 1
    
    '*****************************************************************************************************************
    
    If rst.RecordCount = 0 Then
        MsgBox "La cuenta de banco seleccionada no tiene movimientos", vbInformation + vbOKCancel + vbDefaultButton1, xTitulo
        Set rst = Nothing
        Exit Sub
    Else
        rst.MoveFirst
        xFila = 3
        For A = 1 To rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(xFila, 1) = rst("numasi2") & ""
            Fg1.TextMatrix(xFila, 2) = Format(rst("fchope"), "dd/mm/yy")
            Fg1.TextMatrix(xFila, 3) = rst("descripcion") & ""
            Fg1.TextMatrix(xFila, 4) = rst("numdoc") & ""
            Fg1.TextMatrix(xFila, 7) = Format(NulosN(rst("debe")), FORMAT_MONTO)
            Fg1.TextMatrix(xFila, 8) = Format(NulosN(rst("haber")), FORMAT_MONTO)
            xTotDeb = xTotDeb + NulosN(rst("debe"))
            xTotHab = xTotHab + NulosN(rst("haber"))
            If rst("debe") <> 0 Then
                Saldo = Saldo + NulosN(rst("debe"))
            Else
                Saldo = Saldo - NulosN(rst("haber"))
            End If
            
            Fg1.TextMatrix(xFila, 9) = Format(Saldo, FORMAT_MONTO)
            If rst("conciliado") = True Then
                Fg1.TextMatrix(xFila, 10) = -1
            Else
                Fg1.TextMatrix(xFila, 10) = 0
            End If
            Fg1.TextMatrix(xFila, 11) = rst("id")
            rst.MoveNext
            If rst.EOF = True Then
                Exit For
            End If
            xFila = xFila + 1
        Next A
        
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 6) = "Total del Periodo ==>"
        Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(xTotDeb, FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(xTotHab, FORMAT_MONTO)
    End If
    
    'CARGAMOS LOS CHEQUES GIRADOS PERO NO COBRADOS
    Set rst = Nothing
    RST_Busq rst, "SELECT DISTINCT con_cajabanco.id, con_cajabanco.conciliado, con_diario.numasi, con_cajabanco.fchope, mae_doccajaban.descripcion, con_cajabanco.numdoc, 0 AS debe, " _
        & " con_cajabanco.importe AS haber, con_cajabanco.tipmov FROM (con_cajabanco LEFT JOIN mae_doccajaban ON con_cajabanco.iddoc = mae_doccajaban.id) " _
        & " LEFT JOIN con_diario ON con_cajabanco.id = con_diario.idmov WHERE (((con_cajabanco.tipmov)=2) AND ((con_diario.idlib)=6) AND ((con_cajabanco.tipope)=2) " _
        & " AND ((con_cajabanco.idcueban)=" & NulosN(LblIdCuenta.Caption) & ") AND ((con_cajabanco.fchreg)<CDate('" & TxtFchIni.Valor & "')) AND ((con_cajabanco.conciliado)=0))", xCon

    Fg5.Rows = 1
    
    xTotDeb = 0:    xTotHab = 0
    For A = 1 To rst.RecordCount
        Fg5.Rows = Fg5.Rows + 1
        Fg5.TextMatrix(A, 1) = rst("numasi") & ""
        Fg5.TextMatrix(A, 2) = rst("descripcion") & ""
        Fg5.TextMatrix(A, 3) = Format(rst("fchope"), "dd/mm/yy")
        Fg5.TextMatrix(A, 4) = rst("numdoc") & ""
        Fg5.TextMatrix(A, 5) = Format(NulosN(rst("debe")), FORMAT_MONTO)
        Fg5.TextMatrix(A, 6) = Format(NulosN(rst("haber")), FORMAT_MONTO)
        xTotDeb = xTotDeb + NulosN(rst("debe"))
        xTotHab = xTotHab + NulosN(rst("haber"))
        If rst("debe") <> 0 Then
            Saldo = Saldo + NulosN(rst("debe"))
        Else
            Saldo = Saldo - NulosN(rst("haber"))
        End If
        
        Fg5.TextMatrix(A, 8) = rst("id")
        rst.MoveNext
        If rst.EOF = True Then
            Exit For
        End If
    Next A
    TxtTotDeb5.Text = Format(xTotDeb, FORMAT_MONTO)
    TxtTotHab5.Text = Format(xTotHab, FORMAT_MONTO)
    
    TabOne1.CurrTab = 0
    TabOne2.CurrTab = 0
    
    Set rst = Nothing
    'actualizamos los cheques girados y no cobrados
    xFchIni = "01/" + Format(CDate(TxtFchIni.Valor), "MM") + "/" + AnoTra
    RST_Busq rst, "SELECT * FROM con_cajabancocobrados WHERE fchreg = cdate('" & xFchIni & "') and idcueban = " & NulosN(LblIdCuenta.Caption) & "", xCon
    
    If rst.RecordCount <> 0 Then
        TxtImpEstado.Text = Format(rst("salban"), "0.00")
        rst.MoveFirst
        For A = 1 To rst.RecordCount
            For B = 1 To Fg5.Rows - 1
                If rst("idcajban") = NulosN(Fg5.TextMatrix(B, 8)) Then
                    Fg5.TextMatrix(B, 7) = -1
                    Exit For
                End If
            Next B
            rst.MoveNext
            If rst.EOF = True Then Exit For
        Next A
        HallarTotalNoCobrados
    End If
    
    'borramos los cheques que fueron cobrados en otros mes
    Set rst = Nothing
    RST_Busq rst, "SELECT * From con_cajabancocobrados WHERE (((con_cajabancocobrados.fchreg)<CDate('" & xFchIni & "')) AND ((con_cajabancocobrados.idcueban)=" & NulosN(LblIdCuenta.Caption) & "))", xCon
    If rst.RecordCount <> 0 Then
        rst.MoveFirst
        For A = 1 To rst.RecordCount
            For B = 1 To Fg5.Rows - 1
                If rst("idcajban") = NulosN(Fg5.TextMatrix(B, 8)) Then
                    Fg5.RemoveItem B
                    Exit For
                End If
            Next B
            
            rst.MoveNext
            If rst.EOF = True Then Exit For
        Next A
        HallarTotalNoCobrados
    End If
End Sub



VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmLibroBancos 
   Caption         =   "Contabilidad - Libro Bancos"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6270
      Left            =   0
      TabIndex        =   0
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
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   5850
         Left            =   45
         TabIndex        =   24
         Top             =   375
         Width           =   11805
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   5565
            Left            =   15
            TabIndex        =   25
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
            FormatString    =   $"FrmLibroBancos.frx":0000
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
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   5850
         Left            =   12540
         TabIndex        =   1
         Top             =   375
         Width           =   11805
         Begin VB.CommandButton CmdProcesar 
            Caption         =   "&Ver Conciliacion"
            Height          =   630
            Left            =   9510
            TabIndex        =   19
            Top             =   15
            Width           =   2280
         End
         Begin VB.TextBox TxtImpEstado 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   18
            Text            =   "TxtImpEstado"
            Top             =   105
            Width           =   1200
         End
         Begin VB.TextBox TxtTot1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   8355
            Locked          =   -1  'True
            TabIndex        =   17
            Text            =   "TxtTot1"
            Top             =   2910
            Width           =   1095
         End
         Begin VB.TextBox TxtTot2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   9450
            Locked          =   -1  'True
            TabIndex        =   16
            Text            =   "TxtTot2"
            Top             =   2910
            Width           =   1095
         End
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   2580
            Left            =   30
            TabIndex        =   2
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
            CurrTab         =   1
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
            Begin VB.Frame Frame7 
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   0  'None
               Height          =   2160
               Left            =   -12315
               TabIndex        =   11
               Top             =   45
               Width           =   11670
               Begin VB.TextBox TxtTotDeb5 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Height          =   300
                  Left            =   8370
                  Locked          =   -1  'True
                  TabIndex        =   13
                  Text            =   "TxtTotDeb5"
                  Top             =   1815
                  Width           =   1095
               End
               Begin VB.TextBox TxtTotHab5 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Height          =   300
                  Left            =   9465
                  Locked          =   -1  'True
                  TabIndex        =   12
                  Text            =   "TxtTotHab5"
                  Top             =   1815
                  Width           =   1095
               End
               Begin VSFlex7Ctl.VSFlexGrid Fg5 
                  Height          =   1725
                  Left            =   45
                  TabIndex        =   14
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
                  FormatString    =   $"FrmLibroBancos.frx":0194
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
                  TabIndex        =   15
                  Top             =   1860
                  Width           =   675
               End
            End
            Begin VB.Frame Frame8 
               BackColor       =   &H00C0E0FF&
               BorderStyle     =   0  'None
               Height          =   2160
               Left            =   45
               TabIndex        =   3
               Top             =   45
               Width           =   11670
               Begin VB.TextBox TxtTotDeb 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Height          =   300
                  Left            =   6300
                  Locked          =   -1  'True
                  TabIndex        =   8
                  Text            =   "TxtTotDeb"
                  Top             =   1815
                  Width           =   1095
               End
               Begin VB.TextBox TxtTotHab 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Height          =   300
                  Left            =   7395
                  Locked          =   -1  'True
                  TabIndex        =   7
                  Text            =   "TxtTotHab"
                  Top             =   1815
                  Width           =   1095
               End
               Begin VB.Frame Frame2 
                  Height          =   2115
                  Left            =   10050
                  TabIndex        =   4
                  Top             =   -15
                  Width           =   1590
                  Begin VB.CommandButton CmdDelMov 
                     Caption         =   "Eliminar Movimiento"
                     Enabled         =   0   'False
                     Height          =   495
                     Left            =   165
                     Style           =   1  'Graphical
                     TabIndex        =   6
                     Top             =   1080
                     Width           =   1260
                  End
                  Begin VB.CommandButton CmdAgregaMov 
                     Caption         =   "Agregar Movimiento"
                     Enabled         =   0   'False
                     Height          =   495
                     Left            =   165
                     Style           =   1  'Graphical
                     TabIndex        =   5
                     Top             =   555
                     Width           =   1260
                  End
               End
               Begin VSFlex7Ctl.VSFlexGrid Fg3 
                  Height          =   1725
                  Left            =   45
                  TabIndex        =   9
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
                  FormatString    =   $"FrmLibroBancos.frx":02E0
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
                  TabIndex        =   10
                  Top             =   1860
                  Width           =   675
               End
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg2 
            Height          =   2205
            Left            =   15
            TabIndex        =   20
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
            FormatString    =   $"FrmLibroBancos.frx":03D1
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Importe Estado Cuenta"
            Height          =   195
            Left            =   60
            TabIndex        =   23
            Top             =   150
            Width           =   1620
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Movimientos en Libros y no en Estado de Cuenta"
            Height          =   195
            Left            =   60
            TabIndex        =   22
            Top             =   465
            Width           =   3465
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Total ==>"
            Height          =   195
            Left            =   7380
            TabIndex        =   21
            Top             =   2955
            Width           =   675
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
            Picture         =   "FrmLibroBancos.frx":04E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos.frx":0A2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos.frx":0DBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos.frx":0F18
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos.frx":12AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos.frx":142E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos.frx":1882
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos.frx":199A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos.frx":1EDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos.frx":2422
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos.frx":2536
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos.frx":264A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos.frx":2A9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos.frx":2C0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos.frx":3152
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos.frx":34E4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   1005
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
      Left            =   30
      TabIndex        =   26
      Top             =   285
      Width           =   11865
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
         TabIndex        =   29
         Top             =   165
         Width           =   1890
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
            TabIndex        =   31
            Top             =   270
            Width           =   1560
         End
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
            TabIndex        =   30
            Top             =   525
            Width           =   1560
         End
      End
      Begin VB.CommandButton CmdBusCliPro 
         Height          =   240
         Left            =   3195
         Picture         =   "FrmLibroBancos.frx":3876
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   345
         Width           =   240
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   315
         Left            =   1035
         TabIndex        =   32
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
         TabIndex        =   33
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
      Begin VB.TextBox TxtCuenta 
         Height          =   300
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "TxtCuenta"
         Top             =   315
         Width           =   2430
      End
      Begin VB.Label LblIdMoneda 
         AutoSize        =   -1  'True
         Caption         =   "LblIdMoneda"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   8160
         TabIndex        =   41
         Top             =   840
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label LblIdCuenta 
         AutoSize        =   -1  'True
         Caption         =   "LblIdCuenta"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   8160
         TabIndex        =   40
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Inicio"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   39
         Top             =   690
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Cuenta"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   38
         Top             =   345
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Final"
         Height          =   195
         Index           =   2
         Left            =   2535
         TabIndex        =   37
         Top             =   690
         Width           =   690
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
         TabIndex        =   36
         Top             =   315
         Width           =   2835
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
         TabIndex        =   35
         Top             =   315
         Width           =   1365
      End
      Begin VB.Label LblIdCuentaContable 
         AutoSize        =   -1  'True
         Caption         =   "LblIdCuentaContable"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   8160
         TabIndex        =   34
         Top             =   360
         Visible         =   0   'False
         Width           =   1485
      End
   End
End
Attribute VB_Name = "FrmLibroBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RstLib As New ADODB.Recordset
Dim SeEjecuto As Boolean

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = False
    End If
End Sub

Sub MostrarPeriodo()
    Dim A As Integer
    LblIdMoneda.Caption = "1"
    
    RST_Busq RstLib, "SELECT tes_caja.fchope, tes_cajaorigendet.idmod, tes_documentos.descripcion AS descdoc, tes_caja.numreg, tes_caja.idmon, " _
        & " tes_cajaorigendet.numser, tes_cajaorigendet.numdoc, tes_cajaorigendet.importe, tes_origen.descripcion AS descori, con_planctas.cuenta, " _
        & " tes_mediopago.descripcion AS descmedpag, tes_origen.idcuen FROM (((tes_origen RIGHT JOIN (tes_caja RIGHT JOIN (tes_cajaori RIGHT JOIN " _
        & " tes_cajaorigendet ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) " _
        & " ON tes_caja.id = tes_cajaori.idtes) ON tes_origen.id = tes_cajaori.idori) LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) " _
        & " LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id) LEFT JOIN tes_mediopago ON tes_cajaorigendet.idmedpag = tes_mediopago.id " _
        & " WHERE (((tes_caja.fchope)>=CDate('" & TxtFchIni.Valor & "') And (tes_caja.fchope)<=CDate('" & TxtFchFin.Valor & "')) AND ((tes_cajaorigendet.idmod)=6) " _
        & " AND ((tes_caja.idmon)=" & NulosN(LblIdMoneda.Caption) & ") AND ((tes_origen.idcuen)=9)) ORDER BY tes_caja.fchope", xCon

    Fg1.Rows = 1
    If RstLib.RecordCount <> 0 Then
        RstLib.MoveFirst
        For A = 1 To RstLib.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = RstLib("numreg")
            Fg1.TextMatrix(A, 2) = RstLib("fchope")
            Fg1.TextMatrix(A, 3) = RstLib("descdoc")
            Fg1.TextMatrix(A, 4) = NulosC(RstLib("numser")) & NulosC(RstLib("numdoc"))
            Fg1.TextMatrix(A, 5) = RstLib("descmedpag")
            'Fg1.TextMatrix(A, 6) = 'RstLib("")
            'Fg1.TextMatrix(A, 7) = RstLib("")
            Fg1.TextMatrix(A, 8) = Format(RstLib("importe"), "0.00")
            
            RstLib.MoveNext
            If RstLib.EOF = True Then Exit For
        Next A
    End If

End Sub

Sub Conciliar()

End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Conciliar
    
    If Button.Index = 3 Then MostrarPeriodo
    
    If Button.Index = 10 Then
        Set RstLib = Nothing
        Unload Me
    End If
End Sub

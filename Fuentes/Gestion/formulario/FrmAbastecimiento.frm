VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmAbastecimiento 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Unificado - Plan de Abastecimiento"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7635
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame15 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame15"
      Height          =   315
      Left            =   5760
      TabIndex        =   46
      Top             =   7290
      Width           =   6015
      Begin VB.Shape Shape4 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   180
         Left            =   3195
         Top             =   45
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "= Faltante de Abastecimiento"
         Height          =   195
         Left            =   3870
         TabIndex        =   48
         Top             =   45
         Width           =   2055
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   180
         Left            =   900
         Top             =   45
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "= Sobre Abastecido"
         Height          =   195
         Left            =   1545
         TabIndex        =   47
         Top             =   45
         Width           =   1395
      End
   End
   Begin VB.Frame FrmProgreso 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   1215
      Left            =   2970
      TabIndex        =   39
      Top             =   3600
      Visible         =   0   'False
      Width           =   5625
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   90
         TabIndex        =   40
         Top             =   850
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label LblProducto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LblProducto"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1200
         TabIndex        =   45
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   105
         TabIndex        =   44
         Top             =   75
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         Index           =   3
         X1              =   0
         X2              =   5610
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   2
         X1              =   5610
         X2              =   5610
         Y1              =   15
         Y2              =   1200
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   1035
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   2
         X1              =   0
         X2              =   5610
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label LblProcesa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa: "
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
         Left            =   105
         TabIndex        =   43
         Top             =   345
         Width           =   855
      End
      Begin VB.Label LblEmpresa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LblEmpresa"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1200
         TabIndex        =   42
         Top             =   345
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Producto: "
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
         Left            =   105
         TabIndex        =   41
         Top             =   600
         Width           =   900
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   30
         Top             =   30
         Width           =   5550
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "&H80000009&"
      Height          =   7620
      Left            =   9420
      TabIndex        =   32
      Top             =   -4440
      Visible         =   0   'False
      Width           =   11805
      Begin VB.CommandButton Command5 
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
         Height          =   540
         Left            =   225
         TabIndex        =   38
         ToolTipText     =   "Reducir columnas"
         Top             =   7000
         Width           =   735
      End
      Begin VB.CommandButton Command6 
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
         Height          =   540
         Left            =   990
         TabIndex        =   37
         ToolTipText     =   "Agrandar columnas"
         Top             =   7000
         Width           =   735
      End
      Begin VB.CommandButton CmdPrin 
         Height          =   555
         Left            =   10050
         Picture         =   "FrmAbastecimiento.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Exportar al Excel"
         Top             =   7000
         Width           =   735
      End
      Begin VB.CommandButton CmdSalir 
         Height          =   555
         Left            =   10815
         Picture         =   "FrmAbastecimiento.frx":0B0A
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Salir"
         Top             =   7000
         Width           =   735
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg7 
         Height          =   6480
         Left            =   60
         TabIndex        =   34
         Top             =   420
         Width           =   11640
         _cx             =   20532
         _cy             =   11430
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   128
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
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
         Cols            =   20
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmAbastecimiento.frx":0E14
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
         BackStyle       =   0  'Transparent
         Caption         =   "Consolidado de insumos"
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
         Left            =   105
         TabIndex        =   33
         Top             =   120
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   330
         Left            =   30
         Top             =   45
         Width           =   11700
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   11745
         Y1              =   7600
         Y2              =   7600
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   6645
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   0
         X1              =   11760
         X2              =   11760
         Y1              =   15
         Y2              =   7600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   11745
         Y1              =   15
         Y2              =   15
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   375
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
      Caption         =   "  &Consulta  |   &Detalle  |New Tab|New Tab|New Tab"
      Align           =   0
      CurrTab         =   2
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
         Height          =   6915
         Index           =   4
         Left            =   12840
         TabIndex        =   25
         Top             =   375
         Width           =   11805
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   6675
            Index           =   4
            Left            =   30
            TabIndex        =   26
            Top             =   180
            Width           =   11745
            _cx             =   20717
            _cy             =   11774
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   0
            MousePointer    =   0
            _ConvInfo       =   1
            Version         =   700
            BackColor       =   12632256
            ForeColor       =   -2147483630
            FrontTabColor   =   13160660
            BackTabColor    =   12632256
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483641
            Caption         =   " Terminados  | Intermedios "
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
            Begin VB.Frame Frame6 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   6315
               Index           =   4
               Left            =   15
               TabIndex        =   29
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid Fg6 
                  Height          =   6255
                  Index           =   4
                  Left            =   15
                  TabIndex        =   30
                  Top             =   15
                  Width           =   11670
                  _cx             =   20585
                  _cy             =   11033
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
                  BackColor       =   14613184
                  ForeColor       =   -2147483640
                  BackColorFixed  =   -2147483633
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   128
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483636
                  BackColorAlternate=   14613184
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
                  Cols            =   20
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmAbastecimiento.frx":107E
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
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   6315
               Index           =   4
               Left            =   -12330
               TabIndex        =   27
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid Fg5 
                  Height          =   6255
                  Index           =   4
                  Left            =   15
                  TabIndex        =   28
                  Top             =   15
                  Width           =   11670
                  _cx             =   20585
                  _cy             =   11033
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
                  BackColor       =   14417405
                  ForeColor       =   -2147483640
                  BackColorFixed  =   -2147483633
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   128
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483636
                  BackColorAlternate=   14417405
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
                  Cols            =   20
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmAbastecimiento.frx":12E9
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
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   6915
         Index           =   3
         Left            =   12540
         TabIndex        =   7
         Top             =   375
         Width           =   11805
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   6675
            Index           =   3
            Left            =   30
            TabIndex        =   8
            Top             =   180
            Width           =   11745
            _cx             =   20717
            _cy             =   11774
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   0
            MousePointer    =   0
            _ConvInfo       =   1
            Version         =   700
            BackColor       =   12632256
            ForeColor       =   -2147483630
            FrontTabColor   =   13160660
            BackTabColor    =   12632256
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483641
            Caption         =   " Terminados  | Intermedios "
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
            Begin VB.Frame Frame6 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   6315
               Index           =   3
               Left            =   15
               TabIndex        =   16
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid Fg6 
                  Height          =   6255
                  Index           =   3
                  Left            =   15
                  TabIndex        =   24
                  Top             =   15
                  Width           =   11670
                  _cx             =   20585
                  _cy             =   11033
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
                  BackColor       =   14613184
                  ForeColor       =   -2147483640
                  BackColorFixed  =   -2147483633
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   128
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483636
                  BackColorAlternate=   14613184
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
                  Cols            =   20
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmAbastecimiento.frx":1553
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
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   6315
               Index           =   3
               Left            =   -12330
               TabIndex        =   13
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid Fg5 
                  Height          =   6255
                  Index           =   3
                  Left            =   15
                  TabIndex        =   21
                  Top             =   15
                  Width           =   11670
                  _cx             =   20585
                  _cy             =   11033
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
                  BackColor       =   14417405
                  ForeColor       =   -2147483640
                  BackColorFixed  =   -2147483633
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   128
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483636
                  BackColorAlternate=   14417405
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
                  Cols            =   20
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmAbastecimiento.frx":17BE
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
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6915
         Index           =   2
         Left            =   45
         TabIndex        =   5
         Top             =   375
         Width           =   11805
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   6675
            Index           =   2
            Left            =   30
            TabIndex        =   6
            Top             =   180
            Width           =   11745
            _cx             =   20717
            _cy             =   11774
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   0
            MousePointer    =   0
            _ConvInfo       =   1
            Version         =   700
            BackColor       =   12632256
            ForeColor       =   -2147483630
            FrontTabColor   =   13160660
            BackTabColor    =   12632256
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483641
            Caption         =   " Terminados  | Intermedios "
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
            Begin VB.Frame Frame6 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   6315
               Index           =   2
               Left            =   15
               TabIndex        =   15
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid Fg6 
                  Height          =   6255
                  Index           =   2
                  Left            =   15
                  TabIndex        =   23
                  Top             =   15
                  Width           =   11670
                  _cx             =   20585
                  _cy             =   11033
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
                  BackColor       =   14613184
                  ForeColor       =   -2147483640
                  BackColorFixed  =   -2147483633
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   128
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483636
                  BackColorAlternate=   14613184
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
                  Cols            =   20
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmAbastecimiento.frx":1A28
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
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   6315
               Index           =   2
               Left            =   -12330
               TabIndex        =   12
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid Fg5 
                  Height          =   6255
                  Index           =   2
                  Left            =   15
                  TabIndex        =   20
                  Top             =   15
                  Width           =   11670
                  _cx             =   20585
                  _cy             =   11033
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
                  BackColor       =   14417405
                  ForeColor       =   -2147483640
                  BackColorFixed  =   -2147483633
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   128
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483636
                  BackColorAlternate=   14417405
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
                  Cols            =   20
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmAbastecimiento.frx":1C93
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
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   6915
         Index           =   1
         Left            =   -12450
         TabIndex        =   2
         Top             =   375
         Width           =   11805
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   6675
            Index           =   1
            Left            =   30
            TabIndex        =   3
            Top             =   180
            Width           =   11745
            _cx             =   20717
            _cy             =   11774
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   0
            MousePointer    =   0
            _ConvInfo       =   1
            Version         =   700
            BackColor       =   12632256
            ForeColor       =   -2147483630
            FrontTabColor   =   13160660
            BackTabColor    =   12632256
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483641
            Caption         =   " Terminados  | Intermedios "
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
            Begin VB.Frame Frame6 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   6315
               Index           =   1
               Left            =   15
               TabIndex        =   14
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid Fg6 
                  Height          =   6255
                  Index           =   1
                  Left            =   15
                  TabIndex        =   22
                  Top             =   15
                  Width           =   11670
                  _cx             =   20585
                  _cy             =   11033
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
                  BackColor       =   14613184
                  ForeColor       =   -2147483640
                  BackColorFixed  =   -2147483633
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   128
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483636
                  BackColorAlternate=   14613184
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
                  Cols            =   20
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmAbastecimiento.frx":1EFD
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
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   6315
               Index           =   1
               Left            =   -12330
               TabIndex        =   11
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid Fg5 
                  Height          =   6255
                  Index           =   1
                  Left            =   15
                  TabIndex        =   19
                  Top             =   15
                  Width           =   11670
                  _cx             =   20585
                  _cy             =   11033
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
                  BackColor       =   14417405
                  ForeColor       =   -2147483640
                  BackColorFixed  =   -2147483633
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   128
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483636
                  BackColorAlternate=   14417405
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
                  Cols            =   20
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmAbastecimiento.frx":2168
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
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6915
         Index           =   0
         Left            =   -12750
         TabIndex        =   1
         Top             =   375
         Width           =   11805
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   6675
            Index           =   0
            Left            =   30
            TabIndex        =   4
            Top             =   180
            Width           =   11745
            _cx             =   20717
            _cy             =   11774
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   0
            MousePointer    =   0
            _ConvInfo       =   1
            Version         =   700
            BackColor       =   12632256
            ForeColor       =   -2147483630
            FrontTabColor   =   13160660
            BackTabColor    =   12632256
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483641
            Caption         =   " Terminados  | Intermedios "
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
            Begin VB.Frame Frame6 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   6315
               Index           =   0
               Left            =   12360
               TabIndex        =   10
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid Fg6 
                  Height          =   6255
                  Index           =   0
                  Left            =   15
                  TabIndex        =   18
                  Top             =   15
                  Width           =   11670
                  _cx             =   20585
                  _cy             =   11033
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
                  BackColor       =   14613184
                  ForeColor       =   -2147483640
                  BackColorFixed  =   -2147483633
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   128
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483636
                  BackColorAlternate=   14613184
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
                  Cols            =   20
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmAbastecimiento.frx":23D2
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
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   6315
               Index           =   0
               Left            =   15
               TabIndex        =   9
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid Fg5 
                  Height          =   6255
                  Index           =   0
                  Left            =   15
                  TabIndex        =   17
                  Top             =   15
                  Width           =   11670
                  _cx             =   20585
                  _cy             =   11033
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
                  BackColor       =   14417405
                  ForeColor       =   -2147483640
                  BackColorFixed  =   -2147483633
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   128
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483636
                  BackColorAlternate=   14417405
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
                  Cols            =   20
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmAbastecimiento.frx":263D
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
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5055
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAbastecimiento.frx":28A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAbastecimiento.frx":2DEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAbastecimiento.frx":2F45
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAbastecimiento.frx":32D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAbastecimiento.frx":345B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAbastecimiento.frx":38AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAbastecimiento.frx":39C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAbastecimiento.frx":3F0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAbastecimiento.frx":444F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAbastecimiento.frx":4563
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAbastecimiento.frx":4677
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAbastecimiento.frx":4ACB
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAbastecimiento.frx":4C37
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAbastecimiento.frx":517F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   31
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
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mostrar plan  de abastecimiento unificado"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "exportar excel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmAbastecimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMABASTECIMIENTO
'* Tipo             : MODULO
'* Descripcion      : FORMULARIO QUE VISUALIZA LA INFORMACION DEL PLAN DE ABASTECIMIENTO
'*
'* DISE�ADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 10/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim RstEmp As New ADODB.Recordset            ' RECORDSET PARA CARGAR LA INFORMACION DEL PLAN DE ABASTECIMIENTO ACTIVO
Dim RstInsumos As New ADODB.Recordset        ' RECORDSET PARA CARGAR LOS INSUMOS QUE SE REQUIEREN PARA EL PLAN DE ABASTECIMIENTO
Dim xCon1 As ADODB.Connection                ' CONECCION ADICIONAL A UTILIZAR
Dim xCon2 As ADODB.Connection                ' CONECCION ADICIONAL A UTILIZAR
Dim SeEjecuto As Boolean                     ' VERIFICA QUE EL EVENTO ACTIVA SE EJECUTE UNA SOLA VEZ

Private Sub CmdPrin_Click()
    'ExportarExcelUnif
    Dim xExport As New SGI2_funciones.Formularios
    xExport.VSFlexGrid_Exportar_MSExcel xCon, Fg7, "Consilidado de Insumos y Materia Prima", "Consolidado de Empresas", "", "Plan Unificado de Abastecimiento"
    Set xExport = Nothing
    
End Sub

Private Sub CmdSalir_Click()
    Frame2.Visible = False
    Toolbar1.Enabled = True
    TabOne1.Enabled = True
End Sub

Private Sub Command5_Click()
    ' CAMBIAMOS EL ANCHO DE LAS COLUMNAS DEL CONTRO Fg7 RESTANDOLE 10 PIXEL AL ANCHO ACTUAL
    Fg7.ColWidth(1) = Fg7.ColWidth(1) - 100
    Fg7.ColWidth(5) = Fg7.ColWidth(5) - 10
    Fg7.ColWidth(6) = Fg7.ColWidth(6) - 10
    Fg7.ColWidth(7) = Fg7.ColWidth(7) - 10
    Fg7.ColWidth(8) = Fg7.ColWidth(8) - 10
    Fg7.ColWidth(9) = Fg7.ColWidth(9) - 10
    Fg7.ColWidth(10) = Fg7.ColWidth(10) - 10
    Fg7.ColWidth(11) = Fg7.ColWidth(11) - 10
    Fg7.ColWidth(12) = Fg7.ColWidth(12) - 10
    Fg7.ColWidth(13) = Fg7.ColWidth(13) - 10
    Fg7.ColWidth(14) = Fg7.ColWidth(14) - 10
    Fg7.ColWidth(15) = Fg7.ColWidth(15) - 10
    Fg7.ColWidth(16) = Fg7.ColWidth(16) - 10
End Sub

Private Sub Command6_Click()
    ' CAMBIAMOS EL ANCHO DE LAS COLUMNAS DEL CONTRO Fg7 SUMANDOLE 10 PIXEL AL ANCHO ACTUAL
    Fg7.ColWidth(1) = Fg7.ColWidth(1) + 100
    Fg7.ColWidth(5) = Fg7.ColWidth(5) + 10
    Fg7.ColWidth(6) = Fg7.ColWidth(6) + 10
    Fg7.ColWidth(7) = Fg7.ColWidth(7) + 10
    Fg7.ColWidth(8) = Fg7.ColWidth(8) + 10
    Fg7.ColWidth(9) = Fg7.ColWidth(9) + 10
    Fg7.ColWidth(10) = Fg7.ColWidth(10) + 10
    Fg7.ColWidth(11) = Fg7.ColWidth(11) + 10
    Fg7.ColWidth(12) = Fg7.ColWidth(12) + 10
    Fg7.ColWidth(13) = Fg7.ColWidth(13) + 10
    Fg7.ColWidth(14) = Fg7.ColWidth(14) + 10
    Fg7.ColWidth(15) = Fg7.ColWidth(15) + 10
    Fg7.ColWidth(16) = Fg7.ColWidth(16) + 10

End Sub

Private Sub Form_Activate()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        Dim A As Integer
        Dim xIndex As Integer
        Dim xRuta As String
    
        SeEjecuto = True
        Set xCon1 = AbrirConecciones(AP_RUTABD + "data.mdb")
        RST_Busq RstEmp, "SELECT mae_empresa.* From mae_empresa WHERE (((mae_empresa.anotra)=" & AnoTra & " ) AND ((mae_empresa.activo)=-1))", xCon1
        
        If RstEmp.RecordCount <> 0 Then
            xIndex = 0
            RstEmp.MoveFirst
            Me.Refresh
            FrmProgreso.Visible = True
            For A = 1 To RstEmp.RecordCount
                TabOne1.TabCaption(xIndex) = " " & Trim(RstEmp("abrevia")) & " "
                TabOne1.TabVisible(xIndex) = True
                Me.Refresh
                LblEmpresa.Caption = " " & Trim(RstEmp("abrevia")) & " "
                
                xRuta = AP_RUTABD + Trim(RstEmp("ruta"))
                Set xCon2 = Nothing
                Set xCon2 = AbrirConecciones(xRuta)
                
                If VerificarPlanesActivos > 1 Then
                    MsgBox "No se puede consultar el plan de produccion unificado, existe mas de 1 plan activo en la empresa " + RstEmp("abrevia") + Chr(13) _
                        & "Verifique que solo exista un plan de produccion activo.", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    Set xCon1 = Nothing
                    Set xCon2 = Nothing
                    Unload Me
                    Exit For
                End If
                
                VerInsumos xIndex
                
                RstEmp.MoveNext
                
                If RstEmp.EOF = True Then
                    Exit For
                End If
                xIndex = xIndex + 1
            Next A
            FrmProgreso.Visible = False
        End If
        TabOne1.CurrTab = 0
    End If
End Sub

Function VerificarPlanesActivos() As Integer
    ' VERIFICA QUE HAYA POR LO MENOS UN PLAN ACTIVO
    Dim Rst2 As New ADODB.Recordset
    RST_Busq Rst2, "SELECT * FROM ges_planaba WHERE activo = -1", xCon2
    VerificarPlanesActivos = Rst2.RecordCount
    Set Rst2 = Nothing
End Function

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    Dim A As Integer
    Dim xIndex As Integer
    xIndex = 0
    SeEjecuto = False
    
    For A = 1 To 5
        TabOne1.TabVisible(xIndex) = False
        Frame1(xIndex).BackColor = &H8000000F
        Frame5(xIndex).BackColor = &H8000000F
        Frame6(xIndex).BackColor = &H8000000F
        
        Fg5(xIndex).FrozenCols = 4
        Fg6(xIndex).FrozenCols = 4
        
        Fg5(xIndex).ColWidth(2) = 0
        Fg5(xIndex).ColWidth(3) = 0
        
        Fg6(xIndex).ColWidth(2) = 0
        Fg6(xIndex).ColWidth(3) = 0
        xIndex = xIndex + 1
    Next A
    Frame2.BackColor = &H8000000F
    
    Fg7.ColWidth(2) = 0
    Fg7.ColWidth(3) = 0
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : VerInsumos
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA LA LISTA DE INSUMOS NECESARIOS PARA CUBRIR EL PLAN DE ABASTECIMIENTO
'* DISE�ADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Index        |  Integer   |  ESPECIFICA EL ID DEL CONTROL Fg5
'* DEVUELVE         :
'*****************************************************************************************************
Sub VerInsumos(Index As Integer)
    Dim Rst As New ADODB.Recordset
    Dim Rst2 As New ADODB.Recordset
    Dim A, B, xCol As Integer
    Dim Total As Double
    
    'MOSTRAMOS LOS INSUMOS PARA LOS PRODUCTOS TERMINADOS
    
'SELECT DISTINCT ges_planaba.fchini, alm_inventario.idmae, alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev
'FROM (ges_planaba LEFT JOIN ges_planabadet ON ges_planaba.id = ges_planabadet.idpv) LEFT JOIN (mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed) ON ges_planabadet.codpro = alm_inventario.id
'Where (((ges_planaba.activo) = -1) And ((ges_planabadet.tipo) = 1))
'ORDER BY alm_inventario.descripcion;


    RST_Busq Rst, "SELECT DISTINCT ges_planaba.fchini, alm_inventario.idmae, alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev " _
        + vbCr + "FROM (ges_planaba LEFT JOIN ges_planabadet ON ges_planaba.id = ges_planabadet.idpv) LEFT JOIN (mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed) ON ges_planabadet.codpro = alm_inventario.id " _
        + vbCr + "Where (((ges_planaba.activo) = -1) And ((ges_planabadet.tipo) = 1)) " _
        + vbCr + "ORDER BY alm_inventario.descripcion", xCon2

    Fg5(Index).Rows = 1
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        ProgressBar1.Max = Rst.RecordCount
        For A = 1 To Rst.RecordCount
            ProgressBar1.Value = A
            
            Fg5(Index).Rows = Fg5(Index).Rows + 1
            Total = 0
            Fg5(Index).TextMatrix(A, 0) = Rst("id")
            Fg5(Index).TextMatrix(A, 1) = Rst("descripcion")
            LblProducto.Caption = Rst("descripcion")
            Fg5(Index).TextMatrix(A, 2) = Rst("codpro")
            Fg5(Index).TextMatrix(A, 3) = NulosN(Rst("idmae"))
            Fg5(Index).TextMatrix(A, 4) = NulosC(Rst("abrev"))
            
            RST_Busq Rst2, "SELECT ges_planabadet.idmes, ges_planabadet.cantidad, ges_planabadet.codpro, ges_planabadet.tipo, ges_planaba.activo" _
                + vbCr + "FROM ges_planaba LEFT JOIN ges_planabadet ON ges_planaba.id = ges_planabadet.idpv " _
                + vbCr + "WHERE (((ges_planabadet.codpro)= " & Rst("id") & ") AND ((ges_planabadet.tipo)=1) AND ((ges_planaba.activo)=-1)) " _
                + vbCr + "ORDER BY ges_planabadet.idmes", xCon2

            xCol = 5
            Rst2.MoveFirst

            For B = 1 To 12 'Rst2.RecordCount
                Fg5(Index).TextMatrix(A, xCol) = Format(Rst2("cantidad"), "0.00")
                Total = Total + Rst2("cantidad")

                Rst2.MoveNext
                xCol = xCol + 1
                If Rst2.EOF = True Then
                    Exit For
                End If
            Next B

            Fg5(Index).TextMatrix(A, xCol) = Format(Total, "0.00")
            xCol = xCol + 1
            'Fg5(Index).TextMatrix(A, xCol) = SaldoEntradas(Fg5(Index).TextMatrix(A, 0), "01/01/" & AnoTra, Date, xCon2)
            Fg5(Index).TextMatrix(A, xCol) = SaldoActual(Fg5(Index).TextMatrix(A, 0), Rst("fchini"), Date, xCon2, 1)
            
            Fg5(Index).TextMatrix(A, xCol) = Format(Fg5(Index).TextMatrix(A, xCol), "0.00")
            xCol = xCol + 1
            Fg5(Index).TextMatrix(A, xCol) = Fg5(Index).TextMatrix(A, xCol - 1) - Fg5(Index).TextMatrix(A, xCol - 2)
            Fg5(Index).TextMatrix(A, xCol) = Format(Fg5(Index).TextMatrix(A, xCol), "0.00")
            
            With Fg5(Index)
                .Select A, xCol, A, xCol
                .FillStyle = flexFillRepeat
                If NulosN(.TextMatrix(A, xCol)) >= 0 Then
                    .CellForeColor = &HFF0000
                Else
                    .CellForeColor = &HFF&
                End If
            End With
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
        
        With Fg5(Index)
            .Select 1, .Cols - 3, .Rows - 1, .Cols - 1
            .FillStyle = flexFillRepeat
            .CellBackColor = &H80000013 '&HDDFFFF
            .Select 1, 1, 1, 1
        End With
    End If

    'MOSTRAMOS LOS INSUMOS PARA LOS PRODUCTOS INTERMEDIOS
    
'SELECT DISTINCT ges_planaba.fchini, alm_inventario.idmae, alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev
'FROM (ges_planaba LEFT JOIN ges_planabadet ON ges_planaba.id = ges_planabadet.idpv) LEFT JOIN (mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed) ON ges_planabadet.codpro = alm_inventario.id
'Where (((ges_planaba.activo) = -1) And ((ges_planabadet.tipo) = 2))
'ORDER BY alm_inventario.descripcion;
    
    RST_Busq Rst, "SELECT DISTINCT ges_planaba.fchini, alm_inventario.idmae, alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev " _
        + vbCr + " FROM (ges_planaba LEFT JOIN ges_planabadet ON ges_planaba.id = ges_planabadet.idpv) LEFT JOIN (mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed) ON ges_planabadet.codpro = alm_inventario.id " _
        + vbCr + "Where (((ges_planaba.activo) = -1) And ((ges_planabadet.tipo) = 2)) " _
        + vbCr + "ORDER BY alm_inventario.descripcion", xCon2
        
    Fg6(Index).Rows = 1
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        ProgressBar1.Max = Rst.RecordCount
        For A = 1 To Rst.RecordCount
            ProgressBar1.Value = A
            
            Fg6(Index).Rows = Fg6(Index).Rows + 1
            Total = 0
            Fg6(Index).TextMatrix(A, 0) = Rst("id")
            Fg6(Index).TextMatrix(A, 1) = Rst("descripcion")
            LblProducto.Caption = Rst("descripcion")
            Fg6(Index).TextMatrix(A, 2) = Rst("codpro")
            Fg6(Index).TextMatrix(A, 3) = NulosN(Rst("idmae"))
            Fg6(Index).TextMatrix(A, 4) = Rst("abrev")
            
            RST_Busq Rst2, "SELECT ges_planabadet.idmes, ges_planabadet.cantidad, ges_planabadet.codpro, ges_planabadet.tipo, ges_planaba.activo " _
                        + vbCr + "FROM ges_planaba LEFT JOIN ges_planabadet ON ges_planaba.id = ges_planabadet.idpv " _
                        + vbCr + "WHERE (((ges_planabadet.codpro)= " & Rst("id") & ") AND ((ges_planabadet.tipo)=2) AND ((ges_planaba.activo)=-1)) " _
                        + vbCr + "ORDER BY ges_planabadet.idmes", xCon2
            
            xCol = 5
            If Rst2.RecordCount <> 0 Then Rst2.MoveFirst

            For B = 1 To 12 'Rst2.RecordCount
                Fg6(Index).TextMatrix(A, xCol) = Format(Rst2("cantidad"), "0.00")
                Total = Total + Rst2("cantidad")

                Rst2.MoveNext
                xCol = xCol + 1
                If Rst2.EOF = True Then
                    Exit For
                End If
            Next B

            Fg6(Index).TextMatrix(A, xCol) = Format(Total, "0.00")
            xCol = xCol + 1
            
            Fg6(Index).TextMatrix(A, xCol) = SaldoActual(Fg6(Index).TextMatrix(A, 0), Rst("fchini"), Date, xCon2, 1)
            'Fg6(Index).TextMatrix(A, xCol) = SaldoEntradas(Fg6(Index).TextMatrix(A, 0), "01/01/" & AnoTra, Date, xCon2) '"Stock"
            
            Fg6(Index).TextMatrix(A, xCol) = Format(Fg6(Index).TextMatrix(A, xCol), "0.00")
            xCol = xCol + 1
            Fg6(Index).TextMatrix(A, xCol) = Fg6(Index).TextMatrix(A, xCol - 1) - Fg6(Index).TextMatrix(A, xCol - 2)
            Fg6(Index).TextMatrix(A, xCol) = Format(Fg6(Index).TextMatrix(A, xCol), "0.00")

            With Fg6(Index)
                .Select A, xCol, A, xCol
                .FillStyle = flexFillRepeat
                If NulosN(.TextMatrix(A, xCol)) >= 0 Then
                    .CellForeColor = &HFF0000
                Else
                    .CellForeColor = &HFF&
                End If
            End With
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
        
        With Fg6(Index)
            .Select 1, .Cols - 3, .Rows - 1, .Cols - 1
            .FillStyle = flexFillRepeat
            .CellBackColor = &H80000013 '&HDDFFFF
            .Select 1, 1, 1, 1
        End With
    End If
End Sub


Function SaldoEntradas(IdProducto As Long, FchInicio As String, FchFinal As String, xConAux As ADODB.Connection) As Double
    Dim RstAux As New ADODB.Recordset
    Dim xCad As String
    Dim xTotEnt, xTotSal As Double
    Dim A As Integer
    
    ' CARGAMOS TODAS LAS ENTRADAS
    '--compras+salidas(almacen)+produccion + ventas(nota credito-devoluciones)
    
    xCad = "SELECT Sum([canpro]) AS total " _
        + vbCr + "FROM alm_inventario RIGHT JOIN (mae_prov LEFT JOIN ((mae_documento RIGHT JOIN (com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) LEFT JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) ON mae_prov.id = com_compras.idpro) ON alm_inventario.id = com_comprasdet.iditem " _
        + vbCr + "WHERE (((com_comprasdet.iditem)=" & IdProducto & ") AND ((com_compras.fchdoc)>=CDate('" & FchInicio & "') And (com_compras.fchdoc)<=CDate('" & FchFinal & "')) AND ((com_compras.tipcom)=1))" _
        + vbCr + "Union " _
        + vbCr + "SELECT Sum([cantidad]) AS total " _
        + vbCr + "FROM (alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id) LEFT JOIN (alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) ON alm_ingreso.id = alm_ingresodet.id " _
        + vbCr + "WHERE (((alm_ingresodet.iditem)=" & IdProducto & ") AND ((alm_ingreso.fching)>=CDate('" & FchInicio & "') And (alm_ingreso.fching)<=CDate('" & FchFinal & "')) AND ((alm_ingreso.tipmov)=-1)) " _
        + vbCr + "Union " _
        + vbCr + "SELECT Sum([cantidad]) AS total " _
        + vbCr + "FROM pro_produccion LEFT JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro " _
        + vbCr + "WHERE (((pro_producciondet.iditem)=" & IdProducto & ") AND ((pro_produccion.dia)>=CDate('" & FchInicio & "') And (pro_produccion.dia)<=CDate('" & FchFinal & "')))" _
        + vbCr + "Union " _
        + vbCr + "SELECT Sum(vta_ventasdet.canpro) AS total " _
        + vbCr + "FROM (mae_cliente RIGHT JOIN (vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) ON mae_cliente.id = vta_ventas.idcli) RIGHT JOIN (vta_ventasdet LEFT JOIN alm_inventario ON vta_ventasdet.iditem = alm_inventario.id) ON vta_ventas.id = vta_ventasdet.idvta " _
        + vbCr + "WHERE (((vta_ventasdet.iditem)=" & IdProducto & ") AND ((vta_ventas.fchdoc)>=CDate('" & FchInicio & "') And (vta_ventas.fchdoc)<=CDate('" & FchFinal & "')) AND ((vta_ventas.oriitem)=1 Or (vta_ventas.oriitem)=3) AND ((vta_ventas.iddocref)<>0) AND ((vta_ventas.idmotnotcre)=4))"
    
    RST_Busq RstAux, xCad, xConAux
    xTotEnt = 0
    
    If Not RstAux.EOF Then
        RstAux.MoveFirst
        For A = 1 To RstAux.RecordCount
            xTotEnt = NulosN(RstAux("total"))
            RstAux.MoveNext
            If RstAux.EOF = True Then
                Exit For
            End If
        Next A
    End If
    
    Set RstAux = Nothing
'
'    Dim xSalIni As Double
'    xSalIni = Busca_Codigo(IdProducto, "id", "stckini", "alm_inventario", "N", xCon)
'    SaldoActual = (xSalIni + NulosN(xTotEnt)) - NulosN(xTotSal)

    SaldoEntradas = xTotEnt
End Function

'*****************************************************************************************************
'* Nombre Archivo   : VerUnificado
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA LA LISTA DE INSUMOS UNIFICADOS, PARA ELLO TOMA INFORMACION DE LAS BASE
'*                    DE DATOS ACTIVAS
'* DISE�ADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub VerUnificado()
    Dim A As Integer
    Dim Total As Double
    Dim xIndex As Integer
    TabOne1.Enabled = False
    Toolbar1.Enabled = False
    Frame2.Left = 50
    Frame2.Top = 25
    Frame2.Visible = True
    Fg7.FrozenCols = 4
    
    PreparaRST
    
    xIndex = 0
    For A = 1 To 5
        Dim B As Integer
        If TabOne1.TabVisible(xIndex) = True Then
            For B = 1 To Fg5(xIndex).Rows - 1
                RstInsumos.Filter = adFilterNone
                If RstInsumos.RecordCount <> 0 Then
                    RstInsumos.MoveFirst
                End If
                RstInsumos.Filter = "cod_item = '" & Fg5(xIndex).TextMatrix(B, 3) & "'"
                If RstInsumos.RecordCount = 0 Then
                    RstInsumos.AddNew
                    RstInsumos("descripcion") = Fg5(xIndex).TextMatrix(B, 1)
                    RstInsumos("cod_item") = Fg5(xIndex).TextMatrix(B, 3)
                    RstInsumos("unimed") = Fg5(xIndex).TextMatrix(B, 4)
                    RstInsumos("ene") = Val(Fg5(xIndex).TextMatrix(B, 5))
                    RstInsumos("feb") = Val(Fg5(xIndex).TextMatrix(B, 6))
                    RstInsumos("mar") = Val(Fg5(xIndex).TextMatrix(B, 7))
                    RstInsumos("abr") = Val(Fg5(xIndex).TextMatrix(B, 8))
                    RstInsumos("may") = Val(Fg5(xIndex).TextMatrix(B, 9))
                    RstInsumos("jun") = Val(Fg5(xIndex).TextMatrix(B, 10))
                    RstInsumos("jul") = Val(Fg5(xIndex).TextMatrix(B, 11))
                    RstInsumos("ago") = Val(Fg5(xIndex).TextMatrix(B, 12))
                    RstInsumos("set") = Val(Fg5(xIndex).TextMatrix(B, 13))
                    RstInsumos("oct") = Val(Fg5(xIndex).TextMatrix(B, 14))
                    RstInsumos("nov") = Val(Fg5(xIndex).TextMatrix(B, 15))
                    RstInsumos("dic") = Val(Fg5(xIndex).TextMatrix(B, 16))
                    
                    RstInsumos("prog") = Val(Fg5(xIndex).TextMatrix(B, 17))
                    RstInsumos("comp") = Val(Fg5(xIndex).TextMatrix(B, 18))
                Else
                    If RstInsumos.RecordCount = 1 Then
                        RstInsumos("ene") = RstInsumos("ene") + Val(Fg5(xIndex).TextMatrix(B, 5))
                        RstInsumos("feb") = RstInsumos("feb") + Val(Fg5(xIndex).TextMatrix(B, 6))
                        RstInsumos("mar") = RstInsumos("mar") + Val(Fg5(xIndex).TextMatrix(B, 7))
                        RstInsumos("abr") = RstInsumos("abr") + Val(Fg5(xIndex).TextMatrix(B, 8))
                        RstInsumos("may") = RstInsumos("may") + Val(Fg5(xIndex).TextMatrix(B, 9))
                        RstInsumos("jun") = RstInsumos("jun") + Val(Fg5(xIndex).TextMatrix(B, 10))
                        RstInsumos("jul") = RstInsumos("jul") + Val(Fg5(xIndex).TextMatrix(B, 11))
                        RstInsumos("ago") = RstInsumos("ago") + Val(Fg5(xIndex).TextMatrix(B, 12))
                        RstInsumos("set") = RstInsumos("set") + Val(Fg5(xIndex).TextMatrix(B, 13))
                        RstInsumos("oct") = RstInsumos("oct") + Val(Fg5(xIndex).TextMatrix(B, 14))
                        RstInsumos("nov") = RstInsumos("nov") + Val(Fg5(xIndex).TextMatrix(B, 15))
                        RstInsumos("dic") = RstInsumos("dic") + Val(Fg5(xIndex).TextMatrix(B, 16))
                        
                        RstInsumos("prog") = RstInsumos("prog") + Val(Fg5(xIndex).TextMatrix(B, 17))
                        RstInsumos("comp") = RstInsumos("comp") + Val(Fg5(xIndex).TextMatrix(B, 18))
                    Else
                        'este error nunca debe de ocurrir
                        MsgBox "Hay mas de un items con el mismo codigo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    End If
                End If
            Next B
            
            For B = 1 To Fg6(xIndex).Rows - 1
                RstInsumos.Filter = adFilterNone
                If RstInsumos.RecordCount <> 0 Then
                    RstInsumos.MoveFirst
                End If
                RstInsumos.Filter = "cod_item = '" & Fg6(xIndex).TextMatrix(B, 3) & "'"
                If RstInsumos.RecordCount = 0 Then
                    RstInsumos.AddNew
                    RstInsumos("descripcion") = Fg6(xIndex).TextMatrix(B, 1)
                    RstInsumos("cod_item") = Fg6(xIndex).TextMatrix(B, 3)
                    RstInsumos("unimed") = Fg6(xIndex).TextMatrix(B, 4)
                    RstInsumos("ene") = Val(Fg6(xIndex).TextMatrix(B, 5))
                    RstInsumos("feb") = Val(Fg6(xIndex).TextMatrix(B, 6))
                    RstInsumos("mar") = Val(Fg6(xIndex).TextMatrix(B, 7))
                    RstInsumos("abr") = Val(Fg6(xIndex).TextMatrix(B, 8))
                    RstInsumos("may") = Val(Fg6(xIndex).TextMatrix(B, 9))
                    RstInsumos("jun") = Val(Fg6(xIndex).TextMatrix(B, 10))
                    RstInsumos("jul") = Val(Fg6(xIndex).TextMatrix(B, 11))
                    RstInsumos("ago") = Val(Fg6(xIndex).TextMatrix(B, 12))
                    RstInsumos("set") = Val(Fg6(xIndex).TextMatrix(B, 13))
                    RstInsumos("oct") = Val(Fg6(xIndex).TextMatrix(B, 14))
                    RstInsumos("nov") = Val(Fg6(xIndex).TextMatrix(B, 15))
                    RstInsumos("dic") = Val(Fg6(xIndex).TextMatrix(B, 16))
                    
                    RstInsumos("prog") = Val(Fg6(xIndex).TextMatrix(B, 17))
                    RstInsumos("comp") = Val(Fg6(xIndex).TextMatrix(B, 18))
                Else
                    If RstInsumos.RecordCount = 1 Then
                        RstInsumos("ene") = RstInsumos("ene") + Val(Fg6(xIndex).TextMatrix(B, 5))
                        RstInsumos("feb") = RstInsumos("feb") + Val(Fg6(xIndex).TextMatrix(B, 6))
                        RstInsumos("mar") = RstInsumos("mar") + Val(Fg6(xIndex).TextMatrix(B, 7))
                        RstInsumos("abr") = RstInsumos("abr") + Val(Fg6(xIndex).TextMatrix(B, 8))
                        RstInsumos("may") = RstInsumos("may") + Val(Fg6(xIndex).TextMatrix(B, 9))
                        RstInsumos("jun") = RstInsumos("jun") + Val(Fg6(xIndex).TextMatrix(B, 10))
                        RstInsumos("jul") = RstInsumos("jul") + Val(Fg6(xIndex).TextMatrix(B, 11))
                        RstInsumos("ago") = RstInsumos("ago") + Val(Fg6(xIndex).TextMatrix(B, 12))
                        RstInsumos("set") = RstInsumos("set") + Val(Fg6(xIndex).TextMatrix(B, 13))
                        RstInsumos("oct") = RstInsumos("oct") + Val(Fg6(xIndex).TextMatrix(B, 14))
                        RstInsumos("nov") = RstInsumos("nov") + Val(Fg6(xIndex).TextMatrix(B, 15))
                        RstInsumos("dic") = RstInsumos("dic") + Val(Fg6(xIndex).TextMatrix(B, 16))
                        
                        RstInsumos("prog") = RstInsumos("prog") + Val(Fg6(xIndex).TextMatrix(B, 17))
                        RstInsumos("comp") = RstInsumos("comp") + Val(Fg6(xIndex).TextMatrix(B, 18))
                    Else
                        'este error nunca debe de ocurrir
                        MsgBox "Hay mas de un items con el mismo codigo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    End If
                End If
            Next B
        End If
        xIndex = xIndex + 1
    Next A
    
    RstInsumos.Filter = adFilterNone
    RstInsumos.Sort = "descripcion"
    RstInsumos.MoveFirst
    Fg7.Rows = 1
    For A = 1 To RstInsumos.RecordCount
        Fg7.Rows = Fg7.Rows + 1
        Fg7.TextMatrix(A, 1) = RstInsumos("descripcion")
        Fg7.TextMatrix(A, 2) = RstInsumos("cod_item")
        'Fg7.TextMatrix(A, 3) = RstInsumos("unimed")
        Fg7.TextMatrix(A, 4) = RstInsumos("unimed")
        Fg7.TextMatrix(A, 5) = Format(RstInsumos("ene"), "0.00")
        Fg7.TextMatrix(A, 6) = Format(RstInsumos("feb"), "0.00")
        Fg7.TextMatrix(A, 7) = Format(RstInsumos("mar"), "0.00")
        Fg7.TextMatrix(A, 8) = Format(RstInsumos("abr"), "0.00")
        Fg7.TextMatrix(A, 9) = Format(RstInsumos("may"), "0.00")
        Fg7.TextMatrix(A, 10) = Format(RstInsumos("jun"), "0.00")
        Fg7.TextMatrix(A, 11) = Format(RstInsumos("jul"), "0.00")
        Fg7.TextMatrix(A, 12) = Format(RstInsumos("ago"), "0.00")
        Fg7.TextMatrix(A, 13) = Format(RstInsumos("set"), "0.00")
        Fg7.TextMatrix(A, 14) = Format(RstInsumos("oct"), "0.00")
        Fg7.TextMatrix(A, 15) = Format(RstInsumos("nov"), "0.00")
        Fg7.TextMatrix(A, 16) = Format(RstInsumos("dic"), "0.00")
        
        Fg7.TextMatrix(A, 17) = Format(RstInsumos("prog"), "0.00")
        Fg7.TextMatrix(A, 18) = Format(RstInsumos("comp"), "0.00")
        Fg7.TextMatrix(A, 19) = Fg7.TextMatrix(A, 18) - Fg7.TextMatrix(A, 17)
        
        With Fg7
            .Select A, 19, A, 19
            .FillStyle = flexFillRepeat
            If NulosN(.TextMatrix(A, 19)) >= 0 Then
                .CellForeColor = &HFF0000
            Else
                .CellForeColor = &HFF&
            End If
        End With
        
        RstInsumos.MoveNext
        If RstInsumos.EOF = True Then
            Exit For
        End If
    Next A
    
    With Fg7
        .Select 1, .Cols - 3, .Rows - 1, .Cols - 1
        .FillStyle = flexFillRepeat
        .CellBackColor = &H80000013   '&HE0FEE7 '&HDDFFFF
        .Select 1, 1, 1, 1
    End With
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim A As Integer
    If Button.Index = 1 Then
        VerUnificado
    End If
    If Button.Index = 4 Then
        Unload Me
    End If
    If Button.Index = 3 Then
        For A = 0 To RstEmp.RecordCount - 1
            If TabOne1.CurrTab = A Then
                Dim xExport As New SGI2_funciones.Formularios
                If TabOne2(A).CurrTab = 0 Then
                    'ExportarExcel Fg5(A), TabOne1.TabCaption(A), "Terminado"
                    xExport.VSFlexGrid_Exportar_MSExcel xCon, Fg5(A), "Unificado - Plan de Abastecimiento(" & TabOne1.TabCaption(A) & ") " & AnoTra, "Productos Terminados", "", "Unificado - Plan de Abastecimiento"
                Else
                    'ExportarExcel Fg6(A), TabOne1.TabCaption(A), "Intermedio"
                    xExport.VSFlexGrid_Exportar_MSExcel xCon, Fg6(A), "Unificado - Plan de Abastecimiento(" & TabOne1.TabCaption(A) & ") " & AnoTra, "Productos Intermedios", "", "Unificado - Plan de Abastecimiento"
                    
                End If
                Set xExport = Nothing
                
            End If
        Next A
    End If
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : PreparaRST
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA UN RECORDSET TEMPORAL PARA CARGAR LOS DATOS DE LOS PRODUCTOS A PROCESAR
'* DISE�ADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub PreparaRST()
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(20, 3) As String

    xCampos(0, 0) = "cod_item":     xCampos(0, 1) = "C":      xCampos(0, 2) = "16"
    xCampos(1, 0) = "unimed":       xCampos(1, 1) = "C":      xCampos(1, 2) = "4"
    xCampos(2, 0) = "descripcion":  xCampos(2, 1) = "C":      xCampos(2, 2) = "200"
    xCampos(3, 0) = "ene":          xCampos(3, 1) = "N":      xCampos(3, 2) = "2"
    xCampos(4, 0) = "feb":          xCampos(4, 1) = "N":      xCampos(4, 2) = "2"
    xCampos(5, 0) = "mar":          xCampos(5, 1) = "N":      xCampos(5, 2) = "2"
    xCampos(6, 0) = "abr":          xCampos(6, 1) = "N":      xCampos(6, 2) = "2"
    xCampos(7, 0) = "may":          xCampos(7, 1) = "N":      xCampos(7, 2) = "2"
    xCampos(8, 0) = "jun":          xCampos(8, 1) = "N":      xCampos(8, 2) = "2"
    xCampos(9, 0) = "jul":          xCampos(9, 1) = "N":      xCampos(9, 2) = "2"
    xCampos(10, 0) = "ago":         xCampos(10, 1) = "N":      xCampos(10, 2) = "2"
    xCampos(11, 0) = "set":         xCampos(11, 1) = "N":      xCampos(11, 2) = "2"
    xCampos(12, 0) = "oct":         xCampos(12, 1) = "N":      xCampos(12, 2) = "2"
    xCampos(13, 0) = "nov":         xCampos(13, 1) = "N":      xCampos(13, 2) = "2"
    xCampos(14, 0) = "dic":         xCampos(14, 1) = "N":      xCampos(14, 2) = "2"
    xCampos(15, 0) = "ope":         xCampos(15, 1) = "N":      xCampos(15, 2) = "2"
    xCampos(16, 0) = "idpro":       xCampos(16, 1) = "N":      xCampos(16, 2) = "2"
    xCampos(17, 0) = "tippro":       xCampos(17, 1) = "C":      xCampos(17, 2) = "2"
    
    xCampos(18, 0) = "prog":         xCampos(18, 1) = "N":      xCampos(18, 2) = "2"
    xCampos(19, 0) = "comp":         xCampos(19, 1) = "N":      xCampos(19, 2) = "2"
    
    Set RstInsumos = xFun.CrearRstTMP(xCampos)
    RstInsumos.Open
End Sub

Sub ExportarExcel(FgAux As VSFlexGrid, empresa As String, tipo As String)
'''    Dim A As Integer
'''    Dim B As Integer
'''    Dim xFilas As Integer
'''    Dim xCad As String
'''    Dim objExcel As Object
'''
'''    Set objExcel = CreateObject("Excel.Application")
'''
'''    objExcel.Visible = True
'''    'determina el numero de hojas que se mostrara en el Excel
'''    objExcel.SheetsInNewWorkbook = 1
'''
'''    objExcel.WindowState = 2
'''    objExcel.Workbooks.Add
'''
'''    With objExcel.ActiveSheet
'''        .Cells(1, 2) = "Plan Unificado de Abastecimiento"
'''        .Range("B1", "R1").Merge
'''        .Cells(1, 2).HorizontalAlignment = xlHAlignCenterAcrossSelection
'''        .Cells(1, 2).Font.Bold = True
'''        .Cells(1, 2).Rows(1).Font.Size = 12
'''
'''        .Cells(2, 2) = "Empresa: "
'''        .Cells(2, 2).Font.Bold = True
'''        .Cells(2, 3) = empresa
'''        .Cells(3, 2) = "Tipo de Productos: "
'''        .Cells(3, 2).Font.Bold = True
'''        .Cells(3, 3) = tipo
'''        xFilas = 5
'''        For A = 0 To FgAux.Rows - 1
'''            For B = 1 To FgAux.Cols - 1
'''                If A = 0 Then
'''                    .Cells(xFilas, B + 1).Font.Bold = True
'''                    .Cells(xFilas, B + 1) = "'" + FgAux.TextMatrix(A, B)
'''                Else
'''                    If B <= 4 And B <> 3 Then
'''                        .Cells(xFilas, B + 1) = "'" + FgAux.TextMatrix(A, B)
'''                    Else
'''                        .Cells(xFilas, B + 1) = NulosN(FgAux.TextMatrix(A, B))
'''                    End If
'''                End If
'''            Next B
'''            xFilas = xFilas + 1
'''        Next A
'''    End With
'''    MsgBox "El proceso de exportacion termino con exito", vbInformation + vbOKOnly + vbDefaultButton1, "Registro de Compras y Ventas"
'''    objExcel.WindowState = 1
'''    Set objExcel = Nothing
'''    Exit Sub
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : ExportarExcel
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTA LOS DATOS DEL CONTROL Fg7 A MS EXCEL
'* DISE�ADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub ExportarExcelUnif()
'''    If Fg7.Rows = 1 Then
'''        MsgBox "No se ha procesado registros para el consolidados de insumos", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
'''        Exit Sub
'''    End If
'''
'''    Dim A As Integer
'''    Dim B As Integer
'''    Dim xFilas As Integer
'''    Dim xCad As String
'''    Dim objExcel As Object
'''
'''    Set objExcel = CreateObject("Excel.Application")
'''
'''    objExcel.Visible = True
'''    'determina el numero de hojas que se mostrara en el Excel
'''    objExcel.SheetsInNewWorkbook = 1
'''
'''    objExcel.WindowState = 2
'''    objExcel.Workbooks.Add
'''
'''    With objExcel.ActiveSheet
'''        .Cells(1, 2) = NomEmp
'''        .Cells(1, 10) = Date
'''        .Cells(2, 2) = "N� R.U.C. : " + NumRUC
'''        .Cells(3, 2) = "Consilidado de Insumos y Materia Prima"
'''
'''        For B = B To 4
'''            If TabOne1.TabVisible(B) = True Then
'''                xCad = xCad + TabOne1.TabCaption(B) + ", "
'''            End If
'''        Next B
'''        .Cells(4, 2) = "Empresas Consolidadas : " & xCad
'''
'''        xFilas = 5
'''        For A = 0 To Fg7.Rows - 1
'''            For B = 1 To Fg7.Cols - 1
'''                If A = 0 Then
'''                    .Cells(xFilas, B + 1) = "'" + Fg7.TextMatrix(A, B)
'''                Else
'''                    If B <= 4 Then
'''                        .Cells(xFilas, B + 1) = "'" + Fg7.TextMatrix(A, B)
'''                    Else
'''                        .Cells(xFilas, B + 1) = Val(Fg7.TextMatrix(A, B))
'''                    End If
'''                End If
'''            Next B
'''            xFilas = xFilas + 1
'''        Next A
'''    End With
'''
'''    MsgBox "El proceso de exportacion termino con exito", vbInformation + vbOKOnly + vbDefaultButton1, "Registro de Compras y Ventas"
'''    objExcel.WindowState = 1
'''    Set objExcel = Nothing
'''    Exit Sub
End Sub

